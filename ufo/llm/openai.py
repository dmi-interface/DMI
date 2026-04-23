# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

import functools
import json
import os
import shutil
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any, Callable, Dict, List, Literal, Optional, Tuple

import openai
from openai import AzureOpenAI, OpenAI
from ufo.llm.base import BaseService


class BaseOpenAIService(BaseService):
    def __init__(self, config: Dict[str, Any], agent_type: str, api_provider: str, api_base: str) -> None:
        """
        Create an OpenAI service instance.
        :param config: The configuration for the OpenAI service.
        :param agent_type: The type of the agent.
        :param api_type: The type of the API (e.g., "openai", "aoai", "azure_ad").
        :param api_base: The base URL of the API.
        """
        self.config_llm = config[agent_type]
        self.config = config
        self.api_type = self.config_llm["API_TYPE"].lower()
        self.max_retry = self.config["MAX_RETRY"]
        self.prices = self.config.get("PRICES", {})
        self.agent_type = agent_type
        assert api_provider in ["openai", "aoai", "azure_ad"], "Invalid API Provider"

        self.client: OpenAI = OpenAIService.get_openai_client(
            api_provider,
            api_base,
            self.max_retry,
            self.config["TIMEOUT"],
            self.config_llm.get("API_KEY", ""),
            self.config_llm.get("API_VERSION", ""),
            aad_api_scope_base=self.config_llm.get("AAD_API_SCOPE_BASE", ""),
            aad_tenant_id=self.config_llm.get("AAD_TENANT_ID", ""),
        )

    def _map_response_format_to_text_spec(self, response_format: Optional[Dict[str, Any]]):
        """
        Map the Chat Completions response_format to the Responses API text spec (object).
        Supported mappings:
          - {"type":"json_object"}  -> {"format": {"type": "json_object"}}
          - {"type":"json_schema", "json_schema" or "schema": {...}, "name": "..."}
                                   -> {"format": {"type": "json_schema", "json_schema": {...}, "name": "..."}}
        Other / None -> None (text field not sent)
        """
        if not isinstance(response_format, dict):
            return None

        rft = response_format.get("type")

        if rft == "json_object":
            # Fix: must be json_object, not "json"
            return {"format": {"type": "json_object"}}

        if rft == "json_schema":
            schema = response_format.get("json_schema") or response_format.get("schema")
            name = response_format.get("name") or "structured_output"
            if schema:
                return {"format": {"type": "json_schema", "json_schema": schema, "name": name}}
            return None

        # Optional: allow explicit plain-text request
        if rft == "text":
            return {"format": {"type": "text"}}

        return None

    def _convert_messages_to_responses_input_and_instructions(
            self, messages: List[Dict[str, Any]]
    ) -> Dict[str, Any]:
        """
        Convert Chat Completions messages to Responses API input + instructions.
        - system text is merged into instructions
        - user/assistant messages are kept in input, with content split into typed segments:
            * text   -> {"type":"input_text"/"output_text","text":...}
            * image  -> {"type":"input_image","image_url": "<URL or dataURL string>"} (user only; assistant images are skipped)
        """
        instructions_parts: List[str] = []
        input_msgs: List[Dict[str, Any]] = []

        def _part_text(part: Any) -> str:
            if isinstance(part, dict):
                return part.get("text") or part.get("content") or ""
            return str(part)

        def _extract_image_url_value(image_url_field: Any) -> Optional[str]:
            """
            In Chat Completions, the common form is {"image_url": {"url": "...", "detail": "high"}}.
            Responses API requires a **string**, so extract the url from the object.
            Also accepts a plain string or data URL directly.
            """
            if isinstance(image_url_field, str):
                return image_url_field
            if isinstance(image_url_field, dict):
                # Common field name is url; other fields (e.g. detail) are not supported in Responses and are discarded
                return image_url_field.get("url") or image_url_field.get("data")
            return None

        for m in messages:
            role = m.get("role")
            content = m.get("content", "")

            # 1) system -> instructions (text only)
            if role == "system":
                if isinstance(content, str):
                    if content:
                        instructions_parts.append(content)
                elif isinstance(content, list):
                    for part in content:
                        if isinstance(part, dict) and part.get("type") in ("text", "input_text", "output_text"):
                            txt = _part_text(part)
                            if txt:
                                instructions_parts.append(txt)
                else:
                    instructions_parts.append(str(content))
                continue

            # 2) Keep only user / assistant
            if role not in ("user", "assistant"):
                continue

            content_items: List[Dict[str, Any]] = []
            text_type = "input_text" if role == "user" else "output_text"

            if isinstance(content, str):
                content_items.append({"type": text_type, "text": content})

            elif isinstance(content, list):
                for part in content:
                    if not isinstance(part, dict):
                        content_items.append({"type": text_type, "text": str(part)})
                        continue

                    ptype = part.get("type")

                    if ptype in ("text", "input_text", "output_text"):
                        txt = _part_text(part)
                        content_items.append({"type": text_type, "text": txt})

                    elif ptype in ("image_url", "input_image"):
                        url_val = _extract_image_url_value(part.get("image_url"))
                        # Only user images are treated as input; assistant images are skipped
                        if role == "user" and url_val:
                            content_items.append({"type": "input_image", "image_url": url_val})

                    else:
                        # Fallback for unknown types: treat as text
                        content_items.append({"type": text_type, "text": _part_text(part)})
            else:
                content_items.append({"type": text_type, "text": str(content)})

            input_msgs.append({"role": role, "content": content_items})

        out: Dict[str, Any] = {"input": input_msgs}
        if instructions_parts:
            out["instructions"] = "\n".join(instructions_parts)
        return out

    def _log_reasoning_from_response(self, response: Any) -> None:
        """
        Extract and print reasoning text from a non-streaming response.
        Fault-tolerant data structure handling: first looks for output[*].type == 'reasoning',
        then extracts .text from summary/content/steps/explanations fields.
        """
        try:
            outputs = getattr(response, "output", []) or []
            for item in outputs:
                if getattr(item, "type", None) == "reasoning":
                    blocks = []
                    for field in ("summary", "content", "steps", "explanations"):
                        seq = getattr(item, field, None)
                        if isinstance(seq, list):
                            for b in seq:
                                t = getattr(b, "text", None)
                                if not t and isinstance(b, str):
                                    t = b
                                if t:
                                    blocks.append(t)
                    if blocks:
                        print("=== REASONING ===")
                        for t in blocks:
                            print(t)
                        print("=== END REASONING ===")

        except Exception:
            # Fail silently to avoid disrupting the main flow
            pass

    def _chat_completion(
            self,
            messages: List[Dict[str, str]],
            stream: bool = False,
            temperature: Optional[float] = None,
            max_tokens: Optional[int] = None,
            top_p: Optional[float] = None,
            **kwargs: Any,
    ) -> Tuple[Dict[str, Any], Optional[float]]:
        """
        Use the OpenAI Responses API while preserving behavior equivalent to Chat Completions:
        - Reasoning branch: temperature/top_p/max_tokens are not set; always runs with background=True (store=True).
          In non-streaming mode, internally polls until completion (default 6-minute timeout,
          overridable via background_timeout_s / background_poll_interval_s).
        - Non-reasoning branch: the above parameters are set; background is not handled.
        - Returns ([text], cost).
        """
        import time

        model = self.config_llm["API_MODEL"]

        # Default sampling parameters for the non-reasoning branch (equivalent to original logic)
        temperature = temperature if temperature is not None else self.config["TEMPERATURE"]
        top_p = top_p if top_p is not None else self.config["TOP_P"]
        max_output_tokens = max_tokens if max_tokens is not None else self.config["MAX_TOKENS"]

        # Background polling parameters used only in the reasoning branch (default: 6 min / 2 sec)
        background_timeout_s: int = int(kwargs.pop("background_timeout_s", 360))
        background_poll_interval_s: float = float(kwargs.pop("background_poll_interval_s", 2.0))

        # 1) Convert messages -> input + instructions (including rich content splitting)
        converted = self._convert_messages_to_responses_input_and_instructions(messages)
        input_payload = converted.get("input", [])
        instructions = converted.get("instructions", None)

        # 2) Structured output compatibility: response_format -> text.format (object)
        rf = kwargs.pop("response_format", None)
        text_spec = self._map_response_format_to_text_spec(rf)

        # stream_options from Chat Completions is not valid in Responses API
        kwargs.pop("stream_options", None)

        try:
            # =========================
            #      REASONING BRANCH (always background)
            # =========================
            if self.config_llm.get("REASONING_MODEL", False):
                # Equivalent to original logic: no temperature/length params; always runs in background
                request_kwargs = dict(
                    model=model,
                    reasoning={
                        "summary": "auto",
                        "effort": self.config_llm.get("REASONING_EFFORT", "medium")
                        # "effort": "medium"
                    },
                    input=input_payload,
                    background=True,  # Always background
                    store=True,       # Required for background
                    stream=stream,
                    **kwargs,
                )
                if instructions:
                    request_kwargs["instructions"] = instructions
                if text_spec:
                    request_kwargs["text"] = text_spec

                response: Any = self.client.responses.create(**request_kwargs)

                # Background + streaming: consume events directly
                if stream:
                    collected_content = [""]
                    total_in = total_out = 0
                    for event in response:
                        et = getattr(event, "type", None)
                        if et == "response.output_text.delta":
                            delta = getattr(event, "delta", None)
                            if delta:
                                collected_content[0] += delta
                        elif et == "response.completed":
                            resp_obj = getattr(event, "response", None)
                            if resp_obj and getattr(resp_obj, "usage", None):
                                total_in += getattr(resp_obj.usage, "input_tokens", 0)
                                total_out += getattr(resp_obj.usage, "output_tokens", 0)
                        elif et == "response.error":
                            err = getattr(event, "error", "Unknown error")
                            raise Exception(f"Streaming error: {err}")
                    cost = self.get_cost_estimator(self.api_type, model, self.prices, total_in, total_out)
                    return collected_content, cost

                # Background + non-streaming: poll until completed or timed out
                start = time.time()
                status = getattr(response, "status", None)
                resp_id = getattr(response, "id", None)
                while status in ("queued", "in_progress"):
                    if time.time() - start > background_timeout_s:
                        try:
                            if resp_id:
                                self.client.responses.cancel(resp_id)
                        except Exception:
                            pass
                        raise Exception(f"Background response timed out after {background_timeout_s} seconds.")
                    time.sleep(background_poll_interval_s)
                    response = self.client.responses.retrieve(resp_id)
                    status = getattr(response, "status", None)
                if status not in ("completed", "incomplete"):
                    raise Exception(f"Background response ended with status '{status}'.")

                # Non-streaming: finalize after background polling completes
                output_text = getattr(response, "output_text", None)
                if output_text is None:
                    output_text = ""
                    for out in getattr(response, "output", []):
                        if getattr(out, "type", "") == "message":
                            for c in getattr(out, "content", []):
                                if getattr(c, "type", "") in ("output_text", "input_text"):
                                    output_text += getattr(c, "text", "") or ""
                usage = getattr(response, "usage", None)
                input_tokens = getattr(usage, "input_tokens", 0) if usage else 0
                output_tokens = getattr(usage, "output_tokens", 0) if usage else 0
                cost = self.get_cost_estimator(self.api_type, model, self.prices, input_tokens, output_tokens)
                try:
                    self._log_reasoning_from_response(response)
                except Exception:
                    pass
                return [output_text], cost

            # =========================
            #     NON-REASONING BRANCH (no background)
            # =========================
            # Ensure background-related parameters are not passed to the API
            kwargs.pop("background", None)
            kwargs.pop("background_timeout_s", None)
            kwargs.pop("background_poll_interval_s", None)

            request_kwargs = dict(
                model=model,
                input=input_payload,
                stream=stream,
                temperature=temperature,
                top_p=top_p,
                max_output_tokens=max_output_tokens,
                **kwargs,
            )
            if instructions:
                request_kwargs["instructions"] = instructions
            if text_spec:
                request_kwargs["text"] = text_spec

            response: Any = self.client.responses.create(**request_kwargs)

            # Streaming
            if stream:
                collected_content = [""]
                total_in = total_out = 0
                for event in response:
                    et = getattr(event, "type", None)
                    if et == "response.output_text.delta":
                        delta = getattr(event, "delta", None)
                        if delta:
                            collected_content[0] += delta
                    elif et == "response.completed":
                        resp_obj = getattr(event, "response", None)
                        if resp_obj and getattr(resp_obj, "usage", None):
                            total_in += getattr(resp_obj.usage, "input_tokens", 0)
                            total_out += getattr(resp_obj.usage, "output_tokens", 0)
                    elif et == "response.error":
                        err = getattr(event, "error", "Unknown error")
                        raise Exception(f"Streaming error: {err}")
                cost = self.get_cost_estimator(self.api_type, model, self.prices, total_in, total_out)
                return collected_content, cost

            # Non-streaming
            output_text = getattr(response, "output_text", None)
            if output_text is None:
                output_text = ""
                for out in getattr(response, "output", []):
                    if getattr(out, "type", "") == "message":
                        for c in getattr(out, "content", []):
                            if getattr(c, "type", "") in ("output_text", "input_text"):
                                output_text += getattr(c, "text", "") or ""
            usage = getattr(response, "usage", None)
            input_tokens = getattr(usage, "input_tokens", 0) if usage else 0
            output_tokens = getattr(usage, "output_tokens", 0) if usage else 0
            cost = self.get_cost_estimator(self.api_type, model, self.prices, input_tokens, output_tokens)
            return [output_text], cost

        except openai.APITimeoutError as e:
            # Handle timeout error, e.g. retry or log
            raise Exception(f"OpenAI API request timed out: {e}")
        except openai.APIConnectionError as e:
            # Handle connection error, e.g. check network or log
            raise Exception(f"OpenAI API request failed to connect: {e}")
        except openai.BadRequestError as e:
            # Handle invalid request error, e.g. validate parameters or log
            raise Exception(f"OpenAI API request was invalid: {e}")
        except openai.AuthenticationError as e:
            # Handle authentication error, e.g. check credentials or log
            raise Exception(f"OpenAI API request was not authorized: {e}")
        except openai.PermissionDeniedError as e:
            # Handle permission error, e.g. check scope or log
            raise Exception(f"OpenAI API request was not permitted: {e}")
        except openai.RateLimitError as e:
            # Handle rate limit error, e.g. wait or log
            raise Exception(f"OpenAI API request exceeded rate limit: {e}")
        except openai.APIError as e:
            # Handle API error, e.g. retry or log
            raise Exception(f"OpenAI API returned an API Error: {e}")

    def _chat_completion_operator(
        self,
        message: Dict[str, Any] = None,
        **kwargs: Any,
    ) -> Tuple[Dict[str, Any], Optional[float]]:
        """
        Generates completions for a given conversation using the OpenAI Chat API.
        :param message: The message to send to the API.
        :param n: The number of completions to generate.
        :return: A tuple containing a list of generated completions and the estimated cost.
        """

        inputs = message.get("inputs", [])
        tools = message.get("tools", [])
        previous_response_id = message.get("previous_response_id", None)

        response = self.client.responses.create(
            model=self.config_llm.get("API_MODEL"),
            input=inputs,
            tools=tools,
            previous_response_id=previous_response_id,
            truncation="auto",
            temperature=self.config.get("TEMPERATURE", 0),
            top_p=self.config.get("TOP_P", 0),
            timeout=self.config.get("TIMEOUT", 20),
        ).model_dump()

        if "usage" in response:
            usage = response.get("usage")
            input_tokens = usage.get("input_tokens", 0)
            output_tokens = usage.get("output_tokens", 0)
        else:
            input_tokens = 0
            output_tokens = 0

        cost = self.get_cost_estimator(
            self.api_type,
            self.config_llm["API_MODEL"],
            self.prices,
            input_tokens,
            output_tokens,
        )

        return [response], cost

    @functools.lru_cache()
    @staticmethod
    def get_openai_client(
        api_type: str,
        api_base: str,
        max_retry: int,
        timeout: int,
        api_key: Optional[str] = None,
        api_version: Optional[str] = None,
        aad_api_scope_base: Optional[str] = None,
        aad_tenant_id: Optional[str] = None,
    ) -> OpenAI:
        """
        Create an OpenAI client based on the API type.
        :param api_type: The type of the API, one of "openai", "aoai", or "azure_ad".
        :param api_base: The base URL of the API.
        :param max_retry: The maximum number of retries for the API request.
        :param timeout: The timeout for the API request.
        :param api_key: The API key for the OpenAI API.
        :param api_version: The API version for the Azure OpenAI API.
        :param aad_api_scope_base: The AAD API scope base for the Azure OpenAI API.
        :param aad_tenant_id: The AAD tenant ID for the Azure OpenAI API.
        :return: The OpenAI client.
        """
        if api_type == "openai":
            assert api_key, "OpenAI API key must be specified"
            assert api_base, "OpenAI API base URL must be specified"
            client = OpenAI(
                base_url=api_base,
                api_key=api_key,
                max_retries=max_retry,
                timeout=timeout,
            )
        else:
            assert api_version, "Azure OpenAI API version must be specified"
            if api_type == "aoai":
                assert api_key, "Azure OpenAI API key must be specified"
                client = AzureOpenAI(
                    max_retries=max_retry,
                    timeout=timeout,
                    api_version=api_version,
                    azure_endpoint=api_base,
                    api_key=api_key,
                )
            else:
                assert (
                    aad_api_scope_base and aad_tenant_id
                ), "AAD API scope base and tenant ID must be specified"
                token_provider = OpenAIService.get_aad_token_provider(
                    aad_api_scope_base=aad_api_scope_base,
                    aad_tenant_id=aad_tenant_id,
                )
                client = AzureOpenAI(
                    max_retries=max_retry,
                    timeout=timeout,
                    api_version=api_version,
                    azure_endpoint=api_base,
                    azure_ad_token_provider=token_provider,
                )
        return client

    @functools.lru_cache()
    @staticmethod
    def get_aad_token_provider(
        aad_api_scope_base: str,
        aad_tenant_id: str,
        token_cache_file: str = "aoai-token-cache.bin",
        client_id: Optional[str] = None,
        client_secret: Optional[str] = None,
        use_azure_cli: Optional[bool] = None,
        use_broker_login: Optional[bool] = None,
        use_managed_identity: Optional[bool] = None,
        use_device_code: Optional[bool] = None,
        **kwargs,
    ) -> Callable[[], str]:
        """
        Acquire token from Azure AD for OpenAI.
        :param aad_api_scope_base: The base scope for the Azure AD API.
        :param aad_tenant_id: The tenant ID for the Azure AD API.
        :param token_cache_file: The path to the token cache file.
        :param client_id: The client ID for the AAD app.
        :param client_secret: The client secret for the AAD app.
        :param use_azure_cli: Use Azure CLI for authentication.
        :param use_broker_login: Use broker login for authentication.
        :param use_managed_identity: Use managed identity for authentication.
        :param use_device_code: Use device code for authentication.
        :return: The access token for OpenAI.
        """

        import msal
        from azure.identity import (
            AuthenticationRecord,
            AzureCliCredential,
            ClientSecretCredential,
            DeviceCodeCredential,
            ManagedIdentityCredential,
            TokenCachePersistenceOptions,
            get_bearer_token_provider,
        )
        from azure.identity.broker import InteractiveBrowserBrokerCredential

        api_scope_base = "api://" + aad_api_scope_base

        tenant_id = aad_tenant_id
        scope = api_scope_base + "/.default"

        token_cache_option = TokenCachePersistenceOptions(
            name=token_cache_file,
            enable_persistence=True,
            allow_unencrypted_storage=True,
        )

        def save_auth_record(auth_record: AuthenticationRecord):
            try:
                with open(token_cache_file, "w") as cache_file:
                    cache_file.write(auth_record.serialize())
            except Exception as e:
                print("failed to save auth record", e)

        def load_auth_record() -> Optional[AuthenticationRecord]:
            try:
                if not os.path.exists(token_cache_file):
                    return None
                with open(token_cache_file, "r") as cache_file:
                    return AuthenticationRecord.deserialize(cache_file.read())
            except Exception as e:
                print("failed to load auth record", e)
                return None

        auth_record: Optional[AuthenticationRecord] = load_auth_record()

        current_auth_mode: Literal[
            "client_secret",
            "managed_identity",
            "az_cli",
            "interactive",
            "device_code",
            "none",
        ] = "none"

        implicit_mode = not (
            use_managed_identity or use_azure_cli or use_broker_login or use_device_code
        )

        if use_managed_identity or (implicit_mode and client_id is not None):
            if not use_managed_identity and client_secret is not None:
                assert (
                    client_id is not None
                ), "client_id must be specified with client_secret"
                current_auth_mode = "client_secret"
                identity = ClientSecretCredential(
                    client_id=client_id,
                    client_secret=client_secret,
                    tenant_id=tenant_id,
                    cache_persistence_options=token_cache_option,
                    authentication_record=auth_record,
                )
            else:
                current_auth_mode = "managed_identity"
                if client_id is None:
                    # using default managed identity
                    identity = ManagedIdentityCredential(
                        cache_persistence_options=token_cache_option,
                    )
                else:
                    identity = ManagedIdentityCredential(
                        client_id=client_id,
                        cache_persistence_options=token_cache_option,
                    )
        elif use_azure_cli or (implicit_mode and shutil.which("az") is not None):
            current_auth_mode = "az_cli"
            identity = AzureCliCredential(tenant_id=tenant_id)
        else:
            if implicit_mode:
                # enable broker login for known supported envs if not specified using use_device_code
                if sys.platform.startswith("darwin") or sys.platform.startswith(
                    "win32"
                ):
                    use_broker_login = True
                elif os.environ.get("WSL_DISTRO_NAME", "") != "":
                    use_broker_login = True
                elif os.environ.get("TERM_PROGRAM", "") == "vscode":
                    use_broker_login = True
                else:
                    use_broker_login = False
            if use_broker_login:
                current_auth_mode = "interactive"
                identity = InteractiveBrowserBrokerCredential(
                    tenant_id=tenant_id,
                    cache_persistence_options=token_cache_option,
                    use_default_broker_account=True,
                    parent_window_handle=msal.PublicClientApplication.CONSOLE_WINDOW_HANDLE,
                    authentication_record=auth_record,
                )
            else:
                current_auth_mode = "device_code"
                identity = DeviceCodeCredential(
                    tenant_id=tenant_id,
                    cache_persistence_options=token_cache_option,
                    authentication_record=auth_record,
                )

            try:
                auth_record = identity.authenticate(scopes=[scope])
                if auth_record:
                    save_auth_record(auth_record)

            except Exception as e:
                print(
                    f"failed to acquire token from AAD for OpenAI using {current_auth_mode}",
                    e,
                )
                raise e

        try:
            return get_bearer_token_provider(identity, scope)
        except Exception as e:
            print("failed to acquire token from AAD for OpenAI", e)
            raise e

class OpenAIService(BaseOpenAIService):
    """
    The OpenAI service class to interact with the OpenAI API.
    """

    def __init__(self, config: Dict[str, Any], agent_type: str) -> None:
        """
        Create an OpenAI service instance.
        :param config: The configuration for the OpenAI service.
        :param agent_type: The type of the agent.
        """
        super().__init__(config, agent_type, config[agent_type]["API_TYPE"].lower(), config[agent_type]["API_BASE"])

    def chat_completion(
        self,
        messages: List[Dict[str, str]],
        n: int,
        stream: bool = False,
        temperature: Optional[float] = None,
        max_tokens: Optional[int] = None,
        top_p: Optional[float] = None,
        **kwargs: Any,
    ) -> Tuple[Dict[str, Any], Optional[float]]:
        """
        Generates completions for a given conversation using the OpenAI Chat API.
        :param messages: The list of messages in the conversation.
        :param n: The number of completions to generate.
        :param stream: Whether to stream the API response.
        :param temperature: The temperature parameter for randomness in the output.
        :param max_tokens: The maximum number of tokens in the generated completion.
        :param top_p: The top-p parameter for nucleus sampling.
        :param kwargs: Additional keyword arguments to pass to the OpenAI API.
        :return: A tuple containing a list of generated completions and the estimated cost.
        :raises: Exception if there is an error in the OpenAI API request
        """

        if self.agent_type.lower() != "operator":
            # If the agent type is not "operator", use the OpenAI API directly
            return super()._chat_completion(
                messages,
                False,
                temperature,
                max_tokens,
                top_p,
                response_format={"type": "json_object"},
                **kwargs,
            )
        else:
            # If the agent type is "operator", use the OpenAI Operator API
            return super()._chat_completion_operator(
                messages,
            )


class OpenAIBetaClient:

    Json = Dict[str, Any]

    def __init__(self, endpoint: str, api_version: str):
        """
        The OpenAI Beta client class to interact with the OpenAI API.
        :param endpoint: The OpenAI API endpoint.
        :param api_key: The OpenAI API key.
        :param api_version: The OpenAI API version.
        """

        self.endpoint = endpoint
        self.base_url = endpoint.rstrip("/")

        self.api_version = api_version

    def get_responses(
        self,
        model: str,
        previous_response_id: Optional[str] = None,
        inputs: Optional[list[Json]] = None,  # pylint: disable=redefined-builtin
        tool_output: Optional[list[Json]] = None,
        include: Optional[list[str]] = None,
        tools: Optional[list[Json]] = None,
        metadata: Optional[Json] = None,
        temperature: Optional[float] = None,
        top_p: Optional[float] = None,
        parallel_tool_calls: Optional[bool] = None,
        token_provider: Optional[Callable[[], str]] = None,
    ) -> Json:
        self,

        if self.base_url.endswith("openai.azure.com"):
            url = f"{self.base_url}/openai/responses?api-version={self.api_version}"
        else:
            url = f"{self.base_url}/v1/responses"

        api_key = (
            token_provider if isinstance(token_provider, str) else token_provider()
        )

        headers = {
            "Content-Type": "application/json",
            "x-ms-enable-preview": "true",
            "api-key": api_key,
            "x-ms-enable-preview": "true",
            "Authorization": f"Bearer {api_key}",  # OpenAI
            "OpenAI-Beta": "responses=v1",  # OpenAI
        }

        return self.post_request(
            url,
            data={
                "model": model,
                "previous_response_id": previous_response_id,
                "input": inputs,
                "tool_output": tool_output,
                "include": include,
                "tools": tools,
                "metadata": metadata,
                "temperature": temperature,
                "top_p": top_p,
                "parallel_tool_calls": parallel_tool_calls,
            },
            headers=headers,
        )

    def post_request(self, url: str, data: Json, headers: Json) -> Json:
        """
        Send a POST request to the OpenAI API.
        :param url: The URL of the API endpoint.
        :param data: The data to send in the request.
        :param headers: The headers to send in the request.
        :return: The response from the API.
        """

        headers = {**headers, "content-type": "application/json"}

        data = json.dumps(self.compact(data), ensure_ascii=False).encode("utf-8")

        req = urllib.request.Request(url, data=data, headers=headers, method="POST")

        try:
            with urllib.request.urlopen(req, timeout=20) as response:
                content = response.read().decode("utf-8")
                return json.loads(content)
        except urllib.error.HTTPError as e:
            self._handle_exception(e)
            print("Error:", e)

        return None

    def _handle_exception(self, exception: urllib.error.HTTPError) -> None:
        """
        Handle an exception from the OpenAI API.
        :param exception: The exception from the OpenAI API.
        """
        body = json.loads(exception.file.read().decode("utf-8"))
        request_id = exception.headers.get("x-request-id")

        error = OpenAIError(
            request_id=request_id, status_code=exception.code, message=body
        )
        print("Error:", error)
        raise OpenAIError(
            request_id=request_id, status_code=exception.code, message=body
        )

    @staticmethod
    def compact(data: Json) -> Json:
        """
        Remove None values from a dictionary.
        """
        return {k: v for k, v in data.items() if v is not None}


class OperatorServicePreview(BaseService):
    """
    The Operator service class to interact with the Operator for Computer Using Agent (CUA) API.
    """

    def __init__(
        self, config: Dict[str, Any], agent_type: str = "operator", client=None
    ) -> None:
        """
        Create an Operator service instance.
        :param config: The configuration for the Operator service.
        :param agent_type: The type of the agent.

        """
        self.config_llm = config[agent_type]
        self.config = config
        self.api_type = self.config_llm["API_TYPE"].lower()
        self.api_model = self.config_llm["API_MODEL"].lower()
        self.max_retry = self.config["MAX_RETRY"]
        self.prices = self.config.get("PRICES", {})
        self._agent_type = agent_type

        if client is None:
            self.client = self.get_openai_client()

    def get_openai_client(self):
        """
        Create an OpenAI client based on the API type.
        :return: The OpenAI client.
        """

        # client = OpenAIBetaClient(
        #     endpoint=self.config_llm.get("API_BASE"),
        #     api_version=self.config_llm.get("API_VERSION", ""),
        # )

        token_provider = self.get_token_provider()
        api_key = token_provider()

        client = openai.AzureOpenAI(
            azure_endpoint=self.config_llm.get("API_BASE"),
            api_key=api_key,
            max_retries=self.max_retry,
            timeout=self.config.get("TIMEOUT", 20),
            api_version=self.config_llm.get("API_VERSION"),
            default_headers={"x-ms-enable-preview": "true"},
        )

        return client

    def chat_completion(
        self,
        message: Dict[str, Any] = None,
        n: int = 1,
    ) -> Tuple[Dict[str, Any], Optional[float]]:
        """
        Generates completions for a given conversation using the OpenAI Chat API.
        :param message: The message to send to the API.
        :param n: The number of completions to generate.
        :return: A tuple containing a list of generated completions and the estimated cost.
        """

        inputs = message.get("inputs", [])
        tools = message.get("tools", [])
        previous_response_id = message.get("previous_response_id", None)

        response = self.client.responses.create(
            model=self.config_llm.get("API_MODEL"),
            input=inputs,
            tools=tools,
            previous_response_id=previous_response_id,
            truncation="auto",
            temperature=self.config.get("TEMPERATURE", 0),
            top_p=self.config.get("TOP_P", 0),
            timeout=self.config.get("TIMEOUT", 20),
        ).model_dump()

        if "usage" in response:
            usage = response.get("usage")
            input_tokens = usage.get("input_tokens", 0)
            output_tokens = usage.get("output_tokens", 0)
        else:
            input_tokens = 0
            output_tokens = 0

        cost = self.get_cost_estimator(
            self.api_type,
            self.api_model,
            self.prices,
            input_tokens,
            output_tokens,
        )

        return [response], cost

    def get_token_provider(self):
        """
        Acquire token from Azure AD for OpenAI.
        :return: The access token for OpenAI.
        """

        from azure.identity import AzureCliCredential, get_bearer_token_provider

        tenant_id = self.config_llm.get("AAD_TENANT_ID", "")
        scope = self.config_llm.get("AAD_API_SCOPE", "")

        identity = AzureCliCredential(tenant_id=tenant_id)
        bearer_provider = get_bearer_token_provider(identity, scope)
        return bearer_provider


class OpenAIError(Exception):
    request_id: str
    status_code: int
    message: Dict[str, Any]

    def __init__(self, status_code: int, message: Dict[str, Any], request_id: str):
        """
        The OpenAI API error class.
        :param status_code: The status code of the API response.
        :param message: The error message from the API response.
        :param request_id: The request ID of the API response.
        """
        self.status_code = status_code
        self.message = message
        self.request_id = request_id
        super().__init__(f"OpenAI API error: {status_code} {message}")

    def __str__(self):
        return f"OpenAI API error: {self.request_id} {self.status_code} {json.dumps(self.message, ensure_ascii=False, indent=2)}"
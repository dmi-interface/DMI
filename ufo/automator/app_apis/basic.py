# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.


from abc import abstractmethod
from typing import Dict, List, Type
from typing import Optional
from pywinauto.controls.uiawrapper import UIAWrapper
import win32com.client

from ufo.automator.basic import CommandBasic, ReceiverBasic


class WinCOMReceiverBasic(ReceiverBasic):
    """
    The base class for Windows COM client.
    """

    _command_registry: Dict[str, Type[CommandBasic]] = {}

    def __init__(self, app_root_name: str, process_name: str, clsid: str) -> None:
        """
        Initialize the Windows COM client.
        :param app_root_name: The app root name.
        :param process_name: The process name.
        :param clsid: The CLSID of the COM object.
        """

        self.app_root_name = app_root_name
        self.process_name = process_name

        self.clsid = clsid

        self.client = win32com.client.Dispatch(self.clsid)
        self.com_object = self.get_object_from_process_name()

        self.application_window = None  # Add this attribute

    def set_application_window(self, window: UIAWrapper) -> None:
        """Set the application window reference."""
        self.application_window = window

    def get_application_window(self) -> UIAWrapper:
        """Get the application window."""
        return self.application_window

    @abstractmethod
    def get_object_from_process_name(self) -> win32com.client.CDispatch:
        """
        Get the object from the process name.
        """
        pass

    def get_suffix_mapping(self) -> Dict[str, str]:
        """
        Get the suffix mapping.
        :return: The suffix mapping.
        """
        suffix_mapping = {
            "WINWORD.EXE": "docx",
            "EXCEL.EXE": "xlsx",
            "POWERPNT.EXE": "pptx",
            "olk.exe": "msg",
        }

        return suffix_mapping.get(self.app_root_name, None)

    def app_match(self, object_name_list: List[str]) -> str:
        """
        Check if the process name matches the app root.
        :param object_name_list: The list of object name.
        :return: The matched object name.
        """

        suffix = self.get_suffix_mapping()

        if self.process_name.endswith(suffix):
            clean_process_name = self.process_name[: -len(suffix)]
        else:
            clean_process_name = self.process_name

        if not object_name_list:
            return ""

        return max(
            object_name_list,
            key=lambda x: self.longest_common_substring_length(clean_process_name, x),
        )

    @property
    def full_path(self) -> str:
        """
        Get the full path of the process.
        :return: The full path of the process.
        """
        try:
            full_path = self.com_object.FullName
            return full_path
        except:
            return ""

    def save(self) -> None:
        """
        Save the current state of the app.
        """
        try:
            self.com_object.Save()
        except:
            pass

    def save_to_xml(self, file_path: str) -> None:
        """
        Save the current state of the app to XML.
        :param file_path: The file path to save the XML.
        """
        try:
            self.com_object.SaveAs(file_path, self.xml_format_code)
        except:
            pass

    def close(self) -> None:
        """
        Close the app.
        """
        try:
            self.com_object.Close()
        except:
            pass

    @property
    def type_name(self):
        return "COM"

    @property
    def xml_format_code(self) -> int:
        pass

    @staticmethod
    def longest_common_substring_length(str1: str, str2: str) -> int:
        """
        Get the longest common substring of two strings.
        :param str1: The first string.
        :param str2: The second string.
        :return: The length of the longest common substring.
        """

        m = len(str1)
        n = len(str2)

        dp = [[0] * (n + 1) for _ in range(m + 1)]

        max_length = 0

        for i in range(1, m + 1):
            for j in range(1, n + 1):
                if str1[i - 1] == str2[j - 1]:
                    dp[i][j] = dp[i - 1][j - 1] + 1
                    if dp[i][j] > max_length:
                        max_length = dp[i][j]
                else:
                    dp[i][j] = 0

        return max_length


class WinCOMCommand(CommandBasic):
    """
    The abstract command interface.
    """

    def __init__(self, receiver: WinCOMReceiverBasic, params=None) -> None:
        """
        Initialize the command.
        :param receiver: The receiver of the command.
        """
        self.receiver = receiver
        self.params = params if params is not None else {}

    def get_ui_navigation_files(self, graph_file_key: str, id_map_file_key: str) -> tuple:
        import os
        from ufo.config.config import Config

        configs = Config.get_instance().config_data
        ui_navigation_config = configs.get("UI_NAVIGATION", {})

        # Flatten search: expand all leaf key-value pairs from nested dicts;
        # keys are globally unique, so there are no conflicts
        def flatten(d: dict) -> dict:
            result = {}
            for k, v in d.items():
                if isinstance(v, dict):
                    result.update(flatten(v))  # 递归展开子 dict
                else:
                    result[k] = v
            return result

        flat_config = flatten(ui_navigation_config)

        ui_graph_file = flat_config.get(graph_file_key, "")
        ui_graph_id_map_file = flat_config.get(id_map_file_key, "")

        if not ui_graph_file or not ui_graph_id_map_file:
            raise ValueError(
                f"UI navigation configuration not found. "
                f"Please configure {graph_file_key} and {id_map_file_key} in config file."
            )

        # Convert relative paths to absolute paths
        if not os.path.isabs(ui_graph_file):
            # Locate the project root via the ufo package path
            import ufo
            ufo_package_path = os.path.dirname(ufo.__file__)  # .../UFO/ufo
            project_root = os.path.dirname(ufo_package_path)  # .../UFO

            ui_graph_file = os.path.join(project_root, ui_graph_file.replace('/', os.sep))
            ui_graph_id_map_file = os.path.join(project_root, ui_graph_id_map_file.replace('/', os.sep))

        # Check if the files exist
        if not os.path.exists(ui_graph_file):
            raise FileNotFoundError(f"UI navigation graph file not found: {ui_graph_file}")

        if not os.path.exists(ui_graph_id_map_file):
            raise FileNotFoundError(f"UI navigation ID map file not found: {ui_graph_id_map_file}")

        return ui_graph_file, ui_graph_id_map_file

    @abstractmethod
    def execute(self):
        pass

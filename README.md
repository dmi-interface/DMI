<!-- markdownlint-disable MD033 MD041 -->
<h1 align="center">
  <b>DMI</b> :&nbsp;LLM-friendly OS interfaces for Computer-Use Agents
</h1>

<div align="center">

[![Paper](https://img.shields.io/badge/Paper-arXiv%3A2510.04607-b31b1b)](https://arxiv.org/abs/2510.04607)
[![Python](https://img.shields.io/badge/Python-3.10%20%7C%203.11-blue)](#)
[![License](https://img.shields.io/badge/License-MIT-yellow)](./LICENSE)

</div>
  



## Introduction
OSes have long evolved interfaces to serve different users—from command-line interface (CLI) for experts to GUI for general users. LLM-powered agents are a fundamentally new class of OS users with distinct characteristics (large context-memory, reasoning) and constraints (weak visual capability, latency, token costs, imperfect instruction-following) that existing OS interfaces do not consider or address. As a result, LLMs underperform with GUI.

DMI (Declarative Model Interface) addresses this gap by applying core systems principles (declarative over imperative, separating policy from mechanism, fast-path/slow-path) to build LLM-friendly OS interfaces.

DMI enables API-like declarative interactions for GUI applications.
DMI does **not** require the **application source code** or relying on application programming interfaces (APIs).

For more details, please refer to our paper: [EuroSys'26]: From Imperative to Declarative: Towards LLM-friendly OS Interfaces for Boosted Computer-Use Agents.


## Installation

### Step 1: Installation
DMI interface is integrated into [microsoft/UFO](https://github.com/microsoft/UFO) Windows computer-use agent framework. UFO requires **Python >= 3.10** running on **Windows OS >= 10**. It can be installed by running the following command:
```powershell
# [optional to create conda environment]
# conda create -n ufo python=3.10
# conda activate ufo

# clone the repository
git clone https://github.com/dmi-interface/DMI.git
cd DMI

# install the requirements
pip install -r requirements.txt
```

[//]: # (# note: you need to install Windows SDK &#40;https://learn.microsoft.com/en-us/windows/apps/windows-sdk/downloads&#41; to make sure the dependencies are properly installed.)

### Step 2: Configure the LLMs
Before running UFO agent, you need to provide your LLM configurations **individually for HostAgent and AppAgent**. You can edit config file `ufo/config/config.yaml` for **HOST_AGENT** and **APP_AGENT**. 


#### OpenAI
```yaml
VISUAL_MODE: True, # Whether to use the visual mode
API_TYPE: "openai" , # The API type, "openai" for the OpenAI API.  
API_KEY: "sk-",  # The OpenAI API key, begin with sk-
API_VERSION: "latest",
API_MODEL: "gpt-5",
```

[//]: # ()
[//]: # (#### Azure OpenAI &#40;AOAI&#41;)

[//]: # (```yaml)

[//]: # (VISUAL_MODE: True, # Whether to use the visual mode)

[//]: # (API_TYPE: "aoai" , # The API type, "aoai" for the Azure OpenAI.  )

[//]: # (API_BASE: "YOUR_ENDPOINT", #  The AOAI API address. Format: https://{your-resource-name}.openai.azure.com)

[//]: # (API_KEY: "YOUR_KEY",  # The aoai API key)

[//]: # (API_VERSION: "2024-02-15-preview", # "2024-02-15-preview" by default)

[//]: # (API_MODEL: "gpt-4o",  # The only OpenAI model)

[//]: # (API_DEPLOYMENT_ID: "YOUR_AOAI_DEPLOYMENT", # The deployment id for the AOAI API)

[//]: # (```)

[//]: # ()
[//]: # (> Need Qwen, Gemini, non‑visual GPT‑4, or even **OpenAI CUA Operator** as a AppAgent? See the [model guide]&#40;https://microsoft.github.io/UFO/supported_models/overview/&#41;.)


### Step 3: Set Office Language to English (United States)

DMI has been tested with **Microsoft Office 365 (Version 2604, Build 16.0.19929.20032)**.

Make sure you are using the **English (United States)** display language in Office. Otherwise, DMI may not be able to correctly recognize or interact with Office applications.

You can change the setting by following these steps:

1. Open any Office application, such as **Word**, **Excel**, or **PowerPoint**.
2. Click **File** in the top-left corner.
3. Select **Options**.
4. In the dialog window, go to the **Language** tab.
5. Under **Office display language**, make sure **English (United States)** is selected.
6. Click **Set as Preferred** (or **Set as Default**, depending on your Office version).
7. Click **OK** to save the changes and close the window.
8. Restart the Office application for the changes to take effect.

[//]: # (9. *&#40;Optional&#41;* We also recommend setting your **Windows system language** to **English &#40;United States&#41;** for better compatibility.)



### Step 4: Start UFO Agent

#### ⌨️ You can execute the following on your Windows command Line (CLI):

```powershell
# assume you are in the cloned DMI folder

# [optional] activate ufo environment
# conda activate ufo

python -m ufo --task <your_task_name>
```

This will start the UFO process and you can interact with it through the command line interface. 
If everything goes well, you will see the following message:

```powershell
Welcome to use UFO🛸, A UI-focused Agent for Windows OS Interaction. 
 _   _  _____   ___
| | | ||  ___| / _ \
| | | || |_   | | | |
| |_| ||  _|  | |_| |
 \___/ |_|     \___/
Please enter your request to be completed🛸:
```

Alternatively, you can also directly invoke UFO with a specific task and request by using the following command:

```powershell
python -m ufo --task <your_task_name> -r "<your_request>"
```


###  Step 5: Execution Logs 

You can find the screenshots taken and request & response logs in the following folder:
```
./ufo/logs/<your_task_name>/
```
You may use them to debug, replay, or analyze the agent output.

## What’s Included

This repository focuses on open-sourcing and reproducing the **DMI execution layer**. To make the project directly usable, we provide **prebuilt core English topologies** for executing DMI on supported applications.
The core topologies support task execution in **OSWorld-Windows** for Microsoft Office **Word, Excel, and PowerPoint**:
https://github.com/xlang-ai/OSWorld/tree/main/evaluation_examples/examples_windows. For functionality not covered by the released topologies, the system falls back to UFO2’s default GUI interface.

## Acknowledgements
This project is based on [microsoft/UFO](https://github.com/microsoft/UFO) agent framework,
which is licensed under the MIT License.

We sincerely thank the Microsoft UFO authors and maintainers for open-sourcing their work.
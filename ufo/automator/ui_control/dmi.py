# %% [markdown]
# # Start
# %%
import time
import sys
import os
from pathlib import Path
# from ufo.automator.ui_control.dump_tree_jupyter_latest_new import inspector

# Option 1: use pathlib to locate the UFO/DMI directory (recommended)
current_file = Path(__file__).resolve() if '__file__' in globals() else Path.cwd()

# Define the project keywords to search for
PROJECT_KEYWORDS = ('UFO', 'DMI')

# Walk upward from the current file location until a directory containing any keyword is found
ufo_dir = current_file
while not any(kw in ufo_dir.name for kw in PROJECT_KEYWORDS) and ufo_dir.parent != ufo_dir:
    ufo_dir = ufo_dir.parent

if any(kw in ufo_dir.name for kw in PROJECT_KEYWORDS):
    sys.path.insert(0, str(ufo_dir))
    print(f"✅ Added project directory to sys.path: {ufo_dir}")
else:
    print("⚠️ Project directory not found, using a relative path instead")
    # Fallback: relative path based on the current file
    backup_ufo_dir = Path(__file__).parent.parent.parent.parent
    sys.path.insert(0, str(backup_ufo_dir))

# Get the current working directory
import traceback
import json
import html
import re

import pywinauto
import openai
import pyautogui
import tiktoken
import networkx as nx
import matplotlib.pyplot as plt
import concurrent.futures
import threading
import comtypes
import pywinauto.uia_defines
import comtypes.client


try:
    from comtypes.gen import UIAutomationClient as UAC
except ImportError:
    comtypes.client.GetModule('UIAutomationCore.dll')
    from comtypes.gen import UIAutomationClient as UAC
from pywinauto.uia_element_info import UIAElementInfo
from pywinauto.controls.uiawrapper import UIAWrapper
import win32gui
import win32con
import win32process
import pywinauto.uia_defines as uia

from abc import ABC, abstractmethod
from datetime import datetime
from typing import List, Optional, Dict, Any
from enum import Enum
from collections import defaultdict, deque
from pywinauto import Desktop
from pyvis.network import Network
from copy import deepcopy
from pywinauto.findwindows import find_element
from pywinauto.findwindows import find_elements
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys
from pywinauto.controls.uiawrapper import UIAWrapper
from comtypes import COMError
from comtypes import byref
from pywinauto.uia_defines import NoPatternInterfaceError  # <-- newly added
from pywinauto import mouse
from openai import OpenAI

from functools import lru_cache
from matplotlib.font_manager import FontProperties
# from ufo.automator.ui_control.controller import TextTransformer
from ufo.automator.ui_control.inspector import ControlInspectorFacade
from ufo.automator.ui_control.ui_tree import UITree
# os.environ["OPENAI_API_KEY"] = ""

# %%
# Multilingual label configuration


# ============================================================
# UI control label variants table
# Format: "semantic_key" -> [variant1, variant2, ...]  (any number of languages/dialects/versions)
# To add a new language, simply append it to the corresponding list; no other logic needs to change
# ============================================================
UI_LABEL_VARIANTS: dict[str, list[str]] = {
    # Common dialog buttons
    "cancel":            ["取消", "Cancel"],
    "close":             ["关闭", "Close"],
    "ok":                ["确定", "OK", "Ok"],
    "no":                ["否", "否(N)", "No", "No(N)"],
    "yes":               ["是", "是(Y)", "Yes", "Yes(Y)"],
    # Task pane closing flow
    "task_pane_options": ["任务窗格选项", "Task Pane Options"],
    # Extended wait for wait_for_window_completion
    "replace_ellipsis":  ["替换...", "Replace...", "Replace…"],
    "save_as_ellipsis":  ["另存为...", "Save As..."],
    "save_as":           ["另存为", "Save As"],
    "new_rule_ellipsis": ["新建规则...", "New Rule..."],
    "create_pdf_xps":    ["创建 PDF/XPS", "Create PDF/XPS"],
    "format_ellipsis":   ["格式...", "Format..."],
    "save_s":            ["保存(S)", "Save(S)", "Save"],
    # For recapture_controls_info_in_modal_if_exists
    "save_as_window":    ["另存为", "Save As"],
    "publish_pdf_xps":   ["发布为 PDF 或 XPS", "Publish as PDF or XPS"],
    # LocalizedControlType
    "localized_type_textbox": ["文本框", "textbox"],
    "localized_type_image":   ["图像",  "image"],
    "localized_type_chart":   ["图表",  "chart"],

}


WHITELIST_KEYWORD_VARIANTS: list[str] = [
]


def matches_label(name: str, label_key: str) -> bool:
    """Check whether a control name matches any variant of a semantic label."""
    return name in UI_LABEL_VARIANTS.get(label_key, [])


def find_controls_by_label(window, label_key: str, control_type: str) -> list:
    """Find controls by semantic label, automatically iterating through all language variants."""
    for name in UI_LABEL_VARIANTS.get(label_key, []):
        results = find_controls_by_name_and_type(window, name, control_type)
        if results:
            return results
    return []


def find_controls_by_label_from_all_descendants(window, label_key: str, control_type: str) -> list:
    """Same as above, but uses find_controls_by_name_and_type_from_all_descendants."""
    for name in UI_LABEL_VARIANTS.get(label_key, []):
        results = find_controls_by_name_and_type_from_all_descendants(window, name, control_type)
        if results:
            return results
    return []


# %%
# Set the control types handled in navigation
GLOBAL_INSPECTOR = ControlInspectorFacade(backend="uia")
GLOBAL_UI_NAVIGATION_TYPES = ["Button", "MenuItem", "TabItem", "CheckBox","RadioButton", "ComboBox", "ListItem", "TreeItem","Edit",]
# , "Hyperlink","SplitButton"
UIA_FULL_DESCRIPTION_PROPERTY_ID = pywinauto.uia_defines.IUIA().UIA_dll.UIA_FullDescriptionPropertyId  # Or retrieve it from uia_defines
GLOBAL_TAB_CONTAINER = None
# This variable exists by default. Note: when executing LLM instructions, make sure GLOBAL_TAB_CONTAINER is set to None; otherwise it will search the main window by default, causing Step 2 in navigate_and_execute to make an incorrect judgment
GLOBAL_ANCESTOR_COUNT = 3  # Default is 3; can be adjusted as needed. The values used in exploration and execution within the same app should match
# %%
def get_window_by_title_contains(title_keyword):
    """
    Find the corresponding application window by a substring match in the window title (case-insensitive).
    :param title_keyword: A keyword contained in the window title, e.g. "Document"
    :return: The first matching UIAWrapper window object, or None if not found
    """
    if not title_keyword:
        return None

    # Build a case-insensitive regular expression
    pattern = re.compile(re.escape(title_keyword), re.IGNORECASE)

    inspector = ControlInspectorFacade(backend="uia")
    windows = inspector.get_desktop_windows(remove_empty=True)

    for window in windows:
        try:
            title = window.window_text() or ""
        except Exception:
            continue

        # Search using the regular expression
        if pattern.search(title):
            return window

    return None

def get_supported_patterns(control):
    """Print information about all supported UIA patterns for a control."""
    try:
        # Get the control's IUIAutomationElement object
        element = control.element_info.element

        # Import UIA definitions
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Mapping between UIA pattern IDs and names, using the correct property names
        pattern_checks = [
            (UIA_dll.UIA_IsInvokePatternAvailablePropertyId, "InvokePattern"),
            (UIA_dll.UIA_IsSelectionPatternAvailablePropertyId, "SelectionPattern"),
            (UIA_dll.UIA_IsValuePatternAvailablePropertyId, "ValuePattern"),
            (UIA_dll.UIA_IsRangeValuePatternAvailablePropertyId, "RangeValuePattern"),
            (UIA_dll.UIA_IsScrollPatternAvailablePropertyId, "ScrollPattern"),
            (UIA_dll.UIA_IsExpandCollapsePatternAvailablePropertyId, "ExpandCollapsePattern"),
            (UIA_dll.UIA_IsGridPatternAvailablePropertyId, "GridPattern"),
            (UIA_dll.UIA_IsGridItemPatternAvailablePropertyId, "GridItemPattern"),
            (UIA_dll.UIA_IsMultipleViewPatternAvailablePropertyId, "MultipleViewPattern"),
            (UIA_dll.UIA_IsWindowPatternAvailablePropertyId, "WindowPattern"),
            (UIA_dll.UIA_IsSelectionItemPatternAvailablePropertyId, "SelectionItemPattern"),
            (UIA_dll.UIA_IsDockPatternAvailablePropertyId, "DockPattern"),
            (UIA_dll.UIA_IsTablePatternAvailablePropertyId, "TablePattern"),
            (UIA_dll.UIA_IsTableItemPatternAvailablePropertyId, "TableItemPattern"),
            (UIA_dll.UIA_IsTextPatternAvailablePropertyId, "TextPattern"),
            (UIA_dll.UIA_IsTogglePatternAvailablePropertyId, "TogglePattern"),
            (UIA_dll.UIA_IsTransformPatternAvailablePropertyId, "TransformPattern"),
            (UIA_dll.UIA_IsScrollItemPatternAvailablePropertyId, "ScrollItemPattern"),
            (UIA_dll.UIA_IsLegacyIAccessiblePatternAvailablePropertyId, "LegacyIAccessiblePattern"),
            (UIA_dll.UIA_IsItemContainerPatternAvailablePropertyId, "ItemContainerPattern"),
            (UIA_dll.UIA_IsVirtualizedItemPatternAvailablePropertyId, "VirtualizedItemPattern"),
            (UIA_dll.UIA_IsSynchronizedInputPatternAvailablePropertyId, "SynchronizedInputPattern"),
            (UIA_dll.UIA_IsObjectModelPatternAvailablePropertyId, "ObjectModelPattern"),
            (UIA_dll.UIA_IsAnnotationPatternAvailablePropertyId, "AnnotationPattern"),
            (UIA_dll.UIA_IsTextPattern2AvailablePropertyId, "TextPattern2"),
            (UIA_dll.UIA_IsStylesPatternAvailablePropertyId, "StylesPattern"),
            (UIA_dll.UIA_IsSpreadsheetPatternAvailablePropertyId, "SpreadsheetPattern"),
            (UIA_dll.UIA_IsSpreadsheetItemPatternAvailablePropertyId, "SpreadsheetItemPattern"),
            (UIA_dll.UIA_IsTransformPattern2AvailablePropertyId, "TransformPattern2"),
            (UIA_dll.UIA_IsTextChildPatternAvailablePropertyId, "TextChildPattern"),
            (UIA_dll.UIA_IsDragPatternAvailablePropertyId, "DragPattern"),
            (UIA_dll.UIA_IsDropTargetPatternAvailablePropertyId, "DropTargetPattern"),
            (UIA_dll.UIA_IsTextEditPatternAvailablePropertyId, "TextEditPattern"),
            (UIA_dll.UIA_IsCustomNavigationPatternAvailablePropertyId, "CustomNavigationPattern")
        ]

        # Find the supported patterns
        supported = []

        for property_id, pattern_name in pattern_checks:
            try:
                is_supported = element.GetCurrentPropertyValue(property_id)
                if is_supported:
                    supported.append((property_id, pattern_name))
            except Exception:
                pass  # Ignore unsupported properties

        # Print the results
        if not supported:
            # print(f"Control '{control.window_text()}' does not support any known UIA patterns")
            return []

        # print(f"UIA patterns supported by control '{control.window_text()}':")
        # for property_id, pattern_name in supported:
        #     print(f"  - {pattern_name} (PropertyID: {property_id})")

        return supported

    except Exception as e:
        print(f"Error while getting control pattern information: {str(e)}")
        return []

def get_full_description(control):
    """Get the FullDescription property of a control."""
    try:
        # Access the control's underlying UIA element
        element = control.element_info.element

        # Get the FullDescription property
        # Use UIA_FullDescriptionPropertyId (30159)
        full_description = element.GetCurrentPropertyValue(UIA_FULL_DESCRIPTION_PROPERTY_ID)
        return full_description
    except Exception as e:
        return f"fail to get the FullDescription property: {str(e)}"
# %%
def is_text_pattern_supported(control):
    """
    Check whether a control supports TextPattern.

    Args:
        control: The control object

    Returns:
        bool: True if TextPattern is supported; returns False if retrieval fails
    """
    try:
        # Get the control's IUIAutomationElement object
        element = control.element_info.element
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Check whether the control supports TextPattern
        supports_text = element.GetCurrentPropertyValue(UIA_dll.UIA_IsTextPatternAvailablePropertyId)

        return bool(supports_text)

    except Exception:
        return False

def is_selection_pattern_supported(control):
    """
    Check whether a control supports SelectionPattern.

    Args:
        control: The control object

    Returns:
        bool: True if SelectionPattern is supported; returns False if retrieval fails
    """
    try:
        # Get the control's IUIAutomationElement object
        element = control.element_info.element
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Check whether the control supports SelectionPattern
        supports_selection = element.GetCurrentPropertyValue(UIA_dll.UIA_IsSelectionPatternAvailablePropertyId)

        return bool(supports_selection)

    except Exception:
        return False

def is_expand_collapse_pattern_supported(control):
    """
    Check whether a control supports ExpandCollapsePattern.

    Args:
        control: The control object

    Returns:
        bool: True if ExpandCollapsePattern is supported; returns False if retrieval fails
    """
    try:
        # Get the control's IUIAutomationElement object
        element = control.element_info.element
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Check whether the control supports ExpandCollapsePattern
        supports_expand_collapse = element.GetCurrentPropertyValue(UIA_dll.UIA_IsExpandCollapsePatternAvailablePropertyId)

        return bool(supports_expand_collapse)

    except Exception:
        return False

def is_toggle_pattern_supported(control):
    """
    Check whether a control supports TogglePattern.

    Args:
        control: The control object

    Returns:
        bool: True if TogglePattern is supported; returns False if retrieval fails
    """
    try:
        # Get the control's IUIAutomationElement object
        element = control.element_info.element
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Check whether the control supports TogglePattern
        supports_toggle = element.GetCurrentPropertyValue(UIA_dll.UIA_IsTogglePatternAvailablePropertyId)

        return bool(supports_toggle)

    except Exception:
        return False

# %%
# Determine whether a window is a true modal window
def is_modal_window(window):
    """Check whether a window is a modal window."""
    try:
        return window.iface_window.CurrentIsModal
    except:
        return False

def are_same_controls(control1, control2):
    """
    Determine whether two UIAWrapper controls are the same control.

    Args:
        control1, control2: UIAWrapper objects

    Returns:
        bool: Whether they are the same control
    """
    try:
        # Method 1: compare the underlying UIA Element objects (most reliable)
        if hasattr(control1, 'element_info') and hasattr(control2, 'element_info'):
            elem1 = control1.element_info.element
            elem2 = control2.element_info.element

            # UIA-provided comparison method
            return elem1.Compare(elem2) == 0

    except Exception:
        pass

    # Method 2: compare Process ID + Handle (if available)
    try:
        pid1 = control1.process_id()
        pid2 = control2.process_id()

        if pid1 != pid2:
            return False

        # Try to get the window handle
        try:
            handle1 = control1.handle
            handle2 = control2.handle
            return handle1 == handle2
        except:
            pass
    except:
        pass

    # Method 3: compare a combination of multiple properties
    try:
        # Compare automation_id (if present and unique)
        auto_id1 = getattr(control1.element_info, 'automation_id', '')
        auto_id2 = getattr(control2.element_info, 'automation_id', '')

        if auto_id1 and auto_id2:
            if auto_id1 != auto_id2:
                return False

        # Compare rectangle (position and size)
        rect1 = control1.rectangle()
        rect2 = control2.rectangle()

        if rect1 != rect2:
            return False

        # Compare control type
        type1 = control1.element_info.control_type
        type2 = control2.element_info.control_type

        if type1 != type2:
            return False

        # Compare name
        name1 = control1.element_info.name or control1.window_text()
        name2 = control2.element_info.name or control2.window_text()

        return name1 == name2

    except Exception as e:
        print(f"Error while comparing controls: {e}")
        return False

# is_window_valid is somewhat similar to check_window_controls_still_exist
def is_window_valid(window):
    """Check whether a window is still valid."""
    if window is None:
        return False

    try:
        pid = window.process_id()
        # A process ID of None means the window is no longer valid
        return pid is not None
    except Exception:
        # Any exception indicates the window is no longer valid
        return False
# %%
# Get the topmost window (app-level window) within the app that shares the same PID
# This may be useful for getting the current topmost current_window
def get_top_window_by_zorder(target_pid):
    """
    Get the highest visible window in Z-Order for the specified process.
    # Note: if a child window inside the app-level window is not modal, this function returns the app-level window.
    # If there is a modal child window, it returns the topmost modal child window (nested cases included).

    Args:
        target_pid: Target process ID

    Returns:
        UIAWrapper: The highest window in Z-Order
    """
    windows = []

    def enum_windows_callback(hwnd, lParam):
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)

            if window_pid == target_pid and win32gui.IsWindowVisible(hwnd):
                # Get the window title and filter out system windows
                window_text = win32gui.GetWindowText(hwnd)
                if window_text:  # Keep only windows that have a title
                    windows.append((hwnd, window_text))
        except:
            pass
        return True

    win32gui.EnumWindows(enum_windows_callback, None)

    if windows:
        # EnumWindows returns windows in Z-Order, so the first one is the topmost
        top_hwnd, top_title = windows[0]
        print(f"Top window in Z-Order: {top_title}")
        return UIAWrapper(UIAElementInfo(handle_or_elem=top_hwnd))

    return None

def get_top_level_window_from_current(target_control):
    """
    Traverse upward from any control to find the top-level control of type Window.

    Args:
        target_control: The target_control window object

    Returns:
        UIAWrapper: The top-level Window object
    """
    try:
        current = target_control
        top_level_window = None

        # Traverse upward to find the top-level Window
        while current:
            # Check whether the current control is of type Window
            if current.element_info.control_type == "Window":
                top_level_window = current

            try:
                parent = current.parent()
                # Stop traversing if the parent is the desktop or no parent exists
                if not parent or parent.element_info.control_type == "Desktop":
                    break
                current = parent
            except Exception:
                break

        # Return the top-level Window that was found; if none is found, return the original window
        return top_level_window if top_level_window else target_control

    except Exception as e:
        print(f"Error while getting the top-level Window: {str(e)}")
        return target_control
# %%
def calculate_z_order_rank(hwnd, pid):
    """
    Calculate the Z-Order rank of a window within the same process.

    Args:
        hwnd: Window handle
        pid: Process ID

    Returns:
        int: Z-Order rank (0 = frontmost)
    """
    rank = 0
    h = win32gui.GetWindow(hwnd, win32con.GW_HWNDPREV)
    while h:
        _, p = win32process.GetWindowThreadProcessId(h)
        if p == pid and win32gui.IsWindowVisible(h):
            rank += 1
        h = win32gui.GetWindow(h, win32con.GW_HWNDPREV)
    return rank

def sort_windows_by_z_order(window_controls, strict_mode=True):
    """
    General-purpose window Z-Order sorting function

    Args:
        window_controls: Array of window controls
        strict_mode: Strict mode. If True, all windows must belong to the same process;
                    if False, windows are filtered by the process ID of the first window

    Returns:
        list: Window list sorted by Z-Order (topmost first)

    Raises:
        ValueError: Raised when strict_mode=True and the windows do not belong to the same process
    """
    if not window_controls:
        return []

    try:
        # Get the process IDs of all windows
        window_pids = []
        for window in window_controls:
            try:
                window_pids.append(window.process_id())
            except Exception as e:
                print(f"Failed to get the process ID for a window; skipping this window: {str(e)}")
                continue

        if not window_pids:
            return []

        # Check whether all windows belong to the same process
        unique_pids = set(window_pids)

        if len(unique_pids) > 1:
            if strict_mode:
                raise ValueError(f"The window list contains windows from multiple processes, so a valid Z-order sort cannot be performed. Process IDs: {unique_pids}")
            else:
                # Non-strict mode: use the process ID of the first window and filter windows from the same process
                target_pid = window_pids[0]
                print(f"⚠️ Multiple processes detected ({unique_pids}); filtering by the first window's process ID ({target_pid})")

                filtered_windows = []
                for window in window_controls:
                    try:
                        if window.process_id() == target_pid:
                            filtered_windows.append(window)
                    except Exception:
                        continue

                window_controls = filtered_windows
                if not window_controls:
                    return []

        # Perform Z-Order sorting using the unified process ID
        target_pid = window_pids[0] if len(unique_pids) == 1 else target_pid

        # Sort by Z-Order
        ranked = sorted(
            window_controls,
            key=lambda c: calculate_z_order_rank(c.element_info.handle, target_pid)
        )

        return ranked

    except ValueError:
        # Re-raise ValueError (strict mode error)
        raise
    except Exception as e:
        print(f"Error while sorting windows: {str(e)}")
        # If sorting fails, return the original list
        return window_controls

def filter_valid_accessible_windows(window_controls):
    """
    Filter valid window controls, excluding temporary windows that have no title and are too small.

    Args:
        window_controls: Array of window controls

    Returns:
        list: Filtered list of valid windows
    """
    if not window_controls:
        return []

    filtered_windows = []

    for window in window_controls:
        should_keep = False

        try:
            # Performance optimization: check the window title first
            window_text = window.window_text()

            # Keep it directly if the title is not empty
            if window_text and window_text.strip():
                should_keep = True
            else:
                # If the title is empty, determine whether it is a temporary small window by size
                try:
                    rect = window.rectangle()
                    max_dimension = max(rect.width(), rect.height())
                    # If the largest dimension is smaller than 80, treat it as a temporary small window and discard it
                    should_keep = max_dimension >= 80
                except Exception:
                    # rectangle() raised an error; discard the window
                    should_keep = False

        except Exception:
            # window_text() raised an error; discard the window
            should_keep = False

        if should_keep:
            filtered_windows.append(window)

    return filtered_windows

# %%
# Return the topmost child window of the input window: [] or [window]
# Note: this can detect child windows such as “Clipboard” or “Navigation”
def detect_single_child_window_in_app(main_window, depth=None):
    """
    Detect the first child_window dialog inside the current application window.
    Due to detection instability, only the first detected child_window is returned.
    That window is returned as the one highest in Z-Order.

    Args:
        main_window: Main window object
        depth: Search depth limit (has no effect under high-performance cached UIA, but retained for interface compatibility)

    Returns:
        list: A list containing the first detected dialog; returns an empty list if none is detected
    """
    try:
        # 1. Collect candidate window controls, with exception handling added
        dialogs = []
        try:
            # print("debug, current window is", main_window)
            candidates = main_window.descendants(control_type="Window", depth=depth)
            # print("debug: all candidates are:",candidates)
            for c in candidates:
                try:
                    if c.is_visible() and win32gui.IsWindow(c.element_info.handle):
                        dialogs.append(c)
                except Exception as e:
                    # Ignore visibility check failures for individual controls
                    print(f"Failed to check control visibility; skipping this control: {str(e)}")
                    continue
            # print("debug: all candidates dialogs are:",dialogs)
            # for candidate_window in dialogs:
            #     print(win32gui.IsWindow(candidate_window.element_info.handle))
        except Exception as e:
            print(f"Error while getting descendants of type Window: {str(e)}")
            return []

        if not dialogs:
            print("No visible Window controls found")
            return []
        # for dialog in dialogs:
        #     if not win32gui.IsWindow(dialog.element_info.handle):
        #         print("!!!!!")

        # 2. Filter valid windows first, then sort by Z-Order
        valid_dialogs = filter_valid_accessible_windows(dialogs)
        if not valid_dialogs:
            return []

        ranked = sort_windows_by_z_order(valid_dialogs, strict_mode=False)
        # ranked = sort_windows_by_z_order(dialogs, True)
        if not ranked:
            return []

        # 3. Get the highest window in Z-Order and add logging output
        frontmost_window = ranked[0]
        # print("debug: are same window:",are_same_controls(frontmost_window, main_window))
        try:
            # Quickly get control information from cache
            control_type = frontmost_window.element_info.control_type
            title = frontmost_window.element_info.name or ""

            # If name is empty, try getting window_text
            if not title:
                try:
                    title = frontmost_window.window_text() or ""
                except Exception:
                    title = ""

            # Add debug print output
            if title:
                print(f"Detected child window: {title} ({control_type})")
            else:
                print(f"Detected child window: [untitled] ({control_type})")

        except Exception as e:
            print(f"Error while detecting dialog: {str(e)}")

        return [frontmost_window]

    except Exception as e:
        print(f"Error while executing detect_single_child_window_in_app: {str(e)}")
        return []

# %%
# This function could potentially be replaced with is_window_valid or check_window_controls_still_exist
def is_child_window_still_open_in_main_window(main_window, dialog_title, depth=None):
    """
    Check whether a dialog with a specific title is still open.

    Args:
        main_window: Main window object
        dialog_title: Dialog title

    Returns:
        bool: Returns True if the dialog is still open; otherwise False
    """
    try:
        # Updated to use the new detection function
        dialogs = detect_single_child_window_in_app(main_window, depth=depth)
        for dialog in dialogs:
            if dialog.window_text() == dialog_title:
                return True
        return False
    except Exception:
        # If an exception occurs, conservatively assume the dialog may still be open
        return True

# Close child windows inside the app during the explore phase; by default no save action is taken
# This function is used when Esc cannot close the window during exploration
def close_single_child_window(main_window, depth=None):
    """
    Close the first child window dialog inside the current application.
    # Note: in theory, only one level can be closed at a time, so the caller may need to invoke it multiple times.

    Args:
        main_window: Main window object
        depth: Search depth limit

    Returns:
        bool: Whether the first detected dialog was successfully closed
    """
    try:
        # Detect the first dialog
        dialogs = detect_single_child_window_in_app(main_window, depth=depth)

        if not dialogs:
            print("No child_window dialog detected")
            return True  # If there is no child_window, treat it as already closed

        dialog = dialogs[0]  # Take the first (and only) one
        dialog_title = dialog.window_text()
        print(f"Trying to close child_window dialog: {dialog_title}")

        closed = False

        # Method 1: look for a Cancel button
        # print("Method 1")
        if not closed:
            try:
                all_controls = dialog.descendants()
                cancel_buttons = [c for c in all_controls
                                  if matches_label(c.element_info.name, "cancel")
                                  and c.element_info.control_type == "Button"]

                if cancel_buttons:
                    print("Found a Cancel button, clicking to close...")
                    main_window.set_focus()
                    cancel_buttons[0].set_focus()
                    cancel_buttons[0].click_input()
                    time.sleep(0.2)

                    # Check whether this specific dialog was successfully closed
                    closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                    if closed:
                        print("✓ Successfully closed child_window using the Cancel button")
                        return True
            except Exception as e:
                #print(f"Cancel-button method failed: {str(e)}")
                pass

        # Method 2: look for a Close button
        # print("Method 2")
        if not closed:
            try:
                all_controls = dialog.descendants()
                close_buttons = [c for c in all_controls
                               if matches_label(c.element_info.name, "close") and
                               c.element_info.control_type == "Button"]

                if close_buttons:
                    print("Found a Close button, clicking to close...")
                    main_window.set_focus()
                    close_buttons[0].set_focus()
                    close_buttons[0].click_input()
                    time.sleep(0.2)

                    # Check whether this specific dialog was successfully closed
                    closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                    if closed:
                        print("✓ Successfully closed child_window using the Close button")
                        return True
            except Exception as e:
                #print(f"Close-button method failed: {str(e)}")
                pass

        # Method 3: find and use Task Pane Options and the Close menu item
        # print("Method 3")
        if not closed:
            try:
                task_pane_options = find_controls_by_label_from_all_descendants(dialog.parent(), "task_pane_options", "MenuItem")
                # task_pane_options = find_controls_by_name_and_type_from_all_descendants(main_window, "任务窗格选项", "MenuItem")
                if task_pane_options:
                    print("Found Task Pane Options, clicking...")
                    main_window.set_focus()
                    task_pane_options[0].set_focus()
                    task_pane_options[0].click_input()
                    time.sleep(0.2)

                    # Try clicking the Close menu item
                    close_menu_items  = find_controls_by_label_from_all_descendants(dialog.parent(), "close", "MenuItem")
                    # close_menu_items = find_controls_by_name_and_type_from_all_descendants(main_window, "关闭", "MenuItem")
                    if close_menu_items:
                        print("Found the Close menu item, clicking...")
                        main_window.set_focus()
                        close_menu_items[0].set_focus()
                        close_menu_items[0].click_input()
                        time.sleep(0.2)

                        # Check whether this specific dialog was successfully closed
                        closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                        if closed:
                            print("✓ Successfully closed child_window using Task Pane Options")
                            return True
            except Exception as e:
                print(f"Task Pane Options method failed: {str(e)}")

        # Method 4: find and use the "No" button, then press ESC
        # (This logic is newly added and can even close two child windows)
        # print("Method 4")
        if not closed:
            try:
                all_controls = dialog.descendants()
                no_buttons = [c for c in all_controls
                             if matches_label(c.element_info.name, "no")
                              and c.element_info.control_type == "Button"]

                if no_buttons:
                    print("Found a No(N) button, clicking it and then using ESC to close...")
                    main_window.set_focus()
                    no_buttons[0].set_focus()
                    no_buttons[0].click_input()
                    time.sleep(0.2)  # Slightly extend the wait time

                    # Send the ESC key
                    send_keys("{ESC}")
                    time.sleep(0.1)

                    # Check whether this specific dialog was successfully closed
                    closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                    if closed:
                        print("✓ Successfully closed child_window using the No(N) button + ESC")
                        return True
            except Exception as e:
                print(f"No(N)-button method failed: {str(e)}")


        # All methods failed
        if not closed:
            print(f"✗ Unable to close child_window dialog: {dialog_title}")
            return False

    except Exception as e:
        print(f"Error while closing dialog: {str(e)}")
        return False
# %%
# Build the complete tree via the children() method; used for corner cases
def get_all_descendants_by_calling_children(window, depth=None, max_workers=4):
    """
    Get all descendant controls using the children() approach,
    following the concurrent processing logic of build_ui_tree.

    Args:
        window: The window object whose descendants are to be retrieved
        depth: Maximum depth; None means no depth limit
        max_workers: Maximum number of threads

    Returns:
        list: A list containing all descendant controls
    """
    if not window:
        return []

    # Thread-local storage to avoid cross-thread conflicts
    _thread_local = threading.local()

    def collect_descendants(control, current_depth=0):
        """Recursively collect descendant controls."""
        try:
            result = [control]  # Include the current control

            # Stop searching if the specified depth has been reached
            if depth is not None and current_depth >= depth:
                return result


            # Get child controls
            children = []
            try:
                children = control.children()
            except Exception:
                return result

            # If there are only a few child controls, process them serially
            if len(children) <= 3:
                for child in children:
                    child_descendants = collect_descendants(child, current_depth + 1)
                    result.extend(child_descendants)
            else:
                # Process child controls concurrently
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    # Submit all child tasks
                    future_to_child = {
                        executor.submit(collect_descendants, child, current_depth + 1): child
                        for child in children
                    }

                    # Collect results
                    for future in concurrent.futures.as_completed(future_to_child):
                        try:
                            child_descendants = future.result(timeout=5)  # 5-second timeout
                            result.extend(child_descendants)
                        except Exception as e:
                            print(f"Error while processing child control: {str(e)}")
                            continue

            return result

        except Exception as e:
            print(f"Error while collecting control descendants: {str(e)}")
            return [control] if control else []

    # Start from the window and collect all descendants
    return collect_descendants(window)

# This function is specifically responsible for finding "task pane options"
def find_controls_by_name_and_type_from_all_descendants(window, name, control_type):
    """
    Find controls by name and type, and return a list of all matched controls.
    The controls must exist in the current UI tree to be found; used for corner cases.

    Args:
        window: Window object
        name: Control name
        control_type: Control type

    Returns:
        list: A list of matched controls; returns an empty list if none are found
    """
    if not window:
        return []

    matched_controls = []

    try:
        # Search through all descendant controls
        for ctrl in get_all_descendants_by_calling_children(window, depth=None):
            try:
                # Get the control name (window text)
                # ctrl_name = ctrl.window_text() or ""
                ctrl_name = ctrl.element_info.name or ""  # Previously ctrl.window_text()

                # Get the control type
                ctrl_type = ctrl.element_info.control_type

                # Check whether the name and type match
                if (name == "" or ctrl_name == name) and ctrl_type == control_type:
                    matched_controls.append(ctrl)
            except Exception:
                continue
    except Exception:
        return []

    return matched_controls

# The main performance overhead is in detect_single_child_window_in_app
# %%
# Execute the solving logic, handling edit and similar cases

# Close child windows inside the window, used during the LLM instruction execution phase
def confirm_and_close_single_child_window_with_priority(main_window, depth=None):
    """
    Close child dialog windows in the current application based on priority.
    Priority: OK > Close > Cancel

    Args:
        main_window: Main window object
        depth: Search depth limit

    Returns:
        bool: Whether all dialogs were successfully closed
    """
    try:
        # Detect dialog windows
        dialogs = detect_single_child_window_in_app(main_window, depth=depth)

        if not dialogs:
            return True

        print("Number of dialogs detected:", len(dialogs))
        all_closed = True

        for index, dialog in enumerate(dialogs, 1):
            print(f"Processing dialog {index}")
            dialog_title = dialog.window_text()
            print(f"Attempting to close dialog: {dialog_title}")
            closed = False

            # Priority 1: Look for the OK button
            if not closed:
                try:
                    all_controls = dialog.descendants()
                    ok_buttons = [c for c in all_controls
                                if matches_label(c.element_info.name, "ok") and
                                c.element_info.control_type == "Button"]

                    if ok_buttons:
                        print("Found the OK button, clicking to close...")
                        main_window.set_focus()
                        ok_buttons[0].set_focus()
                        ok_buttons[0].click_input()
                        time.sleep(0.1)

                        closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                        if closed:
                            print("Dialog successfully closed via the OK button")
                            continue
                except Exception as e:
                    print(f"Failed to click the OK button: {str(e)}")

            # Priority 2: Look for the Close button
            if not closed:
                try:
                    all_controls = dialog.descendants()
                    close_buttons = [c for c in all_controls
                                   if matches_label(c.element_info.name, "close") and
                                   c.element_info.control_type == "Button"]

                    if close_buttons:
                        print("Found the Close button, clicking to close...")
                        main_window.set_focus()
                        close_buttons[0].set_focus()
                        close_buttons[0].click_input()
                        time.sleep(0.1)

                        closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                        if closed:
                            print("Dialog successfully closed via the Close button")
                            continue
                except Exception as e:
                    print(f"Failed to click the Close button: {str(e)}")

            # Priority 3: Look for the Cancel button
            if not closed:
                try:
                    all_controls = dialog.descendants()
                    cancel_buttons = [c for c in all_controls
                                    if matches_label(c.element_info.name, "cancel") and
                                    c.element_info.control_type == "Button"]

                    if cancel_buttons:
                        print("Found the Cancel button, clicking to close...")
                        main_window.set_focus()
                        cancel_buttons[0].set_focus()
                        cancel_buttons[0].click_input()
                        time.sleep(0.1)

                        closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                        if closed:
                            print("Dialog successfully closed via the Cancel button")
                            continue
                except Exception as e:
                    print(f"Failed to click the Cancel button: {str(e)}")

            # Fallback method: find and use task pane options and the Close menu item
            if not closed:
                try:
                    task_pane_options = find_controls_by_label_from_all_descendants(dialog.parent(), "task_pane_options", "MenuItem")
                    if task_pane_options:
                        print("Found task pane options, clicking...")
                        main_window.set_focus()
                        task_pane_options[0].set_focus()
                        task_pane_options[0].click_input()
                        time.sleep(0.2)

                        # Try clicking the Close menu item
                        close_menu_items = find_controls_by_label_from_all_descendants(dialog.parent(), "close", "MenuItem")
                        if close_menu_items:
                            print("Found the Close menu item, clicking...")
                            main_window.set_focus()
                            close_menu_items[0].set_focus()
                            close_menu_items[0].click_input()
                            time.sleep(0.1)

                            closed = not is_child_window_still_open_in_main_window(main_window, dialog_title, depth=depth)
                            if closed:
                                print("Dialog successfully closed via task pane options")
                                continue
                except Exception as e:
                    print(f"Task pane options method failed: {str(e)}")

            # If all methods above fail, mark as not fully closed
            if not closed:
                all_closed = False
                print(f"Unable to close dialog: {dialog_title}")

        return all_closed

    except Exception as e:
        print(f"Error while closing dialogs: {str(e)}")
        return False

# Save changes and close the top-level child window itself (preferably by clicking OK)
def confirm_and_close_window(child_window):
    """
    Used during LLM command execution.
    Close the top-level child window with confirmation/save prioritized.
    Priority: OK > Close > ESC > Enter

    Args:
        child_window: The child window to close

    Returns:
        bool: Whether it was successfully closed
    """
    try:
        print(f"Attempting to close window: {child_window.window_text()}")

        # Method 1: Look for the "OK" button
        try:
            ok_buttons = find_controls_by_label(child_window, "ok", "Button")
            if ok_buttons:
                print("Found the OK button, clicking to close")
                child_window.set_focus()
                ok_buttons[0].click_input()
                time.sleep(0.5)
                return True
        except Exception:
            pass

        # Method 2: Look for the "Close" button
        try:
            close_buttons = find_controls_by_label(child_window, "close", "Button")
            if close_buttons:
                print("Found the Close button, clicking to close")
                child_window.set_focus()
                close_buttons[0].click_input()
                time.sleep(0.5)
                return True
        except Exception:
            pass

        # Method 3: Send the ESC key
        try:
            print("Trying to close by sending the ESC key")
            child_window.set_focus()
            send_keys("{ESC}")
            time.sleep(0.5)
            return True
        except Exception:
            pass

        # Method 4: Send the Enter key
        try:
            print("Trying to close by sending the Enter key")
            child_window.set_focus()
            send_keys("{ENTER}")
            time.sleep(0.5)
            return True
        except Exception:
            pass

        return False

    except Exception as e:
        print(f"Error while closing child window: {str(e)}")
        return False

# %%

# %%
# There may be multiple automation ids
# This function returns multiple controls matching the same automation_id
def find_controls_by_automation_id(window, automation_id, control_type=None):
    """
    Find all matching controls by automation ID and return them as a list.

    Args:
        window: Window object
        automation_id: The target control's automation ID
        control_type: Optional control type filter, such as "Edit", "Button", etc.

    Returns:
        list: A list of all matched UIAWrapper control objects; returns an empty list if none are found
    """
    try:
        # Recursively find all matching elements
        elems = find_elements(
            parent=window.element_info,
            auto_id=automation_id,
            control_type=control_type,
            backend="uia",
            top_level_only=False
        )

        # Convert all elements to UIAWrapper objects
        controls = []
        for elem in elems:
            try:
                controls.append(UIAWrapper(elem))
            except Exception:
                continue

        return controls

    except Exception as e:
        print(f"Error finding controls by automation_id={automation_id}: {e}")
        return []

# Consider solution 3 later
# (if it is still not unique enough: when obtaining parent/ancestor, only consider clickable controls)
# %%
# Format generation
# Get the unique identifier of a control
def get_control_identifier(control, ancestor_count=None):
    """
    Get the unique identifier of a control.
    Always use the alt_id format, including automation_id information.

    Args:
        control: Control object
        ancestor_count: Number of ancestor nodes to retrieve, defaults to 3

    Returns:
        str: The control's unique identifier, in the format
             "alt_id:(automation_id) or name|control_type|ancestor_info"
             automation_id is wrapped in parentheses; when absent, it is shown as |(none)|;
        # When automation_id exists (stable identifier):
        "alt_id:(automation_id)|control_type|ancestor_info"

        # When automation_id does not exist:
        "alt_id:name|control_type|ancestor_info"
    """
    if ancestor_count is None:
        ancestor_count = GLOBAL_ANCESTOR_COUNT
    try:
        ctrl_name = control.element_info.name or "[NoText]"
        ctrl_type = control.element_info.control_type

        # Get automation_id
        automation_id = ""
        try:
            automation_id = control.element_info.automation_id or ""
        except Exception:
            pass

        # Get ancestor information
        ancestor_info = get_named_ancestors(control, ancestor_count)

        # Choose the format based on whether automation_id exists
        if automation_id:
            # Use the stable format when automation_id is available, without relying on the mutable name
            unique_id = f"alt_id:({automation_id})|{ctrl_type}|{ancestor_info}"
        else:
            # Use name when automation_id is not available
            unique_id = f"alt_id:{ctrl_name}|{ctrl_type}|{ancestor_info}"

        return unique_id
    except Exception:
        return None

def get_named_ancestors(control, ancestor_count=None, max_depth=8):
    """
    Find the specified number of ancestor nodes for a control,
    including unnamed node information to improve uniqueness
    (nodes with names are considered ancestors).

    Args:
        control: Control object
        ancestor_count: Number of ancestor nodes to retrieve, defaults to 3
        max_depth: Maximum lookup depth to prevent infinite loops

    Returns:
        str: Ancestor node information, separated by "/"
             When automation_id exists: (automation_id)
             When automation_id does not exist: name
             If fewer ancestors are found than requested, returns the actual number found
             If none are found, returns "[NoAncestor]"
    """
    if ancestor_count is None:
        ancestor_count = GLOBAL_ANCESTOR_COUNT

    current = control
    depth = 0
    ancestors = []

    while current and depth < max_depth and len(ancestors) < ancestor_count:
        try:
            # Try to get the parent control
            current = current.parent()
            if not current:
                break

            # Get control information
            name = current.element_info.name or ""
            automation_id = ""

            try:
                automation_id = current.element_info.automation_id or ""
            except Exception:
                pass

            # Build ancestor information using the same format as the main control
            if automation_id:
                # Use the stable format when automation_id exists
                ancestor_info = f"({automation_id})"
            else:
                # Use name when automation_id does not exist
                if name and name.strip():
                    ancestor_info = name.strip()
                else:
                    ancestor_info = "[Unnamed]"

            ancestors.append(ancestor_info)
            depth += 1

        except Exception:
            break

    if not ancestors:
        return "[NoAncestor]"

    # Join ancestor information with "/"
    return "/".join(ancestors)

def get_named_ancestors_novel(control, named_ancestor_count=None, max_depth=8):
    """
    Find the specified number of ancestor nodes for a control,
    including unnamed node information to improve uniqueness.

    Args:
        control: Control object
        named_ancestor_count: Number of named ancestor nodes to retrieve, defaults to 3
        max_depth: Maximum lookup depth to prevent infinite loops, defaults to 8

    Returns:
        str: Ancestor node information, separated by "/"
             When automation_id exists: (automation_id)
             When automation_id does not exist: name
             All ancestors will be included in the return value,
             but only named ancestors are counted toward named_ancestor_count
             If fewer named ancestors are found than requested, returns the actual number found
             If none are found, returns "[NoAncestor]"
    """
    if named_ancestor_count is None:
        named_ancestor_count = GLOBAL_ANCESTOR_COUNT

    current = control
    depth = 0
    ancestors = []
    named_count = 0  # Track the number of named ancestors

    while current and depth < max_depth and named_count < named_ancestor_count:
        try:
            # Try to get the parent control
            current = current.parent()
            if not current:
                break

            # Get control information
            name = current.element_info.name or ""
            automation_id = ""

            try:
                automation_id = current.element_info.automation_id or ""
            except Exception:
                pass

            # Build ancestor information using the same format as the main control
            if automation_id:
                # Use the stable format when automation_id exists
                ancestor_info = f"({automation_id})"
                named_count += 1  # automation_id counts as a named ancestor
            else:
                # Use name when automation_id does not exist
                if name and name.strip():
                    ancestor_info = name.strip()
                    named_count += 1  # A name counts as a named ancestor
                else:
                    ancestor_info = "[Unnamed]"
                    # Unnamed ancestors do not count toward named_count

            ancestors.append(ancestor_info)
            depth += 1

        except Exception:
            break

    if not ancestors:
        return "[NoAncestor]"

    # Join ancestor information with "/"
    return "/".join(ancestors)

# Format parsing
def parse_ancestor_info(ancestor_info):
    """
    Parse the ancestor name string.
    Supports the new formats: (automation_id) or name.

    Args:
        ancestor_names: Ancestor name string, such as "(auto1)/(auto2)/parent_control_name"
    Returns:
        list: A list containing ancestor information,
              where each element is a dict with name and automation_id
    """
    if not ancestor_info or ancestor_info == "[NoAncestor]":
        return []

    ancestors = []
    for ancestor_str in ancestor_info.split("/"):
        if ancestor_str.startswith("(") and ancestor_str.endswith(")"):
            # Format: (automation_id)
            automation_id = ancestor_str[1:-1]  # Remove parentheses
            name = ""  # Do not use name, because it may change
        else:
            # Format: name
            automation_id = ""
            name = ancestor_str

        ancestors.append({
            "name": name,
            "automation_id": automation_id,
            # "depth": depth
        })

    return ancestors

def parse_unique_id(unique_id):
    """
    Parse unique_id. Supports two formats:
    - alt_id:(automation_id)|control_type|ancestor_info
    - alt_id:name|control_type|ancestor_info

    Args:
        unique_id: The full unique_id

    Returns:
        dict: A dictionary containing the parsed result, or None if parsing fails
              {
                  "name": str,  # Empty when automation_id exists
                  "control_type": str,
                  "automation_id": str,  # Empty when automation_id does not exist
                  "ancestor_info": str
              }
    """
    if not unique_id.startswith("alt_id:"):
        return None

    alt_id = unique_id[len("alt_id:"):]
    parts = alt_id.split("|", 2)  # Now there are 3 parts
    if len(parts) != 3:
        return None

    first_part, control_type, ancestor_info = parts

    # Determine whether the first part is automation_id or name
    if first_part.startswith("(") and first_part.endswith(")"):
        # Format with automation_id
        automation_id = first_part[1:-1]  # Remove parentheses
        name = ""  # Do not use name, because it may change
    else:
        # Format without automation_id
        automation_id = ""
        name = first_part

    return {
        "name": name,
        "control_type": control_type,
        "automation_id": automation_id,
        "ancestor_info": ancestor_info
    }

# %%
# Find the specific control
# Note: there may be multiple controls with the same automation_id in the current window
def find_control_by_identifier(window, identifier, control_type=None):
    """
    Find a control by its unique identifier.
    Only the alt_id format is supported.
    """
    if not identifier or not identifier.startswith("alt_id:"):
        return None

    alt_id = identifier[len("alt_id:"):]
    return find_control_by_alt_id(window, alt_id)

def find_control_by_alt_id(window, alt_id):
    """
    Find a control by alt_id using efficient matching logic.
    """
    # Parse alt_id
    parsed = parse_unique_id(f"alt_id:{alt_id}")
    if not parsed:
        return None

    target_name = parsed["name"]
    target_type = parsed["control_type"]
    target_automation_id = parsed["automation_id"]
    target_ancestor_info = parsed["ancestor_info"]

    # Strategy 0: If the control has automation_id, use efficient automation_id lookup
    if target_automation_id:
        return find_control_with_automation_id_with_exclusion(
            window, target_automation_id, target_name, target_type, target_ancestor_info
        )

    # Strategy 1: If there is no automation_id, fall back to the original full-scan method
    # Directly match the full unique_id
    # Get all control information
    controls_info, control_instances = get_control_descendants_with_info(window)

    full_unique_id = f"alt_id:{alt_id}"
    # Direct dictionary lookup, O(1) time complexity
    if full_unique_id in controls_info:
        return control_instances[full_unique_id]

    return None

def find_control_with_automation_id_with_exclusion(window, target_automation_id, target_name, target_type, target_ancestor_info):
    """
    Optimized version:
    quickly find controls by automation_id, then perform detailed comparison only on candidate controls.
    """
    # 1. Quickly get all candidate controls matching automation_id and control_type
    # Note: target_type is parsed from uniqueid, so control_type will be applied inside the search function
    candidate_controls = find_controls_by_automation_id(window, target_automation_id, target_type)

    if not candidate_controls:
        return None

    # 2. Validate ancestor information for all candidate controls, regardless of count
    target_ancestors = parse_ancestor_info(target_ancestor_info)

    valid_candidates = []
    for control in candidate_controls:
        try:
            # Generate unique_id only for this candidate control (small performance cost)
            control_unique_id = get_control_identifier(control)
            if not control_unique_id:
                continue

            # Parse the current control's ancestor information
            parsed = parse_unique_id(control_unique_id)
            if not parsed:
                continue

            current_ancestors = parse_ancestor_info(parsed["ancestor_info"])

            # Exclude deterministically incorrect matches
            # Even if there is only one candidate, it still needs to be validated!
            if is_definitely_wrong_ancestor_match(target_ancestors, current_ancestors):
                continue  # Deterministically wrong, exclude it

            valid_candidates.append(control)

        except Exception:
            continue

    # 3. Return the result
    if not valid_candidates:
        return None  # All candidates were excluded

    if len(valid_candidates) == 1:
        return valid_candidates[0]

    # Multiple candidates still remain; return the first one
    # (or return None to indicate uncertainty)
    print(f"Warning: automation_id '{target_automation_id}' matched multiple candidates; returning the first one", valid_candidates)
    return valid_candidates[0]

def is_definitely_wrong_ancestor_match(target_ancestors, current_ancestors):
    """
    Determine whether this is a definitively incorrect ancestor match
    (used for exclusion, not best-effort matching).

    Returns:
        bool: True means it is definitively incorrect and should be excluded;
              False means it cannot be determined as incorrect and should be kept
    """
    if not target_ancestors or not current_ancestors:
        return False  # Cannot determine, do not exclude

    # Compare total number of ancestors (overall depth)
    if len(target_ancestors) != len(current_ancestors):
        return True  # Different ancestor counts, definitively incorrect

    for i in range(len(target_ancestors)):
        target = target_ancestors[i]
        current = current_ancestors[i]

        # If ancestor automation_id differs, it is definitely not a match
        if (target["automation_id"] and current["automation_id"] and
            target["automation_id"] != current["automation_id"]):
            return True  # Deterministically incorrect, exclude

    return False  # Cannot determine as incorrect, keep it

# More efficient method
def find_controls_by_name_and_type(window, name, control_type):
    """
    Find controls by name and type, and return a list of all matched controls.
    The controls must exist in the current UI tree to be found.

    Args:
        window: Window object
        name: Control name
        control_type: Control type

    Returns:
        list: A list of matched controls; returns an empty list if none are found
    """
    if not window:
        return []

    matched_controls = []

    try:
        # Search through all descendant controls
        for ctrl in window.descendants():
            try:
                # Get the control name (window text)
                # ctrl_name = ctrl.window_text() or ""
                ctrl_name = ctrl.element_info.name or ""  # Previously ctrl.window_text()
                # Get the control type
                ctrl_type = ctrl.element_info.control_type

                # Check whether the name and type match
                if (name == "" or ctrl_name == name) and ctrl_type == control_type:
                    matched_controls.append(ctrl)
            except Exception:
                continue
    except Exception:
        return []

    return matched_controls

# %%
# The UI tree uses control.children instead of descendants.
# (When a popup exists, control.children can obtain the whole app tree)
# whitelist_keywords has a default value inside the function!
def is_blacklisted(name, blacklist_keywords, whitelist_keywords=None):

    """
    Check whether a control name is in the blacklist, with whitelist priority support.

    Args:
        name: Control name
        blacklist_keywords: List of blacklist keywords
        whitelist_keywords: List of whitelist keywords, defaults to None

    Returns:
        bool: Whether it is blacklisted
    """

    if not whitelist_keywords:

        whitelist_keywords = WHITELIST_KEYWORD_VARIANTS

    if not name or not blacklist_keywords:
        return False

    name_lower = name.lower()


    if whitelist_keywords:
        for keyword in whitelist_keywords:
            if keyword in name:
                return False

    # Check blacklist
    for keyword in blacklist_keywords:
        if keyword.lower() in name_lower:
            return True

    return False


# %%
# Top-level child window (desktop-level child window) monitoring logic
class IUIAutomationManager:
    """UIA automation manager - singleton pattern"""
    _instance = None
    _lock = threading.Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if not self._initialized:
            self._uia = comtypes.client.CreateObject(
                UAC.CUIAutomation,
                interface=UAC.IUIAutomation
            )
            self._initialized = True

    @property
    def uia(self):
        """Get the IUIAutomation instance"""
        return self._uia

    def get_root_element(self):
        """Get the root element"""
        return self._uia.GetRootElement()

    def add_event_handler(self, event_id, element, scope, cache_request, handler):
        """Add an event listener"""
        return self._uia.AddAutomationEventHandler(
            event_id, element, scope, cache_request, handler
        )

    def remove_event_handler(self, event_id, element, handler):
        """Remove an event listener"""
        return self._uia.RemoveAutomationEventHandler(event_id, element, handler)

class WindowEventHandler(comtypes.COMObject):
    """Window event handler"""
    _com_interfaces_ = [UAC.IUIAutomationEventHandler]

    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.start_time = time.perf_counter()

    def HandleAutomationEvent(self, sender, eventID):
        try:
            elem_info = UIAElementInfo(sender)
            wrapper = UIAWrapper(elem_info)
            dt = time.perf_counter() - self.start_time
            self.callback(wrapper, dt)
        except Exception as e:
            print(f"Event handling error: {e}")

class TopLevelWindowDetector:
    """Top-level child window (app-level) listener context manager with waiting support"""

    def __init__(self, original_window_pid, timeout=5, control_name=""):
        self.original_window_pid = original_window_pid
        self.timeout = timeout
        self.control_name = control_name

        # Listener results
        self.new_window = None
        self.detection_time = None

        # Synchronization control
        self.stop_event = threading.Event()
        self.ready_event = threading.Event()
        self.monitor_thread = None
        self.click_time = None

    def __enter__(self):
        """Enter the context and start the listener"""
        self.monitor_thread = threading.Thread(target=self._monitor_new_window, daemon=True)
        self.monitor_thread.start()

        # Wait until the listener is ready
        if not self.ready_event.wait(timeout=5.0):
            print(f"⚠️ {self.control_name} listener startup timed out, click will still proceed")
        else:
            print(f"✅ {self.control_name} listener is ready")
            pass

        # Ensure the listener has started (preserve the original sleep logic)
        time.sleep(0.05)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Exit the context and clean up the listener"""
        self.stop_event.set()

        # Simple cleanup; do not wait here
        if self.monitor_thread and self.monitor_thread.is_alive():
            self.monitor_thread.join(0.1)

    def _monitor_new_window(self):
        """Internal method for monitoring new windows"""
        # ❌ Do not set ready here
        # self.ready_event.set()

        try:
            # Call the modified wait_for_new_window_with_stop_ready
            new_window, dt = wait_for_new_window_with_stop_ready(
                timeout=self.timeout,
                target_parent_pid=self.original_window_pid,
                stop_event=self.stop_event,
                ready_event=self.ready_event  # Pass in ready_event
            )

            if new_window and not self.stop_event.is_set():
                self.new_window = new_window
                self.detection_time = dt
                print(f"  {self.control_name} detected a new popup window: {new_window.window_text()}")

        except Exception as e:
            if not self.stop_event.is_set():
                print(f"  {self.control_name} popup listener exception: {e}")

    def wait_for_window_completion(self, control_name="", control_type=""):
        """
        Wait for window detection to complete, reusing the original waiting logic

        Args:
            control_name: Control name
            control_type: Control type

        Returns:
            bool: Whether a new window was detected
        """

        CONTROL_WAIT_CONFIG: dict[tuple[str, str], float] = {
            ("replace_ellipsis", "Button"): 3.0,
            ("save_as_ellipsis", "Button"): 3.0,
            ("save_as", "Button"): 3.0,
            ("new_rule_ellipsis", "MenuItem"): 3.0,
            ("create_pdf_xps", "Button"): 3.0,
            ("format_ellipsis", "Button"): 3.0,  # Excel new rule...
            ("save_s", "Button"): 3.0,
            # ("print",           "Button"):   2.0,
            # ("export",          "Button"):   5.0,
        }

        def _get_control_wait_time(control_name: str, control_type: str) -> float:
            for (label_key, ctype), wait in CONTROL_WAIT_CONFIG.items():
                if control_type == ctype and matches_label(control_name, label_key):
                    return wait
            return 0.0

        wait_time = _get_control_wait_time(control_name, control_type)

        if wait_time > 0:
            # For Replace / Save As... buttons, wait x seconds
            if self.monitor_thread and self.monitor_thread.is_alive():
                self.monitor_thread.join(timeout=wait_time)
                # time.sleep(3)
                print(f"Extra window-listening wait of {wait_time} seconds completed")
            print(f"  {self.control_name} listener task finished, sending stop signal...")
            self.stop_event.set()

        else:
            # Do not wait in other cases
            if self.monitor_thread and self.monitor_thread.is_alive():
                self.monitor_thread.join(0)
            self.stop_event.set()

        return self.new_window is not None

    def get_new_window(self):
        """Get the newly detected window"""
        return self.new_window

# Newly added function with ready-signal support
def wait_for_new_window_with_stop_ready(timeout=5, target_parent_pid=None, stop_event=None, ready_event=None):
    """
    New window listener with ready-signal support

    Args:
        timeout: Timeout duration
        target_parent_pid: Target parent process PID
        stop_event: Stop signal event
        ready_event: Ready signal event

    Returns:
        (new window UIAWrapper, detection elapsed time) or (None, None)
    """
    uia_manager = IUIAutomationManager()
    root = uia_manager.get_root_element()
    result_store = []
    done_event = threading.Event()

    def window_callback(wrapper, dt):
        if stop_event and stop_event.is_set():
            # print("debug: # Received stop signal, return immediately")
            return  # Received stop signal, return immediately

        try:
            # print("debug: new window")

            window_text = wrapper.window_text() or ""
            class_name = wrapper.element_info.class_name or ""

            # Skip system windows
            skip_classes = ["Shell_TrayWnd", "TaskListThumbnailWnd", "MSTaskSwWClass"]
            if class_name in skip_classes:
                return

            # If a parent PID is specified, filter by it
            if target_parent_pid is not None:
                try:
                    if wrapper.process_id() != target_parent_pid:
                        return
                except:
                    return

            print(f"Detected new top-level child window: '{window_text}' (Class: {class_name}, PID: {wrapper.process_id()}, Elapsed: {dt:.3f}s)")
            result_store.extend([wrapper, dt])
            done_event.set()

        except Exception as e:
            if not (stop_event and stop_event.is_set()):
                print(f"Error while handling top-level child window callback: {e}")

    handler = WindowEventHandler(window_callback)

    try:
        # Register the listener
        uia_manager.add_event_handler(
            UAC.UIA_Window_WindowOpenedEventId,
            root,
            UAC.TreeScope_Children,
            None,
            handler
        )

        # ✅ Set ready only after the listener has actually been registered
        if ready_event:
            ready_event.set()

        # Wait for event or stop signal
        start_time = time.time()
        while not done_event.is_set() and timeout > 0:
            if stop_event and stop_event.is_set():
                print("Received stop signal, exiting listener")
                break

            wait_time = min(0.1, timeout)  # Wait at most 0.1 seconds each time
            done_event.wait(wait_time)
            timeout -= wait_time

    finally:
        try:
            uia_manager.remove_event_handler(
                UAC.UIA_Window_WindowOpenedEventId,
                root,
                handler
            )
        except:
            pass

    return (result_store[0], result_store[1]) if result_store else (None, None)

def wait_for_new_window_with_stop(timeout=5, target_parent_pid=None, stop_event=None):
    """
    New window listener with stop-signal support

    Args:
        timeout: Timeout duration
        target_parent_pid: Target parent process PID
        stop_event: Stop signal event

    Returns:
        (new window UIAWrapper, detection elapsed time) or (None, None)
    """
    uia_manager = IUIAutomationManager()
    root = uia_manager.get_root_element()
    result_store = []
    done_event = threading.Event()

    def window_callback(wrapper, dt):
        if stop_event and stop_event.is_set():
            return  # Received stop signal, return immediately

        try:
            window_text = wrapper.window_text() or ""
            class_name = wrapper.element_info.class_name or ""

            # Skip system windows
            skip_classes = ["Shell_TrayWnd", "TaskListThumbnailWnd", "MSTaskSwWClass"]
            if class_name in skip_classes:
                return

            # If a parent PID is specified, filter by it
            if target_parent_pid is not None:
                try:
                    if wrapper.process_id() != target_parent_pid:
                        return
                except:
                    return

            print(f"Detected new top-level child window: '{window_text}' (Class: {class_name}, PID: {wrapper.process_id()}, Elapsed: {dt:.3f}s)")
            result_store.extend([wrapper, dt])
            done_event.set()

        except Exception as e:
            if not (stop_event and stop_event.is_set()):
                print(f"Error while handling top-level child window callback: {e}")

    handler = WindowEventHandler(window_callback)

    try:
        uia_manager.add_event_handler(
            UAC.UIA_Window_WindowOpenedEventId,
            root,
            UAC.TreeScope_Children,
            None,
            handler
        )

        # Wait for event or stop signal
        start_time = time.time()
        while not done_event.is_set() and timeout > 0:
            if stop_event and stop_event.is_set():
                print("Received stop signal, exiting listener")
                break

            wait_time = min(0.1, timeout)  # Wait at most 0.1 seconds each time
            done_event.wait(wait_time)
            timeout -= wait_time

    finally:
        try:
            uia_manager.remove_event_handler(
                UAC.UIA_Window_WindowOpenedEventId,
                root,
                handler
            )
        except:
            pass

    return (result_store[0], result_store[1]) if result_store else (None, None)



# %%
# Versions below 0 use descendants (applicable to both win32 and UIA) instead of the native UIA interface

# Note: In PPT, consider handling the element named "桌面" separately (virtual_control: filter_virtual_controls) to improve performance
# Note: only_visible and only_enabled support three different values; only None means disabled
def get_control_descendants_with_info(window, blacklist_keywords=None, include_windows=False, control_type_list=None, only_visible=None, only_enabled=None, filter_virtual_controls=True):
    """
    Get all descendant controls of a window along with their information, while caching control instances.
    If a control is passed in but is no longer available (for example, a closed window),
    this will return an empty result without raising an error message.

    Args:
        window: Window object
        blacklist_keywords: List of blacklist keywords
        include_windows: Whether to include controls of type Window
        control_type_list: List of control types to retrieve; if None, use the default GLOBAL_UI_NAVIGATION_TYPES
        only_visible: Whether to retrieve only visible controls
        only_enabled: Whether to retrieve only enabled controls
        filter_virtual_controls: Whether to filter out virtual controls (default: True)

    Returns:
        tuple: (control info dict, control instance dict)
    """
    controls_info = {}
    control_instances = {}  # Added: cache control instances

    try:
        # If no control_type_list is specified, use the default global types
        if control_type_list is None:
            control_types = GLOBAL_UI_NAVIGATION_TYPES.copy()
        else:
            control_types = control_type_list.copy()

        if include_windows:
            control_types.append("Window")

        # Convert to a set to improve lookup efficiency
        control_types_set = set(control_types)

        # Use window.descendants() to get all descendant controls
        # descendants() does not support querying multiple control types,
        # so retrieve all controls first and then filter them
        all_controls = window.descendants()

        for control in all_controls:
            try:
                # Get the control type and filter it
                control_type = control.element_info.control_type
                if control_type not in control_types_set:
                    continue

                # Added: virtual control filtering
                if filter_virtual_controls:
                    try:
                        # Only check area when the control is both visible and enabled
                        if control.is_visible() and control.is_enabled():
                            # Check the control's rectangle area
                            rect = control.rectangle()
                            if rect.width() == 0 or rect.height() == 0:
                                print(f"  Filtered out virtual control: {control.element_info.name or '[Unnamed]'} [{control_type}] - area is 0")
                                continue  # Filter out zero-area controls

                        # If the control is invisible or disabled, do not filter it out; keep it
                        # No extra logic is needed here, just continue processing

                    except Exception as e:
                        # If the check cannot be performed, skip this control conservatively
                        print(f"  Error while checking virtual control, skipping: {control.element_info.name or '[Unnamed]'} - {str(e)}")
                        continue

                # Visibility filtering
                if only_visible is not None:
                    try:
                        is_visible = control.is_visible()
                        if only_visible and not is_visible:
                            continue
                        elif not only_visible and is_visible:
                            continue
                    except Exception:
                        # If visibility information cannot be retrieved, skip this filter
                        pass

                # Added: enabled-state filtering
                if only_enabled is not None:
                    try:
                        is_enabled = control.is_enabled()
                        if only_enabled and not is_enabled:
                            continue
                        elif not only_enabled and is_enabled:
                            continue
                    except Exception:
                        # If enabled-state information cannot be retrieved, skip this filter
                        pass

                # Quickly get the control's basic info from cache
                name = control.element_info.name or ""

                # Added: filter out completely empty names
                if not name.strip():
                    continue

                # Check whether the control name is in the blacklist
                if blacklist_keywords and is_blacklisted(name, blacklist_keywords):
                    continue

                # Quickly get automation_id and full_description from cache
                # automation_id = control.element_info.automation_id or ""
                full_description = get_full_description(control) or ""

                # Generate a unique identifier
                unique_id = get_control_identifier(control)

                if unique_id:  # Only process controls with a unique ID
                    controls_info[unique_id] = {
                        "unique_id": unique_id,
                        "name": name,
                        "control_type": control_type,
                        "full_description": full_description
                    }
                    # Cache the control instance
                    control_instances[unique_id] = control

            except Exception:
                continue

    except Exception as e:
        print(f"Error while retrieving descendant controls: {str(e)}")

    return controls_info, control_instances

# %%
# Refactor: single responsibility

def perform_path_repair(current_window, navigation_path, current_index):
    """
    Perform intelligent path-repair logic

    Returns:
        tuple: (found_control: UIAWrapper or None, new_index: int)
    """
    found_alternative = False
    alternative_start_index = -1

    for j in range(current_index + 1, len(navigation_path)):
        alternative_path_id = navigation_path[j]
        alternative_control = find_control_by_identifier(current_window, alternative_path_id)
        if alternative_control:
            print(f"  Found a later control {alternative_path_id}; skipping intermediate steps")
            alternative_start_index = j
            found_alternative = True
            break

    if found_alternative:
        return find_control_by_identifier(current_window, navigation_path[alternative_start_index]), alternative_start_index
    else:
        return None, current_index

# %%
def check_window_controls_still_exist(window_controls):
    """
    Check whether Window control objects still exist

    Args:
        window_controls: List of Window control objects detected earlier

    Returns:
        list: List of Window control objects that still exist
    """
    existing_windows = []
    for control in window_controls:
        try:
            # Try accessing the control's properties to check whether it still exists
            if control.is_enabled() and control.is_visible():
                existing_windows.append(control)
        except Exception:
            # If access fails, the control no longer exists
            continue

    return existing_windows


# %%
# section 2

# %%
import re
import json
from collections import defaultdict

# %%
# Post-processing: convert the navigation forest into a compact text format suitable as LLM output

def build_entry_ref_data(ui_graph_file: str, ui_graph_id_map_file: str) -> dict:
    """
    Build the entry whitelist data structure from the "forest graph" and the numeric ID mapping file:
        {
          "subtree_roots": [<root_num_id>, ...],                # ascending int list
          "entry_ref_whitelist": { "<root_num_id>": [<...>], ... }  # key is str, value is ascending int list
        }

    Identification rules:
      - Subtree root: node attribute subtree_root == True and contains defines (pointing to the original root orig_id)
      - Reference stub (ref): node attribute is_ref == True and ref_to = <orig_id>
      - Use orig_id as the bridge: collect all refs pointing to the same orig into that subtree root's whitelist

    Constraints:
      - id_map must be provided; automatic numbering is not allowed.
      - The id_map file may be in either of the following formats:
          1) {"id_map": { "<unique_id>": <num or str num>, ... }, "reverse_id_map": {...}}
          2) { "<unique_id>": <num or str num>, ... }
      - If a subtree root or its refs cannot be found in id_map, raise ValueError.

    Args:
        ui_graph_file: Path to the forest graph JSON file (containing "nodes" / "edges")
        ui_graph_id_map_file: Path to the numeric ID mapping JSON file

    Returns:
        dict: {"subtree_roots": List[int], "entry_ref_whitelist": Dict[str, List[int]]}
    """
    import json

    # Read the graph
    with open(ui_graph_file, "r", encoding="utf-8") as f:
        graph = json.load(f)
    nodes = graph.get("nodes", {})
    if not isinstance(nodes, dict) or not nodes:
        raise ValueError("ui_graph_file does not contain a valid 'nodes' section.")

    # Read id_map (both container formats are supported)
    with open(ui_graph_id_map_file, "r", encoding="utf-8") as f:
        id_map_raw = json.load(f)
    id_map = id_map_raw["id_map"] if isinstance(id_map_raw, dict) and "id_map" in id_map_raw and isinstance(id_map_raw["id_map"], dict) else id_map_raw
    if not isinstance(id_map, dict) or not id_map:
        raise ValueError("ui_graph_id_map_file does not contain a valid id_map.")

    # Collect subtree roots (root_uid -> defines_orig_uid)
    root_uids = []
    defines_by_root = {}
    for uid, info in nodes.items():
        if info.get("subtree_root") and info.get("defines"):
            root_uids.append(uid)
            defines_by_root[uid] = info["defines"]

    # Collect refs (ref_to_orig_uid -> [ref_uid, ...])
    refs_by_orig = {}
    for uid, info in nodes.items():
        if info.get("is_ref") and info.get("ref_to"):
            orig = info["ref_to"]
            refs_by_orig.setdefault(orig, []).append(uid)

    # Validate id_map coverage (subtree roots + related refs)
    required_uids = set(root_uids)
    for orig in defines_by_root.values():
        for ruid in refs_by_orig.get(orig, []):
            required_uids.add(ruid)
    missing = [u for u in required_uids if u not in id_map]
    if missing:
        preview = ", ".join(missing[:20])
        more = f" and {len(missing)} more" if len(missing) > 20 else ""
        raise ValueError(f"id_map is missing mappings for {len(missing)} required nodes, for example: {preview}{more}")

    # Map to numeric IDs and assemble the result
    subtree_roots_num = []
    entry_ref_whitelist = {}

    for root_uid in root_uids:
        try:
            root_num = int(id_map[root_uid])
        except Exception as e:
            raise ValueError(f"id_map[root] value cannot be converted to an integer: uid={root_uid}, value={id_map[root_uid]!r}") from e

        orig = defines_by_root[root_uid]
        ref_uids = refs_by_orig.get(orig, [])

        ref_nums_set = set()
        for ruid in ref_uids:
            try:
                ref_nums_set.add(int(id_map[ruid]))
            except Exception as e:
                raise ValueError(f"id_map[ref] value cannot be converted to an integer: uid={ruid}, value={id_map[ruid]!r}") from e

        ref_nums = sorted(ref_nums_set)
        subtree_roots_num.append(root_num)
        entry_ref_whitelist[str(root_num)] = ref_nums  # JSON object keys are strings

    subtree_roots_num.sort()

    return {
        "subtree_roots": subtree_roots_num,
        "entry_ref_whitelist": entry_ref_whitelist
    }

def export_compact_paths(graph_file, include_descriptions=True, exclude_leaf_descriptions=True, desc_length=30, description_whitelist=None):
    """
    Convert the UI navigation graph into a compact text format using underscores + numbers for control IDs.
    Output format: name_number[children], with special control types marked as name(Type)_number.
    Supported special markers: Edit, RadioButton, CheckBox.
    Optionally append description info for key control types.
    In a group of controls with the same name, if at least one is an important container type,
    then all controls in that group will carry desc.
    Supports a description whitelist mechanism.

    Args:
        graph_file: Path to the UI navigation graph JSON file
        include_descriptions: Whether to include descriptions, default True
        exclude_leaf_descriptions: Whether to exclude descriptions for leaf nodes, default True
                                  (leaf nodes do not show descriptions)
        desc_length: Maximum description length; if exceeded, it will be truncated with "...", default 30
        description_whitelist: Description whitelist. If a node name contains any keyword in the whitelist,
                               force the full description to be shown (prefer full_description; use unique_id if missing),
                               and do not truncate it. Default is None.

    Returns:
        Compact text representation as a string, using square brackets to represent hierarchy
    """
    import json

    # Read the navigation graph file
    with open(graph_file, encoding='utf-8') as f:
        graph = json.load(f)

    nodes = graph['nodes']
    edges = graph['edges']

    # Set the default whitelist
    if description_whitelist is None:
        description_whitelist = ["横向","纵向"] # PPT customization

    # Define important control types that need descriptions
    important_container_types = [
        "ComboBox", "Group", "Menu", "TabItem",
        "Pane", "MenuBar", "ToolBar","MenuItem","Button"
    ]

    # Define control types that require special markers
    special_marked_types = {
        "Edit": "Edit",
        "RadioButton": "RadioButton",
        "CheckBox": "CheckBox"
    }

    # Check nodes with duplicate names and build the set of nodes that must include descriptions
    name_to_uids = {}
    for uid, node_info in nodes.items():
        name = node_info['name']
        if name not in name_to_uids:
            name_to_uids[name] = []
        name_to_uids[name].append(uid)

    # Identify all duplicate-name nodes and determine whether descriptions should be added
    duplicate_name_nodes = set()
    duplicate_name_nodes_need_desc = set()

    for name, uid_list in name_to_uids.items():
        if len(uid_list) > 1:
            duplicate_name_nodes.update(uid_list)

            # Check whether at least one node in this duplicate-name group is an important container type
            has_important_type = any(
                nodes[uid].get('control_type', '') in important_container_types
                for uid in uid_list
            )

            # If so, all nodes in this group need descriptions
            if has_important_type:
                duplicate_name_nodes_need_desc.update(uid_list)

    # Added: check whitelist nodes
    whitelist_nodes = set()
    for uid, node_info in nodes.items():
        name = node_info['name']
        # Check whether the node name contains any whitelist keyword
        for keyword in description_whitelist:
            if keyword.lower() in name.lower():  # Case-insensitive match
                whitelist_nodes.add(uid)
                break

    # Map original unique_id to consecutive numeric IDs
    id_map = {uid: str(i) for i, uid in enumerate(nodes.keys(), start=1)}

    # Build reverse mapping for later decoding
    reverse_id_map = {str(i): uid for uid, i in id_map.items()}

    def get_control_type_marker(control_type):
        """Get the control type marker; return it for special types, otherwise return an empty string"""
        return f"({special_marked_types[control_type]})" if control_type in special_marked_types else ""

    def should_add_description(uid, control_type, full_description, has_children):
        """Determine whether a description should be added"""

        # If the node is in the whitelist, force description inclusion
        if uid in whitelist_nodes:
            return True

        # If the node is a duplicate-name node that needs a description, return True
        if uid in duplicate_name_nodes_need_desc:
            return True

        # If descriptions are disabled, return False directly
        if not include_descriptions:
            return False

        # If this is a leaf node and leaf descriptions are excluded, return False
        if exclude_leaf_descriptions and not has_children:
            return False

        # Only add descriptions for important container control types with non-empty descriptions
        if control_type in important_container_types and full_description and full_description.strip():
            return True

        return False

    def get_description_for_node(uid, full_description):
        """Get the description for a node"""
        # For duplicate-name nodes that need a description, prefer full_description;
        # if empty, fall back to unique_id
        if uid in duplicate_name_nodes_need_desc  or uid in whitelist_nodes:
            if full_description and full_description.strip():
                return full_description
            else:
                return uid
        else:
            return full_description

    def truncate_description(desc_text, uid):
        """Truncate the description text according to the configured length; whitelist nodes are not truncated"""
        # Do not truncate whitelist nodes
        if uid in whitelist_nodes:
            return desc_text

        # Other nodes follow the original truncation logic
        if len(desc_text) > desc_length:
            return desc_text[:desc_length] + "..."
        return desc_text

    # Recursively build nested text; use a visited set to avoid cyclic references
    def build(uid, visited=None):
        if visited is None:
            visited = set()

        # Detect cyclic references
        if uid in visited:
            # If a cycle is detected, return the node with a special marker
            cid = id_map[uid]
            clean_name = nodes[uid]['name'].replace(' ', '').replace(',', '').replace('，', '')
            control_type = nodes[uid].get('control_type', '')
            full_description = nodes[uid].get('full_description', '')

            # Get the control type marker
            type_marker = get_control_type_marker(control_type)
            base_name = f"{clean_name}{type_marker}_{cid}*"

            # Check whether a description should be added
            # (cycle nodes are treated as leaf nodes)
            if should_add_description(uid, control_type, full_description, False):
                desc_text = get_description_for_node(uid, full_description)
                desc = truncate_description(desc_text, uid)
                return f"{clean_name}{type_marker}(desc:{desc})_{cid}*"
            else:
                return base_name

        visited.add(uid)

        # Clean the name by removing characters that may interfere with formatting
        raw_name = nodes[uid]['name']
        # Only remove spaces and commas; keep parentheses
        clean_name = raw_name.replace(' ', '').replace(',', '').replace('，', '')

        # Get the short ID
        cid = id_map[uid]

        # Get control information
        control_type = nodes[uid].get('control_type', '')
        full_description = nodes[uid].get('full_description', '')

        # Get child nodes
        children = edges.get(uid, [])
        has_children = len(children) > 0

        # Get the control type marker
        type_marker = get_control_type_marker(control_type)

        # Build the base node name
        base_node_name = f"{clean_name}{type_marker}_{cid}"

        # Check whether a description should be added (using parentheses)
        if should_add_description(uid, control_type, full_description, has_children):
            desc_text = get_description_for_node(uid, full_description)
            desc = truncate_description(desc_text, uid)
            node_name = f"{clean_name}{type_marker}(desc:{desc})_{cid}"
        else:
            node_name = base_node_name

        # If there are no child nodes, return the leaf representation directly
        if not has_children:
            return node_name

        # Recursively process all child nodes; pass a copy of visited to keep each path independent
        child_parts = [build(c, visited.copy()) for c in sorted(children)]

        # Use square brackets to represent hierarchy
        children_str = ",".join(child_parts)
        return f"{node_name}[{children_str}]"

    # Find root nodes (nodes that do not appear in any edges values)
    all_targets = {c for lst in edges.values() for c in lst}
    roots = [u for u in nodes if u not in all_targets]

    # If no root nodes are found (possibly a cyclic graph), choose any node as a starting point
    if not roots and nodes:
        roots = [next(iter(nodes.keys()))]

    # Generate the path representation for each root node
    result = "\n".join(build(r) for r in roots)

    # Save the ID mapping for later decoding
    with open(graph_file.replace('.json', '_id_map.json'), 'w', encoding='utf-8') as f:
        json.dump({"id_map": id_map, "reverse_id_map": reverse_id_map}, f, ensure_ascii=False, indent=2)

    return result

def decode_compact_id(compact_id, id_map_file):
    """
    Decode a short ID back to the original unique ID

    Args:
        compact_id: Short numeric ID
        id_map_file: Path to the ID mapping file

    Returns:
        Original unique ID
    """

    # Read ID mappings
    with open(id_map_file, encoding='utf-8') as f:
        mappings = json.load(f)

    reverse_id_map = mappings["reverse_id_map"]

    # Look up the original unique ID
    if compact_id in reverse_id_map:
        return reverse_id_map[compact_id]
    else:
        return None

def count_tokens(text, model="gpt-3.5-turbo"):
    """Calculate the number of tokens in the text"""
    # Get the encoder for the corresponding model
    try:
        encoding = tiktoken.encoding_for_model(model)
    except KeyError:
        # If the specific model is unavailable, use the default encoder
        encoding = tiktoken.get_encoding("cl100k_base")

    # Encode the text into tokens and count them
    tokens = encoding.encode(text)
    return len(tokens)

# %%

# %%
# section 3

#%%
# Post-processing: convert LLM output into structured instructions
def convert_llm_output_to_structured_instructions(llm_output_list, ui_graph_file, ui_graph_id_map_file, filter_leaf_only=False):
    """
    Convert LLM output into a list of structured instructions containing full control information.
    Now supports shortcut instruction format and navigation path computation for forest structures.

    Args:
        llm_output_list: LLM output list, e.g. ["12_($2$)", "16_(2)", 92, 94, 11]
                        or the new shortcut-key format:
                        [{"shortcut_key": "{VK_CONTROL}c"}, {"id": 12, "text": "hello"}]
                        or forest format:
                        [{"id": 20, "entry_ref_id": [109]}, {"id": 12, "text": "hello", "entry_ref_id": [97, 105]}]
        ui_graph_file: Path to the UI navigation graph JSON file
        ui_graph_id_map_file: Path to the ID mapping file
        filter_leaf_only: bool, default False. Whether to keep only leaf nodes

    Returns:
        tuple: (structured_instructions, errors)
            structured_instructions: List of structured instructions; each element contains:
                {
                    "instruction_type": "control" | "edit" | "shortcut",
                    "compact_id": str,
                    "unique_id": str,
                    "name": str,
                    "control_type": str,
                    "navigation_path": list,
                    "text": str (Edit type only),
                    "shortcut_key": str (shortcut type only),
                    "full_description": str,
                    "is_subtree_node": bool (non-shortcut types only),
                    "is_subtree_path_valid": bool (only meaningful when non-shortcut type and is_subtree_node=True)
                }
            errors: List of error messages
    """
    errors = []

    # First parse the LLM output
    # Choose different parsing methods according to filter_leaf_only
    if filter_leaf_only:
        # Use parse_and_filter_llm_output to keep only leaf nodes
        parsed_instructions, parse_errors = parse_and_filter_llm_output(llm_output_list, ui_graph_file)
    else:
        # Use the original parse_llm_output to keep all nodes
        parsed_instructions, parse_errors = parse_llm_output(llm_output_list)

    # Collect parsing errors
    errors.extend(parse_errors)

    # Print parsing errors but continue processing
    for error in parse_errors:
        print(f"Parsing error: {error}")

    # Read UI navigation graph data
    try:
        with open(ui_graph_file, 'r', encoding='utf-8') as f:
            graph_data = json.load(f)
    except Exception as e:
        error_msg = f"Error while reading UI graph file: {str(e)}"
        errors.append(error_msg)
        print(f"Error: {error_msg}")
        return [], errors

    # Read ID mapping data
    try:
        with open(ui_graph_id_map_file, 'r', encoding='utf-8') as f:
            id_map_data = json.load(f)
    except Exception as e:
        error_msg = f"Error while reading ID mapping file: {str(e)}"
        errors.append(error_msg)
        print(f"Error: {error_msg}")
        return [], errors

    structured_instructions = []

    for i, instruction in enumerate(parsed_instructions):
        try:
            instruction_type = instruction["type"]

            # Handle shortcut instructions
            if instruction_type == "shortcut":
                shortcut_key = instruction.get("shortcut_key", "")
                structured_instruction = {
                    "instruction_index": i + 1,
                    "instruction_type": "shortcut",
                    "compact_id": f"shortcut_{i+1}",  # Generate a unique ID for the shortcut
                    "unique_id": f"shortcut_{i+1}",
                    "name": f"shortcut_key: {shortcut_key}",
                    "control_type": "Shortcut",
                    "navigation_path": [],  # No navigation path needed for shortcuts
                    "shortcut_key": shortcut_key,  # Preserve as-is
                    "full_description": f"execute shortcut_key: {shortcut_key}"
                    # Note: shortcut instructions do not include is_subtree_node or is_subtree_path_valid
                }
                structured_instructions.append(structured_instruction)
                continue

            # Handle traditional control instructions
            compact_id = instruction["id"]

            # Decode to get the original unique ID
            unique_id = decode_compact_id(compact_id, ui_graph_id_map_file)
            if not unique_id:
                error_msg = f"Instruction {i+1}: unable to decode control ID {compact_id}; the ID may not exist in the mapping file"
                errors.append(error_msg)
                print(f"Warning: {error_msg}")
                continue

            # Get control information from the navigation graph
            complete_control_info = graph_data.get('nodes', {}).get(unique_id)
            if not complete_control_info:
                error_msg = f"Instruction {i+1}: control {unique_id} not found in the navigation graph; the control may not exist in the UI tree"
                errors.append(error_msg)
                print(f"Warning: {error_msg}")
                continue

            # === Compute navigation path using the forest structure ===
            path_result = calculate_navigation_path_for_forest(graph_data, id_map_data, instruction)
            raw_navigation_path = path_result["navigation_path"]
            path_error = path_result["error"]
            is_subtree_node = path_result["is_subtree_node"]
            is_subtree_path_valid = path_result["is_subtree_path_valid"]

            # If path computation produced a warning or error, record it but continue
            if path_error:
                warning_msg = f"Instruction {i+1} path computation warning: {path_error}"
                errors.append(warning_msg)
                print(f"Warning: {warning_msg}")

            # === Clean copy suffixes from unique_id and navigation_path ===
            cleaned_unique_id = clean_copy_suffix_from_id(unique_id)
            cleaned_navigation_path = clean_copy_suffix_from_path(raw_navigation_path)

            # Build the structured instruction
            structured_instruction = {
                "instruction_index": i + 1,
                "instruction_type": instruction_type,
                "compact_id": compact_id,
                "unique_id": cleaned_unique_id,  # Use cleaned unique_id
                "name": complete_control_info.get('name', '[Unnamed]'),
                "control_type": complete_control_info.get('control_type', '[未知类型]'),
                "navigation_path": cleaned_navigation_path,  # Use computed navigation_path
                "full_description": complete_control_info.get('full_description', ''),
                "is_subtree_node": is_subtree_node,  # Added: whether it is a subtree node
                "is_subtree_path_valid": is_subtree_path_valid  # Added: whether the subtree path is valid
            }

            # If this is an Edit type, add text content
            if instruction_type == "edit":
                structured_instruction["text"] = instruction.get("text", "")

                # Validate that the control type is indeed Edit
                if complete_control_info.get('control_type') != 'Edit':
                    error_msg = f"Instruction {i+1}: control {unique_id} is of type {complete_control_info.get('control_type')}, but was marked as an Edit instruction"
                    errors.append(error_msg)
                    print(f"Warning: {error_msg}")
                    continue

            structured_instructions.append(structured_instruction)

        except Exception as e:
            error_msg = f"Error while processing instruction {i+1}: {str(e)}"
            errors.append(error_msg)
            print(f"Warning: {error_msg}")
            continue

    return structured_instructions, errors

def clean_copy_suffix_from_id(unique_id):
    """
    Remove the copy suffix from unique_id, compatible with the new _copy{n} format

    Args:
        unique_id: An ID that may contain a copy suffix

    Returns:
        str: The cleaned original ID
    """
    if not unique_id:
        return unique_id

    # Handle the new format: original_id_copy1, original_id_copy2, etc.
    if '_copy' in unique_id:
        # Find the position of the last occurrence of _copy
        copy_index = unique_id.rfind('_copy')
        if copy_index != -1:
            # Check whether _copy is followed by digits
            suffix = unique_id[copy_index + 5:]  # 5 is the length of "_copy"
            if suffix.isdigit():
                return unique_id[:copy_index]

    # Handle the old format: original_id[copy_...] (for backward compatibility)
    if '[copy_' in unique_id:
        bracket_index = unique_id.find('[copy_')
        if bracket_index != -1:
            return unique_id[:bracket_index]

    # If there is no copy suffix, return the original ID directly
    return unique_id

def clean_copy_suffix_from_path(navigation_path):
    """
    Remove the copy suffix from all elements in navigation_path,
    compatible with the new _copy{n} format

    Args:
        navigation_path: A path list that may contain copy suffixes

    Returns:
        list: The cleaned path list
    """
    if not navigation_path:
        return navigation_path

    cleaned_path = []
    for path_item in navigation_path:
        cleaned_item = clean_copy_suffix_from_id(path_item)
        cleaned_path.append(cleaned_item)

    return cleaned_path

def print_structured_instructions(structured_instructions):
    """
    Print detailed information for structured instructions

    Args:
        structured_instructions: List of structured instructions
    """
    if not structured_instructions:
        print("No valid structured instructions")
        return

    print(f"Parsed {len(structured_instructions)} structured instructions:")
    print("=" * 80)

    for inst in structured_instructions:
        print(f"Instruction {inst['instruction_index']}:")
        print(f"  Type: {inst['instruction_type'].upper()}")
        print(f"  Compact ID: {inst['compact_id']}")
        print(f"  Unique ID: {inst['unique_id']}")
        print(f"  Control Name: {inst['name']}")
        print(f"  Control Type: {inst['control_type']}")

        if inst['instruction_type'] == 'edit':
            print(f"  Input Text: '{inst['text']}'")
        elif inst['instruction_type'] == 'shortcut':
            print(f"  Shortcut Key: '{inst['shortcut_key']}'")

        print(f"  Navigation Path Length: {len(inst['navigation_path'])}")
        if inst['navigation_path']:
            print(f"  Navigation Path: {' -> '.join(inst['navigation_path'][-3:] if len(inst['navigation_path']) > 3 else inst['navigation_path'])}")

        # Added: show subtree-related information
        if inst['instruction_type'] != 'shortcut':
            print(f"  Is Subtree Node: {inst['is_subtree_node']}")
            if inst['is_subtree_node']:
                validity_status = "✓ Valid" if inst['is_subtree_path_valid'] else "✗ Incomplete"
                print(f"  Subtree Path Status: {validity_status}")

        if inst['full_description']:
            print(f"  Description: {inst['full_description'][:50]}{'...' if len(inst['full_description']) > 50 else ''}")

        print("-" * 60)

def print_structured_instructions_simple(structured_instructions):
    """
    Print simplified information for structured instructions, one line per instruction

    Args:
        structured_instructions: List of structured instructions
    """
    if not structured_instructions:
        print("No valid structured instructions")
        return

    print(f"Parsed {len(structured_instructions)} structured instructions:")

    # Build a simplified instruction list
    simple_instructions = []

    for i, inst in enumerate(structured_instructions, 1):
        simple_inst = {
            'type': inst['instruction_type'],
            'id': inst['compact_id'],
            'name': inst['name']
        }

        # If it is an edit type, add the text field
        if inst['instruction_type'] == 'edit':
            simple_inst['text'] = inst.get('text', '')

        # If it is a shortcut type, add the shortcut key field
        if inst['instruction_type'] == 'shortcut':
            simple_inst['shortcut_key'] = inst.get('shortcut_key', '')

        simple_instructions.append(simple_inst)

        # Build the status tag
        status_info = ""
        if inst['instruction_type'] != 'shortcut':
            if inst['is_subtree_node']:
                if inst['is_subtree_path_valid']:
                    status_info = " [Subtree: Complete]"
                else:
                    status_info = " [Subtree: Incomplete]"
            else:
                status_info = " [Main Tree]"

        # Print each instruction on a single line
        if inst['instruction_type'] == 'edit':
            print(f"{i:2d}. [EDIT    ] {inst['name']} (ID: {inst['compact_id']}) - Input: '{inst.get('text', '')}'{status_info}")
        elif inst['instruction_type'] == 'shortcut':
            print(f"{i:2d}. [SHORTCUT] {inst['name']} (ID: {inst['compact_id']})")
        else:
            print(f"{i:2d}. [CONTROL ] {inst['name']} (ID: {inst['compact_id']}){status_info}")

#%%
# Preprocess LLM output and normalize it into a specific JSON format
def parse_llm_output(llm_output):
    """
    Parse LLM output, supporting three formats:
    1. Legacy list format: ["12_($2$)", "16_(2)", 92, 94, 11]
    2. New-format JSON string: '[{"id":9},{"id":12,"text":"$2$"}]'
    3. Shortcut format: [{"shortcut_key": "{VK_CONTROL}c"}, {"id": 12}]

    Args:
        llm_output: LLM output, which can be a list or a JSON string

    Returns:
        tuple: (parsed_instructions, errors)
            parsed_instructions: The parsed instruction list, where each element is:
                  {"type": "control", "id": id}
                  {"type": "edit", "id": id, "text": text}
                  {"type": "shortcut", "shortcut_key": key}
            errors: A list of error messages, each represented as a string
    """
    parsed_instructions = []
    errors = []

    try:
        # Check the input type
        if isinstance(llm_output, str):
            # Try to parse it as JSON
            try:
                json_data = json.loads(llm_output.strip())
                if isinstance(json_data, list):
                    # New format: JSON array
                    return parse_json_format(json_data)
                else:
                    error_msg = f"Invalid JSON format: expected an array but got {type(json_data)}"
                    errors.append(error_msg)
                    print(f"Warning: {error_msg}")
                    return [], errors
            except json.JSONDecodeError as e:
                # If it is not valid JSON, try to parse it as a single legacy-format string item
                return parse_legacy_format([llm_output])

        elif isinstance(llm_output, list):
            # Check whether the list contains dictionaries (partially parsed new format)
            if llm_output and isinstance(llm_output[0], dict):
                # New format: already-parsed JSON array
                return parse_json_format(llm_output)
            else:
                # Legacy format: mixed-type list
                return parse_legacy_format(llm_output)

        else:
            error_msg = f"Unsupported input type {type(llm_output)}"
            errors.append(error_msg)
            print(f"Warning: {error_msg}")
            return [], errors

    except Exception as e:
        error_msg = f"An error occurred while parsing LLM output: {str(e)}"
        errors.append(error_msg)
        print(f"Warning: {error_msg}")
        return [], errors

def parse_json_format(json_data):
    """
    Parse the new JSON-format output, now supporting shortcuts and entry_ref_id

    Args:
        json_data: Parsed JSON array, e.g.
                   [{"id":9},{"shortcut_key":"{VK_CONTROL}c"},{"id":20,"entry_ref_id":[109]}]

    Returns:
        tuple: (parsed_instructions, errors)
    """
    parsed_instructions = []
    errors = []

    for i, item in enumerate(json_data):
        try:
            if isinstance(item, dict):
                # Check whether this is a shortcut instruction
                if "shortcut_key" in item:
                    shortcut_key = item.get("shortcut_key")
                    if not shortcut_key:
                        error_msg = f"The shortcut_key field in instruction {i+1} is empty"
                        errors.append(error_msg)
                        print(f"Warning: {error_msg}")
                        continue

                    parsed_instructions.append({
                        "type": "shortcut",
                        "shortcut_key": str(shortcut_key)
                    })
                    continue

                # Check the required id field (traditional control)
                if "id" not in item:
                    error_msg = f"Instruction {i+1} is missing the 'id' field: {item}"
                    errors.append(error_msg)
                    print(f"Warning: {error_msg}")
                    continue

                control_id = item["id"]

                # Validate whether the ID can be converted into a numeric string or another valid format
                try:
                    # Only accept int values or strings that can be converted to numbers
                    if isinstance(control_id, int):
                        control_id_str = str(control_id)
                    elif isinstance(control_id, str):
                        # If it is a string, check whether it is purely numeric
                        if control_id.strip().isdigit():
                            control_id_str = control_id.strip()
                        else:
                            error_msg = f"The id '{control_id}' in instruction {i+1} is not a valid numeric format"
                            errors.append(error_msg)
                            print(f"Warning: {error_msg}")
                            continue
                    else:
                        error_msg = f"Unsupported id type in instruction {i+1}: {type(control_id)}"
                        errors.append(error_msg)
                        print(f"Warning: {error_msg}")
                        continue
                except (ValueError, TypeError):
                    error_msg = f"The id '{control_id}' in instruction {i+1} cannot be converted to a valid format"
                    errors.append(error_msg)
                    print(f"Warning: {error_msg}")
                    continue

                # Added: check whether there is an entry_ref_id field (entry references in forest mode)
                entry_ref_ids = None
                if "entry_ref_id" in item:
                    entry_ref_raw = item["entry_ref_id"]

                    # Validate that entry_ref_id is a list
                    if not isinstance(entry_ref_raw, list):
                        error_msg = f"The entry_ref_id in instruction {i+1} must be a list: {entry_ref_raw}"
                        errors.append(error_msg)
                        print(f"Warning: {error_msg}")
                        continue

                    # Validate that each element in the list is a valid numeric ID
                    entry_ref_ids = []
                    validation_failed = False

                    for j, ref_id in enumerate(entry_ref_raw):
                        try:
                            if isinstance(ref_id, int):
                                entry_ref_ids.append(str(ref_id))
                            elif isinstance(ref_id, str):
                                if ref_id.strip().isdigit():
                                    entry_ref_ids.append(ref_id.strip())
                                else:
                                    error_msg = f"The entry_ref_id[{j}] '{ref_id}' in instruction {i+1} is not a valid numeric format"
                                    errors.append(error_msg)
                                    print(f"Warning: {error_msg}")
                                    validation_failed = True
                                    break
                            else:
                                error_msg = f"Unsupported type for entry_ref_id[{j}] in instruction {i+1}: {type(ref_id)}"
                                errors.append(error_msg)
                                print(f"Warning: {error_msg}")
                                validation_failed = True
                                break
                        except (ValueError, TypeError):
                            error_msg = f"The entry_ref_id[{j}] '{ref_id}' in instruction {i+1} cannot be converted to a valid format"
                            errors.append(error_msg)
                            print(f"Warning: {error_msg}")
                            validation_failed = True
                            break

                    if validation_failed:
                        continue

                # Build the base instruction structure
                base_instruction = {
                    "type": "control",  # Default type; it may be changed by the logic below
                    "id": control_id_str
                }

                # If entry_ref_id exists, add it to the instruction
                if entry_ref_ids is not None:
                    base_instruction["entry_ref_id"] = entry_ref_ids

                # Check whether there is a text field (Edit control)
                if "text" in item:
                    base_instruction["type"] = "edit"
                    base_instruction["text"] = str(item["text"])

                parsed_instructions.append(base_instruction)

            elif isinstance(item, int):
                # Compatibility: direct integer ID
                control_id_str = str(item)
                parsed_instructions.append({
                    "type": "control",
                    "id": control_id_str
                })

            else:
                error_msg = f"Instruction {i+1} contains an unsupported item type: {type(item)} - {item}"
                errors.append(error_msg)
                print(f"Warning: {error_msg}")
                continue

        except Exception as e:
            error_msg = f"Error while parsing JSON item {i+1}: {str(e)} - Item content: {item}"
            errors.append(error_msg)
            print(f"Warning: {error_msg}")
            continue

    return parsed_instructions, errors

def parse_legacy_format(llm_output_list):
    """
    Parse the legacy mixed-format output (keeping the original logic)

    Args:
        llm_output_list: Legacy-format output list, e.g. ["12_($2$)", "16_(2)", 92, 94, 11]

    Returns:
        tuple: (parsed_instructions, errors)
    """
    parsed_instructions = []
    errors = []

    for i, item in enumerate(llm_output_list):
        try:
            if isinstance(item, int):
                # Integer type, used directly as the control ID
                control_id = str(item)
                parsed_instructions.append({
                    "type": "control",
                    "id": control_id
                })
            elif isinstance(item, str):
                # Try matching the Edit control format: id_("text") or id_(text)
                edit_pattern = r'^(\d+)_\((.*?)\)$'
                match = re.match(edit_pattern, item)

                if match:
                    control_id = match.group(1)
                    text_content = match.group(2)

                    # Remove surrounding quotes if present
                    if text_content.startswith('"') and text_content.endswith('"'):
                        text_content = text_content[1:-1]
                    elif text_content.startswith("'") and text_content.endswith("'"):
                        text_content = text_content[1:-1]

                    parsed_instructions.append({
                        "type": "edit",
                        "id": control_id,
                        "text": text_content
                    })
                else:
                    # Try parsing it as a pure numeric ID
                    try:
                        control_id = str(int(item))
                        parsed_instructions.append({
                            "type": "control",
                            "id": control_id
                        })
                    except ValueError:
                        error_msg = f"Instruction {i+1} '{item}' cannot be parsed into a valid format"
                        errors.append(error_msg)
                        print(f"Warning: {error_msg}")
                        continue
            else:
                error_msg = f"Unsupported instruction type at position {i+1}: {type(item)} - '{item}'"
                errors.append(error_msg)
                print(f"Warning: {error_msg}")
                continue

        except Exception as e:
            error_msg = f"Error while parsing instruction {i+1}: {str(e)} - Instruction content: {item}"
            errors.append(error_msg)
            print(f"Warning: {error_msg}")
            continue

    return parsed_instructions, errors

def parse_and_filter_llm_output(llm_output, ui_graph_file):
    """
    Convenience function that parses LLM output and filters it to keep only leaf nodes

    Args:
        llm_output: LLM output (JSON string, list, etc.)
        ui_graph_file: Path to the UI navigation graph JSON file

    Returns:
        tuple: (parsed_instructions, errors)
    """
    errors = []

    # Parse first
    parsed_instructions, parse_errors = parse_llm_output(llm_output)

    # Collect parsing errors
    errors.extend(parse_errors)

    # Print parsing errors but continue processing
    for error in parse_errors:
        print(f"Parsing error: {error}")

    if not parsed_instructions:
        error_msg = "Failed to parse LLM output: no valid instructions found"
        errors.append(error_msg)
        print(f"⚠️ {error_msg}")
        return [], errors

    # Then filter leaf nodes (print only here to avoid duplication)
    try:
        filtered_instructions = filter_leaf_nodes_only(parsed_instructions, ui_graph_file, verbose=True)
        return filtered_instructions, errors
    except Exception as e:
        error_msg = f"Error while filtering leaf nodes: {str(e)}"
        errors.append(error_msg)
        return [], errors

def filter_leaf_nodes_only(parsed_instructions, ui_graph_file, verbose=True):
    """
    Filter the output of parse_llm_output and keep only leaf nodes
    (nodes without children)

    Apply special logic to shortcut instructions:
    if a shortcut block is preceded by a non-leaf node, remove it;
    otherwise keep it

    Args:
        parsed_instructions: Output list from parse_llm_output
        ui_graph_file: Path to the UI navigation graph JSON file
        verbose: Whether to print detailed information

    Returns:
        list: Instruction list containing only leaf nodes
    """

    try:
        # Read the UI navigation graph
        with open(ui_graph_file, 'r', encoding='utf-8') as f:
            graph_data = json.load(f)

        edges = graph_data.get('edges', {})

        # Determine which nodes are leaf nodes
        # (nodes not present in edges or whose children are empty)
        leaf_node_ids = set()
        for node_id, children in edges.items():
            if not children:  # No child nodes
                leaf_node_ids.add(node_id)

        # Identify shortcut blocks
        shortcut_blocks = []
        current_block = []

        for i, instruction in enumerate(parsed_instructions):
            if instruction["type"] == "shortcut":
                current_block.append(i)
            else:
                if current_block:
                    shortcut_blocks.append(current_block)
                    current_block = []

        # Handle the last block
        if current_block:
            shortcut_blocks.append(current_block)

        # Filter the instruction list
        filtered_instructions = []
        removed_instructions = []
        indices_to_remove = set()

        # Handle the filtering logic for shortcut blocks
        for block in shortcut_blocks:
            block_start = block[0]
            should_remove_block = False

            # Check whether a non-leaf node appears before the shortcut block
            if block_start > 0:
                prev_instruction = parsed_instructions[block_start - 1]
                if prev_instruction["type"] != "shortcut":
                    # The previous instruction is not a shortcut; check whether it is a non-leaf node
                    try:
                        control_id = prev_instruction["id"]
                        ui_graph_id_map_file = ui_graph_file.replace('.json', '_id_map.json')
                        unique_id = decode_compact_id(control_id, ui_graph_id_map_file)

                        if unique_id and unique_id not in leaf_node_ids:
                            # The previous instruction is a non-leaf node; remove this shortcut block
                            should_remove_block = True
                    except Exception:
                        pass

            if should_remove_block:
                indices_to_remove.update(block)

        # Filter all instructions
        for i, instruction in enumerate(parsed_instructions):
            if instruction["type"] == "shortcut":
                # Shortcut instructions: decide based on block-filtering logic
                if i not in indices_to_remove:
                    filtered_instructions.append(instruction)
                else:
                    removed_instructions.append(instruction)
            else:
                # Regular control instructions: filter based on leaf-node logic
                control_id = instruction["id"]

                # Decode to get the original unique ID
                try:
                    ui_graph_id_map_file = ui_graph_file.replace('.json', '_id_map.json')
                    unique_id = decode_compact_id(control_id, ui_graph_id_map_file)

                    if unique_id and unique_id in leaf_node_ids:
                        # It is a leaf node; keep it
                        filtered_instructions.append(instruction)
                    else:
                        # It is not a leaf node; remove it
                        removed_instructions.append(instruction)

                except Exception as e:
                    # If decoding fails, conservatively keep the instruction
                    filtered_instructions.append(instruction)

        # Print filtering results
        if verbose and removed_instructions:
            print(f"✂️ Filtered out {len(removed_instructions)} non-leaf nodes and shortcut blocks, kept {len(filtered_instructions)} leaf nodes")
            print("Removed instructions:")
            for instruction in removed_instructions:
                if instruction["type"] == "edit":
                    print(f"  - Edit control ID:{instruction['id']} Text:'{instruction['text']}'")
                elif instruction["type"] == "shortcut":
                    print(f"  - Shortcut: {instruction['shortcut_key']}")
                else:
                    print(f"  - {instruction['type'].title()} control ID:{instruction['id']}")
        elif verbose:
            print(f"✓ All {len(filtered_instructions)} nodes are leaf nodes")

        return filtered_instructions

    except Exception as e:
        if verbose:
            print(f"⚠️ Error while filtering leaf nodes: {str(e)}, returning the original instruction list")
        return parsed_instructions
#%%
def calculate_navigation_path_for_forest(graph_data, id_map_data, instruction):
    """
    Calculate the complete navigation path in a forest structure

    Args:
        graph_data: Graph data (forest), containing 'nodes' and 'edges'
        id_map_data: ID mapping data, containing 'reverse_id_map'
        instruction: Parsed instruction, containing id and optional entry_ref_id

    Returns:
        dict: {
            "navigation_path": list,  # Complete navigation path
            "error": str,  # Error message; empty string means no error
            "is_subtree_node": bool,  # Whether it is a subtree node
            "is_subtree_path_valid": bool  # Whether the subtree path is valid
        }
    """
    try:
        # 1. Decode and get the target node's unique_id
        compact_id = instruction["id"]
        reverse_id_map = id_map_data.get("reverse_id_map", {})

        if compact_id not in reverse_id_map:
            return {
                "navigation_path": [],
                "error": f"Unable to decode control ID {compact_id}: it does not exist in the ID mapping",
                "is_subtree_node": False,  # Cannot determine if decoding fails
                "is_subtree_path_valid": False
            }

        target_unique_id = reverse_id_map[compact_id]
        nodes = graph_data.get('nodes', {})

        if target_unique_id not in nodes:
            return {
                "navigation_path": [],
                "error": f"Target node {target_unique_id} does not exist in the graph",
                "is_subtree_node": False,  # Cannot determine if the node does not exist
                "is_subtree_path_valid": False
            }

        target_node = nodes[target_unique_id]

        # 2. Determine whether the node is on the main tree
        tree_root_id = target_node.get('tree_root_id')
        is_subtree_node = False

        # Check whether the node pointed to by tree_root_id is a subtree root
        if tree_root_id and tree_root_id in nodes:
            root_node = nodes[tree_root_id]
            if root_node.get('subtree_root', False):
                is_subtree_node = True

        # 3. If it is on the main tree, return the original navigation_path directly
        if not is_subtree_node:
            original_path = target_node.get('navigation_path', [])
            return {
                "navigation_path": original_path,
                "error": "",
                "is_subtree_node": False,
                "is_subtree_path_valid": True  # Main-tree paths are always valid
            }

        # 4. If it is in a subtree, use entry_ref_id to build the complete path
        entry_ref_ids = instruction.get("entry_ref_id", [])

        if not entry_ref_ids:
            # No entry_ref_id provided; can only return the path within the subtree
            subtree_path = target_node.get('navigation_path', [])
            return {
                "navigation_path": subtree_path,
                "error": f"The target node is in a subtree but no entry_ref_id was provided; only the subtree-internal path is returned",
                "is_subtree_node": True,
                "is_subtree_path_valid": False  # Missing entry_ref_id means the path is incomplete
            }

        # 5. Prefer the full path based on entry_ref_id; if it fails, gradually skip prefixes
        return build_path_with_entry_refs(
            graph_data, reverse_id_map, entry_ref_ids, target_node, target_unique_id
        )

    except Exception as e:
        return {
            "navigation_path": [],
            "error": f"An exception occurred while calculating the navigation path: {str(e)}",
            "is_subtree_node": False,  # Cannot determine in exceptional cases
            "is_subtree_path_valid": False
        }

def build_path_with_entry_refs(graph_data, reverse_id_map, entry_ref_ids, target_node, target_unique_id):
    """
    Build the complete navigation path using the entry_ref_id list.
    Prefer the full chain first; if it fails, gradually skip prefixes.

    Strictly validate chained context consistency:
    ref[i] → subtree_root[i] → ref[i+1]

    Args:
        graph_data: Graph data
        reverse_id_map: Reverse ID mapping
        entry_ref_ids: entry_ref_id list (left to right: main-tree ref → subtree ref → ...)
        target_node: Target node information
        target_unique_id: The target node's unique_id

    Returns:
        dict: {"navigation_path": list, "error": str, "is_subtree_node": bool, "is_subtree_path_valid": bool}
    """
    nodes = graph_data.get('nodes', {})

    # Initialize final_error to avoid undefined-variable errors
    final_error = "unknown error"

    # Try the full chain first, then gradually skip prefixes (correct order)
    for start_index in range(len(entry_ref_ids)):
        try:
            # Try building the path starting from start_index
            path_segments = []
            error_msg = ""

            # Validate and build the path layer by layer
            for i in range(start_index, len(entry_ref_ids)):
                ref_compact_id = entry_ref_ids[i]

                # Decode the ref ID
                if ref_compact_id not in reverse_id_map:
                    error_msg = f"entry_ref_id[{i}] '{ref_compact_id}' cannot be decoded"
                    break

                ref_unique_id = reverse_id_map[ref_compact_id]

                if ref_unique_id not in nodes:
                    error_msg = f"The node corresponding to entry_ref_id[{i}] '{ref_unique_id}' does not exist in the graph"
                    break

                ref_node = nodes[ref_unique_id]

                # Verify that this is a reference stub
                if not ref_node.get('is_ref', False):
                    error_msg = f"The node corresponding to entry_ref_id[{i}] '{ref_unique_id}' is not a reference stub (is_ref=False)"
                    break

                # Validate whether the first-layer ref is in the main forest
                # (only checked for the full path; skipped in fault-tolerant mode)
                if i == start_index and start_index == 0:
                    ref_tree_root_id = ref_node.get('tree_root_id')
                    if ref_tree_root_id and ref_tree_root_id in nodes:
                        ref_tree_root = nodes[ref_tree_root_id]
                        if ref_tree_root.get('subtree_root', False):
                            error_msg = f"The first entry_ref_id[{i}] must be in the main forest, but it is in a subtree (tree_root: {ref_tree_root_id})"
                            break

                # Get the original node ID pointed to by the ref
                ref_to = ref_node.get('ref_to')
                if not ref_to:
                    error_msg = f"The reference stub corresponding to entry_ref_id[{i}] is missing the ref_to field"
                    break

                # Find the subtree root corresponding to ref_to
                # (the subtree root whose defines == ref_to)
                target_subtree_root = None
                for node_id, node_info in nodes.items():
                    if (node_info.get('subtree_root', False) and
                        node_info.get('defines') == ref_to):
                        target_subtree_root = node_id
                        break

                if not target_subtree_root:
                    error_msg = f"entry_ref_id[{i}] points to '{ref_to}', but no matching subtree root was found (defines='{ref_to}')"
                    break

                # Validate whether the next ref is inside the current subtree (if there is a next level)
                if i + 1 < len(entry_ref_ids):
                    next_ref_compact_id = entry_ref_ids[i + 1]
                    if next_ref_compact_id not in reverse_id_map:
                        error_msg = f"entry_ref_id[{i+1}] '{next_ref_compact_id}' cannot be decoded"
                        break

                    next_ref_unique_id = reverse_id_map[next_ref_compact_id]
                    if next_ref_unique_id not in nodes:
                        error_msg = f"The node corresponding to entry_ref_id[{i+1}] '{next_ref_unique_id}' does not exist in the graph"
                        break

                    next_ref_node = nodes[next_ref_unique_id]
                    next_ref_tree_root = next_ref_node.get('tree_root_id')

                    # The next ref must be in the subtree pointed to by the current ref
                    if next_ref_tree_root != target_subtree_root:
                        error_msg = (f"Chain validation failed: entry_ref_id[{i}] points to subtree root '{target_subtree_root}', "
                                   f"but entry_ref_id[{i+1}] is in a different tree (tree_root: {next_ref_tree_root})")
                        break

                # Add the current ref's navigation path to the result
                ref_nav_path = ref_node.get('navigation_path', [])
                path_segments.extend(ref_nav_path)

            # If all refs pass validation, validate the target node last
            if not error_msg:
                # Find the subtree root pointed to by the last ref
                last_ref_compact_id = entry_ref_ids[-1]
                last_ref_unique_id = reverse_id_map[last_ref_compact_id]
                last_ref_node = nodes[last_ref_unique_id]
                last_ref_to = last_ref_node.get('ref_to')

                # Validate that the target node is indeed in the subtree pointed to by the last ref
                target_tree_root = target_node.get('tree_root_id')
                if target_tree_root and target_tree_root in nodes:
                    target_tree_root_node = nodes[target_tree_root]
                    target_defines = target_tree_root_node.get('defines')

                    if last_ref_to != target_defines:
                        error_msg = (f"Final validation failed: the last entry_ref points to '{last_ref_to}', "
                                   f"but the target node is in the subtree with defines='{target_defines}', which does not match")
                        final_error = error_msg
                    else:
                        # Success! Add the target node's path inside the subtree
                        target_subtree_path = target_node.get('navigation_path', [])
                        path_segments.extend(target_subtree_path)

                        # Determine whether this is a complete path
                        is_complete_path = (start_index == 0)

                        if is_complete_path:
                            # Full path succeeded, no error
                            return {
                                "navigation_path": path_segments,
                                "error": "",  # Empty error when the full path succeeds
                                "is_subtree_node": True,
                                "is_subtree_path_valid": True
                            }
                        else:
                            # Partial path succeeded; record the skipped prefix and failure reason
                            skipped_refs = entry_ref_ids[:start_index]
                            partial_error = f"Skipped the first {start_index} entry_ref_id values: {skipped_refs}. Reason: {final_error}"
                            return {
                                "navigation_path": path_segments,
                                "error": partial_error,
                                "is_subtree_node": True,
                                "is_subtree_path_valid": False  # Partial path is not considered valid
                            }
                else:
                    error_msg = f"The target node is missing a valid tree_root_id or a corresponding subtree root node"
                    final_error = error_msg

            # If this round failed, record the error and try the next round
            # (skipping more prefixes)
            if error_msg:
                final_error = error_msg

        except Exception as e:
            # An exception occurred in this round; continue trying the next round
            final_error = f"An exception occurred while processing entry_ref_ids: {str(e)}"
            continue

    # All attempts failed; return only the subtree-internal path and the error message
    subtree_path = target_node.get('navigation_path', [])
    return {
        "navigation_path": subtree_path,
        "error": f"Unable to build the full path via entry_ref_id: {final_error}. Returning only the subtree-internal path",
        "is_subtree_node": True,
        "is_subtree_path_valid": False  # Failed to build a complete path
    }

#%%
# Get the app's true topmost operable window
def get_top_operatable_window(target_pid):
    """
    Get the topmost operable window (top_operatable_window) in the specified process

    Logic flow:
    1. Get the top HWND window in Z-order
       (could be an app-level window or a modal child window)
    2. Get the app-level window
       (to ensure a correct application-level window reference)
    3. Detect whether there is a child window inside the app-level window
    4. If a child window exists, return it as top_operatable_window;
       otherwise, return the top HWND window as top_operatable_window

    Args:
        target_pid: Target process ID

    Returns:
        UIAWrapper: The top_operatable_window object, or None if retrieval fails
    """
    try:
        # print(f"🎯 Starting to get the top_operatable_window for process {target_pid}...")
        # print("=" * 60)

        # Step 1: Get the top HWND window in Z-order using get_top_window_by_zorder
        # print(f"Step 1: Getting the top HWND window in Z-order...")
        hwnd_top_window = get_top_window_by_zorder(target_pid)

        if not hwnd_top_window:
            print("❌ Failed to get the top HWND window; failed to retrieve top_operatable_window")
            return None

        try:
            hwnd_window_title = hwnd_top_window.window_text() or "[Untitled]"
            # print(f"  Top HWND window: {hwnd_window_title}")
        except Exception:
            print("  Top HWND window: [Unable to get title]")

        # Step 2: Get the app-level window using get_top_level_window_from_current
        # print("Step 2: Getting the app-level window...")
        app_level_window = get_top_level_window_from_current(hwnd_top_window)

        if not app_level_window:
            print("❌ Failed to get the app-level window; failed to retrieve top_operatable_window")
            return None
        # print("----app level-debug: are same window ",are_same_controls(app_level_window,main_window))
        # print("----app level-debug: main_window_windowchild are: ",main_window.descendants(control_type="Window", depth=None))
        # print("----app level-debug: app_level_window_windowchild are: ",app_level_window.descendants(control_type="Window", depth=None))

        try:
            app_window_title = app_level_window.window_text() or "[Untitled]"
            # print(f"  App-level window: {app_window_title}")
        except Exception:
            print("  App-level window: [Unable to get title]")

        # Step 3: Use detect_single_child_window_in_app to determine whether a child exists
        # print("Step 3: Detecting child windows inside the app-level window...")
        child_windows = detect_single_child_window_in_app(app_level_window)

        # Step 4: Decide top_operatable_window based on whether a child window exists
        # print("Step 4: Determining top_operatable_window...")
        if child_windows and len(child_windows) > 0:
            # A child window exists; return it as top_operatable_window
            child_window = child_windows[0]  # detect_single_child_window_in_app returns a z-order-sorted list

            try:
                child_window_title = child_window.window_text() or "[Untitled]"
                print(f"✅ Child window detected; returning it as top_operatable_window")
                print(f"  top_operatable_window: {child_window_title}")
            except Exception:
                print("✅ Child window detected; returning it as top_operatable_window")
                print("  top_operatable_window: [Unable to get title]")

            return child_window
        else:
            # No child window exists; return the top HWND window as top_operatable_window
            try:
                hwnd_window_title = hwnd_top_window.window_text() or "[Untitled]"
                print("✅ No child window detected; returning the top HWND window as top_operatable_window")
                print(f"  top_operatable_window: {hwnd_window_title}")
            except Exception:
                print("✅ No child window detected; returning the top HWND window as top_operatable_window")
                print("  top_operatable_window: [Unable to get title]")

            return hwnd_top_window

    except Exception as e:
        print(f"❌ Error while executing get_top_operatable_window: {str(e)}")
        return None

def close_top_operatable_window_by_instance(top_operatable_window):
    """
    Close the specified top_operatable_window instance

    Logic flow:
    1. Check whether the window still exists
    2. Get the app-level window
    3. Determine whether it is the same window and choose the corresponding closing strategy
    4. Verify the closing result

    Args:
        top_operatable_window: top_operatable_window instance

    Returns:
        bool: Whether the close operation succeeded
    """
    try:
        if not top_operatable_window:
            print("❌ top_operatable_window is None, canceling close operation")
            return False

        # Get the window title for logging
        try:
            window_title = top_operatable_window.window_text() or "[Untitled]"
        except Exception:
            window_title = "[Title unavailable]"

        print(f"🚀 Starting to close top_operatable_window: {window_title}")

        # a. First check whether the window still exists
        if not check_window_controls_still_exist([top_operatable_window]):
            print("✅ Window no longer exists, no need to close")
            return True

        # b. Get the app-level window
        app_level_window = get_top_level_window_from_current(top_operatable_window)
        if not app_level_window:
            print("❌ Failed to get app-level window, close operation failed")
            return False

        # c. Check whether it is the same window
        is_same_window = are_same_controls(top_operatable_window, app_level_window)

        # Perform the close operation
        close_success = False
        if is_same_window:
            # e. It is the app-level window, use confirm_and_close_window
            print("🎯 Detected app-level window, using confirmation-based close")
            close_success = confirm_and_close_window(app_level_window)
        else:
            # d. It is a child window, use confirm_and_close_single_child_window_with_priority
            print("🎯 Detected child window, using dialog close strategy")
            close_success = confirm_and_close_single_child_window_with_priority(app_level_window)

        if not close_success:
            print("❌ Window close operation failed")
            return False

        # f. Verify the close result
        if check_window_controls_still_exist([top_operatable_window]):
            print("❌ Window close verification failed, the window still exists")
            return False
        else:
            print("✅ top_operatable_window closed successfully")
            return True

    except Exception as e:
        print(f"❌ Error occurred while closing top_operatable_window: {str(e)}")
        return False

def reverse_search_in_window(structured_instruction, top_operatable_window):
    """
    Reverse-search the controls in structured_instruction within the given window

    Args:
        structured_instruction: Structured instruction containing unique_id and navigation_path
        top_operatable_window: The window to search in

    Returns:
        tuple: (found: bool, deepest_index: int or None, first_control: UIAWrapper or None, error_message: str)
               - found: Whether the control was found
               - deepest_index: The index of the deepest available control, or None if not found
               - first_control: The first found control instance, used for performance optimization
               - error_message: Error message
    """
    try:
        unique_id = structured_instruction["unique_id"]
        navigation_path = structured_instruction["navigation_path"]
        control_name = structured_instruction.get("name", "[Unknown]")

        print(f"🔍 Reverse-searching for control in window: {control_name}")

        # Build the control cache
        try:
            controls_info, control_instances = get_control_descendants_with_info(top_operatable_window)
        except Exception as e:
            return False, None, None, f"Failed to build control cache: {str(e)}"

        # First, directly search for the target control
        if unique_id in control_instances:
            # Return the actual index position of the target control in navigation_path
            target_index = len(navigation_path) - 1  # The target control is usually the last item in the path
            target_control = control_instances[unique_id]
            print(f"✅ Found target control directly: {unique_id} ({control_name}) (target_index: {target_index})")
            return True, target_index, target_control, ""

        # Reverse-scan the navigation path
        optimized_path, deepest_index = reverse_scan_path(navigation_path, control_instances)

        if deepest_index is not None:
            # Get the found control ID and instance
            found_control_id = navigation_path[deepest_index]
            first_control = control_instances[found_control_id]
            print(f"✅ Reverse scan found an available control: {found_control_id} (deepest_index: {deepest_index})")
            return True, deepest_index, first_control, ""
        else:
            error_msg = f"None of the controls in the path exist in the current window"
            print(f"❌ {error_msg}")
            return False, None, None, error_msg

    except Exception as e:
        error_msg = f"Error occurred during reverse search: {str(e)}"
        print(f"❌ {error_msg}")
        return False, None, None, error_msg

# Note: this function will close top-level windows one by one until the condition is met
def transition_suitable_operatable_window(target_pid, structured_instruction, max_attempts=6):
    """
    Repeatedly get the operatable_window and determine whether navigation
    should start from the current window.
    If not, close the current window and get the next one until a suitable
    window is found.

    Args:
        target_pid: Target process ID
        structured_instruction: Structured instruction used to determine whether navigation can start from the current window
        max_attempts: Maximum number of attempts to prevent infinite loops

    Returns:
        tuple: (success: bool, top_window: UIAWrapper or None, app_level_window: UIAWrapper or None,
                deepest_index: int or None, first_control: UIAWrapper or None, error_message: str)
               - success: Whether a suitable window was successfully found
               - top_window: The found top_operatable_window, or None on failure
               - app_level_window: The corresponding app-level window, or None on failure
               - deepest_index: The index of the deepest available control, indicating where to start navigating in navigation_path; None on failure
               - first_control: The first found control instance (performance optimization), or None on failure
               - error_message: Error message
    """
    try:
        control_name = structured_instruction.get("name", "[Unknown]")
        print(f"🎯 Starting to find an operatable_window suitable for navigating to control '{control_name}'")

        for attempt in range(max_attempts):
            print(f"📍 Attempt {attempt + 1} to get operatable_window...")

            # Get the current top_operatable_window
            current_top_operatable_window = get_top_operatable_window(target_pid)

            if not current_top_operatable_window:
                error_msg = f"Attempt {attempt + 1}: Failed to get operatable_window"
                print(f"❌ {error_msg}")
                return False, None, None, None, None, error_msg

            try:
                window_title = current_top_operatable_window.window_text() or "[Untitled]"
                print(f"  Current top window: {window_title}")
            except Exception:
                print(f"  Current top window: [Title unavailable]")

            # Reverse-search the target control in top_operatable_window
            found, deepest_index, first_control, search_error = reverse_search_in_window(
                structured_instruction, current_top_operatable_window
            )

            if found:
                # A suitable window was found; navigation can start from this window
                found_control_id = structured_instruction["navigation_path"][deepest_index]
                print(f"✅ Found a suitable window combination - available control: {found_control_id} (deepest_index: {deepest_index})")

                # Get the corresponding app-level window
                app_level_window = get_top_level_window_from_current(current_top_operatable_window)
                if not app_level_window:
                    error_msg = f"Attempt {attempt + 1}: Failed to get app-level window"
                    print(f"❌ {error_msg}")
                    return False, None, None, None, None, error_msg

                return True, current_top_operatable_window, app_level_window, deepest_index, first_control, ""

            # Not found; need to close the current top window
            print(f"❌ Current window is not suitable for navigation: {search_error}")
            print(f"🚪 Attempting to close the current operatable_window...")

            # Record the window state before closing for verification
            window_before_close = current_top_operatable_window

            # Close the current window
            close_success = close_top_operatable_window_by_instance(current_top_operatable_window)

            if not close_success:
                error_msg = f"Attempt {attempt + 1}: Failed to close window"
                print(f"❌ {error_msg}")
                return False, None, None, None, None, error_msg

            # Verify whether the window was actually closed
            if is_window_valid(window_before_close):
                error_msg = f"Attempt {attempt + 1}: Window close verification failed; the window still exists"
                print(f"❌ {error_msg}")
                return False, None, None, None, None, error_msg

            print(f"✅ Window closed successfully, preparing to get the next operatable_window")

            # Briefly wait for the UI to stabilize
            import time
            time.sleep(0.3)

        # Reached the maximum number of attempts without finding a suitable window
        error_msg = f"Still failed to find a suitable operatable_window after {max_attempts} attempts"
        print(f"❌ {error_msg}")
        return False, None, None, None, None, error_msg

    except Exception as e:
        error_msg = f"Error occurred while finding a suitable window: {str(e)}"
        print(f"❌ {error_msg}")
        return False, None, None, None, None, error_msg


    except Exception as e:
        error_msg = f"Error occurred while finding a suitable window: {str(e)}"
        print(f"❌ {error_msg}")
        return False, None, None, None, error_msg
#%%
# visit  Allow retry if the first UI modeling attempt fails
def execute_llm_instructions(llm_output, current_window, main_window, ui_graph_file, ui_graph_id_map_file, filter_leaf_only=False, max_retries=3):
    """
    Execute the instructions output by the LLM, handling all window scenarios and UI modeling issues.
    Execution of any subsequent instructions stops as soon as one instruction fails.

    Args:
        llm_output: LLM output list, e.g. ["12_($2$)", "16_(2)", 92, 94, 11]
        current_window: Current operating window (may be the main window or a top-level child window; modal windows are not allowed)
        main_window: Main application window
        ui_graph_file: Path to the UI navigation graph JSON file
        ui_graph_id_map_file: Path to the ID mapping file
        max_retries: Maximum retry count

    Returns:
        dict: Execution result
    """
    print("=" * 80)
    print("Starting execution of LLM instructions")
    print("=" * 80)

    # Parse the LLM output into structured instructions
    structured_instructions, conversion_errors = convert_llm_output_to_structured_instructions(
        llm_output, ui_graph_file, ui_graph_id_map_file, filter_leaf_only
    )

    # If there are errors during parsing, return immediately
    if conversion_errors:
        error_message = "Instruction parsing failed with the following errors:\n" + "\n".join(f"- {error}" for error in conversion_errors) + "\n\nNo instructions were executed."
        print("❌ " + error_message)
        return {
            "success": False,
            "error": error_message
        }

    if not structured_instructions:
        error_message = "There are no valid instructions to execute, possibly because non-leaf nodes were output instead of leaf nodes. No instructions were executed."
        return {
            "success": False,
            "error": error_message
        }

    # Newly added: replace placeholders with actual values
    structured_instructions = resolve_template_placeholders(
        structured_instructions, main_window
    )

    print_structured_instructions(structured_instructions)

    # Execution result statistics
    execution_results = {
        "total_instructions": len(structured_instructions),
        "successful_instructions": 0,
        "failed_instructions": 0,
        "execution_details": [],
        "final_window": current_window,
        "ui_modeling_issue_detected": False,
        "stopped_due_to_failure": False  # Indicates whether execution stopped because of a failure
    }

    working_window = current_window

    # Execute instructions one by one
    for i, instruction in enumerate(structured_instructions):
        print(f"\n{'='*60}")
        print(f"Executing instruction {i+1}/{len(structured_instructions)}: {instruction['name']}")
        print(f"{'='*60}")

        success = False
        retries = 0
        error_message = ""
        ui_modeling_issue_retry_attempted = False

        # Check whether this is a shortcut instruction
        is_shortcut_instruction = instruction.get('instruction_type') == 'shortcut'

        # Shortcut instructions have a max retry count of 0 (no retries)
        current_max_retries = 0 if is_shortcut_instruction else max_retries

        if is_shortcut_instruction:
            print("📝 Shortcut instructions are not allowed to retry")

        # Retry loop
        while not success and retries <= current_max_retries:
            if retries > 0:
                print(f"Retry attempt {retries + 1}...")

            # Execute a single instruction
            success, working_window, error_message = execute_single_instruction(
                instruction, main_window)

            # Handle execution failure
            if not success:
                # Shortcut instruction failure handling (checked first: no retry for any reason)
                if is_shortcut_instruction:
                    print(f"Shortcut instruction execution failed; no retry will be performed: {error_message}")
                    break

                # Check whether this is a disabled-control error (new)
                elif error_message.startswith("CONTROL_DISABLED:"):
                    print("🚫 Detected that the target control is disabled and cannot be operated")
                    # Disabled-control errors do not retry; exit directly
                    break

                # Check whether this is a UI modeling issue
                elif error_message.startswith("UI_MODELING_ISSUE:"):
                    print("🚫 Detected a UI navigation tree modeling issue")
                    execution_results["ui_modeling_issue_detected"] = True

                    # If a UI modeling retry has not been attempted yet, do it once
                    if not ui_modeling_issue_retry_attempted:
                        ui_modeling_issue_retry_attempted = True
                        print("🔄 Performing a dedicated retry for the UI modeling issue (1 attempt)...")
                        retries += 1
                        time.sleep(2)  # Wait 2 seconds before retrying a UI modeling issue
                        continue
                    else:
                        # Already retried once; confirm it is a UI modeling issue
                        print("🚫 UI modeling issue retry failed")
                        break

                # Normal instruction retry logic
                else:
                    retries += 1
                    if retries <= current_max_retries:
                        print(f"Instruction execution failed, preparing to retry: {error_message}")
                        time.sleep(1)  # Wait 1 second before retrying

        # ============== Retry loop ended; handle final execution result ==============

        final_success = success  # Save the final execution result to avoid confusion with the loop-local success

        if final_success:
            # Instruction executed successfully
            execution_results["successful_instructions"] += 1

            instruction_result = {
                "instruction_index": i + 1,
                "instruction_name": instruction['name'],
                "instruction_type": instruction['instruction_type'],
                "success": True,
                "retries": retries,
                "error": None,
                "is_ui_modeling_issue": False
            }

            # If a UI modeling retry was attempted and eventually succeeded, add a marker
            if ui_modeling_issue_retry_attempted:
                instruction_result["ui_modeling_retry_successful"] = True
                print(f"✓ Instruction {i+1} executed successfully (after UI modeling retry)")
            elif is_shortcut_instruction:
                instruction_result["is_shortcut_no_retry"] = True
                print(f"✓ Instruction {i+1} executed successfully (shortcut instruction)")
            else:
                print(f"✓ Instruction {i+1} executed successfully")

            execution_results["execution_details"].append(instruction_result)

        else:
            # Instruction failed; stop executing all subsequent instructions
            execution_results["failed_instructions"] += 1
            execution_results["stopped_due_to_failure"] = True

            print(f"🚫 Instruction {i+1} failed, stopping execution of all remaining instructions")

            # Record the currently failed instruction
            current_instruction_result = {
                "instruction_index": i + 1,
                "instruction_name": instruction['name'],
                "instruction_type": instruction['instruction_type'],
                "success": False,
                "retries": retries,
                "error": error_message,
                "is_ui_modeling_issue": error_message.startswith("UI_MODELING_ISSUE:") if error_message else False,
                "is_control_disabled": error_message.startswith("CONTROL_DISABLED:") if error_message else False  # New
            }

            # Add special markers
            if ui_modeling_issue_retry_attempted:
                current_instruction_result["ui_modeling_retry_attempted"] = True
            if is_shortcut_instruction:
                current_instruction_result["is_shortcut_no_retry"] = True

            execution_results["execution_details"].append(current_instruction_result)

            # Record all subsequent instructions as skipped
            for j in range(i + 1, len(structured_instructions)):
                skipped_instruction = {
                    "instruction_index": j + 1,
                    "instruction_name": structured_instructions[j]['name'],
                    "instruction_type": structured_instructions[j]['instruction_type'],
                    "success": False,
                    "retries": 0,
                    "error": f"Skipped because a previous instruction ({i+1}) failed",
                    "is_ui_modeling_issue": False,
                    "skipped": True,
                    "skipped_reason": "previous_instruction_failed"
                }
                execution_results["execution_details"].append(skipped_instruction)
                execution_results["failed_instructions"] += 1

            # Break out of the main loop; do not execute subsequent instructions
            break

    execution_results["final_window"] = working_window

    # Print execution summary
    print("\n" + "=" * 80)
    print("Instruction Execution Summary")
    print("=" * 80)
    print(f"Total instructions: {execution_results['total_instructions']}")
    print(f"Executed successfully: {execution_results['successful_instructions']}")
    print(f"Failed: {execution_results['failed_instructions']}")

    if execution_results['total_instructions'] > 0:
        success_rate = execution_results['successful_instructions']/execution_results['total_instructions']*100
        print(f"Success rate: {success_rate:.1f}%")
    else:
        print("Success rate: 0.0%")

    # Special notices
    if execution_results["stopped_due_to_failure"]:
        print("⚠️  Execution of the remaining instructions was stopped early due to a failure")

    if execution_results["ui_modeling_issue_detected"]:
        print("⚠️  A UI navigation tree modeling issue was detected; please check the accuracy of the UI navigation graph")

    # Detailed execution results
    print("\nDetailed execution results:")
    for detail in execution_results["execution_details"]:
        status = "✓" if detail["success"] else "✗"
        instruction_info = f"{status} Instruction {detail['instruction_index']}: {detail['instruction_name']} ({detail['instruction_type']})"

        if detail["success"]:
            print(f"  {instruction_info}")
            if detail.get("ui_modeling_retry_successful"):
                print(f"    Note: Succeeded after a UI modeling retry")
            if detail.get("is_shortcut_no_retry"):
                print(f"    Note: Shortcut instruction (no retry mechanism)")
            if detail["retries"] > 0:
                print(f"    Retry count: {detail['retries']}")
        else:
            print(f"  {instruction_info}")
            if detail.get("skipped"):
                if detail.get("skipped_reason") == "previous_instruction_failed":
                    print(f"    Status: Skipped (a previous instruction failed)")
                else:
                    print(f"    Status: Skipped")
            else:
                # This is an actual failed instruction
                if detail.get("is_control_disabled"):
                    print(f"    Error type: Control disabled")
                elif detail.get("is_ui_modeling_issue"):
                    print(f"    Error type: UI modeling issue")
                    if detail.get("ui_modeling_retry_attempted"):
                        print(f"    Note: A UI modeling retry was attempted but still failed")
                    elif detail.get("is_shortcut_no_retry"):
                        print(f"    Note: Shortcut instructions are not allowed to retry")
                elif detail.get("is_shortcut_no_retry"):
                    print(f"    Note: Shortcut instruction (no retry mechanism)")
                else:
                    print(f"    Retry count: {detail['retries']}")

            if detail["error"]:
                # Truncate the error message to the first 100 characters to avoid excessive length
                error_preview = detail["error"][:100] + "..." if len(detail["error"]) > 100 else detail["error"]
                print(f"    Error: {error_preview}")

    # Overall success is true only if all instructions succeed
    overall_success = execution_results['failed_instructions'] == 0
    execution_results["success"] = overall_success

    # Remove non-serializable fields before returning
    execution_results.pop("final_window", None)

    try:
        json.dumps(execution_results, ensure_ascii=False)
        print("✅ The return result can be JSON-serialized")
    except (TypeError, ValueError) as e:
        print(f"❌ JSON serialization failed: {e}")
        # Perform further cleanup if needed

    return execution_results
#%%

#%%
def execute_single_instruction(structured_instruction, main_window):
    """
    Execute a single instruction - using the new window management and navigation capabilities
    Now supports shortcut instructions

    Args:
        structured_instruction: Structured instruction
        main_window: Main window (used to get the process ID and as a fallback)

    Returns:
        tuple: (success: bool, updated_window: UIAWrapper, error_message: str)
    """
    try:
        instruction_type = structured_instruction.get("instruction_type", "")
        control_name = structured_instruction.get("name", "[Unknown]")
        print(f"🎯 Executing instruction: {control_name} (type: {instruction_type})")

        # Newly added: handle shortcut instructions
        if instruction_type == "shortcut":
            return execute_shortcut_instruction(structured_instruction, main_window)

        # Original control operation logic
        # Get the process ID from the main window
        try:
            target_pid = main_window.process_id()
        except Exception as e:
            error_msg = f"Failed to get process ID: {str(e)}"
            print(f"❌ {error_msg}")
            return False, main_window, error_msg

        # Step 1: Get a suitable operating window and navigation info
        success, top_window, app_window, deepest_index, first_control, error = transition_suitable_operatable_window(
            target_pid, structured_instruction
        )

        if not success:
            print(f"❌ Failed to get a suitable operating window: {error}")
            return False, main_window, error

        # Step 2: Use the optimized path to perform navigation and operation (passing the cached control instance)
        nav_success, updated_window, nav_error, reachable = navigate_and_execute_with_optimized_path(
            structured_instruction, app_window, deepest_index, first_control
        )

        if nav_success:
            print(f"✅ Instruction executed successfully: {control_name}")
            return True, updated_window, ""
        else:
            print(f"❌ Instruction execution failed: {nav_error}")

            # If it is a UI modeling issue, return the error directly
            if nav_error.startswith("UI_MODELING_ISSUE:"):
                return False, updated_window, nav_error
            else:
                return False, main_window, nav_error

    except Exception as e:
        error_msg = f"Exception occurred while executing instruction: {str(e)}"
        print(f"❌ {error_msg}")
        return False, main_window, error_msg

#%%
def resolve_template_placeholders(structured_instructions, main_window):
    """
    Replace placeholders in structured instructions with actual values
    """
    try:
        # Get actual values
        current_values = extract_current_environment_values(main_window)

        if not current_values:
            print("⚠️ Failed to get current environment values; skipping placeholder replacement")
            return structured_instructions

        print(f"🔄 Starting placeholder replacement, current environment values: {current_values}")

        # Convert the entire instruction list to a JSON string, perform global replacement, then convert it back
        import json

        # Convert to JSON string
        instructions_json = json.dumps(structured_instructions, ensure_ascii=False)

        # Replace placeholders globally
        for placeholder, actual_value in current_values.items():
            instructions_json = instructions_json.replace(placeholder, actual_value)

        # Convert back to a Python object
        resolved_instructions = json.loads(instructions_json)

        print("✅ Placeholder replacement completed")
        return resolved_instructions

    except Exception as e:
        print(f"⚠️ Error occurred during placeholder replacement: {str(e)}; using original instructions")
        return structured_instructions

def extract_current_environment_values(main_window):
    """
    Extract actual values from the current environment based on the main window
    """
    current_values = {}

    try:
        # Get the application's own name
        app_name = main_window.element_info.name or ""
        if app_name.strip():
            current_values["{APP_NAME}"] = app_name.strip()

        # Get the parent's name
        parent = main_window.parent()
        if parent:
            parent_name = parent.element_info.name or ""
            if parent_name.strip():
                current_values["{PARENT_NAME}"] = parent_name.strip()

        return current_values

    except Exception as e:
        print(f"Error occurred while extracting environment values: {str(e)}")
        return {}
#%%
def _collect_child_rects(window):
    """Collect the rectangles of all visible child controls (screen coordinates)"""
    rects = []
    for child in window.descendants():
        try:
            r = child.rectangle()
            # Exclude controls with zero size or hidden controls
            if r.width() > 0 and r.height() > 0:
                rects.append(r)
        except Exception:
            pass
    return rects

def _is_point_in_rects(x, y, rects):
    """Check whether (x, y) falls within any child-control rectangle"""
    for r in rects:
        if r.left <= x <= r.right and r.top <= y <= r.bottom:
            return True
    return False
# After entering text into an edit control, click a safe blank area to exit edit mode and apply the text
def click_blank_spot(window, granularity=20, margin=5):
    """
    Find a "lazy blank spot" in the client area of the window and click it.

    Args:
        window (pywinauto.WindowSpecification / UIAWrapper): Top-level window or dialog
        granularity (int): Grid sampling step size in pixels; smaller is more precise but slower
        margin (int): Safety pixels to keep away from the window edge

    Returns:
        (x, y): The actual screen coordinates clicked

    Raises:
        RuntimeError: If no blank spot can be found after scanning the whole area
    """
    win_rect = window.rectangle()

    # Grid sampling within the client area
    start_x = win_rect.left + margin
    start_y = win_rect.top  + margin
    end_x   = win_rect.right  - margin
    end_y   = win_rect.bottom - margin

    # First collect all child control rectangles to speed up later checks
    child_rects = _collect_child_rects(window)

    # Search from bottom-right to top-left, which usually avoids menus/toolbars in the upper-left
    for y in range(end_y, start_y, -granularity):
        for x in range(end_x, start_x, -granularity):
            if not _is_point_in_rects(x, y, child_rects):
                # Found a blank spot ➜ click it
                mouse.click(button='left', coords=(x, y))
                return (x, y)

    raise RuntimeError("No safe lazy blank area found to click")
#%%
# For operations like replace that require extra wait time; otherwise the UI may not finish loading
# current_top_window input requirement:
#   It must be the topmost window within current_window
#   (for example, a child_window if one exists; otherwise current_window itself)
# Note: if current_window contains a child_window and you pass current_window directly,
# the function may mistakenly think the click succeeded when in fact it did not
#%%

#%%
def navigate_and_execute_with_optimized_path(structured_instruction, app_level_window, deepest_index, first_control=None):
    """
    Navigate to the specified control using the optimized path and perform the operation
    Assumes that the passed-in app-level window definitely contains controls on the path,
    so the path is simplified directly based on deepest_index

    Args:
        structured_instruction: Structured instruction containing navigation path and operation info
        app_level_window: app-level window object
        deepest_index: Index of the deepest available control, used to simplify the path
        first_control: The first found control instance (performance optimization); if None, the default lookup logic is used

    Returns:
        tuple: (success: bool, updated_window: UIAWrapper or None, error_message: str, should_be_reachable: bool)
    """
    try:
        unique_id = structured_instruction["unique_id"]
        navigation_path = structured_instruction["navigation_path"]
        instruction_type = structured_instruction["instruction_type"]
        control_name = structured_instruction.get("name", "[Unknown]")
        control_type = structured_instruction.get("control_type", None)

        print(f"🚀 Starting optimized-path navigation to control: {control_name} ({instruction_type})")

        # Get the process ID of the original window, used to monitor new windows
        try:
            original_window_pid = app_level_window.process_id()
        except Exception as e:
            return False, None, f"Failed to get the window process ID: {str(e)}", False

        working_window = app_level_window
        should_be_reachable = True  # Assume it should be reachable
        fuzzy_matching_mode = False  # Indicates whether fuzzy matching mode has been entered

        # Simplify the navigation path based on deepest_index
        optimized_path = navigation_path[deepest_index:]
        print(f"📍 Using simplified path: starting from index {deepest_index}, {len(navigation_path)} -> {len(optimized_path)} nodes")

        # If the simplified path has only one element (the target control), use the cache directly
        if len(optimized_path) == 1:
            if first_control is not None:
                print("✅ Simplified path contains only the target control; using cached control (performance optimization)")
                target_control = first_control
            else:
                print("✅ Simplified path contains only the target control; locating it directly")
                target_control = find_control_by_identifier(working_window, unique_id, control_type)
        else:
            target_control = None

        # If navigation is needed, perform it step by step
        if len(optimized_path) > 1:
            print("🧭 Starting step-by-step navigation...")

            # Step-by-step navigation to the target control location
            i = 0
            while i < len(optimized_path):
                path_id = optimized_path[i]
                print(f"  Navigation step {i+1}: {path_id}")

                # Find the control at the current step on the path
                if i == 0 and first_control is not None:
                    # Performance optimization: use the already found first control instance
                    current_control = first_control
                    print(f"    ⚡ Using the cached first control (performance optimization)")
                elif i == 0:
                    # First control without cache: use default lookup logic
                    current_control = find_control_by_identifier(working_window, path_id)
                    print(f"    Using the default lookup method (first control)")
                else:
                    current_control = find_control_by_identifier(working_window, path_id)

                if not current_control:
                    print(f"  ✗ Failed to find path control: {path_id}")

                    # Intelligent path repair: use the unified repair mechanism
                    current_control, i = perform_path_repair(working_window, optimized_path, i)

                    if not current_control:
                        # If path repair fails and it should be reachable, start fuzzy matching
                        if should_be_reachable:
                            print(f"  ⚠️ It should be reachable, but path repair failed; starting fuzzy matching...")
                            fuzzy_matching_mode = True

                            # Perform fuzzy-matching navigation
                            success, updated_window, fuzzy_error = fuzzy_navigate_and_execute(
                                structured_instruction, working_window, optimized_path[i:],
                                original_window_pid
                            )

                            if success:
                                success_msg = f"✓ Control operation completed: {control_name} (using fuzzy matching)"
                                print(success_msg)
                                return True, updated_window, "", should_be_reachable
                            else:
                                error_msg = f"UI navigation tree modeling issue: {fuzzy_error}. Controls on the path were found, which means the target control should theoretically be reachable, but fuzzy matching also failed."
                                return False, working_window, error_msg, should_be_reachable
                        else:
                            # If subsequent controls also cannot be found, attempt UI recovery and exit
                            # The error handling logic here may need to be unified with the final exception handling
                            return False, working_window, f"Navigation failed: unable to find control {path_id}", should_be_reachable

                # If the path has not yet reached the last control (the target control), click this control to expand the UI
                if i < len(optimized_path) - 1:
                    print(f"  Clicking intermediate control: {current_control.element_info.name or '[Unnamed]'}")

                    # Use a context manager to handle window monitoring
                    with TopLevelWindowDetector(original_window_pid, timeout=5, control_name="Navigation") as monitor:
                        # Click the control
                        working_window.set_focus()
                        current_control.click_input()
                        time.sleep(0.2)  # Wait for UI response

                        # Wait for monitoring to complete
                        control_display_name = current_control.element_info.name or ""
                        control_display_type = current_control.element_info.control_type or ""
                        monitor.wait_for_window_completion(control_display_name, control_display_type)

                        if monitor.get_new_window():
                            working_window = monitor.get_new_window()
                            print(f"    Switched to new popup window: {working_window.window_text()}")

                # Move to the next control
                i += 1

            # Only set target_control after traversing the full path successfully
            if i == len(optimized_path):
                target_control = current_control
                print(f"🎯 Navigation completed, target control: {control_name}")
        else:
            # The simplified path contains only the target control; navigation is complete
            print("🎯 Navigation completed (direct access)")

        # After navigation completes, check the target control again
        if not target_control:  # Note: this should not be triggered
            print("-----Note: this branch should not be triggered. Navigation completed; locating target control")
            if len(optimized_path) == 1:
                # If there is only the target control, locate it directly
                target_control = find_control_by_identifier(working_window, unique_id, control_type)

        if not target_control:
            if should_be_reachable and not fuzzy_matching_mode:
                # Try one final fuzzy match
                print("  ⚠️ Navigation completed but target control not found; attempting final fuzzy matching...")
                print(f"  ⚠️ It should be reachable, but the target control could not be found; starting fuzzy matching...")
                fuzzy_matching_mode = True

                # Perform fuzzy matching for the target control
                success, updated_window, fuzzy_error = fuzzy_navigate_and_execute(
                    structured_instruction, working_window, [unique_id],
                    original_window_pid
                )

                if success:
                    success_msg = f"✓ Control operation completed: {control_name} (using fuzzy matching)"
                    print(success_msg)
                    return True, updated_window, "", should_be_reachable
                else:
                    error_msg = f"UI navigation tree modeling issue: {fuzzy_error}. The target control should theoretically be reachable, but fuzzy matching also failed."
                    return False, working_window, error_msg, should_be_reachable
            else:
                error_msg = f"Unable to get target control: {control_name}"
                print(f"❌ {error_msg}")
                return False, working_window, error_msg, should_be_reachable

        # ===== Newly added: check whether the target control is enabled =====
        print(f"🔍 Checking whether the target control is enabled: {control_name}")
        try:
            if not target_control.is_enabled():
                error_msg = f"CONTROL_DISABLED: The target control '{control_name}' was found but is disabled and cannot be operated. Consider prerequisite dependencies—this control may become available only after other controls are accessed first."
                print(f"❌ {error_msg}")
                return False, working_window, error_msg, should_be_reachable
            else:
                print(f"✅ The target control is enabled; proceeding with the operation")
        except Exception as e:
            # If enabled status cannot be checked, log a warning and continue
            print(f"⚠️ Unable to check whether the target control is enabled: {str(e)}; continuing with the operation")

        # Execute the target control operation
        print(f"🎯 Executing target control operation: {instruction_type}")

        with TopLevelWindowDetector(original_window_pid, timeout=5, control_name="Target") as monitor:

            # Execute the specific operation
            working_window.set_focus()

            # === Execute the target control operation using execute_control_with_interaction ===
            text_to_input = structured_instruction.get("text") if instruction_type == "edit" else None
            success = execute_control_with_interaction(target_control, working_window, text_to_input)

            if not success:
                return False, working_window, f"Target control operation failed: {control_name}", should_be_reachable

            time.sleep(0.2)  # Wait for the operation to complete

            # Wait for target-control monitoring to complete
            target_control_name = target_control.element_info.name or ""
            target_control_type = target_control.element_info.control_type or ""
            monitor.wait_for_window_completion(target_control_name, target_control_type)

            if monitor.get_new_window():
                working_window = monitor.get_new_window()
                print(f"  Switched to new popup window after operating on target control: {working_window.window_text()}")

        success_msg = f"✓ Control operation completed: {control_name}"
        if fuzzy_matching_mode:
            success_msg += " (using fuzzy matching)"
        print(success_msg)
        return True, working_window, "", should_be_reachable

    except Exception as e:
        error_msg = f"Error occurred while navigating and operating on the control: {str(e)}"
        print(f"✗ {error_msg}")

        # Try to restore the UI state
        try:
            send_keys("{ESC}")
            time.sleep(0.1)
        except Exception:
            pass

        return False, app_level_window, error_msg, False

def reverse_scan_path(navigation_path, control_instances):
    """
    Reverse-scan the navigation path to find the deepest available control

    Args:
        navigation_path: Full navigation path
        control_instances: Dictionary cache of control instances

    Returns:
        tuple: (optimized path list, deepest available control index or None)
    """
    # Scan backward from the end of the path
    for i in reversed(range(len(navigation_path))):
        path_id = navigation_path[i]
        if path_id in control_instances:
            # Found the deepest available control; return the path starting from this point
            optimized_path = navigation_path[i:]
            return optimized_path, i

    # No available control was found
    return navigation_path, None
#%%
def select_all_text_by_text_pattern(target_control):
    """
    Use TextPattern to select all text in the control

    Args:
        target_control: Target control

    Returns:
        bool: Whether the selection succeeded
    """
    try:
        # Use the unified check function
        if not is_text_pattern_supported(target_control):
            print("The control does not support TextPattern")
            return False

        # Get the underlying UIA element
        element = target_control.element_info.element
        UIA_dll = pywinauto.uia_defines.IUIA().UIA_dll

        # Get the TextPattern interface
        text_pattern = element.GetCurrentPattern(UIA_dll.UIA_TextPatternId)
        text_pattern_interface = text_pattern.QueryInterface(UAC.IUIAutomationTextPattern)

        # Get the full document range
        document_range = text_pattern_interface.DocumentRange

        # Directly select the full document range (equivalent to Ctrl+A)
        document_range.Select()

        return True
    except Exception as e:
        print(f"TextPattern selection failed: {e}")
        return False

# Control interaction: click or edit
def execute_control_with_interaction(control, current_top_window, text=None):
    from ufo.automator.ui_control.controller import TextTransformer
    """
    Perform an action on the specified control (click or enter text)

    Args:
        control: Control object
        current_top_window: Current window
        text: Text content (only needed for Edit controls)

    Returns:
        bool: Whether the operation succeeded
    """
    try:
        control_type = control.element_info.control_type
        control_name = control.element_info.name or control.window_text() or "[Unnamed]"

        print(f"Executing control action: {control_name} ({control_type})")

        print("current top window:",current_top_window)
        # current_top_window.set_focus()

        if control_type == "Edit" and text is not None:
            current_top_window.set_focus()
            # Edit control: clear and enter text
            print(f"  Entering text: '{text}'")

            control.click_input()  # Set focus
            time.sleep(0.1)

            # Clear existing content
            # Prefer using TextPattern for Select All
            if is_text_pattern_supported(control):
                print("  The control supports TextPattern; using TextPattern for Select All")
                if select_all_text_by_text_pattern(control):
                    print("  ✓ TextPattern Select All succeeded")
                else:
                    print("  TextPattern Select All failed; falling back to the original approach")
                    # Fall back to the original Ctrl+A + DELETE method
                    control.type_keys("^a", with_spaces=True)  # Ctrl+A
                    time.sleep(0.05)
                    control.type_keys("{DELETE}", with_spaces=True)
                    time.sleep(0.15)
            else:
                print("  The control does not support TextPattern; using the original Ctrl+A + DELETE method")
                # Clear existing content
                control.type_keys("^a", with_spaces=True)  # Ctrl+A
                time.sleep(0.5)
                control.type_keys("{DELETE}", with_spaces=True)
                time.sleep(0.5)

            # Enter new text
            if text:
                # The type_keys method requires certain characters to be escaped
                text = TextTransformer.transform_text(text, "all")

                control.type_keys(text, with_spaces=True)
                # control.iface_value.SetValue(text) # This is not ideal; it breaks automatic case behavior
                # To make the edit take effect, it may be necessary to click an inert control
                # try:
                #     blank_xy = click_blank_spot(current_top_window)
                #     print("✓ Successfully clicked a blank area:", blank_xy)
                # except RuntimeError as err:
                #     print("✗", err)
        elif control_type == "ListItem":
            # current_top_window.set_focus()


            # ListItem control: prefer invoke first, and fall back to click_input if it fails # Scenario? (invoke on png may trigger a popup, which is different from clicking the png)
            # print(f"  Clicking ListItem control (prefer invoke)")
            # try:
            #     # Try the invoke method
            #     control.set_focus()
            #     control.invoke()
            #     print(f"  ✓ invoke succeeded")
            # except Exception as invoke_error:
            #     print(f"  ⚠️ invoke failed; falling back to click_input: {str(invoke_error)}")
            #     control.click_input()
            #     print(f"  ✓ click_input succeeded")
            print(f"  Clicking ListItem control")
            control_visible = control.is_visible()
            if not control_visible:
                print("  Control is not visible; trying invoke")
                try:
                    # Try the invoke method
                    control.set_focus()
                    control.invoke()
                    print(f"  ✓ invoke succeeded")
                except Exception as invoke_error:
                    print(f"  ⚠️ invoke failed; falling back to click_input: {str(invoke_error)}")
                    control.click_input()
                    print(f"  ✓ click_input succeeded")
            else:
                    # current_top_window.set_focus()
                    # In Excel, clicking again after select on a frozen pane may cause it to stop taking effect
                    # select_controls(control)
                    control.set_focus()
                    control.click_input()
                    print(f"  ✓ click_input succeeded")

        else:
            current_top_window.set_focus()

            control_visible = control.is_visible()
            if not control_visible:
                print("  Control is not visible; trying invoke")
                try:
                    current_top_window.set_focus()
                    control.invoke()
                    print("  ✓ invoke succeeded")
                except Exception as invoke_error:
                    print(f"  ⚠️ invoke failed ({invoke_error}); falling back to click_input")
                    # print("control.is_visible()", control.is_visible())
                    # print("control.is_enabled()", control.is_enabled())
                    print(control.get_properties())
                    # print(get_control_identifier(control))
                    # print(control.rectangle())
                    # Needed for "Desktop"
                    control.set_focus()
                    control.click_input()
                    print("  ✓ click_input succeeded")
            else:
                print("  Control is visible; using click_input")
                control.click_input()
            # print(f"  Clicking control")
            # control.click_input()

        time.sleep(0.2)
        print(f"✓ Control action completed")
        return True

    except Exception as e:
        print(f"✗ Control action failed: {str(e)}")
        return False
# Execute shortcut keys
def execute_shortcut_instruction(structured_instruction, main_window):
    from ufo.automator.ui_control.controller import TextTransformer
    """
    Execute a shortcut-key instruction

    Args:
        structured_instruction: Structured instruction containing shortcut-key information
        main_window: Main window (app-level window)

    Returns:
        tuple: (success: bool, updated_window: UIAWrapper, error_message: str)
    """
    try:
        shortcut_key = structured_instruction.get("shortcut_key", "")
        control_name = structured_instruction.get("name", f"shortcut_key: {shortcut_key}")

        print(f"⌨️ Executing shortcut instruction: {shortcut_key}")

        if not shortcut_key:
            error_msg = "Shortcut key is empty"
            print(f"❌ {error_msg}")
            return False, main_window, error_msg

        # Get the process ID for window monitoring
        try:
            original_window_pid = main_window.process_id()
        except Exception as e:
            error_msg = f"Failed to get the window process ID: {str(e)}"
            print(f"❌ {error_msg}")
            return False, main_window, error_msg

        # Get the actual top_operatable_window used to execute the shortcut
        top_operatable_window = get_top_operatable_window(original_window_pid)
        if not top_operatable_window:
            error_msg = "Failed to get top_operatable_window"
            print(f"❌ {error_msg}")
            return False, main_window, error_msg

        # Use the window detector to monitor new windows that may be triggered by the shortcut
        with TopLevelWindowDetector(original_window_pid, timeout=5, control_name="shortcut_key") as monitor:
            # Ensure top_operatable_window gets focus
            top_operatable_window.set_focus()
            time.sleep(0.1)  # Wait for focus to be set

            # Escape the shortcut text (using the escape method you provided)
            keys = TextTransformer.transform_text(shortcut_key, "all")
            print(f"  Escaped keys: {keys}")

            # Execute the shortcut on top_operatable_window
            top_operatable_window.type_keys(keys)
            print(f"  ✓ Shortcut sent to top_operatable_window")

            # Wait briefly for the shortcut to take effect
            time.sleep(0.3)

            # Wait for monitoring to complete
            monitor.wait_for_window_completion(control_name, "Shortcut")

            # Check whether a new window was opened
            updated_window = main_window
            if monitor.get_new_window():
                updated_window = monitor.get_new_window()
                print(f"  ✓ Switched to new window after shortcut action: {updated_window.window_text()}")
            else:
                # If no new window was opened, return the current top_operatable_window
                # Reacquire the latest top_operatable_window, because the shortcut may have changed the window state
                updated_top_operatable_window = get_top_operatable_window(original_window_pid)
                if updated_top_operatable_window:
                    updated_window = updated_top_operatable_window

        print(f"✅ Shortcut executed successfully: {shortcut_key}")
        time.sleep(0.8)
        return True, updated_window, ""

    except Exception as e:
        error_msg = f"An exception occurred while executing the shortcut: {str(e)}"
        print(f"❌ {error_msg}")
        return False, main_window, error_msg
#%%
# Fuzzy matching   For consecutive fuzzy matching, this may cause the same UI to be opened repeatedly

def fuzzy_navigate_and_execute(structured_instruction, working_window, remaining_path, original_window_pid):
    """
    Perform fuzzy-matched navigation and execution.
    Note: this function is responsible not only for navigation,
    but also for executing the final target control.

    Args:
        structured_instruction: Structured instruction
        working_window: Current working window
        remaining_path: Remaining navigation path (the part that needs fuzzy matching)
        original_window_pid: Original window process ID

    Returns:
        tuple: (success: bool, updated_window: UIAWrapper or None, error_message: str)
    """
    print("🔍 Starting fuzzy-match navigation...")

    try:
        current_window = working_window

        # Process the remaining path step by step using fuzzy matching
        for i, path_id in enumerate(remaining_path):
            print(f"  🔍 Fuzzy-matching path control {i+1}/{len(remaining_path)}: {path_id}")

            # Try fuzzy-matching the current path control
            matched_control = fuzzy_match_control(current_window, path_id)

            if not matched_control:
                print(f"  ✗ Fuzzy match failed: {path_id}")

                # Try skipping the current control and look for a later control
                found_alternative = False
                for j in range(i + 1, len(remaining_path)):
                    alternative_path_id = remaining_path[j]
                    alternative_control = fuzzy_match_control(current_window, alternative_path_id)
                    if alternative_control:
                        print(f"  ✓ Found a later control via fuzzy matching; skipping intermediate steps: {alternative_path_id}")
                        matched_control = alternative_control
                        i = j  # Update index
                        found_alternative = True
                        break

                if not found_alternative:
                    return False, current_window, f"Fuzzy matching could not find the path control: {path_id}"
            else:
                print(f"  ✓ Fuzzy match succeeded: {matched_control.element_info.name or '[Unnamed]'}")

            # If this is not the last control, click to expand the UI
            if i < len(remaining_path) - 1:
                print(f"  Clicking the intermediate control matched by fuzzy matching: {matched_control.element_info.name or '[NoText]'}")

                # Use a context manager to replace the original monitoring logic
                with TopLevelWindowDetector(original_window_pid, timeout=5,
                                        control_name="Fuzzy Match Navigation") as monitor:
                    # Click the control
                    current_window.set_focus()
                    matched_control.click_input()
                    time.sleep(0.2)

                    # Wait for monitoring to complete (pass in the control name and type)
                    control_display_name = matched_control.element_info.name or ""
                    control_display_type = matched_control.element_info.control_type or ""

                    monitor.wait_for_window_completion(control_display_name, control_display_type)

                    if monitor.get_new_window():
                        current_window = monitor.get_new_window()
                        print(f"  Fuzzy matching switched to a new popup window: {current_window.window_text()}")
            else:
                # Final control: execute the action directly
                print(f"  Executing the action on the target control matched by fuzzy matching...")

                # ===== Check whether the target control is enabled (fuzzy-match path) =====
                control_name = structured_instruction.get("name", "[Unknown]")
                print(f"🔍 Checking whether the target control from fuzzy matching is enabled: {control_name}")
                try:
                    if not matched_control.is_enabled():
                        error_msg = f"CONTROL_DISABLED: The target control found by fuzzy matching, '{control_name}', is disabled and cannot be operated"
                        print(f"❌ {error_msg}")
                        return False, current_window, error_msg
                    else:
                        print(f"✅ The target control from fuzzy matching is enabled; continuing with execution")
                except Exception as e:
                    # If the enabled state cannot be checked, log a warning and continue
                    print(f"⚠️ Unable to check whether the target control from fuzzy matching is enabled: {str(e)}. Continuing with execution")

                # Execute the target control action
                instruction_type = structured_instruction["instruction_type"]

                current_window.set_focus()

                if instruction_type == "edit":
                    # Edit control: get the text to enter
                    text_to_input = structured_instruction.get("text", "")
                    print(f"    Entering text into the Edit control via fuzzy matching: '{text_to_input}'")

                    # Use the unified control interaction method
                    success = execute_control_with_interaction(matched_control, current_window, text=text_to_input)
                else:
                    # Non-Edit control: click directly
                    print(f"    Clicking control via fuzzy matching: {control_name}")

                    # Use the unified control interaction method
                    success = execute_control_with_interaction(matched_control, current_window)

                if not success:
                    return False, current_window, f"Failed to operate the target control found by fuzzy matching: {control_name}"

        print("✓ Fuzzy-match navigation and execution completed")
        return True, current_window, ""

    except Exception as e:
        error_msg = f"An error occurred during fuzzy-match navigation: {str(e)}"
        print(f"✗ {error_msg}")
        return False, working_window, error_msg

def fuzzy_match_control(window, path_id):
    """
    Fuzzy-match a control in the specified window.

    Args:
        window: Target window
        path_id: Path identifier, e.g. "alt_id:Subscript(B)|CheckBox|Find Font/Find and Replace/Desktop 1"

    Returns:
        UIAWrapper or None: The matched control, or None if not found
    """
    try:
        # Parse path_id to get key information
        control_info = parse_path_identifier(path_id)
        if not control_info:
            print(f"    Unable to parse path identifier: {path_id}")
            return None

        # Extract key matching information
        identifier_type = control_info.get("type")  # alt_id, automation_id, etc. # Note: internally unified as alt_id
        control_name = control_info.get("name", "")
        control_type = control_info.get("control_type", "")

        print(f"    Attempting fuzzy match: name='{control_name}', type='{control_type}', identifier_type='{identifier_type}'")

        # Strategy 1: exact match (as a fallback)
        exact_match = find_control_by_identifier(window, path_id)
        if exact_match:
            print(f"    ✓ Exact match succeeded")
            return exact_match

        # Strategy 2: fuzzy match by name and type
        if control_name and control_type:
            # print("------------------Strategy 2: fuzzy match by name and type")
            fuzzy_matches = fuzzy_find_controls_by_name_and_type(window, control_name, control_type)
            if fuzzy_matches:
                best_match = fuzzy_matches[0]  # Take the most similar one
                print(f"    ✓ Fuzzy match by name + type succeeded: {best_match.element_info.name or '[Unnamed]'}")
                return best_match

        print(f"    ✗ All fuzzy-matching strategies failed")
        return None

    except Exception as e:
        print(f"    Error during fuzzy matching: {str(e)}")
        return None

def fuzzy_find_target_control(window, structured_instruction):
    """
    Fuzzy-match the target control.

    Args:
        window: Target window
        structured_instruction: Structured instruction

    Returns:
        UIAWrapper or None: The matched control
    """
    try:
        unique_id = structured_instruction["unique_id"]
        control_name = structured_instruction.get("name", "")
        control_type = structured_instruction.get("control_type", "")

        print(f"    Fuzzy-matching target control: name='{control_name}', type='{control_type}'")

        # Try fuzzy matching
        return fuzzy_match_control(window, unique_id)

    except Exception as e:
        print(f"    Error while fuzzy-matching the target control: {str(e)}")
        return None

def parse_path_identifier(path_id):
    """
    Parse a path identifier and extract key information.

    Args:
        path_id: Path identifier, such as "alt_id:name|control_type|(automation_id)|ancestor_info"

    Returns:
        dict: A dictionary containing the parsed information
    """
    try:
        # Directly use the existing unified parsing function
        parsed = parse_unique_id(path_id)
        if not parsed:
            return None

        # Convert to the format expected by fuzzy matching
        result = {
            "type": "alt_id",  # Currently unified to support alt_id
            "name": parsed["name"],
            "control_type": parsed["control_type"],
            "automation_id": parsed["automation_id"],
            "ancestor_info": parsed["ancestor_info"]
        }

        return result

    except Exception as e:
        print(f"Error while parsing path identifier: {str(e)}")
        return None

def fuzzy_find_controls_by_name_and_type(window, target_name, target_type, similarity_threshold=0.6, keyword_match_groups=["填充颜色"]):
    """
    Fuzzy-find controls by name and type.

    Args:
        window: Target window
        target_name: Target name
        target_type: Target type
        similarity_threshold: Similarity threshold
        keyword_match_groups: A list of keyword groups
                             A match is considered successful as long as both the target name
                             and the control name contain any item from the same group.

    Returns:
        list: A list of matched controls sorted by similarity
    """
    try:
        from difflib import SequenceMatcher

        matches = []

        # Traverse all controls in the window
        def walk_controls(element):
            try:
                element_name = element.element_info.name or ""
                element_type = element.element_info.control_type or ""

                # Calculate name similarity
                name_similarity = SequenceMatcher(None, target_name.lower(), element_name.lower()).ratio()

                # Type must match
                if element_type == target_type:
                    # Check whether the match conditions are satisfied
                    is_match = False
                    match_reason = ""

                    # Condition 1: traditional similarity match
                    if name_similarity >= similarity_threshold:
                        is_match = True
                        match_reason = f"similarity match ({name_similarity:.2f})"

                    # Condition 2: keyword match group
                    if keyword_match_groups and not is_match:
                        target_lower = target_name.lower()
                        element_lower = element_name.lower()

                        for keyword in keyword_match_groups:
                            keyword_lower = keyword.lower()
                            # Check whether both the target name and the control name contain the keyword
                            if keyword_lower in target_lower and keyword_lower in element_lower:
                                is_match = True
                                match_reason = f"keyword match ('{keyword}')"
                                # Give keyword matches a higher similarity score
                                # so they rank higher in the results
                                name_similarity = max(name_similarity, 0.8)
                                break

                    # If matched successfully, add to the results
                    if is_match:
                        matches.append({
                            "control": element,
                            "name_similarity": name_similarity,
                            "element_name": element_name,
                            "match_reason": match_reason
                        })

                # Recursively process child controls
                try:
                    for child in element.children():
                        walk_controls(child)
                except:
                    pass

            except Exception:
                pass

        walk_controls(window)

        # Sort by similarity
        matches.sort(key=lambda x: x["name_similarity"], reverse=True)

        # Return the control list
        result = [match["control"] for match in matches]

        if result:
            print(f"      Found {len(result)} fuzzy matches by name + type")
            for i, match in enumerate(matches[:3]):  # Show only the top 3
                print(f"        {i+1}. '{match['element_name']}' ({match['match_reason']})")

        return result

    except Exception as e:
        print(f"      Error while fuzzy-finding by name + type: {str(e)}")
        return []

#%%
# Target selection
def select_controls(controls):
    """
    Select multiple controls - supports either a single control or a list of controls.

    Args:
        controls: A single control object or a list of controls

    Returns:
        dict: Detailed selection result information
    """
    # Handle input format - normalize everything into a list
    if controls is None:
        return {"success": False, "error": "The input control is None"}

    # If the input is not a list, convert it to a list
    if not isinstance(controls, (list, tuple)):
        controls = [controls]

    # Validate that the list is not empty
    if len(controls) == 0:
        return {"success": False, "error": "The control list is empty"}

    # Validate that all elements are control objects
    # for i, control in enumerate(controls):
    #     if not hasattr(control, 'iface_selection_item'):
    #         return {
    #             "success": False,
    #             "error": f"Element {i+1} is not a valid control object; missing selection interface"
    #         }
    # Validate that all elements are control objects
    for i, control in enumerate(controls):
        try:
            # Try accessing the interface to verify support
            _ = control.iface_selection_item
        except (AttributeError, NoPatternInterfaceError):
            return {
                "success": False,
                "error": f"Element {i+1} is not a valid control object; missing the selection interface"
            }
    # Single-control case
    if len(controls) == 1:
        control = controls[0]
        control_name = getattr(control.element_info, 'name', None) or "[Unnamed]"

        try:
            control.iface_selection_item.Select()
            success = control.is_selected()
            return {
                "success": success,
                "selected_count": 1 if success else 0,
                "total_count": 1,
                "method": "single_select",
                "details": [
                    {
                        "control_name": control_name,
                        "operation": "Select",
                        "success": success,
                        "is_selected": success
                    }
                ]
            }
        except Exception as e:
            return {
                "success": False,
                "selected_count": 0,
                "total_count": 1,
                "method": "single_select",
                "error": f"Failed to select the single control: {str(e)}",
                "details": [
                    {
                        "control_name": control_name,
                        "operation": "Select",
                        "success": False,
                        "error": str(e)
                    }
                ]
            }

    # Multiple-controls case
    operation_details = []

    for i, control in enumerate(controls):
        # Safely get the control name
        try:
            control_name = getattr(control.element_info, 'name', None) or f"[Control {i+1}]"
        except Exception:
            control_name = f"[Control {i+1}]"

        try:
            # Use Select for the first control, and AddToSelection for the others
            if i == 0:
                control.iface_selection_item.Select()
                operation = "Select"
                print(f"✓ Selected the first control: {control_name}")
            else:
                control.iface_selection_item.AddToSelection()
                operation = "AddToSelection"
                print(f"✓ Added control to selection: {control_name}")

            # Check whether the current control is selected
            try:
                current_is_selected = control.is_selected()
            except Exception:
                current_is_selected = False

            # Verify whether all controls that should be selected (from index 0 to i) are selected
            expected_selected_controls = controls[:i+1]
            actual_selection_status = []
            all_expected_selected = True

            for j, expected_control in enumerate(expected_selected_controls):
                try:
                    expected_name = getattr(expected_control.element_info, 'name', None) or f"[Control {j+1}]"
                except Exception:
                    expected_name = f"[Control {j+1}]"

                try:
                    is_actually_selected = expected_control.is_selected()
                    actual_selection_status.append({
                        "control_name": expected_name,
                        "expected_selected": True,
                        "actually_selected": is_actually_selected
                    })
                    if not is_actually_selected:
                        all_expected_selected = False
                except Exception as e:
                    actual_selection_status.append({
                        "control_name": expected_name,
                        "expected_selected": True,
                        "actually_selected": False,
                        "check_error": str(e)
                    })
                    all_expected_selected = False

            # Record detailed information about the current operation
            operation_detail = {
                "step": i + 1,
                "control_name": control_name,
                "operation": operation,
                "operation_success": True,
                "current_control_selected": current_is_selected,
                "all_expected_selected": all_expected_selected,
                "selection_status": actual_selection_status
            }

            # If validation fails, record the failure info and return
            if not all_expected_selected:
                operation_detail["validation_failed"] = True

                # Find which controls are not selected
                unselected_controls = [
                    status["control_name"] for status in actual_selection_status
                    if not status["actually_selected"]
                ]

                operation_detail["unselected_controls"] = unselected_controls
                operation_details.append(operation_detail)

                print(f"✗ Validation failed at step {i+1}: the following controls were not selected: {unselected_controls}")

                return {
                    "success": False,
                    "selected_count": sum(1 for status in actual_selection_status if status["actually_selected"]),
                    "total_count": len(controls),
                    "method": "multi_select",
                    "failed_at_step": i + 1,
                    "failed_control": control_name,
                    "failure_reason": f"Validation failed: {len(unselected_controls)} control(s) were not selected",
                    "details": operation_details
                }

            operation_details.append(operation_detail)

        except Exception as e:
            # The operation itself failed
            error_msg = str(e)
            print(f"✗ Operation failed at step {i+1}: {error_msg}")

            # Still check the current selection status
            actual_selection_status = []
            for j, check_control in enumerate(controls[:i+1]):
                try:
                    check_name = getattr(check_control.element_info, 'name', None) or f"[Control {j+1}]"
                except Exception:
                    check_name = f"[Control {j+1}]"

                try:
                    is_selected = check_control.is_selected()
                    actual_selection_status.append({
                        "control_name": check_name,
                        "expected_selected": True,
                        "actually_selected": is_selected
                    })
                except Exception:
                    actual_selection_status.append({
                        "control_name": check_name,
                        "expected_selected": True,
                        "actually_selected": False
                    })

            operation_detail = {
                "step": i + 1,
                "control_name": control_name,
                "operation": "Select" if i == 0 else "AddToSelection",
                "operation_success": False,
                "operation_error": error_msg,
                "selection_status": actual_selection_status
            }
            operation_details.append(operation_detail)

            return {
                "success": False,
                "selected_count": sum(1 for status in actual_selection_status if status["actually_selected"]),
                "total_count": len(controls),
                "method": "multi_select",
                "failed_at_step": i + 1,
                "failed_control": control_name,
                "failure_reason": f"Operation failed: {error_msg}",
                "details": operation_details
            }

    # All operations succeeded
    final_selected_count = len([detail for detail in operation_details
                               if detail.get("current_control_selected", False)])

    print(f"✓ Multi-select operation completed: successfully selected {final_selected_count}/{len(controls)} controls")

    return {
        "success": True,
        "selected_count": final_selected_count,
        "total_count": len(controls),
        "method": "multi_select" if len(controls) > 1 else "single_select",
        "details": operation_details
    }
# For a specific control type, whether multi-selection is supported can be checked with parent().can_select_multiple()
# Basic capabilities for locating context (operation target) controls:
#control.iface_selection_item.Select()
#control.is_selected()


#%%

#%%
from comtypes.gen.UIAutomationClient import (
    TextUnit_Line, TextUnit_Paragraph,
    TextPatternRangeEndpoint_Start, TextPatternRangeEndpoint_End,
)

def _find_text_provider(control):
    """Find a control that supports TextPattern (the control itself or one of its descendants)."""
    if getattr(control, "iface_text", None):
        return control
    for descendant in control.descendants():
        if getattr(descendant, "iface_text", None):
            return descendant
    return None


def _collapse_to_start(rng):
    """Collapse a TextRange to the start of the document (Start==End==DocStart)"""
    rng.MoveEndpointByRange(TextPatternRangeEndpoint_End, rng, TextPatternRangeEndpoint_Start)
    return rng

def _segment_text_is_empty(seg_range):
    """Check whether a text segment is empty (only \r\n / spaces / tabs, etc. are considered empty)"""
    # Note: GetText(-1) may include a trailing \r\n, so remove those before checking
    txt = seg_range.GetText(-1)
    return txt.strip() == ""

def _nth_nonempty_to_physical_index(text_pattern, unit, nth_nonempty):
    """
    Map the "nth non-empty segment (line/paragraph)" to the "physical segment index" (0-based)
    Returns: (phys_index or None, total_nonempty_found)
    """
    doc = text_pattern.DocumentRange
    cur_start = _collapse_to_start(doc.Clone())
    phys_index = 0
    count_nonempty = 0

    while True:
        # Compute [cur_start, next_start) as the current segment
        next_start = cur_start.Clone()
        moved = next_start.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, unit, 1)

        seg = cur_start.Clone()
        seg.MoveEndpointByRange(TextPatternRangeEndpoint_End, next_start, TextPatternRangeEndpoint_Start)

        if not _segment_text_is_empty(seg):
            count_nonempty += 1
            if count_nonempty == nth_nonempty:
                return phys_index, count_nonempty  # Hit

        if moved == 0:
            # Reached the end of the document
            break

        cur_start = next_start
        phys_index += 1

    return None, count_nonempty  # Fewer than nth_nonempty segments

def select_line(control, start_line, end_line=None, non_empty=True):
    """
    Precisely select the specified line or line range.
    When non_empty=True, start_line / end_line represent the "Nth non-empty line".
    All other behavior remains the same as the original version.
    """
    try:

        control.set_focus()
        control.click_input()

        if end_line is None:
            end_line = start_line

        if start_line < 1 or end_line < 1:
            return {
                "success": False,
                "start_line": start_line,
                "end_line": end_line,
                "selected_text": "",
                "error": f"Invalid line number: start line {start_line}, end line {end_line}. Line numbers must start from 1"
            }

        if start_line > end_line:
            return {
                "success": False,
                "start_line": start_line,
                "end_line": end_line,
                "selected_text": "",
                "error": f"Invalid line range: start line {start_line} > end line {end_line}"
            }

        text_control = _find_text_provider(control)
        if not text_control:
            return {
                "success": False,
                "start_line": start_line,
                "end_line": end_line,
                "selected_text": "",
                "error": "Neither the control nor any of its descendants supports UIA TextPattern"
            }
        text_pattern = text_control.iface_text

        # —— Compute the physical line indices (0-based) ——
        if non_empty:
            start_phys, cnt_nonempty = _nth_nonempty_to_physical_index(text_pattern, TextUnit_Line, start_line)
            if start_phys is None:
                return {
                    "success": False,
                    "start_line": start_line,
                    "end_line": end_line,
                    "selected_text": "",
                    "error": f"Start line out of range: requested the {start_line}th non-empty line, but the document has only {cnt_nonempty} non-empty lines"
                }
            end_phys, cnt_nonempty2 = _nth_nonempty_to_physical_index(text_pattern, TextUnit_Line, end_line)
            if end_phys is None:
                return {
                    "success": False,
                    "start_line": start_line,
                    "end_line": end_line,
                    "selected_text": "",
                    "error": f"End line out of range: requested the {end_line}th non-empty line, but the document has only {cnt_nonempty2} non-empty lines"
                }
        else:
            # Original bounds checking (counted by physical line numbers, including empty lines)
            test_range = text_pattern.DocumentRange.Clone()
            _collapse_to_start(test_range)
            actual_moved_start = test_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, start_line - 1)
            if actual_moved_start < start_line - 1:
                return {
                    "success": False,
                    "start_line": start_line,
                    "end_line": end_line,
                    "selected_text": "",
                    "error": f"Start line out of range: requested line {start_line}, but the document has only {actual_moved_start + 1} lines"
                }

            test_range2 = text_pattern.DocumentRange.Clone()
            _collapse_to_start(test_range2)
            actual_moved_end = test_range2.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, end_line - 1)
            if actual_moved_end < end_line - 1:
                return {
                    "success": False,
                    "start_line": start_line,
                    "end_line": end_line,
                    "selected_text": "",
                    "error": f"End line out of range: requested line {end_line}, but the document has only {actual_moved_end + 1} lines"
                }

            start_phys = start_line - 1
            end_phys = end_line - 1

        # —— Build the selection range: from start_phys to the start of the line after end_phys ——
        target_range = text_pattern.DocumentRange.Clone()
        target_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, start_phys)
        target_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, target_range, TextPatternRangeEndpoint_Start)

        next_line_range = text_pattern.DocumentRange.Clone()
        next_line_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Line, end_phys + 1)
        next_line_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, next_line_range, TextPatternRangeEndpoint_Start)

        target_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, next_line_range, TextPatternRangeEndpoint_Start)

        selected_text = target_range.GetText(-1).rstrip('\r\n')
        target_range.Select()

        return {
            "success": True,
            "start_line": start_line,   # Keep the user-passed semantics here; when non_empty=True, this means the Nth non-empty line
            "end_line": end_line,
            "selected_text": selected_text,
            "error": None
        }

    except Exception as e:
        return {
            "success": False,
            "start_line": start_line,
            "end_line": end_line if end_line is not None else start_line,
            "selected_text": "",
            "error": f"Error while selecting lines {start_line}-{end_line if end_line is not None else start_line}: {str(e)}"
        }

def select_paragraph(control, start_paragraph, end_paragraph=None, non_empty=True):
    """
    Precisely select the specified paragraph or paragraph range.
    When non_empty=True, start/end represent the Nth non-empty paragraph.
    """
    try:

        control.set_focus()
        control.click_input()

        if end_paragraph is None:
            end_paragraph = start_paragraph

        if start_paragraph < 1 or end_paragraph < 1:
            return {
                "success": False,
                "start_paragraph": start_paragraph,
                "end_paragraph": end_paragraph,
                "selected_text": "",
                "error": f"Invalid paragraph number: start paragraph {start_paragraph}, end paragraph {end_paragraph}. Paragraph numbers must start from 1"
            }

        if start_paragraph > end_paragraph:
            return {
                "success": False,
                "start_paragraph": start_paragraph,
                "end_paragraph": end_paragraph,
                "selected_text": "",
                "error": f"Invalid paragraph range: start paragraph {start_paragraph} > end paragraph {end_paragraph}"
            }

        text_control = _find_text_provider(control)
        if not text_control:
            return {
                "success": False,
                "start_paragraph": start_paragraph,
                "end_paragraph": end_paragraph,
                "selected_text": "",
                "error": "Neither the control nor any of its descendants supports UIA TextPattern"
            }
        text_pattern = text_control.iface_text

        if non_empty:
            start_phys, cnt_nonempty = _nth_nonempty_to_physical_index(text_pattern, TextUnit_Paragraph, start_paragraph)
            if start_phys is None:
                return {
                    "success": False,
                    "start_paragraph": start_paragraph,
                    "end_paragraph": end_paragraph,
                    "selected_text": "",
                    "error": f"Start paragraph out of range: requested the {start_paragraph}th non-empty paragraph, but the document has only {cnt_nonempty} non-empty paragraphs"
                }
            end_phys, cnt_nonempty2 = _nth_nonempty_to_physical_index(text_pattern, TextUnit_Paragraph, end_paragraph)
            if end_phys is None:
                return {
                    "success": False,
                    "start_paragraph": start_paragraph,
                    "end_paragraph": end_paragraph,
                    "selected_text": "",
                    "error": f"End paragraph out of range: requested the {end_paragraph}th non-empty paragraph, but the document has only {cnt_nonempty2} non-empty paragraphs"
                }
        else:
            test_range = text_pattern.DocumentRange.Clone()
            _collapse_to_start(test_range)
            actual_moved_start = test_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Paragraph, start_paragraph - 1)
            if actual_moved_start < start_paragraph - 1:
                return {
                    "success": False,
                    "start_paragraph": start_paragraph,
                    "end_paragraph": end_paragraph,
                    "selected_text": "",
                    "error": f"Start paragraph out of range: requested paragraph {start_paragraph}, but the document has only {actual_moved_start + 1} paragraphs"
                }

            test_range2 = text_pattern.DocumentRange.Clone()
            _collapse_to_start(test_range2)
            actual_moved_end = test_range2.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Paragraph, end_paragraph - 1)
            if actual_moved_end < end_paragraph - 1:
                return {
                    "success": False,
                    "start_paragraph": start_paragraph,
                    "end_paragraph": end_paragraph,
                    "selected_text": "",
                    "error": f"End paragraph out of range: requested paragraph {end_paragraph}, but the document has only {actual_moved_end + 1} paragraphs"
                }

            start_phys = start_paragraph - 1
            end_phys = end_paragraph - 1

        target_range = text_pattern.DocumentRange.Clone()
        target_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Paragraph, start_phys)
        target_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, target_range, TextPatternRangeEndpoint_Start)

        next_par_range = text_pattern.DocumentRange.Clone()
        next_par_range.MoveEndpointByUnit(TextPatternRangeEndpoint_Start, TextUnit_Paragraph, end_phys + 1)
        next_par_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, next_par_range, TextPatternRangeEndpoint_Start)

        target_range.MoveEndpointByRange(TextPatternRangeEndpoint_End, next_par_range, TextPatternRangeEndpoint_Start)

        selected_text = target_range.GetText(-1).rstrip('\r\n')
        target_range.Select()

        return {
            "success": True,
            "start_paragraph": start_paragraph,  # When non_empty=True, this means the Nth non-empty paragraph
            "end_paragraph": end_paragraph,
            "selected_text": selected_text,
            "error": None
        }

    except Exception as e:
        return {
            "success": False,
            "start_paragraph": start_paragraph,
            "end_paragraph": end_paragraph if end_paragraph is not None else start_paragraph,
            "selected_text": "",
            "error": f"Error while selecting paragraphs {start_paragraph}-{end_paragraph if end_paragraph is not None else start_paragraph}: {str(e)}"
        }



#%%
# Scrollbar handling  Later, scroll_to_percent should be enhanced with return values, logging, exception handling, etc.
def scroll_to_percent(
        ctrl: UIAWrapper,
        horiz_percent: Optional[float] = -1,
        vert_percent: Optional[float] = -1,
        max_parent_hops: int = 4,
) -> None:
    """
    Move the scrollbar / scrollable container to the specified percentage.
    ─ horiz_percent, vert_percent: 0–100; -1 or None means "leave unchanged"
    ─ max_parent_hops: maximum number of levels to walk upward when searching for a scrollable container
    Raises RuntimeError if precise scrolling is not possible.
    """
    hp, vp = _clamp_percent(horiz_percent), _clamp_percent(vert_percent)

    # Does the target control itself support ScrollPattern?
    if _try_scrollpattern(ctrl, hp, vp):
        return

    # Is the target control the scrollbar itself?
    ok_h = _try_rangevalue(ctrl, 'h', hp)
    ok_v = _try_rangevalue(ctrl, 'v', vp)
    if ok_h and ok_v:
        return

    # Walk upward to find the nearest ancestor that supports ScrollPattern
    parent = ctrl.parent()
    hops = 0
    while parent and hops < max_parent_hops:
        if _try_scrollpattern(parent, hp, vp):
            return
        parent = parent.parent()
        hops += 1

    # 4️⃣ Fallback: only perform top / bottom scrolling
    if vp in (0, 100):
        try:
            ctrl.scroll('begin' if vp == 0 else 'end')
            return
        except Exception:
            pass

    raise RuntimeError("Unable to precisely scroll this control to the specified percentage; "
                       "please consider direction issues or fall back to mouse / keyboard simulation.")

def _clamp_percent(p: Optional[float]) -> float:
    """None → -1, valid range [0,100]; otherwise raise ValueError."""
    if p is None or p == -1:
        return -1.0
    if not (0 <= p <= 100):
        raise ValueError("percent must be ∈[0,100] or -1/None")
    return float(p)

def _try_scrollpattern(ctrl: UIAWrapper, hp: float, vp: float) -> bool:
    try:
        iface = ctrl.iface_scroll  # Raises NoPatternInterfaceError if missing
        iface.SetScrollPercent(hp, vp)
        return True
    except (AttributeError, COMError, NoPatternInterfaceError):
        return False

def _try_rangevalue(ctrl: UIAWrapper, axis: str, percent: float) -> bool:
    """axis='h' or 'v'; percent cannot be -1."""
    if percent < 0:
        return True  # -1 == unchanged
    try:
        rv = ctrl.iface_range_value
    except (AttributeError, COMError, NoPatternInterfaceError):
        return False

    # Scrollbars usually have only one dimension; checking OrientationProperty would be more rigorous,
    # but here we use a simple heuristic: width > height ≈ horizontal, otherwise vertical
    bounds = ctrl.rectangle()
    is_vert = bounds.height() > bounds.width()
    if axis == 'h' and is_vert:
        return False  # Not a horizontal bar
    if axis == 'v' and not is_vert:
        return False  # Not a vertical bar

    lo, hi = rv.CurrentMinimum, rv.CurrentMaximum
    value = lo + (hi - lo) * percent / 100.0
    try:
        rv.SetValue(value)
        return True
    except COMError:
        return False

#%%


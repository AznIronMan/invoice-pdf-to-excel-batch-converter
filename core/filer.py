import os
import platform
import tkinter as tk


from tkinter import filedialog
from typing import Optional, Union, Tuple


from core.logger import zlog as log


Detailed_Result = Tuple[
    bool,
    Optional[
        Union[str, tuple[float, float], tuple[int, int], tuple[str, str]]
    ],
]


def is_gui_available():
    try:
        root = tk.Tk()
        root.withdraw()
        return True
    except:
        return False


def path_to_module(path: str) -> str:
    try:
        relative_path = os.path.relpath(path, os.getcwd())

        if relative_path.startswith(".\\") or relative_path.startswith("./"):
            relative_path = relative_path[2:]
        relative_path, _ = os.path.splitext(relative_path)

        return relative_path.replace("\\", ".").replace("/", ".")
    except Exception as e:
        error_message = f"Failed to convert path to module. Exception: {e}"
        log(error_message, "WARNING")
        return path


def select_folder(title: str = "Select a folder to process"):
    current_working_directory = os.getcwd()
    os_type = platform.system()

    if os_type in ['Windows', 'Darwin'] or (os_type == 'Linux' and is_gui_available()):
        root = tk.Tk()
        root.withdraw()
        folder_selected = filedialog.askdirectory(
            title=title,
            initialdir=current_working_directory)
        if not folder_selected:
            folder_selected = current_working_directory
        return folder_selected
    else:
        if title == "Select a folder to process":
            title = title.replace("Select", "Enter")
        folder_selected = input(
            f"{title} (default: {current_working_directory}): ")
        if not folder_selected:
            folder_selected = current_working_directory
        return folder_selected

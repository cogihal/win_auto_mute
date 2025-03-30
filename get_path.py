"""
Get the folder path of the running Python script or freezed exe file by pyinstaller.
"""

from pathlib import Path
import sys


def get_runtime_folder_path() -> Path:
    """
    Get the runtime folder path.

    - If the script is run as a script, it will return the script folder path.
    - If the script is run as a freezed exe file by pyinstaller, it will return temporary folder path.
    """
    runtime_path = Path(__file__).resolve().parent
    return runtime_path


def get_script_folder_path() -> Path:
    """
    Get the script folder path.

    - If the script is run as a script, it will return the script folder path.
    - If the script is run as a freezed exe file by pyinstaller, it will return the exe file folder path.
    """
    script_path = Path(sys.argv[0]).resolve().parent
    return script_path


def get_script_basename() -> str:
    """
    Get the script name without extension.

    - If the script is run as a script, it will return the script basename.
    - If the script is run as a freezed exe file by pyinstaller, it will return the exe file basename.
    """
    script_name = Path(sys.argv[0]).stem
    return script_name

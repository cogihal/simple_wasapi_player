import os
import sys


def get_module_path() -> str:
    """
    Return the path of the main module.
    """
    full_path_name = os.path.abspath(sys.argv[0])
    folder_name = os.path.dirname(full_path_name)
    return folder_name



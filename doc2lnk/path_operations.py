import wslPath
import platform
import os

def convert_to_os_specific_path(directory):
    try:
        return wslPath.to_windows(directory)
    except Exception:
        if platform.system() == 'Linux':
            return wslPath.to_posix(directory)

        return directory

def get_directory(file_path):
    directory = os.path.dirname(file_path)

    if not directory:
        directory = os.getcwd()

    return convert_to_os_specific_path(directory)
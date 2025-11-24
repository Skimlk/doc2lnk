from pylnk3 import PathSegmentEntry, TYPE_FOLDER, TYPE_FILE
from datetime import datetime
from . import path_operations
import ntpath
import os

# Modify PathSegmentEntry's create_for_path to check paths based on OS being used
def create_for_path_os_agnostic(cls, path):
    entry = cls()
    entry.type = os.path.isdir(path_operations.convert_to_os_specific_path(path)) and TYPE_FOLDER or TYPE_FILE
    try:
        st = os.stat(path_operations.convert_to_os_specific_path(path))
        entry.file_size = st.st_size
        entry.modified = datetime.fromtimestamp(st.st_mtime)
        entry.created = datetime.fromtimestamp(st.st_ctime)
        entry.accessed = datetime.fromtimestamp(st.st_atime)
    except FileNotFoundError:
        now = datetime.now()
        entry.file_size = 0
        entry.modified = now
        entry.created = now
        entry.accessed = now
    entry.short_name = ntpath.split(path)[1]
    entry.full_name = entry.short_name
    return entry

PathSegmentEntry.create_for_path = classmethod(create_for_path_os_agnostic)
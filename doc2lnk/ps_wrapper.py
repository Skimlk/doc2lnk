import textwrap

EDIT_LNKDOCUMENT_PATH = 'doc2lnk/Edit-LnkDocument.ps1'

def wrap_powershell_script(inner_script, document_name, delimiter):
    return textwrap.dedent(rf'''
        {open(EDIT_LNKDOCUMENT_PATH, "r", encoding="utf-8-sig").read()}

        {inner_script};
        Edit-LnkDocument '{document_name}' '{delimiter}';
    ''')
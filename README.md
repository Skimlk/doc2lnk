# doc2lnk

## Overview
`doc2lnk` is a Python tool that converts standard document files into `.lnk` shortcut files. The generated `.lnk` files can execute any PowerShell script provided as an argument while still allowing the original document to be opened and saved normally. It does this by appending the content of the original document to the shortcut file.

## Requirements
- Python 3.13 or later

## Installation

### 1. Create and activate a virtual environment
```bash
python3 -m venv venv
```

**Linux/macOS:**
```bash
source venv/bin/activate
```

**Windows:**
```powershell
.\venv\Scripts\activate
```

### 2. Install dependencies
```bash
pip3 install -r requirements.txt
```

## Usage
```bash
# Run with a script file
python3 -m doc2lnk <document> <script>

# Run with a script string
python3 -m doc2lnk <document> --string <VALUE>
```

### Options
- `-o, --overwrite`: Overwrite the original document file with the created shortcut file.
- `-r, --restore`: Restore the original document after executing previously created `.lnk` file.
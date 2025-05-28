# Dynamic File Generator

This Python script generates arbitrary, valid files in various formats (DOCX, XLSX, PPTX, PDF, PST, ZIP) to a specified size. It ensures each file is correctly structured so that it opens in its respective application.

## Supported Formats

* **DOCX** (Microsoft Word)
* **XLSX** (Microsoft Excel)
* **PPTX** (Microsoft PowerPoint)
* **PDF** (any PDF reader)
* **PST** (Outlook Personal Storage Table)
* **ZIP** (Archive)

## Requirements

All dependencies are listed in [requirements.txt](requirements.txt). You can install them in a virtual environment:

```bash
python -m venv venv
# macOS/Linux
source venv/bin/activate
# Windows
powershell -ep bypass
venv\Scripts\activate

pip install -r requirements.txt
```

- Ensure you have Python 3.7 and above  

## Usage

```bash
usage: script.py [-h] [-f {docx,xlsx,pptx,pdf,pst,zip}] size [output]

positional arguments:
  size            target size (e.g. 150KB, 2.5MB, or 10 for 10MB)
  output          optional filename (default: <format>_<size>.<ext>)

options:
  -h, --help      show this help message and exit
  -f, --format    file format to produce (docx, xlsx, pptx, pdf, pst, zip) (default: docx)
```

### Examples

#### Generating in KB

```
# 100 KB Word document
python script.py 100KB -f docx small.docx

# 512 KB PDF
python script.py 512KB -f pdf report.pdf

# 8 KB ZIP archive
python script.py 8KB -f zip archive.zip
```

#### Generating in MB

```
# 1.5 MB Excel workbook
python script.py 1.5MB -f xlsx workbook.xlsx

# 10 MB PowerPoint deck (default name: pptx_10.pptx)
python script.py 10 -f pptx

# 5 MB Outlook PST mailbox
python script.py 5MB -f pst mailbox.pst
```

## Notes

* **PST generation** requires [Aspose.Email-for-Python-via-NET](https://pypi.org/project/Aspose.Email-for-Python-via-NET/) and a Windows environment.
* All other formats work cross-platform.

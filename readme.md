# Dynamic File Generator Script

This Python script generates large, valid files in various formats (DOCX, XLSX, PPTX, PDF, PST, ZIP) to a specified size. It ensures each file is correctly structured so that it opens in its respective application.

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

## Usage

```bash
python script.py SIZE_MB -f FORMAT
```

* `SIZE_MB`: Target file size in megabytes (integer).
* `-f FORMAT`: File format (`docx`, `xlsx`, `pptx`, `pdf`, `pst`, `zip`). Defaults to `docx`.
* `OUTPUT`: (optional) Output filename. Defaults to `output.<format>`.

### Examples

* Generate a 50 MB Word document:

  ```bash
  python script.py 50 -f docx
  ```

* Generate a 100 MB Excel spreadsheet:

  ```bash
  python script.py 100 -f xlsx
  ```

* Generate a 10 MB PowerPoint presentation:

  ```bash
  python script.py 10 -f pptx
  ```

* Generate a 5 MB PDF file:

  ```bash
  python script.py 5 -f pdf
  ```

* Generate a 7 MB Outlook PST file:

  ```bash
  python script.py 7 -f pst
  ```

* Generate a 20 MB ZIP archive:

  ```bash
  python script.py 20 -f zip
  ```

## Notes

* **PST generation** requires [Aspose.Email-for-Python-via-NET](https://pypi.org/project/Aspose.Email-for-Python-via-NET/) and a Windows environment.
* All other formats work cross-platform.

## License

MIT License. Feel free to modify and distribute.

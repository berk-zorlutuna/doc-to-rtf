# doc-to-rtf

Convert Microsoft Word `.doc` and `.docx` files to `.rtf` with Python, using LibreOffice in headless mode (keeps formatting, tables and images).

## Requirements
- Python 3.9+
- [LibreOffice](https://www.libreoffice.org/) installed (`soffice` on PATH, or set `SOFFICE_PATH`)

## Usage
```bash
python doc2rtf.py report.docx                 # -> report.rtf next to the input
python doc2rtf.py a.doc b.docx -o converted   # into a folder
python doc2rtf.py my_docs -r -o converted     # a whole folder, recursively
python doc2rtf.py my_docs --no-overwrite      # skip existing .rtf files
```

From Python:
```python
from doc2rtf import convert_file
convert_file("report.doc", "out_dir")
```

## Tests
```bash
pip install -r requirements-dev.txt
pytest
```

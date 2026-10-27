# Receipt Processor

Reads JPEG, PNG, HEIC, and TIFF receipts from `Processing/`. macOS Vision OCR
recognizes English and Traditional Chinese locally. DeepSeek `deepseek-chat`
extracts the shop, purchase category, transaction date, total, and currency from
the OCR text. If the shop name or total is unknown, the receipt is resized to a
maximum 1600-pixel edge and sent to DeepSeek `deepseek-flash` for image-based
extraction. Results are appended to `Processed/YYYY-MM-DD.xlsx`, using the
current date for the workbook filename. The `Receipt Date` column contains the
date read from each receipt, with the image's file-modified date used if it
cannot be read. The model name and source image are also recorded. Successfully
processed images are moved to `Processed/`.
Each moved image receives a unique numeric suffix before its extension, and the
renamed filename is stored in the workbook's `Source` column.

## macOS

Install dependencies:

```sh
python3 -m venv .venv
.venv/bin/pip install -r requirements.txt
```

Set the API key and run:

```sh
export DEEPSEEK_API_KEY="your-api-key"
.venv/bin/python receipt_processor.py
```

## Windows

Install Python 3.10 or later and Tesseract OCR 5. During Tesseract setup,
include the English (`eng`) and Traditional Chinese (`chi_tra`) language data
and ensure the Tesseract installation directory is on `PATH`.

From PowerShell, install the Python dependencies:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Set the API key and run the processor:

```powershell
$env:DEEPSEEK_API_KEY = "your-api-key"
.\.venv\Scripts\python.exe receipt_processor.py
```

Put receipt images in `Processing/`. Supported formats are JPEG, PNG, HEIC,
and TIFF. Processed images and the Excel workbook are written to `Processed/`.

.DS_Store
.venv/
__pycache__/
*.py[cod]
Processed/*
Processing/*

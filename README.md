# Custom Zonal OCR Project

A Python script that extracts structured data from medical record PDFs using zonal OCR — cropping to specific x/y coordinates on the page — and writes the results to an Excel file.

Designed for CMS-1500 and similarly formatted medical billing forms where fields always appear at fixed positions across all documents.

---

## Features

- **Zonal OCR** — crops to exact pixel coordinates per field rather than scanning the full page, improving accuracy and speed
- **cv2 pre-processing** — denoising, adaptive thresholding, deskewing, and sharpening before OCR
- **PSM fallback** — automatically tries multiple Tesseract page segmentation modes and picks the richest result
- **Field cleaners** — post-processes raw OCR output to normalize names, dates (`MM/DD/YYYY`), and currency values
- **Excel output** — writes results to a formatted `.xlsx` file, overwriting it on each run
- **Selective extraction** — choose which fields to extract per run via `--zones`
- **Calibration tool** — interactive window to find pixel coordinates for any field on your PDF

---

## Requirements

### Python dependencies

```bash
pip install -r requirements.txt
```

### System dependencies

| Platform | Command |
|----------|---------|
| Ubuntu/Debian | `sudo apt install tesseract-ocr poppler-utils` |
| macOS | `brew install tesseract poppler` |
| Windows | Install [Tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and [Poppler](https://github.com/oschwartz10612/poppler-windows/releases), then add both to PATH |

On Windows you may also need to set the Tesseract path explicitly at the top of the script:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

---

## Usage

### Basic — process all PDFs in the `./pdfs/` folder

```bash
python extract_medical_records.py
```

### Specify a folder or individual files

```bash
python extract_medical_records.py path/to/folder
python extract_medical_records.py file1.pdf file2.pdf
```

### Extract only specific fields

```bash
# Only name and total charge
python extract_medical_records.py --zones name,total_charge

# Only date of service
python extract_medical_records.py --zones date_of_service
```

Available zone keys: `name`, `date_of_service`, `total_charge`

### Calibration — find pixel coordinates for your PDF

```bash
python extract_medical_records.py --calibrate path/to/sample.pdf
```

Opens an interactive window with a live crosshair. Hover over any field to read its `x, y` pixel coordinates at 300 DPI. Existing zones are drawn in green.

---

## Configuration

All configuration is at the top of `extract_medical_records.py`.

### Zone coordinates

```python
ZONES: dict[str, tuple[int, int, int, int]] = {
    #           x     y     w     h
    "name":            (150, 210, 600,  55),
    "date_of_service": (150, 310, 400,  55),
    "total_charge":    (150, 410, 300,  55),
}
```

Each value is `(x, y, width, height)` in pixels at 300 DPI, measured from the top-left corner of page 1. Use `--calibrate` to find the correct values for your specific form.

### Adding a new field

1. Add an entry to `ZONES` and `ZONE_PSM` with its coordinates and preferred Tesseract PSM mode
2. Add a matching entry to `ZONE_FIELDS` with its Excel column label and cleaner (`"name"`, `"date"`, or `"charge"`)
3. If needed, write a new cleaner function and register it in the `CLEANERS` dict

### Debug mode

```python
DEBUG = False  # Set to True to save cropped zone images and print raw OCR text
```

When enabled, saves `debug_<field>.png` for each zone and prints raw Tesseract output for every PSM attempt — useful for diagnosing extraction issues.

---

## Output

Results are saved to `medical_records_output.xlsx` in the working directory. The file is overwritten on each run. If the file is open in Excel, the script will prompt you to close it before retrying.

| Column | Description |
|--------|-------------|
| File | Source PDF filename |
| Name | Patient name |
| Date of Service | Date formatted as `MM/DD/YYYY` |
| Total Charge | Billing amount |

---

## Project Structure

```
.
├── extract_medical_records.py  # Main script
├── requirements.txt
├── README.md
├── pdfs/                       # Place input PDFs here
└── medical_records_output.xlsx # Generated on each run
```

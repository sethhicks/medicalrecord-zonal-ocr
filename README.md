# Custom Zonal OCR Project

A Python script that extracts structured data from medical record PDFs (CMS-1500 and UB-04) using zonal OCR with template-based alignment, and writes the results to an Excel file.

Each page of the input PDF is automatically identified as a CMS or UB form, aligned to a reference template to correct for scan variation, and then OCR'd at specific field coordinates.

---

## Features

- **Automatic form detection** — identifies CMS-1500 and UB-04 forms per page without manual input
- **Template alignment** — ORB feature matching + homography corrects translation, rotation, scale, and perspective distortion between scans
- **Confidence scoring** — low-confidence alignments are flagged in the console and highlighted yellow in Excel
- **Zonal OCR** — crops to exact pixel coordinates per field for accuracy and speed
- **cv2 pre-processing** — denoising, Otsu thresholding, deskewing, and optional sharpening
- **PSM fallback** — tries multiple Tesseract page segmentation modes and picks the best result
- **Field cleaners** — normalises names, dates (`MM/DD/YYYY`), and currency values from raw OCR output
- **Multi-page support** — processes every page of a single PDF, one Excel row per page
- **Selective extraction** — choose which fields to extract via `--zones`
- **Calibration tool** — interactive window to find pixel coordinates for any field

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

On Windows you may also need to set the Tesseract path explicitly near the top of the script:

```python
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
```

---

## Usage

### Basic — process all pages of a PDF

```bash
python extract_medical_records.py path/to/claims.pdf
```

### Extract only specific fields

```bash
python extract_medical_records.py claims.pdf --zones name,total_charge
python extract_medical_records.py claims.pdf --zones date_of_service
```

Available zone keys: `name`, `date_of_service`, `total_charge`, `dob`, `cpt_hcpcs`
(Note: `cpt_hcpcs` is only available for CMS forms)

### Force a specific form type

```bash
python extract_medical_records.py claims.pdf --form cms
python extract_medical_records.py claims.pdf --form ub
```

By default the form type is auto-detected per page.

### Calibration — find pixel coordinates for a form

```bash
python extract_medical_records.py --calibrate templates/template_cms.png --form cms
python extract_medical_records.py --calibrate templates/template_ub.png --form ub
```

Opens an interactive window. Hover over any field to read its `x, y` pixel coordinates at 300 DPI. OCR zones are drawn in green, detection regions in yellow. Left-click to print coordinates to the console.

### Save a template from a PDF

```bash
python extract_medical_records.py --save-template pdfs/sample_cms.pdf --form cms
python extract_medical_records.py --save-template pdfs/sample_ub.pdf --form ub
```

Or simply drop a clean PNG scan directly into the `templates/` folder.

---

## Configuration

All configuration is near the top of `extract_medical_records.py`.

### Zone coordinates

```python
ZONES_1500: dict[str, tuple[int, int, int, int]] = {
    #           x     y     w     h
    "name":            (85,  535, 830, 60),
    "date_of_service": (82, 2200, 260, 90),
    ...
}
```

Each value is `(x, y, width, height)` in pixels at 300 DPI, measured from the top-left corner of the **template image**. Use `--calibrate` to find the correct values.

### Confidence threshold

```python
LOW_CONFIDENCE_THRESHOLD = 50  # inliers below this → flagged as low confidence
```

Raise this value to flag more results for review, lower it to be more permissive.

### Adding a new zone

1. Add an entry to `ZONES_CMS` and/or `ZONES_UB` with its coordinates
2. Add a matching entry to `ZONE_PSM` with the preferred Tesseract PSM mode
3. Add a matching entry to `ZONE_FIELDS` with its Excel column label and cleaner key (`"name"`, `"date"`, `"charge"`, or `"cpt"`)
4. If a new cleaner is needed, write it and register it in the `CLEANERS` dict

### Debug mode

```python
DEBUG = False  # Set to True to enable
```

When enabled, saves cropped zone images (`debug_<form>_<field>.png`) and prints raw Tesseract output for every PSM attempt — useful for diagnosing extraction issues.

---

## Output

Results are saved to `medical_records_output.xlsx`, overwriting on each run. If the file is open in Excel the script will prompt you to close it.

| Column | Description |
|--------|-------------|
| File | Source PDF filename |
| Page | Page number within the PDF |
| Form Type | `CMS` or `UB` |
| Confidence | `OK` or `LOW` — low-confidence rows are highlighted yellow |
| Name | Patient name |
| Date of Service | Formatted as `MM/DD/YYYY` |
| Total Charge | Whole-dollar billing amount |
| Date of Birth | Formatted as `MM/DD/YYYY` |
| CPT/HCPCS | 5-digit procedure code (CMS forms only) |

---

## Project Structure

```
.
├── extract_medical_records.py   # Main script
├── requirements.txt
├── README.md
├── templates/
│   ├── template_cms.png         # Reference scan for CMS-1500 alignment
│   └── template_ub.png          # Reference scan for UB-04 alignment
├── pdfs/                        # Place input PDFs here
└── medical_records_output.xlsx  # Generated on each run
```

"""
Medical Record PDF Extractor
-----------------------------
Extracts Name, Date of Service, and Total Charge from
identically-formatted medical record PDFs using zonal OCR
(crops to specific x/y coordinates per field) and writes results to Excel.

Usage:
    python extract_medical_records.py                        # scans ./pdfs/
    python extract_medical_records.py /path/to/pdf/folder    # custom folder
    python extract_medical_records.py file1.pdf file2.pdf    # specific files

Output:
    medical_records_output.xlsx  (created/appended in the working directory)

Dependencies:
    pip install pdf2image pytesseract opencv-python numpy openpyxl
    System: tesseract-ocr and poppler-utils must also be installed.
      Ubuntu/Debian: sudo apt install tesseract-ocr poppler-utils
      macOS:         brew install tesseract poppler
      Windows:       install Tesseract from https://github.com/UB-Mannheim/tesseract/wiki
                     install Poppler from https://github.com/oschwartz10612/poppler-windows/releases
                     and add both to PATH
"""

import sys
import os
import glob
import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path
import openpyxl
import re
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

#Create zonal OCR zones
ZONES: dict[str, tuple[int, int, int, int]] = {
    #           x     y     w     h
    "name":            (78, 506, 830,  100),
    "date_of_service": (92, 2130, 250, 85),
    "total_charge":    (1550, 2725, 280,  75),
    "dob":             (920, 550, 310,  50),
    "cpt_hcpcs":       (790, 2140, 190,  75),
}

# Tesseract page-segmentation mode per zone.
# 7 = single line  |  6 = uniform block of text  |  8 = single word
ZONE_PSM: dict[str, str] = {
    "name":            "--psm 7",
    "date_of_service": "--psm 7",
    "total_charge":    "--psm 7",
    "dob":             "--psm 7",
    "cpt_hcpcs":       "--psm 8",  # single word — 5-digit code
}
 
# Maps each zone key to its Excel column label and cleaner function name.
# Add a new entry here whenever you add a zone to ZONES/ZONE_PSM above.
ZONE_FIELDS: dict[str, dict] = {
    "name":            {"label": "Name",            "cleaner": "name"},
    "date_of_service": {"label": "Date of Service", "cleaner": "date"},
    "total_charge":    {"label": "Total Charge",    "cleaner": "charge"},
    "dob":             {"label": "Date of Birth",   "cleaner": "date"},
    "cpt_hcpcs":       {"label": "CPT/HCPCS",       "cleaner": "cpt"},
}
 
OUTPUT_FILE = "medical_records_output.xlsx"
DEBUG = False  # Set to True to save cropped zone images and print raw OCR text
 
# ---------------------------------------------------------------------------
# Image pre-processing
# ---------------------------------------------------------------------------
 
def preprocess_image(gray: np.ndarray) -> np.ndarray:
    """Denoise → adaptive threshold → deskew → sharpen."""
    denoised = cv2.fastNlMeansDenoising(gray, h=10)
 
    binary = cv2.adaptiveThreshold(
        denoised, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY,
        blockSize=31,
        C=15,
    )
 
    # Deskew (small angles only)
    coords = np.column_stack(np.where(binary < 255))
    if len(coords) > 0:
        angle = cv2.minAreaRect(coords)[-1]
        angle = -(90 + angle) if angle < -45 else -angle
        if abs(angle) < 10:
            h, w = binary.shape
            M = cv2.getRotationMatrix2D((w / 2, h / 2), angle, 1.0)
            binary = cv2.warpAffine(binary, M, (w, h),
                                    flags=cv2.INTER_CUBIC,
                                    borderMode=cv2.BORDER_REPLICATE)
 
    kernel = np.array([[0, -1, 0], [-1, 5, -1], [0, -1, 0]])
    return cv2.filter2D(binary, -1, kernel)
 
# ---------------------------------------------------------------------------
# Zonal OCR
# ---------------------------------------------------------------------------
 
def crop_zone(img: np.ndarray, zone: tuple[int, int, int, int]) -> np.ndarray:
    """Crop a region from a full-page cv2 image. zone = (x, y, w, h)."""
    x, y, w, h = zone
    img_h, img_w = img.shape[:2]
    x1, y1 = max(x, 0), max(y, 0)
    x2, y2 = min(x + w, img_w), min(y + h, img_h)
    return img[y1:y2, x1:x2]
 
 
def ocr_zone(img_bgr: np.ndarray, zone: tuple[int, int, int, int], psm: str,
             field_name: str = "") -> str:
    """Crop → grayscale → preprocess → OCR a single zone.
    Falls back through PSM modes 6, 4, and 3 if the primary PSM returns nothing.
    """
    cropped   = crop_zone(img_bgr, zone)
    gray      = cv2.cvtColor(cropped, cv2.COLOR_BGR2GRAY)
    processed = preprocess_image(gray)
 
    if DEBUG and field_name:
        debug_path = f"debug_{field_name}.png"
        cv2.imwrite(debug_path, processed)
        print(f"    [debug] saved cropped zone → {debug_path}")
 
    # Try all PSM modes and return the result with the most digit/letter content
    candidates = []
    for mode in [psm, "--psm 6", "--psm 4", "--psm 3"]:
        text = pytesseract.image_to_string(processed, config=mode).strip()
        if DEBUG:
            print(f"    [debug] {field_name} psm={mode!r:12s} → {text!r}")
        if text:
            candidates.append(text)
    if not candidates:
        return ""
    # Pick the candidate with the most alphanumeric characters
    return max(candidates, key=lambda t: len(re.findall(r'[A-Za-z0-9]', t)))
 
 
def clean_name(raw: str) -> str:
    """Extract the patient name from OCR output that may include a form label above it.
    Strategy: drop lines containing label keywords or box-drawing characters,
    then return the last remaining line (the actual data row).
    """
    LABEL_KEYWORDS = ("patient", "name", "first", "middle", "initial", "last")
    lines = [l.strip() for l in raw.splitlines() if l.strip()]
    data_lines = [
        l for l in lines
        if not any(kw in l.lower() for kw in LABEL_KEYWORDS)
        and not re.match(r'^[\|\-\+\s]+$', l)
    ]
    if data_lines:
        return re.sub(r'^\|?\s*|\s*\|?$', '', data_lines[-1]).strip()
    return re.sub(r'^\|?\s*|\s*\|?$', '', lines[-1]).strip() if lines else ""
 
 
def clean_date(raw: str) -> str:
    """Normalise OCR noise in date values.
    Handles separators misread as | or : between MM, DD, YY(YY) parts.
    Always outputs MM/DD/YYYY, expanding 2-digit years (e.g. 25 -> 2025).
    """
    # Extract all digit groups (ignores |, :, spaces, and other noise)
    parts = re.findall(r'\d+', raw)
    if len(parts) >= 3:
        mm, dd, yy = parts[0], parts[1], parts[2]
        if len(yy) == 2:
            yy = "20" + yy  # expand YY -> YYYY (e.g. 25 -> 2025)
        return f"{mm.zfill(2)}/{dd.zfill(2)}/{yy}"
    elif len(parts) == 2:
        return f"{parts[0].zfill(2)}/{parts[1].zfill(2)}"
    elif len(parts) == 1:
        return parts[0]
    return raw.strip()
 
 
def clean_charge(raw: str) -> str:
    """Normalise OCR noise in currency values: '361 :00' → '361.00'."""
    # Keep only digits, dots, commas, and dollar signs
    cleaned = re.sub(r"[^\d.,$]", "", raw.replace(" :", ".").replace(" ", ""))
    return cleaned
 
 
def clean_cpt(raw: str) -> str:
    """Extract a 5-digit CPT/HCPCS code from OCR output."""
    # First try to find an explicit 5-digit sequence
    match = re.search(r'\b(\d{5})\b', raw)
    if match:
        return match.group(1)
    # Fallback: strip all non-digits and zero-pad or truncate to 5
    digits = re.sub(r'\D', '', raw)
    if digits:
        return digits[:5].zfill(5)
    return raw.strip()
 
 
CLEANERS = {
    "name":   clean_name,
    "date":   clean_date,
    "charge": clean_charge,
    "cpt":    clean_cpt,
}
 
 
def extract_fields(pdf_path: str, active_zones: list[str]) -> dict:
    """Convert page 1 to an image and OCR the requested zones only."""
    pages = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
    img_bgr = cv2.cvtColor(np.array(pages[0]), cv2.COLOR_RGB2BGR)
 
    record = {"File": os.path.basename(pdf_path)}
    for zone_key in active_zones:
        meta  = ZONE_FIELDS[zone_key]
        raw   = ocr_zone(img_bgr, ZONES[zone_key], ZONE_PSM[zone_key], zone_key)
        value = CLEANERS[meta["cleaner"]](raw)
        record[meta["label"]] = value
    return record
 
# ---------------------------------------------------------------------------
# Calibration helper  (python extract_medical_records.py --calibrate file.pdf)
# ---------------------------------------------------------------------------
 
def run_calibration(pdf_path: str):
    """Open an interactive window showing pixel coordinates on hover."""
    print(f"Calibrating against: {pdf_path}")
    print("Hover over each field to read its (x, y) pixel coordinates.")
    print("Press 'q' or Escape to quit.\n")
 
    pages = convert_from_path(pdf_path, dpi=300, first_page=1, last_page=1)
    img   = cv2.cvtColor(np.array(pages[0]), cv2.COLOR_RGB2BGR)
 
    # Draw existing zones in green so you can see where they currently land
    for name, (x, y, w, h) in ZONES.items():
        cv2.rectangle(img, (x, y), (x + w, y + h), (0, 200, 0), 2)
        cv2.putText(img, name, (x + 4, y + 20),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 200, 0), 2)
 
    display = img.copy()
    coords_label = ["x=0, y=0"]
 
    def on_mouse(event, x, y, flags, param):
        overlay = img.copy()
        # crosshair
        cv2.line(overlay, (x, 0), (x, overlay.shape[0]), (0, 0, 255), 1)
        cv2.line(overlay, (0, y), (overlay.shape[1], y), (0, 0, 255), 1)
        label = f"x={x}, y={y}"
        cv2.putText(overlay, label, (x + 8, y - 8),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 0, 255), 2)
        cv2.imshow("Calibration — press Q to quit", overlay)
 
    win = "Calibration — press Q to quit"
    # Scale down for display if the image is very tall
    scale = min(1.0, 1200 / img.shape[0])
    if scale < 1.0:
        disp = cv2.resize(img, (int(img.shape[1] * scale), int(img.shape[0] * scale)))
    else:
        disp = img.copy()
 
    cv2.namedWindow(win, cv2.WINDOW_NORMAL)
    cv2.imshow(win, disp)
    cv2.setMouseCallback(win, on_mouse)
 
    while True:
        key = cv2.waitKey(20) & 0xFF
        if key in (ord('q'), 27):
            break
    cv2.destroyAllWindows()
    print("Calibration closed.")
 
# ---------------------------------------------------------------------------
# Excel writing
# ---------------------------------------------------------------------------
 
HEADER_FILL  = PatternFill("solid", start_color="1F4E79")
HEADER_FONT  = Font(bold=True, color="FFFFFF", name="Arial", size=11)
DATA_FONT    = Font(name="Arial", size=10)
CENTER       = Alignment(horizontal="center", vertical="center")
LEFT         = Alignment(horizontal="left",   vertical="center")
THIN         = Side(style="thin", color="CCCCCC")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
# Base column widths — extended dynamically from ZONE_FIELDS
_BASE_COL_WIDTHS = {"File": 30, "Name": 25, "Date of Service": 18, "Total Charge": 16, "Date of Birth": 18, "CPT/HCPCS": 14}
 
def build_header(active_zones: list[str]) -> list[str]:
    """Return the ordered Excel column list for the given active zones."""
    return ["File"] + [ZONE_FIELDS[z]["label"] for z in active_zones]
 
def col_width(label: str) -> int:
    return _BASE_COL_WIDTHS.get(label, 20)
 
 
def setup_workbook(header: list[str]):
    if os.path.exists(OUTPUT_FILE):
        try:
            os.remove(OUTPUT_FILE)
            print(f"  Overwriting existing file: {OUTPUT_FILE}")
        except PermissionError:
            print(f"\n❌  Cannot overwrite {OUTPUT_FILE} — the file is open in another program.")
            print("    Please close it in Excel and run the script again.")
            sys.exit(1)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Medical Records"
    for col_idx, col_name in enumerate(header, start=1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = CENTER
        cell.border    = BORDER
        ws.column_dimensions[cell.column_letter].width = col_width(col_name)
        ws.row_dimensions[1].height = 20
        ws.freeze_panes = "A2"
    return wb, ws
 
 
def append_row(ws, record: dict, row: int, header: list[str]):
    alt_fill = PatternFill("solid", start_color="EBF3FB") if row % 2 == 0 else None
    for col_idx, col_name in enumerate(header, start=1):
        cell = ws.cell(row=row, column=col_idx, value=record.get(col_name, ""))
        cell.font      = DATA_FONT
        cell.border    = BORDER
        cell.alignment = CENTER if col_name in ("Date of Service", "Total Charge") else LEFT
        if alt_fill:
            cell.fill = alt_fill
 
 
def save_to_excel(records: list[dict], active_zones: list[str]):
    header = build_header(active_zones)
    wb, ws = setup_workbook(header)
    start_row = ws.max_row + 1
    for i, record in enumerate(records):
        append_row(ws, record, start_row + i, header)
    wb.save(OUTPUT_FILE)
    print(f"\n✅  Saved {len(records)} record(s) → {OUTPUT_FILE}")
 
# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
 
def collect_pdf_paths(args: list[str]) -> list[str]:
    if not args:
        folder = "pdfs"
        paths = sorted(glob.glob(os.path.join(folder, "*.pdf")))
        if not paths:
            print(f"No PDFs found in './{folder}/'. Pass paths as arguments or place PDFs there.")
            sys.exit(1)
        return paths
    paths = []
    for arg in args:
        if os.path.isdir(arg):
            paths.extend(sorted(glob.glob(os.path.join(arg, "*.pdf"))))
        elif arg.endswith(".pdf") and os.path.isfile(arg):
            paths.append(arg)
        else:
            print(f"  ⚠  Skipping unrecognised argument: {arg}")
    return paths
 
 
def parse_args(argv: list[str]):
    """Parse CLI arguments and return (pdf_args, active_zones).
 
    --zones name,total_charge   Comma-separated zone keys to extract.
                                Defaults to all zones in ZONE_FIELDS.
    --calibrate <pdf>           Open calibration window (handled separately).
 
    All other arguments are treated as PDF paths or folders.
    """
    import argparse
    parser = argparse.ArgumentParser(
        prog="extract_medical_records.py",
        description="Extract fields from medical record PDFs into Excel.",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "pdfs", nargs="*",
        help="PDF files or folders to process (default: ./pdfs/)",
    )
    parser.add_argument(
        "--calibrate", metavar="PDF",
        help="Open calibration window for the given PDF.",
    )
    parser.add_argument(
        "--zones", metavar="ZONE1,ZONE2,...",
        help=(
            "Comma-separated list of zone keys to extract.\n"
            f"Available: {', '.join(ZONE_FIELDS.keys())}\n"
            "Default: all zones."
        ),
    )
    args = parser.parse_args(argv)
 
    # Resolve active zones
    if args.zones:
        requested = [z.strip() for z in args.zones.split(",")]
        invalid = [z for z in requested if z not in ZONE_FIELDS]
        if invalid:
            print(f"❌  Unknown zone(s): {', '.join(invalid)}")
            print(f"   Available zones: {', '.join(ZONE_FIELDS.keys())}")
            sys.exit(1)
        active_zones = requested
    else:
        active_zones = list(ZONE_FIELDS.keys())
 
    return args, active_zones
 
 
def main():
    args, active_zones = parse_args(sys.argv[1:])
 
    if args.calibrate:
        run_calibration(args.calibrate)
        return
 
    print(f"Zones selected: {', '.join(active_zones)}\n")
 
    pdf_paths = collect_pdf_paths(args.pdfs)
    print(f"Processing {len(pdf_paths)} PDF(s)...\n")
 
    records = []
    for path in pdf_paths:
        print(f"  Reading: {path}")
        try:
            record = extract_fields(path, active_zones)
            records.append(record)
            for zone_key in active_zones:
                label = ZONE_FIELDS[zone_key]["label"]
                print(f"    {label:<20} {record.get(label) or '(not found)'}")
        except Exception as e:
            print(f"    ❌  Error reading {path}: {e}")
 
    if records:
        save_to_excel(records, active_zones)
    else:
        print("No records extracted.")
 
 
if __name__ == "__main__":
    main()
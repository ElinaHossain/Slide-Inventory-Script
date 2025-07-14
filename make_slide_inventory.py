#!/usr/bin/env python3
"""
make_slide_inventory.py
Creates an .xlsx inventory with drop-down lists and (optionally) auto‑filled
Format/Extent columns.

Usage:
    python make_slide_inventory.py <folder> <output.xlsx>
"""
import re, sys
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation

# regex helper to extract box, strip, and image numbers
def extract_key(p: Path):
    """
    Turn 'UAC099_30-02_13.tif' into (2, 13)
    So sorting happens in numeric order.
    """
    match = re.search(r"-(\d+)_0*(\d+)", p.stem)
    if match:
        return (int(match.group(1)), int(match.group(2)))
    return (999, 999)  # fallback if pattern doesn't match

# ---------- rules (edit as needed) ----------
PATTERN_MAP = {
    r"_bw_":  ("b&w slide",            "24 x 36mm (35mm/slides)"),
    r"_neg_": ("color negative",       "24 x 36mm (35mm/slides)"),
    r"4x6":   ("color photograph print","4 x 6 in (print)"),
    r"8x10":  ("color photograph print","8 x 10 in (print)"),
}
DEFAULT_FORMAT = "color slide"
DEFAULT_EXTENT = "24 x 36mm (35mm/slides)"

FMT_LIST = [
    "color slide", "b&w slide", "color negative", "b&w negative",
    "color photograph print", "b&w photograph print",
]
EXT_LIST = [
    "24 x 36mm (35mm/slides)", "2 x 3 in (polaroids)",
    "2.3 x 3.5 polaroids", "5 x 5 6x5", "4 x 6 in (print)",
    "5 x 7 in (print)", "8 x 10 in (print)",
    "8.5 x 11 in (print)", "12 x 18 in (print)",
]

def choose_fmt_extent(text: str):
    for patt, (fmt, ext) in PATTERN_MAP.items():
        if re.search(patt, text, flags=re.I):
            return fmt, ext
    return DEFAULT_FORMAT, DEFAULT_EXTENT

def main():
    if len(sys.argv) != 3:
        print("Usage: python make_slide_inventory.py <folder> <output.xlsx>")
        sys.exit(1)

    src = Path(sys.argv[1]).resolve()
    out = Path(sys.argv[2]).with_suffix(".xlsx").resolve()
    if not src.is_dir():
        print("Folder not found:", src); sys.exit(1)

    tiffs = sorted(
    (p for p in src.rglob("*") if p.suffix.lower() in {".tif", ".tiff"}),
    key=extract_key
    )
    if not tiffs:
        print("No TIFF files found"); sys.exit(0)

    wb = Workbook()
    ws = wb.active
    ws.title = "Inventory"
    ws.append(["File Name", "Format", "Extent", "Scanning Notes", "Description"])

    for p in tiffs:
        fmt, ext = choose_fmt_extent(str(p))
        ws.append([p.stem, fmt, ext, "", ""])

    lists = wb.create_sheet("_lists"); lists.sheet_state = "hidden"
    for i, v in enumerate(FMT_LIST, 2): lists[f"A{i}"] = v
    for i, v in enumerate(EXT_LIST, 2): lists[f"B{i}"] = v

    dv_fmt = DataValidation(type="list",
                            formula1=f"_lists!$A$2:$A${len(FMT_LIST)+1}")
    dv_ext = DataValidation(type="list",
                            formula1=f"_lists!$B$2:$B${len(EXT_LIST)+1}")
    dv_fmt.ranges.add(f"B2:B{len(tiffs)+1}")
    dv_ext.ranges.add(f"C2:C{len(tiffs)+1}")
    ws.add_data_validation(dv_fmt); ws.add_data_validation(dv_ext)
    ws.freeze_panes = "A2"

    wb.save(out)
    print(f"Saved {len(tiffs)} rows to {out}")

if __name__ == "__main__":
    main()

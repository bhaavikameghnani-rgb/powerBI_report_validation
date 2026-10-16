"""
Power BI File-Based Migration Validator
=========================================
Validates two MANUALLY provided .pbix files against each other — no Power BI
Service connection required for this flow.

Compares (parsed straight out of the .pbix via pbixray):
  - Tables + column counts, and per-column datatypes
  - Measures (DAX)
  - Calculated columns and calculated tables (DAX)
  - Relationships
  - Power Query (M) code per table
  - M Parameters
  - Row-Level Security roles
  - Report pages + visual types (best-effort JSON parse)

Optionally also runs the same pixel + text visual diff as the Tableau
validator, against two exported PDFs/screenshots you supply.

NOTE: if a .pbix is a live/thin report (pointing at a Power BI Service
dataset or Analysis Services model rather than storing its own data model),
there's nothing to extract for the model-based sheets — you'll get a note
in the Summary sheet, and only the report-page comparison will be populated.

Tableau and Power BI are intentionally NOT mixed in one script — use
tableau_file_validator.py for .twb/.twbx files.
"""

import os
import shutil
import tempfile
from datetime import datetime

import powerbi_validator_core as core

# ─────────────────────────────────────────────
# CONFIG — edit before running
# ─────────────────────────────────────────────
SRC_FILE = r"path\to\source_report.pbix"
DST_FILE = r"path\to\destination_report.pbix"

# Optional — leave as "" to skip the visual/KPI comparison step
SRC_VISUAL = r"path\to\source_export.pdf"
DST_VISUAL = r"path\to\destination_export.pdf"

OUTPUT_FOLDER = "Migration_Output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def main():
    for label, path in [("SRC_FILE", SRC_FILE), ("DST_FILE", DST_FILE)]:
        if not os.path.exists(path):
            raise FileNotFoundError(f"{label} not found: {path}")
        if not path.lower().endswith(".pbix"):
            raise ValueError(f"{label} must be .pbix (got: {path})")

    print("Parsing source report …")
    src_meta = core.parse_workbook(SRC_FILE)
    if src_meta["notes"]:
        print(f"  ℹ {src_meta['notes']}")
    print("Parsing destination report …")
    dst_meta = core.parse_workbook(DST_FILE)
    if dst_meta["notes"]:
        print(f"  ℹ {dst_meta['notes']}")

    print("Comparing metadata …")
    diff = core.compare_metadata(src_meta, dst_meta)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_file = os.path.join(OUTPUT_FOLDER, f"powerbi_comparison_{ts}.xlsx")

    page_results = None
    if SRC_VISUAL and DST_VISUAL:
        for label, path in [("SRC_VISUAL", SRC_VISUAL), ("DST_VISUAL", DST_VISUAL)]:
            if not os.path.exists(path):
                raise FileNotFoundError(f"{label} not found: {path}")
        print("Running visual/KPI comparison from uploaded exports …")
        tmp_dir = tempfile.mkdtemp(prefix="powerbi_file_diff_")
        try:
            page_results = core.run_visual_comparison(SRC_VISUAL, DST_VISUAL, tmp_dir)
            core.build_excel_report(src_meta, dst_meta, diff, out_file, page_results)
            mismatches = sum(1 for pr in page_results
                              if pr["pixel_changed_pct"] > 2 or pr["text_status"] == "Mismatch")
            print(f"✅ Visual comparison added — {mismatches}/{len(page_results)} page(s) differ.")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)
    else:
        core.build_excel_report(src_meta, dst_meta, diff, out_file)
        print("ℹ No SRC_VISUAL/DST_VISUAL provided — skipping visual/KPI step.")

    print(f"\n✅ Comparison saved: {out_file}")


if __name__ == "__main__":
    main()

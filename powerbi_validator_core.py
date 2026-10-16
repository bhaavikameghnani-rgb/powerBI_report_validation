"""
Core logic for the Power BI file-based validator: pbix metadata parsing/diffing
and visual/pixel diffing. Shared by the CLI script and the Streamlit app —
no Power BI Service connection required.

Metadata extraction uses `pbixray`, which reads the compressed data model
directly out of the .pbix (works even without Power BI Desktop installed).
Report page/visual extraction is a best-effort parser of the Report/Layout
JSON inside the .pbix, since Microsoft has shipped two different report
storage formats (legacy single-file JSON, and the newer PBIR folder format) —
see parse_report_pages() below if a file doesn't parse cleanly.
"""

import json
import os
import re
import zipfile

import numpy as np
import pandas as pd
from PIL import Image, ImageDraw

try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from pbixray import PBIXRay, LiveConnectionError, NoEmbeddedModelError
except ImportError:
    PBIXRay = None
    LiveConnectionError = NoEmbeddedModelError = Exception

from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

_ILLEGAL_CHARS_RE = re.compile(r"[\000-\010]|[\013-\014]|[\016-\037]")


def clean(value):
    if isinstance(value, str):
        return _ILLEGAL_CHARS_RE.sub("", value)
    return value


# ── Data model metadata (via pbixray) ────────────────────────
def parse_model_metadata(path):
    """
    Returns a dict of DataFrames for the embedded data model, or empty
    DataFrames (with a note) if the file is a live/thin report with no
    embedded model.
    """
    empty = {
        "tables": pd.DataFrame(columns=["Table Name", "Column Count"]),
        "columns": pd.DataFrame(columns=["Table.Column", "Datatype"]),
        "measures": pd.DataFrame(columns=["Measure Name", "Table Name", "Expression", "Display Folder"]),
        "calc_columns": pd.DataFrame(columns=["Column Name", "Table Name", "Expression"]),
        "calc_tables": pd.DataFrame(columns=["Table Name", "Expression"]),
        "relationships": pd.DataFrame(columns=["Relationship", "Cardinality", "Is Active", "Cross Filter"]),
        "power_query": pd.DataFrame(columns=["Table Name", "Expression"]),
        "parameters": pd.DataFrame(columns=["Parameter Name", "Expression"]),
        "rls": pd.DataFrame(columns=["Role Name", "Table Name", "Filter Expression"]),
        "note": "",
    }
    if PBIXRay is None:
        empty["note"] = "pbixray is not installed — run: pip install pbixray"
        return empty

    try:
        model = PBIXRay(path)
    except (LiveConnectionError, NoEmbeddedModelError) as e:
        empty["note"] = f"No embedded data model (live/thin report): {e}"
        return empty

    schema = model.schema if model.schema is not None else pd.DataFrame()
    tables_df = pd.DataFrame({"Table Name": model.tables}) if model.tables is not None else pd.DataFrame()
    if not schema.empty:
        col_counts = schema.groupby("TableName").size().reset_index(name="Column Count")
        tables_df = pd.merge(tables_df, col_counts, left_on="Table Name", right_on="TableName", how="left")
        tables_df = tables_df[["Table Name", "Column Count"]].fillna(0)

    columns_df = pd.DataFrame()
    if not schema.empty:
        columns_df = pd.DataFrame({
            "Table.Column": schema["TableName"] + "." + schema["ColumnName"],
            "Datatype": schema["PandasDataType"].astype(str),
        })

    measures_df = pd.DataFrame()
    dm = getattr(model, "dax_measures", pd.DataFrame())
    if dm is not None and not dm.empty:
        measures_df = pd.DataFrame({
            "Measure Name": dm["Name"],
            "Table Name": dm["TableName"],
            "Expression": dm["Expression"].astype(str),
            "Display Folder": dm.get("DisplayFolder", pd.Series([""] * len(dm))).astype(str),
        })

    calc_cols_df = pd.DataFrame()
    dc = getattr(model, "dax_columns", pd.DataFrame())
    if dc is not None and not dc.empty:
        calc_cols_df = pd.DataFrame({
            "Column Name": dc["ColumnName"],
            "Table Name": dc["TableName"],
            "Expression": dc["Expression"].astype(str),
        })

    calc_tables_df = pd.DataFrame()
    dt = getattr(model, "dax_tables", pd.DataFrame())
    if dt is not None and not dt.empty:
        calc_tables_df = pd.DataFrame({
            "Table Name": dt["TableName"],
            "Expression": dt["Expression"].astype(str),
        })

    rel_df = pd.DataFrame()
    rel = getattr(model, "relationships", pd.DataFrame())
    if rel is not None and not rel.empty:
        rel_df = pd.DataFrame({
            "Relationship": (rel["FromTableName"] + "[" + rel["FromColumnName"] + "] -> " +
                              rel["ToTableName"] + "[" + rel["ToColumnName"] + "]"),
            "Cardinality": rel.get("Cardinality", "").astype(str),
            "Is Active": rel.get("IsActive", "").astype(str),
            "Cross Filter": rel.get("CrossFilteringBehavior", "").astype(str),
        })

    pq_df = pd.DataFrame()
    pq = getattr(model, "power_query", pd.DataFrame())
    if pq is not None and not pq.empty:
        pq_df = pd.DataFrame({
            "Table Name": pq["TableName"],
            "Expression": pq["Expression"].astype(str),
        })

    params_df = pd.DataFrame()
    mp = getattr(model, "m_parameters", pd.DataFrame())
    if mp is not None and not mp.empty:
        params_df = pd.DataFrame({
            "Parameter Name": mp["ParameterName"],
            "Expression": mp["Expression"].astype(str),
        })

    rls_df = pd.DataFrame()
    rls = getattr(model, "rls", pd.DataFrame())
    if rls is not None and not rls.empty:
        rls_df = pd.DataFrame({
            "Role Name": rls["RoleName"],
            "Table Name": rls["TableName"],
            "Filter Expression": rls.get("FilterExpression", "").astype(str),
        })

    return {
        "tables": tables_df.fillna(""),
        "columns": columns_df.fillna(""),
        "measures": measures_df.fillna(""),
        "calc_columns": calc_cols_df.fillna(""),
        "calc_tables": calc_tables_df.fillna(""),
        "relationships": rel_df.fillna(""),
        "power_query": pq_df.fillna(""),
        "parameters": params_df.fillna(""),
        "rls": rls_df.fillna(""),
        "note": "",
    }


# ── Report page/visual metadata (best-effort JSON parsing) ──
def _decode_json_entry(raw_bytes):
    for enc in ("utf-16", "utf-8-sig", "utf-8"):
        try:
            return json.loads(raw_bytes.decode(enc))
        except Exception:
            continue
    return None


def parse_report_pages(path):
    """
    Best-effort extraction of page names + visual types from the pbix's
    Report/Layout. Handles the legacy single-file JSON layout; the newer
    PBIR folder format (Report/definition/pages/...) may need a small
    tweak per file — this returns an empty list with a note if neither
    parses, rather than failing the whole run.
    """
    rows = []
    note = ""
    try:
        with zipfile.ZipFile(path, "r") as z:
            names = z.namelist()

            # Legacy format: single "Report/Layout" JSON file
            if "Report/Layout" in names:
                data = _decode_json_entry(z.read("Report/Layout"))
                if data:
                    for section in data.get("sections", []):
                        vis_types = []
                        for vc in section.get("visualContainers", []):
                            try:
                                cfg = json.loads(vc.get("config", "{}"))
                                vt = cfg.get("singleVisual", {}).get("visualType", "unknown")
                                vis_types.append(vt)
                            except Exception:
                                vis_types.append("unknown")
                        rows.append({
                            "Page Name": section.get("displayName", section.get("name", "")),
                            "Visual Count": len(section.get("visualContainers", [])),
                            "Visual Types": ", ".join(sorted(set(vis_types))),
                        })
                    return rows, note
                note = "Found Report/Layout but couldn't decode it as JSON."

            # Newer PBIR format: Report/definition/pages/<page>/page.json + visuals/*/visual.json
            page_dirs = sorted({n.split("/")[3] for n in names
                                 if n.startswith("Report/definition/pages/") and len(n.split("/")) > 3})
            if page_dirs:
                for pdir in page_dirs:
                    prefix = f"Report/definition/pages/{pdir}/"
                    page_name = pdir
                    if prefix + "page.json" in names:
                        pdata = _decode_json_entry(z.read(prefix + "page.json"))
                        if pdata:
                            page_name = pdata.get("displayName", pdir)
                    vis_types = []
                    visual_files = [n for n in names
                                    if n.startswith(prefix + "visuals/") and n.endswith("visual.json")]
                    for vf in visual_files:
                        vdata = _decode_json_entry(z.read(vf))
                        if vdata:
                            vt = vdata.get("visual", {}).get("visualType", "unknown")
                            vis_types.append(vt)
                    rows.append({
                        "Page Name": page_name,
                        "Visual Count": len(visual_files),
                        "Visual Types": ", ".join(sorted(set(vis_types))),
                    })
                return rows, note

            note = "No Report/Layout or PBIR pages folder found in this file."
    except Exception as e:
        note = f"Could not parse report pages: {e}"
    return rows, note

# ── PBIP project parsing ─────────────────────────────────────

def _find_first_file(root, filename):
    for current_root, _, files in os.walk(root):
        for name in files:
            if name.lower() == filename.lower():
                return os.path.join(current_root, name)
    return None


def _find_files(root, extensions):
    matches = []

    for current_root, _, files in os.walk(root):
        for name in files:
            if any(name.lower().endswith(ext) for ext in extensions):
                matches.append(os.path.join(current_root, name))

    return matches


def _parse_tmdl_files(project_root):
    """
    Best-effort parser for PBIP semantic-model TMDL files.

    Extracts:
      - tables
      - columns
      - measures
      - calculated columns
      - calculated tables
      - Power Query expressions
      - relationships
      - RLS roles

    TMDL is indentation-sensitive, so this parser intentionally focuses
    on the common metadata declarations rather than trying to implement
    the complete TMDL language.
    """

    table_rows = []
    column_rows = []
    measure_rows = []
    calc_column_rows = []
    calc_table_rows = []
    relationship_rows = []
    power_query_rows = []
    rls_rows = []

    tmdl_files = _find_files(project_root, [".tmdl"])

    current_table = None
    current_column = None
    current_measure = None
    current_partition = None
    current_role = None

    expression_lines = []
    expression_target = None

    def flush_expression():
        nonlocal expression_lines, expression_target

        if expression_target is None:
            expression_lines = []
            return

        expression = "\n".join(expression_lines).strip()

        if expression:
            expression_target["Expression"] = expression

        expression_lines = []
        expression_target = None

    for path in tmdl_files:
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                lines = f.readlines()
        except UnicodeDecodeError:
            with open(path, "r", encoding="utf-8") as f:
                lines = f.readlines()
        except Exception:
            continue

        for raw_line in lines:
            line = raw_line.rstrip("\r\n")

            if not line.strip():
                continue

            stripped = line.strip()
            indent = len(line) - len(line.lstrip())

            # New table
            m = re.match(r"^table\s+(.+?)\s*$", stripped)
            if m and indent <= 4:
                flush_expression()

                table_name = m.group(1).strip().strip("'")

                current_table = table_name
                current_column = None
                current_measure = None
                current_partition = None
                current_role = None

                table_rows.append({
                    "Table Name": table_name,
                    "Column Count": 0,
                })

                continue

            # New column
            m = re.match(r"^column\s+(.+?)\s*$", stripped)
            if m and current_table is not None:
                flush_expression()

                column_name = m.group(1).strip().strip("'")

                current_column = {
                    "Column Name": column_name,
                    "Table Name": current_table,
                    "Expression": "",
                }

                column_rows.append({
                    "Table.Column":
                        f"{current_table}.{column_name}",
                    "Datatype": "",
                })

                current_measure = None
                current_partition = None
                continue

            # New measure
            m = re.match(
                r"^measure\s+(.+?)(?:\s*=\s*(.*))?$",
                stripped,
            )

            if m and current_table is not None:
                flush_expression()

                measure_name = m.group(1).strip().strip("'")
                measure_expression = (m.group(2) or "").strip()

                current_measure = {
                    "Measure Name": measure_name,
                    "Table Name": current_table,
                    "Expression": measure_expression,
                    "Display Folder": "",
                }

                measure_rows.append(current_measure)

                current_column = None
                current_partition = None
                continue

            # New partition / Power Query source
            m = re.match(r"^partition\s+(.+?)\s*=\s*(.*)$", stripped)

            if m and current_table is not None:
                flush_expression()

                current_partition = {
                    "Table Name": current_table,
                    "Expression": m.group(2).strip(),
                }

                power_query_rows.append(current_partition)

                current_column = None
                current_measure = None
                continue

            # Relationship
            m = re.match(r"^relationship\s+(.+?)\s*$", stripped)

            if m:
                flush_expression()

                relationship_rows.append({
                    "Relationship": m.group(1).strip().strip("'"),
                    "Cardinality": "",
                    "Is Active": "",
                    "Cross Filter": "",
                })

                continue

            # Role / RLS
            m = re.match(r"^role\s+(.+?)\s*$", stripped)

            if m:
                flush_expression()

                current_role = {
                    "Role Name": m.group(1).strip().strip("'"),
                    "Table Name": "",
                    "Filter Expression": "",
                }

                rls_rows.append(current_role)
                continue

            # Datatype
            if current_column is not None:
                m = re.match(r"^dataType:\s*(.+)$", stripped)

                if m:
                    datatype = m.group(1).strip()

                    for row in reversed(column_rows):
                        if row["Table.Column"] == (
                            f"{current_table}.{current_column['Column Name']}"
                        ):
                            row["Datatype"] = datatype
                            break

                    continue

            # DAX / M expression continuation
            if stripped.startswith("expression ="):
                flush_expression()

                expression_target = (
                    current_measure
                    or current_column
                    or current_partition
                    or current_role
                )

                expression_lines = [
                    stripped.split("=", 1)[1].strip()
                ]

                continue

            if expression_target is not None:
                # Indented continuation of an expression
                if indent > 4:
                    expression_lines.append(stripped)
                    continue

        flush_expression()

    # Update table column counts
    for table_row in table_rows:
        table_name = table_row["Table Name"]

        count = sum(
            1
            for row in column_rows
            if row["Table.Column"].startswith(table_name + ".")
        )

        table_row["Column Count"] = count

    return {
        "tables": pd.DataFrame(
            table_rows,
            columns=["Table Name", "Column Count"],
        ),
        "columns": pd.DataFrame(
            column_rows,
            columns=["Table.Column", "Datatype"],
        ),
        "measures": pd.DataFrame(
            measure_rows,
            columns=[
                "Measure Name",
                "Table Name",
                "Expression",
                "Display Folder",
            ],
        ),
        "calc_columns": pd.DataFrame(
            calc_column_rows,
            columns=[
                "Column Name",
                "Table Name",
                "Expression",
            ],
        ),
        "calc_tables": pd.DataFrame(
            calc_table_rows,
            columns=["Table Name", "Expression"],
        ),
        "relationships": pd.DataFrame(
            relationship_rows,
            columns=[
                "Relationship",
                "Cardinality",
                "Is Active",
                "Cross Filter",
            ],
        ),
        "power_query": pd.DataFrame(
            power_query_rows,
            columns=["Table Name", "Expression"],
        ),
        "parameters": pd.DataFrame(
            columns=["Parameter Name", "Expression"],
        ),
        "rls": pd.DataFrame(
            rls_rows,
            columns=[
                "Role Name",
                "Table Name",
                "Filter Expression",
            ],
        ),
    }
def _parse_pbip_report(project_root):
    """
    Extract report pages and visual types from PBIR project files.
    Supports the newer PBIR folder structure and PBIR-Legacy report.json.
    """

    rows = []

    # ── PBIR format ──────────────────────────────────────────

    page_json_files = _find_files(project_root, ["page.json"])

    for page_path in page_json_files:
        try:
            with open(page_path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
        except Exception:
            continue

        page_name = (
            data.get("displayName")
            or data.get("name")
            or os.path.basename(os.path.dirname(page_path))
        )

        page_dir = os.path.dirname(page_path)

        visual_types = []

        for root, _, files in os.walk(page_dir):
            for filename in files:
                if filename.lower() != "visual.json":
                    continue

                visual_path = os.path.join(root, filename)

                try:
                    with open(
                        visual_path,
                        "r",
                        encoding="utf-8-sig",
                    ) as f:
                        visual_data = json.load(f)

                    visual = visual_data.get("visual", visual_data)

                    visual_type = (
                        visual.get("visualType")
                        or visual.get("singleVisual", {}).get("visualType")
                        or "unknown"
                    )

                    visual_types.append(visual_type)

                except Exception:
                    visual_types.append("unknown")

        rows.append({
            "Page Name": page_name,
            "Visual Count": len(visual_types),
            "Visual Types": ", ".join(sorted(set(visual_types))),
        })

    if rows:
        return rows, ""

    # ── PBIR-Legacy ──────────────────────────────────────────

    report_json = _find_first_file(project_root, "report.json")

    if report_json:
        try:
            with open(
                report_json,
                "r",
                encoding="utf-8-sig",
            ) as f:
                data = json.load(f)

            for section in data.get("sections", []):
                visual_types = []

                for visual in section.get("visualContainers", []):
                    try:
                        config = json.loads(
                            visual.get("config", "{}")
                        )

                        visual_type = (
                            config
                            .get("singleVisual", {})
                            .get("visualType", "unknown")
                        )

                    except Exception:
                        visual_type = "unknown"

                    visual_types.append(visual_type)

                rows.append({
                    "Page Name": (
                        section.get("displayName")
                        or section.get("name", "")
                    ),
                    "Visual Count": len(
                        section.get("visualContainers", [])
                    ),
                    "Visual Types": ", ".join(
                        sorted(set(visual_types))
                    ),
                })

            return rows, ""

        except Exception as e:
            return rows, f"Could not parse report.json: {e}"

    return rows, (
        "No PBIR page definitions or report.json "
        "were found in the PBIP project."
    )
def parse_workbook(path):
    """
    Parse either:
      - a .pbix file
      - a PBIP project directory

    Both return the same metadata structure so the existing
    compare_metadata() function can be reused.
    """

    # ── PBIP project ──────────────────────────────────────────
    if os.path.isdir(path):
        model_meta = _parse_tmdl_files(path)
        page_rows, page_note = _parse_pbip_report(path)

        return {
            "tables": model_meta["tables"].fillna(""),
            "columns": model_meta["columns"].fillna(""),
            "measures": model_meta["measures"].fillna(""),
            "calc_columns": model_meta["calc_columns"].fillna(""),
            "calc_tables": model_meta["calc_tables"].fillna(""),
            "relationships": model_meta["relationships"].fillna(""),
            "power_query": model_meta["power_query"].fillna(""),
            "parameters": model_meta["parameters"].fillna(""),
            "rls": model_meta["rls"].fillna(""),
            "pages": pd.DataFrame(page_rows).fillna(""),
            "notes": "; ".join(
                n for n in [page_note] if n
            ),
        }

    # ── PBIX file ─────────────────────────────────────────────
    model_meta = parse_model_metadata(path)
    page_rows, page_note = parse_report_pages(path)

    return {
        "tables": model_meta["tables"],
        "columns": model_meta["columns"],
        "measures": model_meta["measures"],
        "calc_columns": model_meta["calc_columns"],
        "calc_tables": model_meta["calc_tables"],
        "relationships": model_meta["relationships"],
        "power_query": model_meta["power_query"],
        "parameters": model_meta["parameters"],
        "rls": model_meta["rls"],
        "pages": pd.DataFrame(page_rows).fillna(""),
        "notes": "; ".join(
            n for n in [model_meta["note"], page_note]
            if n
        ),
    }

def _values_equal(a, b):
    """
    Compare metadata values without treating harmless numeric formatting
    differences such as 3 vs 3.0 as mismatches.

    Examples:
        3 == 3.0
        "3" == "3.0"
        16 == "16.0"

    Non-numeric strings are still compared as strings.
    """
    # Both missing / NaN
    if pd.isna(a) and pd.isna(b):
        return True

    # One missing, one populated
    if pd.isna(a) or pd.isna(b):
        return False

    # First try numeric comparison, including numpy numeric types
    try:
        a_num = pd.to_numeric(a)
        b_num = pd.to_numeric(b)

        if not pd.isna(a_num) and not pd.isna(b_num):
            return float(a_num) == float(b_num)
    except (TypeError, ValueError):
        pass

    # Fall back to normal string comparison
    return str(a).strip() == str(b).strip()


def _display_value(value):
    """
    Format comparison values cleanly for the Streamlit table and Excel
    report. Numeric values that are mathematically whole are displayed
    without an unnecessary .0.
    """
    if pd.isna(value):
        return ""

    try:
        numeric = pd.to_numeric(value)

        if not pd.isna(numeric):
            numeric = float(numeric)

            if numeric.is_integer():
                return str(int(numeric))

            return str(numeric)
    except (TypeError, ValueError):
        pass

    return str(value)
# ── Metadata comparison ──────────────────────────────────────
def _diff_on_key(src_df, dst_df, key_col, compare_cols, label):
    if src_df.empty and dst_df.empty:
        return pd.DataFrame()

    # Safety net: duplicate keys on either side would multiply rows (N x M)
    # instead of comparing 1:1 — collapse to one row per key before merging.
    src_df = src_df.drop_duplicates(subset=[key_col], keep="first")
    dst_df = dst_df.drop_duplicates(subset=[key_col], keep="first")

    merged = pd.merge(
        src_df,
        dst_df,
        on=key_col,
        how="outer",
        suffixes=("_SRC", "_DST"),
    )

    rows = []

    for _, r in merged.iterrows():
        for col in compare_cols:
            s_raw = r.get(f"{col}_SRC", "")
            d_raw = r.get(f"{col}_DST", "")

            s_missing = pd.isna(s_raw) or str(s_raw).strip() == ""
            d_missing = pd.isna(d_raw) or str(d_raw).strip() == ""

            # If both values are equivalent, there is no difference.
            # This correctly treats 3 and 3.0 as equal.
            if _values_equal(s_raw, d_raw):
                continue

            if d_missing:
                status = "Missing in Destination"
            elif s_missing:
                status = "Extra in Destination"
            else:
                status = "Mismatch"

            rows.append({
                label: r[key_col],
                "Field": col,
                "Source Value": _display_value(s_raw),
                "Target Value": _display_value(d_raw),
                "Status": status,
            })

    return pd.DataFrame(rows)


def compare_metadata(src, dst):
    return {
        "Table_Mismatch": _diff_on_key(
            src["tables"], dst["tables"], "Table Name", ["Column Count"], "Table Name"),
        "Column_Mismatch": _diff_on_key(
            src["columns"], dst["columns"], "Table.Column", ["Datatype"], "Table.Column"),
        "Measure_Mismatch": _diff_on_key(
            src["measures"], dst["measures"], "Measure Name",
            ["Table Name", "Expression", "Display Folder"], "Measure Name"),
        "CalcColumn_Mismatch": _diff_on_key(
            src["calc_columns"], dst["calc_columns"], "Column Name",
            ["Table Name", "Expression"], "Column Name"),
        "CalcTable_Mismatch": _diff_on_key(
            src["calc_tables"], dst["calc_tables"], "Table Name", ["Expression"], "Table Name"),
        "Relationship_Mismatch": _diff_on_key(
            src["relationships"], dst["relationships"], "Relationship",
            ["Cardinality", "Is Active", "Cross Filter"], "Relationship"),
        "PowerQuery_Mismatch": _diff_on_key(
            src["power_query"], dst["power_query"], "Table Name", ["Expression"], "Table Name"),
        "Parameter_Mismatch": _diff_on_key(
            src["parameters"], dst["parameters"], "Parameter Name", ["Expression"], "Parameter Name"),
        "RLS_Mismatch": _diff_on_key(
            src["rls"], dst["rls"], "Role Name", ["Table Name", "Filter Expression"], "Role Name"),
        "Page_Mismatch": _diff_on_key(
            src["pages"], dst["pages"], "Page Name", ["Visual Count", "Visual Types"], "Page Name"),
    }


# ── Visual / KPI comparison (identical approach to the Tableau validator) ──
def load_pages(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        if fitz is None:
            raise RuntimeError("PyMuPDF (fitz) is required to read PDFs — pip install pymupdf")
        doc = fitz.open(path)
        images, texts = [], []
        zoom = 150 / 72
        mat = fitz.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat, alpha=False)
            images.append(Image.frombytes("RGB", [pix.width, pix.height], pix.samples))
            texts.append(page.get_text("text").strip())
        doc.close()
        return images, texts
    return [Image.open(path).convert("RGB")], [""]


def pixel_diff(img_src, img_dst, highlight_color=(255, 0, 0)):
    w, h = max(img_src.width, img_dst.width), max(img_src.height, img_dst.height)
    a = img_src.resize((w, h), Image.LANCZOS).convert("RGB")
    b = img_dst.resize((w, h), Image.LANCZOS).convert("RGB")
    diff = np.abs(np.array(a, dtype=np.int16) - np.array(b, dtype=np.int16)).sum(axis=2)
    mask = diff > 30
    changed_pct = mask.sum() / mask.size * 100
    blend = Image.blend(a, b, alpha=0.5)
    draw = ImageDraw.Draw(blend)
    coords = np.argwhere(mask)
    if len(coords):
        for y, x in coords[::4]:
            draw.point((x, y), fill=highlight_color)
    return blend, round(changed_pct, 2)


def text_diff_summary(text_src, text_dst):
    if not text_src and not text_dst:
        return "N/A (image only)", "", ""
    lines_s, lines_d = set(text_src.splitlines()), set(text_dst.splitlines())
    only_s, only_d = lines_s - lines_d, lines_d - lines_s
    if not only_s and not only_d:
        return "Identical", "", ""
    return ("Mismatch", " | ".join(list(only_s)[:5])[:500], " | ".join(list(only_d)[:5])[:500])


def run_visual_comparison(src_visual, dst_visual, tmp_dir):
    src_images, src_texts = load_pages(src_visual)
    dst_images, dst_texts = load_pages(dst_visual)
    max_pages = max(len(src_images), len(dst_images))
    page_results = []
    for i in range(max_pages):
        p_num = i + 1
        if i >= len(src_images) or i >= len(dst_images):
            page_results.append({"page": p_num, "pixel_changed_pct": 100.0,
                                  "text_status": "Missing page", "src_only_text": "",
                                  "dst_only_text": "", "diff_img_path": None})
            continue
        diff_img, pct = pixel_diff(src_images[i], dst_images[i])
        diff_path = os.path.join(tmp_dir, f"page{p_num}_diff.png")
        diff_img.save(diff_path, format="PNG")
        t_stat, s_only, d_only = text_diff_summary(
            src_texts[i] if i < len(src_texts) else "",
            dst_texts[i] if i < len(dst_texts) else "")
        page_results.append({"page": p_num, "pixel_changed_pct": pct, "text_status": t_stat,
                              "src_only_text": s_only, "dst_only_text": d_only,
                              "diff_img_path": diff_path})
    return page_results


# ── Excel export ─────────────────────────────────────────────
_HEADER_FILL = PatternFill("solid", start_color="1F4E79")
_MISMATCH_FILL = PatternFill("solid", start_color="FFE0E0")
_OK_FILL = PatternFill("solid", start_color="E2EFDA")
_HEADER_FONT = Font(bold=True, color="FFFFFF", name="Arial", size=11)
_BODY_FONT = Font(name="Arial", size=10)
_THIN = Side(style="thin", color="AAAAAA")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _style_header_row(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border, cell.alignment = _HEADER_FILL, _HEADER_FONT, _BORDER, _CENTER


def _style_data_row(ws, row, cols, is_mismatch=False):
    fill = _MISMATCH_FILL if is_mismatch else _OK_FILL
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill, cell.font, cell.border, cell.alignment = fill, _BODY_FONT, _BORDER, _CENTER


def _clean_df(df):
    """Sanitize every cell without relying on DataFrame.applymap (removed in
    newer pandas) or DataFrame.map (not present in older pandas) — Series.map
    via apply() works across all pandas versions."""
    return df.apply(lambda col: col.map(clean))


def write_metadata_sheet(writer, sheet_name, df):
    if df.empty:
        return
    _clean_df(df).to_excel(writer, sheet_name=sheet_name, index=False)


def style_metadata_sheets(xl_path, sheet_names):
    wb = load_workbook(xl_path)
    for sn in sheet_names:
        if sn not in wb.sheetnames:
            continue
        ws = wb[sn]
        if ws.max_row < 1:
            continue
        _style_header_row(ws, 1, ws.max_column)
        for row in range(2, ws.max_row + 1):
            status_val = str(ws.cell(row=row, column=ws.max_column).value or "")
            is_bad = any(k in status_val for k in ("Mismatch", "Missing", "Extra"))
            _style_data_row(ws, row, ws.max_column, is_mismatch=is_bad)
        for col in ws.columns:
            max_len = max((len(str(c.value or "")) for c in col), default=10)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 60)
        ws.freeze_panes = "A2"
    wb.save(xl_path)


def embed_visual_comparison(xl_path, sheet_name, page_results):
    wb = load_workbook(xl_path)
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)
    headers = ["Page #", "Pixel Change %", "Visual Status", "Text Status",
               "Source Only Text (sample)", "Dest Only Text (sample)", "Diff Image"]
    for i, h in enumerate(headers, 1):
        ws.cell(row=1, column=i, value=h)
    _style_header_row(ws, 1, len(headers))
    ws.row_dimensions[1].height = 30
    for i, w in enumerate([8, 16, 16, 18, 50, 50, 80], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    row = 2
    for pr in page_results:
        is_bad = pr["pixel_changed_pct"] > 2 or pr["text_status"] == "Mismatch"
        vis_str = f"⚠ {pr['pixel_changed_pct']}% changed" if pr["pixel_changed_pct"] > 2 else "✓ OK"
        ws.cell(row=row, column=1, value=pr["page"])
        ws.cell(row=row, column=2, value=pr["pixel_changed_pct"])
        ws.cell(row=row, column=3, value=vis_str)
        ws.cell(row=row, column=4, value=clean(pr["text_status"]))
        ws.cell(row=row, column=5, value=clean(pr["src_only_text"]))
        ws.cell(row=row, column=6, value=clean(pr["dst_only_text"]))
        _style_data_row(ws, row, 6, is_mismatch=is_bad)
        if pr.get("diff_img_path") and os.path.exists(pr["diff_img_path"]):
            try:
                img_h_px = 200
                pil_img = Image.open(pr["diff_img_path"])
                ratio = img_h_px / pil_img.height
                xl_img = XLImage(pr["diff_img_path"])
                xl_img.height = img_h_px
                xl_img.width = int(pil_img.width * ratio)
                ws.add_image(xl_img, f"G{row}")
                ws.row_dimensions[row].height = img_h_px * 0.75
            except Exception as e:
                ws.cell(row=row, column=7, value=clean(f"Image error: {e}"))
        else:
            ws.cell(row=row, column=7, value="No image")
            ws.row_dimensions[row].height = 30
        row += 1
    wb.save(xl_path)


def build_excel_report(src_meta, dst_meta, diff, out_path, page_results=None):
    sheet_order = ["Table_Mismatch", "Column_Mismatch", "Measure_Mismatch",
                   "CalcColumn_Mismatch", "CalcTable_Mismatch", "Relationship_Mismatch",
                   "PowerQuery_Mismatch", "Parameter_Mismatch", "RLS_Mismatch", "Page_Mismatch"]
    raw_keys = ["tables", "columns", "measures", "calc_columns", "calc_tables",
                "relationships", "power_query", "parameters", "rls", "pages"]
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for k in sheet_order:
            write_metadata_sheet(writer, k, diff.get(k, pd.DataFrame()))
        for k in raw_keys:
            write_metadata_sheet(writer, f"SRC_{k}", src_meta.get(k, pd.DataFrame()))
            write_metadata_sheet(writer, f"DST_{k}", dst_meta.get(k, pd.DataFrame()))
        summary_rows = [{"Sheet": k, "Mismatch Count": len(diff.get(k, pd.DataFrame()))} for k in sheet_order]
        notes = "; ".join(n for n in [src_meta.get("notes", ""), dst_meta.get("notes", "")] if n)
        if notes:
            summary_rows.append({"Sheet": "NOTES", "Mismatch Count": notes})
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
    style_metadata_sheets(out_path, sheet_order + ["Summary"])
    if page_results:
        embed_visual_comparison(out_path, "Visual_Comparison", page_results)
    return out_path

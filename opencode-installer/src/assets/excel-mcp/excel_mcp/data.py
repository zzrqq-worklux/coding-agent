from pathlib import Path
from typing import Any, Dict, List, Optional
import logging

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter

from .exceptions import DataError
from .cell_utils import parse_cell_range
from .cell_validation import get_data_validation_for_cell
from .backup import backup_file

logger = logging.getLogger(__name__)

import re

VALUE_MODES = ("both", "value", "formula")


def _is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def _resolve_cell_value(cell, cached_cell, value_mode: str = "both") -> Dict[str, Any]:
    """Resolve a cell into {value, formula?, value_source?}.

    - formula: value is the formula text (legacy behaviour)
    - value:   value is the cached result when the cell holds a formula
    - both:    value is the cached result, formula is also returned
    """
    raw = cell.value
    is_f = _is_formula(raw)
    cached = cached_cell.value if cached_cell is not None else None

    if value_mode == "formula":
        value = raw
    elif is_f:
        value = cached
    else:
        value = raw

    out: Dict[str, Any] = {"value": value}
    if is_f:
        out["formula"] = raw
        out["value_source"] = "cached" if cached is not None else "formula_only"
    else:
        out["value_source"] = "literal"
    return out


def _merged_value_map(ws, cached_ws, value_mode: str) -> Dict[tuple, Dict[str, Any]]:
    """Map cells of *vertical* merged ranges to the anchor value.

    Only vertical merges (one column, several rows) are filled — that is the
    "category spans rows" pattern. Horizontal/title merges are left alone so
    a merged instruction banner does not become five identical column headers.
    """
    mapping: Dict[tuple, Dict[str, Any]] = {}
    for rng in ws.merged_cells.ranges:
        if rng.max_row <= rng.min_row:
            continue
        anchor = ws.cell(row=rng.min_row, column=rng.min_col)
        cached_anchor = cached_ws.cell(row=rng.min_row, column=rng.min_col) if cached_ws is not None else None
        resolved = _resolve_cell_value(anchor, cached_anchor, value_mode)
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                if (r, c) != (rng.min_row, rng.min_col):
                    mapping[(r, c)] = resolved
    return mapping


def _unique_headers(headers: List[str]) -> List[str]:
    """Deduplicate repeated column headers (merged headers repeat after fill-forward)."""
    seen: Dict[str, int] = {}
    out: List[str] = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            out.append(f"{h}_{seen[h]}")
        else:
            seen[h] = 1
            out.append(h)
    return out


_REF_RE = re.compile(
    r"(?:(?:'([^']+)'|([A-Za-z0-9_\u4e00-\u9fa5]+))!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)"
)

PLACEHOLDERS = {"", "/", "-", "—", "–", "n/a", "na", "none"}
KEY_HEADERS = [
    ("项目名称", 100), ("项目名", 95), ("项目", 90),
    ("部件", 85), ("物料名称", 82), ("物料编码", 80), ("料号", 78), ("名称", 70),
    ("序号", 60), ("no", 55), ("#", 50),
]


def _pick_header_row(rows, max_scan: int = 10) -> int:
    """Pick the earliest row with >=3 string cells (header); fallback to the richest row."""
    best_i, best_score = 0, -1
    for i, row in enumerate(rows[:max_scan]):
        score = sum(1 for c in row if isinstance(c, str) and c.strip())
        if score >= 3:
            return i
        if score > best_score:
            best_i, best_score = i, score
    return best_i


def _pick_key_col(headers):
    """Return (index, header) of the best business key column, or (None, '')."""
    best_idx, best_score, best_name = None, -1, ""
    for idx, h in enumerate(headers):
        if not isinstance(h, str):
            continue
        low = h.strip().lower()
        for kw, score in KEY_HEADERS:
            if kw.lower() in low and score > best_score:
                best_idx, best_score, best_name = idx, score, h.strip()
    return best_idx, best_name


def _norm_key(value) -> str:
    if value is None:
        return ""
    s = str(value).replace("\n", " ").strip()
    s = re.sub(r"[（(][^）)]*[）)]\s*$", "", s).strip()
    s = re.sub(r"^新\s+", "", s).strip()
    return re.sub(r"\s+", " ", s)


def _is_placeholder(value) -> bool:
    return value is None or (isinstance(value, str) and value.strip().lower() in PLACEHOLDERS)


def _values_equal(a, b, tolerance: float) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) <= tolerance
    if isinstance(a, str) and isinstance(b, str):
        return a.strip() == b.strip()
    return a == b


def _text_change(old, new, text_mode: str):
    """Return a compact representation of a string change (append-only when possible)."""
    if text_mode == "none":
        return {"old_len": len(old), "new_len": len(new)}
    if text_mode == "tail":
        if new.startswith(old):
            return {"appended": new[len(old):]}
        prefix = 0
        limit = min(len(old), len(new))
        while prefix < limit and old[prefix] == new[prefix]:
            prefix += 1
        if prefix >= 20:
            return {"appended": new[prefix:], "prefix_trimmed": prefix}
    return {"old": old, "new": new}


def diff_workbooks(
    old_file: Path | str,
    new_file: Path | str,
    sheet_name: str = "",
    key_column: str = "",
    normalize_key: bool = True,
    text_mode: str = "tail",
    ignore_placeholders: bool = True,
    tolerance: float = 1e-6,
    rename_similarity: float = 0.6,
    max_rows: int = 200,
) -> Dict[str, Any]:
    """Row-aligned semantic diff of two workbooks (values + formulas, no formatting).

    Args:
        old_file/new_file: versions to compare
        sheet_name: compare only this sheet (empty = all common sheets)
        key_column: business key column letter or header name (empty = auto-detect)
        normalize_key: strip newlines / trailing parentheticals / leading "新 " before matching
        text_mode: tail (only appended text) | full | none
        ignore_placeholders: treat "/" "-" etc. as "cleared" and collapse into a count
        tolerance: numeric comparison tolerance
        rename_similarity: similarity threshold to report a key as renamed
        max_rows: cap detail rows per sheet
    """
    import difflib

    def load(path):
        wf = load_workbook(path, read_only=False)
        wv = load_workbook(path, data_only=True)
        return wf, wv

    wf_old, wv_old = load(old_file)
    wf_new, wv_new = load(new_file)
    try:
        names_old, names_new = wf_old.sheetnames, wf_new.sheetnames
        common = [s for s in names_new if s in names_old]
        if sheet_name:
            if sheet_name not in common:
                raise DataError(f"Sheet '{sheet_name}' not found in both workbooks")
            common = [sheet_name]

        out = {
            "old": str(old_file),
            "new": str(new_file),
            "sheets_added": [s for s in names_new if s not in names_old],
            "sheets_removed": [s for s in names_old if s not in names_new],
            "summary": {"sheets_compared": len(common), "changed_rows": 0, "added": 0,
                        "removed": 0, "renamed": 0, "formula_changes": 0,
                        "cleared_cells": 0, "truncated": False},
            "sheets": [],
        }

        for sname in common:
            ws_o, ws_n = wf_old[sname], wf_new[sname]
            wv_o, wv_n = wv_old[sname], wv_new[sname]

            def grid(ws, wsv):
                rows = []
                for r in range(1, ws.max_row + 1):
                    rows.append([(ws.cell(row=r, column=c).value, wsv.cell(row=r, column=c).value)
                                 for c in range(1, ws.max_column + 1)])
                return rows

            go, gn = grid(ws_o, wv_o), grid(ws_n, wv_n)
            if not go or not gn:
                continue

            hdr = _pick_header_row([[raw for raw, _ in row] for row in gn])
            headers = [h for h, _ in gn[hdr]]
            headers = _unique_headers([str(h) if h is not None else f"Column{i+1}"
                                       for i, h in enumerate(headers)])

            # default: compare by row index
            key_idx, key_name = None, ""
            if key_column:
                key_column = key_column.strip()
                if len(key_column) == 1 and key_column.isalpha():
                    key_idx = ord(key_column.upper()) - 65
                    key_name = headers[key_idx] if key_idx < len(headers) else key_column
                else:
                    for i, h in enumerate(headers):
                        if h == key_column:
                            key_idx, key_name = i, h
                            break
            else:
                key_idx, key_name = _pick_key_col(headers)

            def row_dict(row, idx):
                if idx is None:
                    return None
                cell = row[idx] if idx < len(row) else (None, None)
                return cell[1] if cell[1] is not None else cell[0]

            def build(rows, data_start):
                m: Dict[str, Dict[str, Any]] = {}
                for ri in range(data_start, len(rows)):
                    row = rows[ri]
                    if all((v is None or (isinstance(v, str) and not v.strip())) for v, _ in row):
                        continue
                    raw_key = row_dict(row, key_idx)
                    k = _norm_key(raw_key) if normalize_key else str(raw_key).strip()
                    if not k:
                        k = f"__row_{ri+1}"
                    if k in m:
                        n = 2
                        while f"{k}#{n}" in m:
                            n += 1
                        k = f"{k}#{n}"
                    m[k] = {"row": ri + 1, "cells": row}
                return m

            old_rows = build(go, hdr + 1)
            new_rows = build(gn, hdr + 1)

            matched, renamed = [], []
            old_keys, new_keys = set(old_rows), set(new_rows)
            exact = old_keys & new_keys
            for k in exact:
                matched.append((k, k))

            rest_old = sorted(old_keys - exact)
            rest_new = sorted(new_keys - exact)
            candidates = []
            for ok in rest_old:
                for nk in rest_new:
                    r = difflib.SequenceMatcher(None, ok, nk).ratio()
                    if r >= rename_similarity:
                        candidates.append((r, ok, nk))
            candidates.sort(reverse=True)
            used_old, used_new = set(), set()
            for r, ok, nk in candidates:
                if ok in used_old or nk in used_new:
                    continue
                used_old.add(ok)
                used_new.add(nk)
                matched.append((ok, nk))
                renamed.append({"old": ok, "new": nk, "ratio": round(r, 3)})

            added = [k for k in rest_new if k not in used_new]
            removed = [k for k in rest_old if k not in used_old]

            changes = []
            formula_changes = 0
            cleared = 0
            for ok, nk in matched:
                orow, nrow = old_rows[ok], new_rows[nk]
                cell_changes = {}
                width = max(len(orow["cells"]), len(nrow["cells"]))
                for ci in range(width):
                    ov, oc = orow["cells"][ci] if ci < len(orow["cells"]) else (None, None)
                    nv, nc = nrow["cells"][ci] if ci < len(nrow["cells"]) else (None, None)
                    ov_eff = oc if oc is not None else ov
                    nv_eff = nc if nc is not None else nv

                    if _values_equal(ov_eff, nv_eff, tolerance):
                        if isinstance(ov, str) and isinstance(nv, str) and ov.startswith("=") and nv.startswith("=") and ov != nv:
                            name = headers[ci] if ci < len(headers) else f"Column{ci+1}"
                            cell_changes[name] = {"kind": "formula", "old": ov, "new": nv}
                            formula_changes += 1
                        continue

                    if ignore_placeholders and _is_placeholder(nv_eff) and not _is_placeholder(ov_eff):
                        cleared += 1
                        continue

                    name = headers[ci] if ci < len(headers) else f"Column{ci+1}"
                    if isinstance(ov_eff, str) and isinstance(nv_eff, str):
                        cell_changes[name] = {"kind": "text", **_text_change(ov_eff, nv_eff, text_mode)}
                    else:
                        cell_changes[name] = {"kind": "value", "old": ov_eff, "new": nv_eff}

                if cell_changes:
                    changes.append({"key": nk, "old_row": orow["row"], "new_row": nrow["row"],
                                    "changes": cell_changes})

            def brief(keys):
                out_rows = []
                for k in keys[:max_rows]:
                    r = new_rows.get(k) or old_rows.get(k)
                    out_rows.append({"key": k, "row": r["row"] if r else None})
                return out_rows

            sheet_out = {
                "sheet": sname,
                "key": {"column": get_column_letter(key_idx + 1) if key_idx is not None else None,
                        "header": key_name, "mode": "aligned" if key_idx is not None else "row_index"},
                "changed_rows": len(changes),
                "changes": changes[:max_rows],
                "added": brief(added),
                "removed": brief(removed),
                "renamed": renamed,
                "formula_changes": formula_changes,
                "cleared_cells": cleared,
            }
            out["sheets"].append(sheet_out)
            out["summary"]["changed_rows"] += len(changes)
            out["summary"]["added"] += len(added)
            out["summary"]["removed"] += len(removed)
            out["summary"]["renamed"] += len(renamed)
            out["summary"]["formula_changes"] += formula_changes
            out["summary"]["cleared_cells"] += cleared
            if len(changes) > max_rows:
                out["summary"]["truncated"] = True

        return out
    finally:
        wf_old.close()
        wv_old.close()
        wf_new.close()
        wv_new.close()



def get_formula_map(filepath: Path | str, sheet_name: Optional[str] = None) -> Dict[str, Any]:
    """List every formula cell and the ranges it references (local or cross-sheet)."""
    wb = load_workbook(filepath, read_only=False)
    try:
        targets = [sheet_name] if sheet_name else list(wb.sheetnames)
        sheets_out = []
        for name in targets:
            if name not in wb.sheetnames:
                raise DataError(f"Sheet '{name}' not found")
            ws = wb[name]
            cells = []
            for row in ws.iter_rows():
                for cell in row:
                    if not _is_formula(cell.value):
                        continue
                    refs = []
                    for m in _REF_RE.finditer(cell.value):
                        refs.append({"sheet": m.group(1) or m.group(2), "range": m.group(3)})
                    cells.append({"cell": cell.coordinate, "formula": cell.value, "refs": refs})
            if cells:
                sheets_out.append({"sheet": name, "formula_cells": len(cells), "cells": cells})
        return {"filepath": str(filepath), "sheet_count": len(sheets_out), "sheets": sheets_out}
    finally:
        wb.close()

def _safe_color(color):
    """Safely extract color value from openpyxl Color object."""
    if color is None:
        return None
    
    def _ok(v):
        if v is None:
            return False
        s = str(v)
        return s and 'Values must be' not in s and s not in ('00000000', '0', 'None')
    
    try:
        if hasattr(color, 'rgb') and _ok(color.rgb):
            return str(color.rgb)
    except Exception:
        pass
    try:
        if hasattr(color, 'theme') and isinstance(color.theme, int) and color.theme > 0:
            return f"theme:{color.theme}"
    except Exception:
        pass
    try:
        if hasattr(color, 'indexed') and _ok(color.indexed) and color.indexed != 0:
            return f"indexed:{color.indexed}"
    except Exception:
        pass
    return None

def _get_cell_formatting(cell):
    """Extract formatting metadata from an openpyxl cell."""
    fmt = {}
    
    if cell.fill and cell.fill.fgColor:
        c = _safe_color(cell.fill.fgColor)
        if c:
            fmt["fill_color"] = c
    if cell.fill and cell.fill.bgColor:
        c = _safe_color(cell.fill.bgColor)
        if c and c != fmt.get("fill_color"):
            fmt["bg_color"] = c
    
    if cell.font:
        fc = _safe_color(cell.font.color)
        if fc:
            fmt["font_color"] = fc
        if cell.font.bold:
            fmt["bold"] = True
        if cell.font.italic:
            fmt["italic"] = True
        if cell.font.underline and cell.font.underline != 'none':
            fmt["underline"] = cell.font.underline
        if cell.font.size:
            fmt["font_size"] = cell.font.size
        if cell.font.name:
            fmt["font_name"] = cell.font.name
    
    if cell.number_format and cell.number_format != 'General':
        fmt["number_format"] = cell.number_format
    
    if cell.alignment:
        al = cell.alignment
        align = {}
        if al.horizontal:
            align["horizontal"] = al.horizontal
        if al.vertical:
            align["vertical"] = al.vertical
        if al.wrap_text:
            align["wrap_text"] = True
        if align:
            fmt["alignment"] = align
    
    if cell.border:
        borders = {}
        for side in ['left', 'right', 'top', 'bottom']:
            b = getattr(cell.border, side, None)
            if b and b.style:
                borders[side] = b.style
        if borders:
            fmt["border"] = borders
    
    return fmt if fmt else None

def read_excel_range(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    preview_only: bool = False
) -> List[Dict[str, Any]]:
    """Read data from Excel range with optional preview mode"""
    wb = None
    try:
        wb = load_workbook(filepath, read_only=False)
        
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found")
            
        ws = wb[sheet_name]

        # Parse start cell
        if ':' in start_cell:
            start_cell, end_cell = start_cell.split(':')
            
        # Get start coordinates
        try:
            start_coords = parse_cell_range(f"{start_cell}:{start_cell}")
            if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Determine end coordinates
        if end_cell:
            try:
                end_coords = parse_cell_range(f"{end_cell}:{end_cell}")
                if not end_coords or not all(coord is not None for coord in end_coords[:2]):
                    raise DataError(f"Invalid end cell reference: {end_cell}")
                end_row, end_col = end_coords[0], end_coords[1]
            except ValueError as e:
                raise DataError(f"Invalid end cell format: {str(e)}")
        else:
            # If no end_cell, expand the end to the sheet's data boundary.
            # Respect the caller's start_cell (only the default A1 expands
            # the start to the true used range, matching
            # read_excel_range_with_metadata behavior).
            if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
                # Handle empty sheet
                end_row, end_col = start_row, start_col
            else:
                if start_cell.upper() == "A1":
                    start_row, start_col = ws.min_row, ws.min_column
                end_row, end_col = ws.max_row, ws.max_column

        # Validate range bounds
        if start_row > ws.max_row or start_col > ws.max_column:
            # This case can happen if start_cell is outside the used area on a sheet with data
            # or on a completely empty sheet.
            logger.warning(
                f"Start cell {start_cell} is outside the sheet's data boundary "
                f"({get_column_letter(ws.min_column)}{ws.min_row}:{get_column_letter(ws.max_column)}{ws.max_row}). "
                f"No data will be read."
            )
            return []

        data = []
        for row in range(start_row, end_row + 1):
            row_data = []
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                row_data.append(cell.value)
            if any(v is not None for v in row_data):
                data.append(row_data)

        if preview_only:
            data = data[:50]

        return data
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range: {e}")
        raise DataError(str(e))
    finally:
        if wb is not None:
            wb.close()

def write_data(
    filepath: str,
    sheet_name: Optional[str],
    data: Optional[List[List]],
    start_cell: str = "A1",
) -> Dict[str, str]:
    """Write data to Excel sheet with workbook handling
    
    Headers are handled intelligently based on context.
    """
    wb = None
    try:
        if not data:
            raise DataError("No data provided to write")

        backup_file(filepath)
        wb = load_workbook(filepath)

        # If no sheet specified, use active sheet
        if not sheet_name:
            active_sheet = wb.active
            if active_sheet is None:
                raise DataError("No active sheet found in workbook")
            sheet_name = active_sheet.title
        elif sheet_name not in wb.sheetnames:
            wb.create_sheet(sheet_name)

        ws = wb[sheet_name]

        # Validate start cell
        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        if len(data) > 0:
            _write_data_to_worksheet(ws, data, start_cell)

        wb.save(filepath)

        return {"message": f"Data written to {sheet_name}", "active_sheet": sheet_name}
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write data: {e}")
        raise DataError(str(e))
    finally:
        if wb is not None:
            wb.close()

def _write_data_to_worksheet(
    worksheet: Worksheet, 
    data: List[List], 
    start_cell: str = "A1",
) -> None:
    """Write data to worksheet with intelligent header handling"""
    try:
        if not data:
            raise DataError("No data provided to write")

        try:
            start_coords = parse_cell_range(start_cell)
            if not start_coords or not all(x is not None for x in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Write data
        for i, row in enumerate(data):
            for j, val in enumerate(row):
                worksheet.cell(row=start_row + i, column=start_col + j, value=val)
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to write worksheet data: {e}")
        raise DataError(str(e))

def read_excel_range_with_metadata(
    filepath: Path | str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    include_validation: bool = True,
    preview_only: bool = False,
    value_mode: str = "both",
    merged_fill: bool = True
) -> Dict[str, Any]:
    """Read data from Excel range with cell metadata including validation rules.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell address
        end_cell: Ending cell address (optional)
        include_validation: Whether to include validation metadata
        preview_only: Limit the read to the first 50 non-empty rows
        value_mode: both (default) | value | formula
        merged_fill: fill merged ranges with the anchor value
        
    Returns:
        Dictionary containing structured cell data with metadata
    """
    if value_mode not in VALUE_MODES:
        raise DataError(f"value_mode must be one of {VALUE_MODES}")
    wb = None
    wb_cached = None
    try:
        wb = load_workbook(filepath, read_only=False)
        wb_cached = load_workbook(filepath, data_only=True) if value_mode != "formula" else None
        
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found")
            
        ws = wb[sheet_name]

        # Parse start cell
        if ':' in start_cell:
            start_cell, end_cell = start_cell.split(':')
            
        # Get start coordinates
        try:
            start_coords = parse_cell_range(f"{start_cell}:{start_cell}")
            if not start_coords or not all(coord is not None for coord in start_coords[:2]):
                raise DataError(f"Invalid start cell reference: {start_cell}")
            start_row, start_col = start_coords[0], start_coords[1]
        except ValueError as e:
            raise DataError(f"Invalid start cell format: {str(e)}")

        # Determine end coordinates
        if end_cell:
            try:
                end_coords = parse_cell_range(f"{end_cell}:{end_cell}")
                if not end_coords or not all(coord is not None for coord in end_coords[:2]):
                    raise DataError(f"Invalid end cell reference: {end_cell}")
                end_row, end_col = end_coords[0], end_coords[1]
            except ValueError as e:
                raise DataError(f"Invalid end cell format: {str(e)}")
        else:
            # If no end_cell, use the full data range of the sheet
            if ws.max_row == 1 and ws.max_column == 1 and ws.cell(1, 1).value is None:
                # Handle empty sheet
                end_row, end_col = start_row, start_col
            else:
                # Use the sheet's own boundaries, but respect the provided start_cell
                end_row, end_col = ws.max_row, ws.max_column
                # If start_cell is 'A1' (default), we should find the true start
                if start_cell == 'A1':
                    start_row, start_col = ws.min_row, ws.min_column

        # Validate range bounds
        if start_row > ws.max_row or start_col > ws.max_column:
            # This case can happen if start_cell is outside the used area on a sheet with data
            # or on a completely empty sheet.
            logger.warning(
                f"Start cell {start_cell} is outside the sheet's data boundary "
                f"({get_column_letter(ws.min_column)}{ws.min_row}:{get_column_letter(ws.max_column)}{ws.max_row}). "
                f"No data will be read."
            )
            return {"range": f"{start_cell}:", "sheet_name": sheet_name, "cells": []}

        # Build structured cell data
        range_str = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
        range_data = {
            "range": range_str,
            "sheet_name": sheet_name,
            "value_mode": value_mode,
            "cells": []
        }

        cached_ws = wb_cached[sheet_name] if wb_cached is not None else None
        merged_map = _merged_value_map(ws, cached_ws, value_mode) if merged_fill else {}
        
        if preview_only:
            # Preview caps the number of rows (data rows only, header kept)
            max_row = min(end_row, start_row + 49)
            end_row = max_row

        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = ws.cell(row=row, column=col)
                cell_address = f"{get_column_letter(col)}{row}"

                resolved = merged_map.get((row, col))
                if resolved is None:
                    cached_cell = cached_ws.cell(row=row, column=col) if cached_ws is not None else None
                    resolved = _resolve_cell_value(cell, cached_cell, value_mode)

                cell_data = {
                    "address": cell_address,
                    **resolved,
                    "row": row,
                    "column": col
                }
                
                formatting = _get_cell_formatting(cell)
                if formatting:
                    cell_data["formatting"] = formatting
                
                # Add validation metadata if requested
                if include_validation:
                    validation_info = get_data_validation_for_cell(ws, cell_address)
                    if validation_info:
                        cell_data["validation"] = validation_info
                    else:
                        cell_data["validation"] = {"has_validation": False}
                
                range_data["cells"].append(cell_data)

        return range_data
        
    except DataError as e:
        logger.error(str(e))
        raise
    except Exception as e:
        logger.error(f"Failed to read Excel range with metadata: {e}")
        raise DataError(str(e))
    finally:
        if wb is not None:
            wb.close()
        if wb_cached is not None:
            wb_cached.close()


def _is_empty_row(ws, row, min_col, max_col) -> bool:
    """Check if a row is completely empty within the given column range."""
    for col in range(min_col, max_col + 1):
        if ws.cell(row=row, column=col).value is not None:
            return False
    return True


def _detect_data_regions(ws, start_row, start_col, end_row, end_col, min_gap_rows: int = 2) -> list:
    """Detect separate data regions within a sheet range, split by empty rows.
    
    A data region is a contiguous block of rows where no gap of min_gap_rows 
    or more consecutive empty rows exists.
    
    Returns list of (region_start_row, region_end_row) tuples.
    """
    regions = []
    region_start = None
    consecutive_empty = 0
    
    for row in range(start_row, end_row + 1):
        if _is_empty_row(ws, row, start_col, end_col):
            consecutive_empty += 1
            if consecutive_empty >= min_gap_rows and region_start is not None:
                regions.append((region_start, row - consecutive_empty))
                region_start = None
                consecutive_empty = 0
        else:
            consecutive_empty = 0
            if region_start is None:
                region_start = row
    
    if region_start is not None:
        regions.append((region_start, end_row))
    
    return regions


def _detect_header_row(ws, region_start, region_end, start_col, end_col) -> int:
    """Detect which row in a region is the header row.
    
    Returns the header row number, or region_start if no header detected.
    Uses heuristics: bold formatting, different data types from following rows.
    """
    # Check if first row has different formatting (bold) from second row
    first_bold = 0
    second_bold = 0
    for col in range(start_col, end_col + 1):
        c1 = ws.cell(row=region_start, column=col)
        c2 = ws.cell(row=min(region_start + 1, region_end), column=col)
        if c1.font and c1.font.bold:
            first_bold += 1
        if c2.font and c2.font.bold:
            second_bold += 1
    
    # If first row is predominantly bold and second is not, first row is header
    if first_bold > 0 and first_bold > second_bold:
        return region_start
    
    # If first row values look like strings and subsequent rows look like data
    non_empty_cols = 0
    string_cols = 0
    for col in range(start_col, end_col + 1):
        v1 = ws.cell(row=region_start, column=col).value
        v2 = ws.cell(row=min(region_start + 1, region_end), column=col).value if region_end > region_start else None
        if v1 is not None:
            non_empty_cols += 1
            if isinstance(v1, str) and not str(v1).startswith('='):
                string_cols += 1
    
    # If most non-empty cells in first row are strings, treat as header
    if non_empty_cols > 0 and string_cols / non_empty_cols > 0.5:
        return region_start
    
    return region_start


def read_sheet_regions(
    filepath: str,
    sheet_name: str,
    start_cell: str = "A1",
    end_cell: Optional[str] = None,
    min_gap_rows: int = 2,
    max_rows_per_region: Optional[int] = None,
    value_mode: str = "both",
    merged_fill: bool = True,
) -> dict:
    """Read a worksheet split into logical data regions (sub-tables).
    
    Automatically detects separate data regions divided by blank rows,
    identifies header rows, and returns each region as a named table.
    
    Args:
        filepath: Path to Excel file
        sheet_name: Name of worksheet
        start_cell: Starting cell (default A1)
        end_cell: Ending cell (optional, auto-detects)
        min_gap_rows: Minimum consecutive empty rows to split regions (default 2)
        max_rows_per_region: Limit rows per region for preview (None = all)
        value_mode: both (default) | value | formula
        merged_fill: fill merged ranges with the anchor value
        
    Returns:
        dict with sheet info and list of regions, each containing headers and data
    """
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter

    if value_mode not in VALUE_MODES:
        raise DataError(f"value_mode must be one of {VALUE_MODES}")

    wb = None
    wb_cached = None
    try:
        wb = load_workbook(filepath, read_only=False)
        wb_cached = load_workbook(filepath, data_only=True) if value_mode != "formula" else None
        
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found")
        
        ws = wb[sheet_name]
        
        # Parse range
        if ':' in start_cell:
            start_cell, end_cell = start_cell.split(':')
        
        coords = parse_cell_range(f"{start_cell}:{start_cell}")
        if not coords or coords[0] is None or coords[1] is None:
            raise DataError(f"Invalid start cell: {start_cell}")
        s_row, s_col = coords[0], coords[1]
        
        if end_cell:
            ecoords = parse_cell_range(f"{end_cell}:{end_cell}")
            if not ecoords or ecoords[2] is None or ecoords[3] is None:
                raise DataError(f"Invalid end cell: {end_cell}")
            e_row, e_col = ecoords[2], ecoords[3]
        else:
            e_row, e_col = ws.max_row, ws.max_column
        
        # Detect regions
        raw_regions = _detect_data_regions(ws, s_row, s_col, e_row, e_col, min_gap_rows)
        
        if not raw_regions:
            return {
                "sheet_name": sheet_name,
                "region_count": 0,
                "regions": [],
                "message": "No data regions found"
            }

        cached_ws = wb_cached[sheet_name] if wb_cached is not None else None
        merged_map = _merged_value_map(ws, cached_ws, value_mode) if merged_fill else {}

        def _value_at(row, col):
            hit = merged_map.get((row, col))
            if hit is not None:
                return hit.get("value")
            cell = ws.cell(row=row, column=col)
            cached_cell = cached_ws.cell(row=row, column=col) if cached_ws is not None else None
            return _resolve_cell_value(cell, cached_cell, value_mode).get("value")

        # Build region data
        regions_output = []
        for idx, (r_start, r_end) in enumerate(raw_regions):
            # Detect header
            header_row = _detect_header_row(ws, r_start, r_end, s_col, e_col)
            has_header = (header_row == r_start)
            
            # Collect headers (raw anchor values — do NOT fill horizontal merges)
            headers = []
            for col in range(s_col, e_col + 1):
                cell = ws.cell(row=header_row, column=col)
                cached_cell = cached_ws.cell(row=header_row, column=col) if cached_ws is not None else None
                v = _resolve_cell_value(cell, cached_cell, value_mode).get("value")
                h = str(v) if v is not None else f"Column{get_column_letter(col)}"
                headers.append(h)
            headers = _unique_headers(headers)
            
            # Collect data
            data_start = (header_row + 1) if has_header else r_start
            data_end = r_end
            if max_rows_per_region:
                data_end = min(data_end, data_start + max_rows_per_region - 1)
            
            rows = []
            for row in range(data_start, data_end + 1):
                if _is_empty_row(ws, row, s_col, e_col):
                    continue
                row_data = {}
                for col_idx, col in enumerate(range(s_col, e_col + 1)):
                    row_data[headers[col_idx]] = _value_at(row, col)
                rows.append(row_data)
            
            region_range = f"{get_column_letter(s_col)}{r_start}:{get_column_letter(e_col)}{r_end}"
            regions_output.append({
                "index": idx,
                "range": region_range,
                "has_header": has_header,
                "headers": headers,
                "row_count": len(rows),
                "data": rows
            })
        
        # Also list native Excel tables
        native_tables = []
        for tbl in ws.tables.values():
            native_tables.append({
                "name": tbl.name,
                "display_name": tbl.displayName,
                "range": tbl.ref,
            })
        
        return {
            "sheet_name": sheet_name,
            "value_mode": value_mode,
            "region_count": len(regions_output),
            "native_tables": native_tables,
            "regions": regions_output
        }
        
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to read sheet regions: {e}")
        raise DataError(str(e))
    finally:
        if wb is not None:
            wb.close()
        if wb_cached is not None:
            wb_cached.close()

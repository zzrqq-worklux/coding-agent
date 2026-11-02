import re

from openpyxl.utils import column_index_from_string

# Excel limits (2007+)
_MAX_COLUMNS = 16384  # XFD
_MAX_ROWS = 1048576

# A cell reference: optional absolute markers, column letters, row digits.
# Strictly anchored so 'A1B2', 'A1.5', 'A1 B2' etc. are rejected.
_CELL_RE = re.compile(r"^\$?([A-Za-z]+)\$?([0-9]+)$")


def _parse_cell(ref: str) -> tuple[int, int]:
    """Parse a single cell reference into (row, col) 1-based indices."""
    if not isinstance(ref, str) or not ref.strip():
        raise ValueError(f"Invalid cell reference: {ref!r}")

    match = _CELL_RE.match(ref.strip())
    if not match:
        raise ValueError(f"Invalid cell reference: {ref}")

    col_str, row_str = match.groups()
    col = column_index_from_string(col_str.upper())
    row = int(row_str)

    if col > _MAX_COLUMNS:
        raise ValueError(f"Column out of range (max {_MAX_COLUMNS}): {ref}")
    if row > _MAX_ROWS:
        raise ValueError(f"Row out of range (max {_MAX_ROWS}): {ref}")

    return row, col


def parse_cell_range(
    cell_ref: str,
    end_ref: str | None = None
) -> tuple[int, int, int | None, int | None]:
    """Parse Excel cell reference into row and column indices.

    Accepts 'A1', '$A$1', 'A1:B2', '$A$1:$B$2', or a separate end reference.
    Returns (start_row, start_col, end_row, end_col); end coordinates are
    None when only a single cell is given. Raises ValueError on invalid input.
    """
    if end_ref is not None:
        start_cell, end_cell = cell_ref, end_ref
    elif ":" in cell_ref:
        parts = cell_ref.split(":")
        if len(parts) != 2 or not parts[0] or not parts[1]:
            raise ValueError(f"Invalid cell range: {cell_ref}")
        start_cell, end_cell = parts
    else:
        start_cell, end_cell = cell_ref, None

    start_row, start_col = _parse_cell(start_cell)

    if end_cell is not None:
        end_row, end_col = _parse_cell(end_cell)
    else:
        end_row = end_col = None

    return start_row, start_col, end_row, end_col


def validate_cell_reference(cell_ref: str) -> bool:
    """Validate Excel cell reference format (e.g., 'A1', '$BC$123')"""
    if not cell_ref:
        return False
    try:
        _parse_cell(cell_ref)
        return True
    except ValueError:
        return False

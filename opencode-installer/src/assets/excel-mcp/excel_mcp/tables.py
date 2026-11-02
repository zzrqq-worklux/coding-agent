import uuid
import logging

from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from .exceptions import DataError

logger = logging.getLogger(__name__)

def create_excel_table(
    filepath: str,
    sheet_name: str,
    data_range: str,
    table_name: str | None = None,
    table_style: str = "TableStyleMedium9"
) -> dict:
    """Creates a native Excel table for the given data range.
    
    Args:
        filepath: Path to the Excel file.
        sheet_name: Name of the worksheet.
        data_range: The cell range for the table (e.g., "A1:D5").
        table_name: A unique name for the table. If not provided, a unique name is generated.
        table_style: The visual style to apply to the table.
        
    Returns:
        A dictionary with a success message and table details.
    """
    wb = None
    try:
        wb = load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found.")
            
        ws = wb[sheet_name]

        # If no table name is provided, generate a unique one
        if not table_name:
            table_name = f"Table_{uuid.uuid4().hex[:8]}"

        # Check if table name already exists
        if table_name in ws.parent.defined_names:
            raise DataError(f"Table name '{table_name}' already exists.")

        # Create the table
        table = Table(displayName=table_name, ref=data_range)
        
        # Apply style
        style = TableStyleInfo(
            name=table_style, 
            showFirstColumn=False,
            showLastColumn=False, 
            showRowStripes=True, 
            showColumnStripes=False
        )
        table.tableStyleInfo = style
        
        ws.add_table(table)
        
        wb.save(filepath)
        
        return {
            "message": f"Successfully created table '{table_name}' in sheet '{sheet_name}'.",
            "table_name": table_name,
            "range": data_range
        }

    except Exception as e:
        logger.error(f"Failed to create table: {e}")
        raise DataError(str(e))
    finally:
        if wb is not None:
            wb.close()


def list_tables_in_sheet(filepath: str, sheet_name: str) -> dict:
    """List all native Excel tables in a worksheet.
    
    Args:
        filepath: Path to the Excel file.
        sheet_name: Name of the worksheet.
        
    Returns:
        dict with table list containing name, range, and style info
    """
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found.")
        
        ws = wb[sheet_name]
        tables = []
        
        for tbl in ws.tables.values():
            info = {
                "name": tbl.name,
                "display_name": tbl.displayName,
                "range": tbl.ref,
                "header_row_count": tbl.headerRowCount,
                "totals_row_count": tbl.totalsRowCount or 0,
            }
            if tbl.tableStyleInfo:
                info["style"] = tbl.tableStyleInfo.name
            tables.append(info)
        
        wb.close()
        return {
            "sheet_name": sheet_name,
            "table_count": len(tables),
            "tables": tables
        }
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to list tables: {e}")
        raise DataError(str(e))


def read_table_data(filepath: str, sheet_name: str, table_name: str) -> dict:
    """Read data from a specific native Excel table.
    
    Args:
        filepath: Path to the Excel file.
        sheet_name: Name of the worksheet.
        table_name: Name or display name of the table.
        
    Returns:
        dict with table data as list of rows
    """
    from openpyxl import load_workbook
    
    try:
        wb = load_workbook(filepath)
        if sheet_name not in wb.sheetnames:
            raise DataError(f"Sheet '{sheet_name}' not found.")
        
        ws = wb[sheet_name]
        
        target_table = None
        for tbl in ws.tables.values():
            if tbl.name == table_name or tbl.displayName == table_name:
                target_table = tbl
                break
        
        if not target_table:
            available = [t.name for t in ws.tables.values()]
            raise DataError(
                f"Table '{table_name}' not found. Available: {available or 'none'}"
            )
        
        # Parse table range to get data
        from .cell_utils import parse_cell_range
        from openpyxl.utils import get_column_letter
        
        ref = target_table.ref
        if ':' in ref:
            start_cell, end_cell = ref.split(':')
        else:
            start_cell = end_cell = ref
        
        coords = parse_cell_range(f"{start_cell}:{end_cell}")
        if not coords:
            raise DataError(f"Invalid table range: {ref}")
        
        start_row, start_col, end_row, end_col = coords[0], coords[1], coords[2], coords[3]
        
        # Read headers and data
        headers = []
        for col in range(start_col, end_col + 1):
            cell = ws.cell(row=start_row, column=col)
            headers.append(str(cell.value) if cell.value is not None else f"Column{col}")
        
        rows = []
        for row in range(start_row + 1, end_row + 1):
            row_data = {}
            for col_idx, col in enumerate(range(start_col, end_col + 1)):
                cell = ws.cell(row=row, column=col)
                row_data[headers[col_idx]] = cell.value
            rows.append(row_data)
        
        wb.close()
        return {
            "table_name": target_table.name,
            "display_name": target_table.displayName,
            "range": ref,
            "headers": headers,
            "row_count": len(rows),
            "data": rows
        }
    except DataError:
        raise
    except Exception as e:
        logger.error(f"Failed to read table: {e}")
        raise DataError(str(e)) 

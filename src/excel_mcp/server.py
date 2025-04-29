from mcp.server.fastmcp import FastMCP
import os
import uuid
from typing import Dict, List, Optional, Any
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, ScatterChart, AreaChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.table import Table, TableStyleInfo

mcp = FastMCP("excel")

USER_AGENT = "excel-app/1.0"

class ExcelAutomation:
    def __init__(self):
        self.active_workbook = None
        self.workbook_path = None
        
    def initialize(self):
        """Initialize the Excel automation - no app instance needed with openpyxl"""
        return True
                
    def get_active_workbook(self):
        """Get information about the currently active workbook"""
        if self.active_workbook is None:
            return None
        
        sheet_names = self.active_workbook.sheetnames
        
        return {
            "name": os.path.basename(self.workbook_path) if self.workbook_path else "Untitled",
            "path": self.workbook_path,
            "sheet_count": len(sheet_names),
            "sheet_names": sheet_names
        }

# Create a global instance of our automation class
excel_automation = ExcelAutomation()

@mcp.tool()
def initialize_excel() -> bool:
    """Initialize connection to Excel."""
    return excel_automation.initialize()

@mcp.tool()
def get_workbook() -> Dict[str, Any]:
    """Get information about the currently active workbook."""
    workbook_info = excel_automation.get_active_workbook()
    if workbook_info is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    return workbook_info

@mcp.tool()
def open_workbook(file_path: str) -> Dict[str, Any]:
    """
    Open an Excel workbook from the specified path.
    
    Args:
        file_path: Full path to the Excel file (.xlsx)
        
    Returns:
        Dictionary with workbook metadata
    """
    if not os.path.exists(file_path):
        return {"error": f"File not found: {file_path}"}
    
    try:
        excel_automation.active_workbook = openpyxl.load_workbook(file_path)
        excel_automation.workbook_path = file_path
        
        return {
            "name": os.path.basename(file_path),
            "path": file_path,
            "sheet_count": len(excel_automation.active_workbook.sheetnames),
            "sheet_names": excel_automation.active_workbook.sheetnames
        }
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def get_sheets() -> List[Dict[str, Any]]:
    """
    Get a list of all sheets in the active workbook.
    
    Returns:
        List of sheet metadata
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    sheets = []
    
    try:
        for i, sheet_name in enumerate(wb.sheetnames):
            sheet = wb[sheet_name]
            
            # Get sheet dimensions
            max_row = sheet.max_row
            max_col = sheet.max_column
            
            sheets.append({
                "id": str(i),
                "index": i,
                "name": sheet_name,
                "rows": max_row,
                "columns": max_col
            })
        
        return sheets
    except Exception as e:
        return {"error": f"Error getting sheets: {str(e)}"}

@mcp.tool()
def create_workbook() -> Dict[str, Any]:
    """
    Create a new Excel workbook.
    
    Returns:
        Dictionary containing new workbook metadata
    """
    try:
        excel_automation.active_workbook = openpyxl.Workbook()
        excel_automation.workbook_path = ""
        
        return {
            "name": "New Workbook",
            "path": "",
            "sheet_count": len(excel_automation.active_workbook.sheetnames),
            "sheet_names": excel_automation.active_workbook.sheetnames
        }
    except Exception as e:
        return {"error": str(e)}

@mcp.tool()
def save_workbook(path: str = None) -> Dict[str, Any]:
    """
    Save the active workbook to disk.
    
    Args:
        path: Optional path to save the file (if None, save to current location)
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    try:
        save_path = path if path else excel_automation.workbook_path
        
        # If this is a new workbook without a path, we need a path
        if not save_path:
            return {"error": "Save path must be specified for new workbooks"}
        
        excel_automation.active_workbook.save(save_path)
        
        # Update the path in our records
        excel_automation.workbook_path = save_path
        
        return {
            "success": True, 
            "path": save_path
        }
    except Exception as e:
        return {"success": False, "error": str(e)}

@mcp.tool()
def close_workbook() -> Dict[str, Any]:
    """
    Close the active workbook.
    
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    try:
        # With openpyxl, we just remove it from our tracking
        excel_automation.active_workbook = None
        excel_automation.workbook_path = None
        return {"success": True}
    except Exception as e:
        return {"success": False, "error": str(e)}

@mcp.tool()
def add_sheet(name: str = None) -> Dict[str, Any]:
    """
    Add a new sheet to the workbook.
    
    Args:
        name: Name for the new sheet (optional)
            
    Returns:
        Information about the new sheet
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Generate a default name if none provided
        if name is None:
            base_name = "Sheet"
            counter = 1
            name = f"{base_name}{counter}"
            
            # Find a unique name
            while name in wb.sheetnames:
                counter += 1
                name = f"{base_name}{counter}"
        
        # Add new sheet
        sheet = wb.create_sheet(name)
        sheet_index = wb.sheetnames.index(name)
        
        return {
            "id": str(sheet_index),
            "index": sheet_index,
            "name": name,
            "rows": 0,
            "columns": 0
        }
    except Exception as e:
        return {"error": f"Error adding sheet: {str(e)}"}

@mcp.tool()
def get_sheet_data(sheet_index: int = None, sheet_name: str = None, 
                  start_row: int = 1, end_row: int = None,
                  start_col: int = 1, end_col: int = None) -> Dict[str, Any]:
    """
    Get data from a sheet.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        start_row: First row to retrieve (1-based)
        end_row: Last row to retrieve (1-based, optional)
        start_col: First column to retrieve (1-based)
        end_col: Last column to retrieve (1-based, optional)
        
    Returns:
        Dictionary containing sheet data
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Determine data range
        max_row = sheet.max_row
        max_col = sheet.max_column
        
        if end_row is None or end_row > max_row:
            end_row = max_row
            
        if end_col is None or end_col > max_col:
            end_col = max_col
        
        # Extract data
        data = []
        for row in range(start_row, end_row + 1):
            row_data = []
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                row_data.append(str(cell.value) if cell.value is not None else "")
            data.append(row_data)
        
        # Get column headers (if available)
        headers = []
        if start_row > 1:
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=1, column=col)
                headers.append(str(cell.value) if cell.value is not None else f"Column {get_column_letter(col)}")
        
        return {
            "success": True,
            "sheet_name": sheet.title,
            "sheet_index": wb.sheetnames.index(sheet.title),
            "start_row": start_row,
            "end_row": end_row,
            "start_col": start_col,
            "end_col": end_col,
            "headers": headers if headers else None,
            "data": data
        }
    except Exception as e:
        return {"error": f"Error retrieving sheet data: {str(e)}"}

@mcp.tool()
def update_cell(sheet_index: int = None, sheet_name: str = None,
               row: int = 1, col: int = 1, value: str = "") -> Dict[str, Any]:
    """
    Update the value of a cell.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        row: Row index (1-based)
        col: Column index (1-based)
        value: New cell value
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Update cell value
        sheet.cell(row=row, column=col).value = value
        
        return {
            "success": True,
            "message": f"Cell {get_column_letter(col)}{row} updated successfully"
        }
    except Exception as e:
        return {"error": f"Error updating cell: {str(e)}"}

@mcp.tool()
def update_range(sheet_index: int = None, sheet_name: str = None,
                start_row: int = 1, start_col: int = 1, 
                data: List[List[str]] = None) -> Dict[str, Any]:
    """
    Update a range of cells with data.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        start_row: Starting row index (1-based)
        start_col: Starting column index (1-based)
        data: 2D array of values to insert
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    if data is None or not data:
        return {"error": "No data provided"}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Update range with data
        for i, row_data in enumerate(data):
            for j, value in enumerate(row_data):
                sheet.cell(row=start_row + i, column=start_col + j).value = value
        
        end_row = start_row + len(data) - 1
        end_col = start_col + len(data[0]) - 1
        
        return {
            "success": True,
            "message": f"Range {get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row} updated successfully"
        }
    except Exception as e:
        return {"error": f"Error updating range: {str(e)}"}

@mcp.tool()
def format_cells(sheet_index: int = None, sheet_name: str = None,
                start_row: int = 1, end_row: int = 1,
                start_col: int = 1, end_col: int = 1,
                bold: bool = None, italic: bool = None,
                font_size: int = None, font_color: str = None,
                bg_color: str = None, alignment: str = None) -> Dict[str, Any]:
    """
    Format a range of cells.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        start_row: Starting row index (1-based)
        end_row: Ending row index (1-based)
        start_col: Starting column index (1-based)
        end_col: Ending column index (1-based)
        bold: Set text to bold
        italic: Set text to italic
        font_size: Set font size
        font_color: Set font color (hex code)
        bg_color: Set background color (hex code)
        alignment: Set text alignment ('left', 'center', 'right')
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Map alignment string to openpyxl alignment
        alignment_map = {
            'left': Alignment(horizontal='left'),
            'center': Alignment(horizontal='center'),
            'right': Alignment(horizontal='right')
        }
        
        # Apply formatting to each cell in the range
        for row in range(start_row, end_row + 1):
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                
                # Create a new font based on the current one
                font = Font(
                    name=cell.font.name,
                    size=font_size if font_size is not None else cell.font.size,
                    bold=bold if bold is not None else cell.font.bold,
                    italic=italic if italic is not None else cell.font.italic,
                    color=font_color if font_color is not None else cell.font.color
                )
                
                cell.font = font
                
                # Set background color if specified
                if bg_color is not None:
                    cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                
                # Set alignment if specified
                if alignment is not None and alignment.lower() in alignment_map:
                    cell.alignment = alignment_map[alignment.lower()]
        
        return {
            "success": True,
            "message": f"Range {get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row} formatted successfully"
        }
    except Exception as e:
        return {"error": f"Error formatting cells: {str(e)}"}

@mcp.tool()
def add_chart(sheet_index: int = None, sheet_name: str = None,
             chart_type: str = "bar", title: str = "Chart",
             data_range_start_row: int = 1, data_range_end_row: int = 10,
             data_range_start_col: int = 1, data_range_end_col: int = 2,
             categories_in_first_column: bool = True,
             target_cell: str = "E5") -> Dict[str, Any]:
    """
    Add a chart to a worksheet.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        chart_type: Type of chart ('bar', 'line', 'pie', 'scatter', 'area')
        title: Chart title
        data_range_start_row: Starting row of data range (1-based)
        data_range_end_row: Ending row of data range (1-based)
        data_range_start_col: Starting column of data range (1-based)
        data_range_end_col: Ending column of data range (1-based)
        categories_in_first_column: Whether categories are in the first column
        target_cell: Cell reference where to place the chart (e.g., 'E5')
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Create the appropriate chart type
        chart_type_lower = chart_type.lower()
        if chart_type_lower == 'bar':
            chart = BarChart()
        elif chart_type_lower == 'line':
            chart = LineChart()
        elif chart_type_lower == 'pie':
            chart = PieChart()
        elif chart_type_lower == 'scatter':
            chart = ScatterChart()
        elif chart_type_lower == 'area':
            chart = AreaChart()
        else:
            return {"error": f"Unsupported chart type: {chart_type}. Supported types: bar, line, pie, scatter, area"}
        
        # Set chart title
        chart.title = title
        
        # Define data and categories references
        if chart_type_lower == 'scatter':
            # For scatter charts, we need to handle X and Y values differently
            for col_idx in range(data_range_start_col + 1, data_range_end_col + 1):
                # X values (from first column)
                x_values = Reference(
                    sheet, 
                    min_row=data_range_start_row, 
                    max_row=data_range_end_row, 
                    min_col=data_range_start_col
                )
                
                # Y values (from subsequent columns)
                y_values = Reference(
                    sheet, 
                    min_row=data_range_start_row, 
                    max_row=data_range_end_row, 
                    min_col=col_idx
                )
                
                # Create series
                series = chart.series.append(y_values, x_values)
                
                # Try to get series name from header row if available
                if data_range_start_row > 1:
                    header_cell = sheet.cell(row=data_range_start_row-1, column=col_idx)
                    if header_cell.value:
                        series.title = header_cell.value
        else:
            # For other chart types
            if categories_in_first_column:
                # Categories are in the first column
                categories = Reference(sheet, min_row=data_range_start_row, max_row=data_range_end_row, min_col=data_range_start_col)
                data = Reference(sheet, min_row=data_range_start_row-1, max_row=data_range_end_row, 
                                min_col=data_range_start_col+1, max_col=data_range_end_col)
                # Add data to chart
                chart.add_data(data, titles_from_data=True)
                chart.set_categories(categories)
            else:
                # First row contains categories
                data = Reference(sheet, min_row=data_range_start_row, max_row=data_range_end_row, 
                               min_col=data_range_start_col, max_col=data_range_end_col)
                # Add data to chart
                chart.add_data(data, titles_from_data=True)
        
        # Add chart to worksheet
        sheet.add_chart(chart, target_cell)
        
        return {
            "success": True,
            "message": f"{chart_type.capitalize()} chart added successfully to {sheet.title} at position {target_cell}",
            "chart_info": {
                "type": chart_type,
                "title": title,
                "location": target_cell
            }
        }
    except Exception as e:
        return {"error": f"Error adding chart: {str(e)}"}
        
@mcp.tool()
def rename_sheet(sheet_index: int = None, sheet_name: str = None, new_name: str = None) -> Dict[str, Any]:
    """
    Rename a worksheet.
    
    Args:
        sheet_index: Index of the sheet to rename (0-based, optional if sheet_name provided)
        sheet_name: Current name of the sheet (optional if sheet_index provided)
        new_name: New name for the sheet
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    if new_name is None:
        return {"error": "New sheet name must be provided"}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Check if the new name already exists
        if new_name in wb.sheetnames:
            return {"error": f"Sheet name '{new_name}' already exists"}
        
        # Rename the sheet
        sheet.title = new_name
        
        return {
            "success": True,
            "message": f"Sheet renamed to '{new_name}' successfully"
        }
    except Exception as e:
        return {"error": f"Error renaming sheet: {str(e)}"}

@mcp.tool()
def delete_sheet(sheet_index: int = None, sheet_name: str = None) -> Dict[str, Any]:
    """
    Delete a worksheet.
    
    Args:
        sheet_index: Index of the sheet to delete (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet to delete (optional if sheet_index provided)
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Check if this is the only sheet
        if len(wb.sheetnames) <= 1:
            return {"error": "Cannot delete the only sheet in the workbook"}
        
        # Delete the sheet
        wb.remove(sheet)
        
        return {
            "success": True,
            "message": "Sheet deleted successfully",
            "remaining_sheets": wb.sheetnames
        }
    except Exception as e:
        return {"error": f"Error deleting sheet: {str(e)}"}

@mcp.tool()
def add_formula(sheet_index: int = None, sheet_name: str = None,
               row: int = 1, col: int = 1, formula: str = "") -> Dict[str, Any]:
    """
    Add a formula to a cell.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        row: Row index (1-based)
        col: Column index (1-based)
        formula: Excel formula to add (e.g., "=SUM(A1:A10)")
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    if not formula:
        return {"error": "Formula must be provided"}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Make sure formula starts with =
        if not formula.startswith('='):
            formula = '=' + formula
        
        # Add formula to cell
        sheet.cell(row=row, column=col).value = formula
        
        return {
            "success": True,
            "message": f"Formula added to cell {get_column_letter(col)}{row} successfully"
        }
    except Exception as e:
        return {"error": f"Error adding formula: {str(e)}"}

@mcp.tool()
def merge_cells(sheet_index: int = None, sheet_name: str = None,
               start_row: int = 1, end_row: int = 1,
               start_col: int = 1, end_col: int = 2) -> Dict[str, Any]:
    """
    Merge a range of cells.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        start_row: Starting row index (1-based)
        end_row: Ending row index (1-based)
        start_col: Starting column index (1-based)
        end_col: Ending column index (1-based)
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Create range string
        range_str = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
        
        # Merge cells
        sheet.merge_cells(range_str)
        
        return {
            "success": True,
            "message": f"Cells {range_str} merged successfully"
        }
    except Exception as e:
        return {"error": f"Error merging cells: {str(e)}"}

@mcp.tool()
def unmerge_cells(sheet_index: int = None, sheet_name: str = None,
                 start_row: int = 1, end_row: int = 1,
                 start_col: int = 1, end_col: int = 2) -> Dict[str, Any]:
    """
    Unmerge a range of cells.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        start_row: Starting row index (1-based)
        end_row: Ending row index (1-based)
        start_col: Starting column index (1-based)
        end_col: Ending column index (1-based)
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Create range string
        range_str = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
        
        # Unmerge cells
        sheet.unmerge_cells(range_str)
        
        return {
            "success": True,
            "message": f"Cells {range_str} unmerged successfully"
        }
    except Exception as e:
        return {"error": f"Error unmerging cells: {str(e)}"}

@mcp.tool()
def set_column_width(sheet_index: int = None, sheet_name: str = None,
                    col: int = 1, width: float = 10) -> Dict[str, Any]:
    """
    Set the width of a column.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        col: Column index (1-based)
        width: Column width in characters
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Set column width
        sheet.column_dimensions[get_column_letter(col)].width = width
        
        return {
            "success": True,
            "message": f"Column {get_column_letter(col)} width set to {width} successfully"
        }
    except Exception as e:
        return {"error": f"Error setting column width: {str(e)}"}

@mcp.tool()
def set_row_height(sheet_index: int = None, sheet_name: str = None,
                  row: int = 1, height: float = 15) -> Dict[str, Any]:
    """
    Set the height of a row.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        row: Row index (1-based)
        height: Row height in points
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # Set row height
        sheet.row_dimensions[row].height = height
        
        return {
            "success": True,
            "message": f"Row {row} height set to {height} successfully"
        }
    except Exception as e:
        return {"error": f"Error setting row height: {str(e)}"}

@mcp.tool()
def format_as_table(sheet_index: int = None, sheet_name: str = None,
                   table_name: str = "Table1", style_name: str = "TableStyleLight9",
                   start_row: int = 1, end_row: int = None,
                   start_col: int = 1, end_col: int = None,
                   show_first_column: bool = False,
                   show_last_column: bool = False,
                   show_row_stripes: bool = True,
                   show_column_stripes: bool = False) -> Dict[str, Any]:
    """
    Format a range of cells as an Excel table.
    
    Args:
        sheet_index: Index of the sheet (0-based, optional if sheet_name provided)
        sheet_name: Name of the sheet (optional if sheet_index provided)
        table_name: Name for the table (must be unique in workbook)
        style_name: Excel table style name (e.g., TableStyleLight9, TableStyleMedium2)
        start_row: Starting row index (1-based)
        end_row: Ending row index (1-based, optional - defaults to max row with data)
        start_col: Starting column index (1-based)
        end_col: Ending column index (1-based, optional - defaults to max column with data)
        show_first_column: Whether to highlight the first column
        show_last_column: Whether to highlight the last column
        show_row_stripes: Whether to show row stripes
        show_column_stripes: Whether to show column stripes
        
    Returns:
        Status of the operation
    """
    if excel_automation.active_workbook is None:
        return {"error": "No active workbook. Please open or create a workbook first."}
    
    wb = excel_automation.active_workbook
    
    try:
        # Get the sheet by index or name
        if sheet_name is not None:
            if sheet_name not in wb.sheetnames:
                return {"error": f"Sheet not found: {sheet_name}"}
            sheet = wb[sheet_name]
        elif sheet_index is not None:
            if sheet_index < 0 or sheet_index >= len(wb.sheetnames):
                return {"error": f"Invalid sheet index: {sheet_index}. Valid range is 0-{len(wb.sheetnames)-1}"}
            sheet = wb[wb.sheetnames[sheet_index]]
        else:
            # Default to active sheet
            sheet = wb.active
        
        # If end_row or end_col not provided, find the last row/column with data
        if end_row is None:
            # Find the last row with data starting from start_row
            for row in range(sheet.max_row, start_row - 1, -1):
                has_data = False
                for col in range(start_col, sheet.max_column + 1):
                    if sheet.cell(row=row, column=col).value is not None:
                        has_data = True
                        break
                if has_data:
                    end_row = row
                    break
            # If no data found, use start_row
            if end_row is None:
                end_row = start_row
        
        if end_col is None:
            # Find the last column with data
            for col in range(sheet.max_column, start_col - 1, -1):
                has_data = False
                for row in range(start_row, end_row + 1):
                    if sheet.cell(row=row, column=col).value is not None:
                        has_data = True
                        break
                if has_data:
                    end_col = col
                    break
            # If no data found, use start_col
            if end_col is None:
                end_col = start_col
        
        # Create the reference string for the table
        ref = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
        
        # Check if a table with this name already exists
        for existing_table in sheet.tables.values():
            if existing_table.name == table_name:
                return {"error": f"Table with name '{table_name}' already exists in this workbook"}
        
        # Create the table
        tab = Table(displayName=table_name, ref=ref)
        
        # Set table style
        style = TableStyleInfo(
            name=style_name,
            showFirstColumn=show_first_column,
            showLastColumn=show_last_column,
            showRowStripes=show_row_stripes,
            showColumnStripes=show_column_stripes
        )
        tab.tableStyleInfo = style
        
        # Add the table to the worksheet
        sheet.add_table(tab)
        
        return {
            "success": True,
            "message": f"Range {ref} formatted as table '{table_name}' successfully",
            "table_info": {
                "name": table_name,
                "range": ref,
                "style": style_name
            }
        }
    except Exception as e:
        return {"error": f"Error formatting as table: {str(e)}"}

@mcp.tool()
def create_xlwings_pivot_table(
    filepath: str = None,
    sheet_name: str = None,
    data_range: str = None,
    row_fields: List[str] = None,
    value_fields: List[str] = None,
    column_fields: List[str] = None,
    filter_fields: List[str] = None,
    agg_func: str = "Sum",
    target_sheet_name: str = None,
    target_cell: str = "A3",
    pivot_name: str = None,
    include_totals: bool = True,
    style_name: str = "PivotStyleMedium9"
) -> Dict[str, Any]:
    """
    Create a pivot table using xlwings (requires Excel installation).
    
    Args:
        filepath: Path to Excel file (required if no active workbook)
        sheet_name: Name of source worksheet
        data_range: Source data range reference (e.g., "A1:D100")
        row_fields: Fields for row labels
        value_fields: Fields for values
        column_fields: Optional fields for column labels
        filter_fields: Optional fields for report filters
        agg_func: Aggregation function (Sum, Count, Average, Min, Max, etc.)
        target_sheet_name: Name of worksheet to place pivot table (creates new if None)
        target_cell: Cell reference for pivot table position
        pivot_name: Name for the pivot table (optional)
        include_totals: Whether to include row and column totals
        style_name: Excel pivot table style name
        
    Returns:
        Status of the operation
    """
    try:
        from .pivot_xlwings import create_pivot_table_xlwings, is_xlwings_available
    except ImportError:
        return {"error": "Failed to import pivot_xlwings module"}
    
    if not is_xlwings_available():
        return {"error": "xlwings is not installed or Excel is not available. Please install xlwings with 'pip install xlwings'"}
    
    # If no filepath provided but we have an active workbook, save it to a temp file
    if not filepath and excel_automation.active_workbook:
        import tempfile
        temp_dir = tempfile.mkdtemp()
        filepath = os.path.join(temp_dir, "temp_workbook.xlsx")
        excel_automation.active_workbook.save(filepath)
    
    if not filepath:
        return {"error": "No filepath provided and no active workbook"}
    
    if not sheet_name:
        return {"error": "Sheet name must be provided"}
    
    if not data_range:
        return {"error": "Data range must be provided"}
    
    if not row_fields:
        return {"error": "Row fields must be provided"}
    
    if not value_fields:
        return {"error": "Value fields must be provided"}
    
    try:
        result = create_pivot_table_xlwings(
            filepath=filepath,
            sheet_name=sheet_name,
            data_range=data_range,
            rows=row_fields,
            values=value_fields,
            columns=column_fields,
            filters=filter_fields,
            agg_func=agg_func,
            target_sheet_name=target_sheet_name,
            target_cell=target_cell,
            pivot_name=pivot_name,
            include_totals=include_totals,
            style_name=style_name,
            visible=False  # Run Excel in background
        )
        
        # If we created a temp file and have an active workbook, reload it
        if excel_automation.active_workbook and filepath.endswith("temp_workbook.xlsx"):
            excel_automation.active_workbook = openpyxl.load_workbook(filepath)
        
        return {
            "success": True,
            "message": result["message"],
            "pivot_info": result["details"]
        }
    except Exception as e:
        return {"error": f"Error creating pivot table with xlwings: {str(e)}"}


@mcp.tool()
def refresh_xlwings_pivot_table(
    filepath: str = None,
    pivot_name: str = None
) -> Dict[str, Any]:
    """
    Refresh an existing pivot table using xlwings.
    
    Args:
        filepath: Path to Excel file (required if no active workbook)
        pivot_name: Name of the pivot table to refresh
        
    Returns:
        Status of the operation
    """
    try:
        from .pivot_xlwings import refresh_pivot_table_xlwings, is_xlwings_available
    except ImportError:
        return {"error": "Failed to import pivot_xlwings module"}
    
    if not is_xlwings_available():
        return {"error": "xlwings is not installed or Excel is not available. Please install xlwings with 'pip install xlwings'"}
    
    # If no filepath provided but we have an active workbook, save it to a temp file
    if not filepath and excel_automation.active_workbook:
        import tempfile
        temp_dir = tempfile.mkdtemp()
        filepath = os.path.join(temp_dir, "temp_workbook.xlsx")
        excel_automation.active_workbook.save(filepath)
    
    if not filepath:
        return {"error": "No filepath provided and no active workbook"}
    
    if not pivot_name:
        return {"error": "Pivot table name must be provided"}
    
    try:
        result = refresh_pivot_table_xlwings(
            filepath=filepath,
            pivot_name=pivot_name,
            visible=False  # Run Excel in background
        )
        
        # If we created a temp file and have an active workbook, reload it
        if excel_automation.active_workbook and filepath.endswith("temp_workbook.xlsx"):
            excel_automation.active_workbook = openpyxl.load_workbook(filepath)
        
        return {
            "success": True,
            "message": result["message"]
        }
    except Exception as e:
        return {"error": f"Error refreshing pivot table with xlwings: {str(e)}"}


@mcp.tool()
def modify_xlwings_pivot_table(
    filepath: str = None,
    pivot_name: str = None,
    add_row_fields: List[str] = None,
    remove_row_fields: List[str] = None,
    add_value_fields: List[str] = None,
    remove_value_fields: List[str] = None,
    add_column_fields: List[str] = None,
    remove_column_fields: List[str] = None,
    add_filter_fields: List[str] = None,
    remove_filter_fields: List[str] = None,
    change_agg_funcs: Dict[str, str] = None,
    include_totals: bool = None,
    style_name: str = None
) -> Dict[str, Any]:
    """
    Modify an existing pivot table using xlwings.
    
    Args:
        filepath: Path to Excel file (required if no active workbook)
        pivot_name: Name of the pivot table to modify
        add_row_fields: Fields to add as row labels
        remove_row_fields: Fields to remove from row labels
        add_value_fields: Fields to add as values
        remove_value_fields: Fields to remove from values
        add_column_fields: Fields to add as column labels
        remove_column_fields: Fields to remove from column labels
        add_filter_fields: Fields to add as report filters
        remove_filter_fields: Fields to remove from report filters
        change_agg_funcs: Dict mapping field names to new aggregation functions
        include_totals: Whether to include row and column totals
        style_name: Excel pivot table style name
        
    Returns:
        Status of the operation
    """
    try:
        from .pivot_xlwings import modify_pivot_table_xlwings, is_xlwings_available
    except ImportError:
        return {"error": "Failed to import pivot_xlwings module"}
    
    if not is_xlwings_available():
        return {"error": "xlwings is not installed or Excel is not available. Please install xlwings with 'pip install xlwings'"}
    
    # If no filepath provided but we have an active workbook, save it to a temp file
    if not filepath and excel_automation.active_workbook:
        import tempfile
        temp_dir = tempfile.mkdtemp()
        filepath = os.path.join(temp_dir, "temp_workbook.xlsx")
        excel_automation.active_workbook.save(filepath)
    
    if not filepath:
        return {"error": "No filepath provided and no active workbook"}
    
    if not pivot_name:
        return {"error": "Pivot table name must be provided"}
    
    try:
        result = modify_pivot_table_xlwings(
            filepath=filepath,
            pivot_name=pivot_name,
            add_rows=add_row_fields,
            remove_rows=remove_row_fields,
            add_values=add_value_fields,
            remove_values=remove_value_fields,
            add_columns=add_column_fields,
            remove_columns=remove_column_fields,
            add_filters=add_filter_fields,
            remove_filters=remove_filter_fields,
            change_agg_func=change_agg_funcs,
            include_totals=include_totals,
            style_name=style_name,
            visible=False  # Run Excel in background
        )
        
        # If we created a temp file and have an active workbook, reload it
        if excel_automation.active_workbook and filepath.endswith("temp_workbook.xlsx"):
            excel_automation.active_workbook = openpyxl.load_workbook(filepath)
        
        return {
            "success": True,
            "message": result["message"]
        }
    except Exception as e:
        return {"error": f"Error modifying pivot table with xlwings: {str(e)}"}

def main():
    mcp.run(transport="stdio")

if __name__ == "__main__":
    main()

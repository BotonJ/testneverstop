# table_injector.py (formerly table_injector_fixed.py): 强制复制表格并注入Summary

from openpyxl import load_workbook
from openpyxl.styles import Font, Border, Side, PatternFill, Alignment # Import necessary styles for robust copying

def force_copy(ws_tpl, ws_combined, start_row, start_col, nrows, ncols, target_col):
    """
    Forces copy of cell values and styles from template to target worksheet,
    bypassing issues with empty cells.
    """
    for i in range(nrows):
        for j in range(ncols):
            source_cell = ws_tpl.cell(start_row + i, start_col + j)
            target_cell = ws_combined.cell(i + 1, target_col + j)

            # Copy value
            target_cell.value = source_cell.value

            # Copy styles safely
            if source_cell.has_style:
                target_cell.font = Font(
                    name=source_cell.font.name,
                    size=source_cell.font.size,
                    bold=source_cell.font.bold,
                    italic=source_cell.font.italic,
                    vertAlign=source_cell.font.vertAlign,
                    underline=source_cell.font.underline,
                    strike=source_cell.font.strike,
                    color=source_cell.font.color
                )
                target_cell.border = Border(
                    left=Side(style=source_cell.border.left.style, color=source_cell.border.left.color),
                    right=Side(style=source_cell.border.right.style, color=source_cell.border.right.color),
                    top=Side(style=source_cell.border.top.style, color=source_cell.border.top.color),
                    bottom=Side(style=source_cell.border.bottom.style, color=source_cell.border.bottom.color)
                )
                target_cell.fill = PatternFill(
                    patternType=source_cell.fill.patternType,
                    fgColor=source_cell.fill.fgColor,
                    bgColor=source_cell.fill.bgColor
                )
                target_cell.number_format = source_cell.number_format
                target_cell.alignment = Alignment(
                    horizontal=source_cell.alignment.horizontal,
                    vertical=source_cell.alignment.vertical,
                    text_rotation=source_cell.alignment.text_rotation,
                    wrap_text=source_cell.alignment.wrap_text,
                    shrink_to_fit=source_cell.alignment.shrink_to_fit,
                    indent=source_cell.alignment.indent
                )
            # Copy dimensions if needed (e.g., column width, row height) - more complex, might be handled by template

def inject_tables_and_summary(output_path, template_path, summary_values, dest_path):
    wb_out = load_workbook(output_path,data_only=True)
    wb_tpl = load_workbook(template_path)

    # Ensure 'Sheet1' exists in the template
    if "Sheet1" not in wb_tpl.sheetnames:
        raise ValueError(f"Template file '{template_path}' must contain a sheet named 'Sheet1'.")
    ws_tpl = wb_tpl["Sheet1"]

    # Create or get the combined sheet
    if "汇总区块" in wb_out.sheetnames:
        ws_combined = wb_out["汇总区块"]
    else:
        ws_combined = wb_out.create_sheet("汇总区块")

    # Force copy table regions from the template.
    # The current hardcoded ranges (10 rows, 4 cols) need to be accurate to the template structure.
    force_copy(ws_tpl, ws_combined, start_row=1, start_col=1, nrows=10, ncols=4, target_col=1)    # table1
    force_copy(ws_tpl, ws_combined, start_row=1, start_col=6, nrows=10, ncols=4, target_col=9)    # table2
    force_copy(ws_tpl, ws_combined, start_row=10, start_col=1, nrows=10, ncols=4, target_col=17)  # table3

    # Inject summary values into the combined sheet
    base_row = 30
    base_col = 1
    ws_combined.cell(base_row, base_col, "文字说明字段：")
    for idx, (k, v) in enumerate(summary_values.items()):
        ws_combined.cell(base_row + idx + 1, base_col, k)
        # Ensure values are not None before setting, or convert to string if desired
        ws_combined.cell(base_row + idx + 1, base_col + 1, v if v is not None else "N/A")

    wb_out.save(dest_path)
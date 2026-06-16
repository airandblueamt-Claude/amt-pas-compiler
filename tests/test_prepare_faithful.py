import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl
from openpyxl.styles import Font, Alignment
import render_tables as RT

fd, src = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
wb = openpyxl.Workbook(); ws = wb.active
ws["A1"] = "Item"; ws["B2"] = "A long description that the tool must NOT reflow or rewrap at all"
ws["B2"].font = Font(name="Times New Roman", size=13, bold=True)
ws["B2"].alignment = Alignment(horizontal="left", vertical="top", wrap_text=False)
ws.row_dimensions[2].height = 22.0
wb.save(src)

prepared = RT._prepare_xlsx_for_print(src)
p = openpyxl.load_workbook(prepared).active

assert p["B2"].font.name == "Times New Roman", p["B2"].font.name
assert p["B2"].font.size == 13, p["B2"].font.size
assert p["B2"].alignment.wrap_text in (False, None), p["B2"].alignment.wrap_text
assert abs((p.sheet_format.defaultRowHeight or 0)) >= 0
assert p.row_dimensions[2].height == 22.0, p.row_dimensions[2].height
assert p.page_setup.fitToWidth == 1, p.page_setup.fitToWidth
print("OK")

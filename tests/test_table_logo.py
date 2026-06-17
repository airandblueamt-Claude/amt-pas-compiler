"""
A BOQ table renders faithfully (content untouched) AND carries the stamped AMT
logo header. Run: python3 tests/test_table_logo.py
"""
import os, sys, inspect, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import openpyxl, render_tables as RT

# render_table takes no 'brand' param (faithful: branding is overlaid, not re-typeset)
assert "brand" not in inspect.signature(RT.render_table).parameters, "no brand param"

fd, src = tempfile.mkstemp(suffix=".xlsx"); os.close(fd)
wb = openpyxl.Workbook(); ws = wb.active
ws["A1"] = "Item"; ws["B1"] = "Desc"; ws["A2"] = 1; ws["B2"] = "Ceiling speaker 6W"
wb.save(src)

out = tempfile.mkstemp(suffix=".pdf")[1]
res, eng = RT.render_table(src, out, "BOQ", "REF-001 v.00")

import fitz
d = fitz.open(res)
assert d.page_count >= 1
# table content is preserved (faithful)
assert "Ceiling speaker 6W" in d[0].get_text(), "table text must survive"
# the AMT logo header is stamped on the page (an image is present)...
assert len(d[0].get_images()) >= 1, "logo header should be stamped on the table page"
# ...and it sits in the reserved top band, not over the table body
top_band = RT.A.logo_header_reserve() + 8
in_band = [r for img in d[0].get_images(full=True)
           for r in d[0].get_image_rects(img[0])
           if r.y0 < top_band]
assert in_band, "stamped logo should sit in the reserved top margin"
print("OK", eng)

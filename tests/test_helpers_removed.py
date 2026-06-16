import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import render_tables as RT
for name in ("_col_points", "_rows_with_images", "_estimate_row_height",
             "TABLE_FONT", "TABLE_FONT_SIZE"):
    assert not hasattr(RT, name), f"{name} should be removed"
print("OK")

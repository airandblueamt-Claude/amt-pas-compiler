import os, sys, inspect
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import amt_common as A, normalize as NORM
for name in ("stamp_pdf", "_stamp_header", "HEADER_BAND",
             "TABLE_TOP_RESERVE", "TABLE_BOTTOM_RESERVE"):
    assert not hasattr(A, name), f"amt_common.{name} should be removed"
assert "top_reserve" not in inspect.signature(NORM.to_a4_portrait).parameters
print("OK")

import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import assemble as ASM
assert not hasattr(ASM, "_is_branded"), "remove _is_branded"
assert not hasattr(ASM, "DEFAULT_UNBRANDED"), "remove DEFAULT_UNBRANDED"
print("OK")

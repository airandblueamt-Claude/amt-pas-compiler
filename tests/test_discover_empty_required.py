"""
Regression test for discover.py empty-section handling.

A *required* append section that exists but is empty must be a hard ERROR
(so the build refuses), while an *optional* empty section is only a WARNING
(placeholder inserted). Guards against the reintroduced bug where both paths
appended to warnings. Run: python3 tests/test_discover_empty_required.py
"""
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "scripts"))
import discover  # noqa: E402


def _make_section(root: str, name: str, files: dict[str, str]) -> None:
    d = os.path.join(root, name)
    os.makedirs(d, exist_ok=True)
    for fname, content in files.items():
        with open(os.path.join(d, fname), "w") as fh:
            fh.write(content)


def main() -> int:
    sections = [
        {"no": 1, "prefix": "1", "kind": "append", "optional": False,
         "en": "Required Docs", "ar": "—"},
        {"no": 2, "prefix": "2", "kind": "append", "optional": True,
         "en": "Optional Docs", "ar": "—"},
    ]

    with tempfile.TemporaryDirectory() as root:
        # Section 1 (required) present but EMPTY; section 2 (optional) EMPTY.
        os.makedirs(os.path.join(root, "1-Required Docs"))
        os.makedirs(os.path.join(root, "2-Optional Docs"))

        m = discover.discover(root, sections)
        errs = " | ".join(m["errors"])
        warns = " | ".join(m["warnings"])

        assert "Section 1" in errs, f"required-empty should ERROR, got errors={m['errors']}"
        assert "Section 1" not in warns, f"required-empty must not be only a warning: {warns}"
        assert "Section 2" in warns, f"optional-empty should WARN, got warnings={m['warnings']}"
        assert "Section 2" not in errs, f"optional-empty must not ERROR: {m['errors']}"

        # Sanity: a required section WITH content produces no error.
        _make_section(root, "1-Required Docs", {"x.pdf": "%PDF-1.4 stub"})
        m2 = discover.discover(root, sections)
        assert not any("Section 1" in e for e in m2["errors"]), \
            f"required-with-content must not error: {m2['errors']}"

    print("PASS — required empty -> error, optional empty -> warning, content -> ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())

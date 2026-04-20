"""
test_overflow.py
────────────────
Tests for the overflow safety check mechanism.

Covers:
  1. Clean slide  → no violations reported
  2. Overflowing slide → violations detected with correct edge/px info
  3. on_overflow="warn"  → warning issued, save() still completes
  4. on_overflow="raise" → OverflowError raised before .pptx is written
  5. on_overflow="ignore" → checks skipped entirely, save() completes
  6. check_slide() standalone → works independently of save()
  7. check_all_slides() → batch check before committing to save()
"""

import warnings
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from src.html_to_pptx import HtmlToPptxConverter, SlideCheckResult


# ── Sample HTML ───────────────────────────────────────────────────────────────

CLEAN_HTML = """<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { width:1280px; height:720px; overflow:hidden;
         background:#1E2761; display:flex; align-items:center;
         justify-content:center; font-family:Georgia,serif; color:#fff; }
  h1 { font-size:52px; }
</style></head>
<body><h1>Clean Slide — no overflow</h1></body>
</html>"""

# This slide has a div that is intentionally taller than 720px
OVERFLOW_HTML = """<!DOCTYPE html>
<html><head><meta charset="UTF-8">
<style>
  * { margin:0; padding:0; box-sizing:border-box; }
  body { width:1280px; height:720px; overflow:hidden;
         background:#f2f2f2; font-family:Arial,sans-serif; }
  .tall-box {
    width: 600px;
    height: 900px;   /* overflows by 180px */
    background: #cc3333;
    color: white;
    padding: 20px;
    font-size: 24px;
  }
  .wide-box {
    position: absolute;
    top: 100px;
    left: 1100px;    /* right edge = 1100+300 = 1400px, overflows by 120px */
    width: 300px;
    height: 100px;
    background: #3366cc;
    color: white;
  }
</style></head>
<body>
  <div class="tall-box">I overflow bottom by ~180px</div>
  <div class="wide-box">I overflow right by ~120px</div>
</body>
</html>"""


def print_result(result: SlideCheckResult) -> None:
    print(result.summary())
    if result.has_overflow:
        for v in result.violations:
            print(f"    tag={v.tag!r}  edges={v.overflow_edges}  px={v.overflow_px}")


# ── Tests ─────────────────────────────────────────────────────────────────────

def test_clean_slide():
    print("\n── Test 1: clean slide ──────────────────────────────────────────")
    c = HtmlToPptxConverter(on_overflow="raise")
    c.add_slide_from_string(CLEAN_HTML, title="Clean")
    results = c.check_all_slides()
    assert not results[0].has_overflow, "Expected no overflow on clean slide"
    print_result(results[0])
    print("PASS")


def test_overflow_detected():
    print("\n── Test 2: overflow detected ────────────────────────────────────")
    c = HtmlToPptxConverter(on_overflow="ignore")   # don't raise, just inspect
    c.add_slide_from_string(OVERFLOW_HTML, title="Overflow")
    results = c.check_all_slides()
    r = results[0]
    assert r.has_overflow, "Expected overflow violations"
    assert any("bottom" in v.overflow_edges for v in r.violations), \
        "Expected a bottom overflow"
    assert any("right" in v.overflow_edges for v in r.violations), \
        "Expected a right overflow"
    print_result(r)
    print("PASS")


def test_on_overflow_warn(tmp_path):
    print("\n── Test 3: on_overflow='warn' ───────────────────────────────────")
    c = HtmlToPptxConverter(on_overflow="warn")
    c.add_slide_from_string(OVERFLOW_HTML, title="Overflow warn")
    out = tmp_path / "warn_test.pptx"
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        saved = c.save(str(out))
    assert out.exists(), "save() should still write the file"
    assert any("overflow" in str(w.message).lower() for w in caught), \
        "Expected a warning to be issued"
    print(f"  Warning issued: {caught[0].message}")
    print(f"  File written:   {saved}")
    print("PASS")


def test_on_overflow_raise(tmp_path):
    print("\n── Test 4: on_overflow='raise' ──────────────────────────────────")
    c = HtmlToPptxConverter(on_overflow="raise")
    c.add_slide_from_string(OVERFLOW_HTML, title="Overflow raise")
    out = tmp_path / "raise_test.pptx"
    try:
        c.save(str(out))
        assert False, "Expected OverflowError to be raised"
    except OverflowError as e:
        print(f"  OverflowError raised as expected:\n  {str(e)[:200]}")
    print("PASS")


def test_on_overflow_ignore(tmp_path):
    print("\n── Test 5: on_overflow='ignore' ─────────────────────────────────")
    c = HtmlToPptxConverter(on_overflow="ignore")
    c.add_slide_from_string(OVERFLOW_HTML, title="Overflow ignore")
    out = tmp_path / "ignore_test.pptx"
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        saved = c.save(str(out))
    assert out.exists(), "save() should write the file"
    assert not any("overflow" in str(w.message).lower() for w in caught), \
        "Expected no warning when on_overflow='ignore'"
    print(f"  File written (no warning): {saved}")
    print("PASS")


def test_check_slide_standalone():
    print("\n── Test 6: check_slide() standalone ────────────────────────────")
    c = HtmlToPptxConverter()
    c.add_slide_from_string(CLEAN_HTML,    title="Clean")
    c.add_slide_from_string(OVERFLOW_HTML, title="Overflow")

    r0 = c.check_slide(0)
    r1 = c.check_slide(1)

    assert not r0.has_overflow, "Slide 0 should be clean"
    assert r1.has_overflow,     "Slide 1 should overflow"
    print_result(r0)
    print_result(r1)
    print("PASS")


def test_check_all_before_save():
    print("\n── Test 7: check_all_slides() before save ───────────────────────")
    c = HtmlToPptxConverter(on_overflow="ignore")   # skip auto-check in save()
    c.add_slide_from_string(CLEAN_HTML,    title="Good")
    c.add_slide_from_string(OVERFLOW_HTML, title="Bad")
    c.add_slide_from_string(CLEAN_HTML,    title="Good again")

    results = c.check_all_slides()
    overflowing = [r for r in results if r.has_overflow]
    print(f"  {len(results)} slides checked, {len(overflowing)} overflowing")
    for r in results:
        print_result(r)

    assert len(overflowing) == 1
    assert overflowing[0].slide_index == 1
    print("PASS")


if __name__ == "__main__":
    import tempfile
    tmp = Path(tempfile.mkdtemp())

    test_clean_slide()
    test_overflow_detected()
    test_on_overflow_warn(tmp)
    test_on_overflow_raise(tmp)
    test_on_overflow_ignore(tmp)
    test_check_slide_standalone()
    test_check_all_before_save()

    print("\n══ All tests passed ══")

"""
example_usage.py
────────────────
Demonstrates the HtmlToPptxConverter library end-to-end.

Run:
    cd html_to_pptx
    python example_usage.py
"""

import os
import argparse
from pathlib import Path
from html_to_pptx import HtmlToPptxConverter


def main():
    print("=== HtmlToPptxConverter ===\n")

    converter = HtmlToPptxConverter(on_overflow="raise")

    # ── Configurations ────────────────────────────────────────────────────────────
    INPUT_DIR = Path(__file__).parent / "inputs"
    OUTPUT_DIR = Path(__file__).parent / "outputs"

    parser = argparse.ArgumentParser()
    parser.add_argument("--target-dir", required=True)
    parser.add_argument("--output-filename", default="sample.pptx")
    args = parser.parse_args()

    if args.target_dir is None:
        raise Exception("Provide target_dir")
    elif args.output_filename is None:
        raise Exception("Provide output_filename")
    else:
        SLIDES_DIR = INPUT_DIR / args.target_dir / "slides"
        OUTPUT = OUTPUT_DIR / args.output_filename

    # ── Add slides by files────────────────────────────────────────────────────────────
    for f in os.listdir(SLIDES_DIR):
        if os.path.isfile(os.path.join(SLIDES_DIR, f)):
            converter.add_slide_from_file(
                os.path.join(SLIDES_DIR, f),
                # title="Title Slide",
                # notes="This library walks through the HTML-to-PPTX pipeline.",
            )

    # ── Add slides by string────────────────────────────────────────────────────────────
    # From string (simulates HTML returned by an LLM and confirmed by the user)
    slide2_html = (SLIDES_DIR / "slide2.html").read_text(encoding="utf-8")
    converter.add_slide_from_string(
        slide2_html,
        title="Architecture",
        notes="Walk through the 4 steps: HTML → Playwright → python-pptx → .pptx",
    )

    print(f"Slides queued : {len(converter)}")

    # ── Optional: demonstrate reorder / remove API ────────────────────────────
    # e.g. move slide 2 to position 0:
    #   converter.reorder_slides([1, 0, 2])
    # e.g. remove slide at index 1:
    #   converter.remove_slide(1)

    # ── Export ────────────────────────────────────────────────────────────────
    print(f"Rendering & saving → {OUTPUT}")
    saved = converter.save(OUTPUT)
    print(f"✓ Saved: {saved}\n")


if __name__ == "__main__":
    main()

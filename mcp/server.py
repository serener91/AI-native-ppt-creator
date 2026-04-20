"""
html_to_ppt MCP server
──────────────────────
Remote FastMCP HTTP server that exposes the HtmlToPptxConverter library.

Workflow for an LLM client:
  1. Read   resource  html://guidelines        → learn slide constraints
  2. Call   tool      check_slides([...])      → validate HTML (optional)
  3. Call   tool      generate_pptx([...])     → render & save .pptx
  4. Call   tool      get_pptx(path)           → retrieve file as base64

Run:
    python -m src.html_to_ppt.server
"""

from __future__ import annotations

import asyncio
import base64
from pathlib import Path
from typing import Any

from fastmcp import FastMCP
from pydantic import BaseModel

from src.html_to_pptx import HtmlToPptxConverter
from config import settings

# ── Server instance ────────────────────────────────────────────────────────────

mcp = FastMCP(
    name="html_to_ppt_mcp",
    instructions=(
        "Convert HTML slides to a Microsoft PowerPoint (.pptx) file. "
        "Start by reading the html://guidelines resource to understand slide constraints, "
        "then use check_slides to validate your HTML, and finally generate_pptx to produce the file."
    ),
)

# ── Helpers ────────────────────────────────────────────────────────────────────

_GUIDELINES_PATH = Path(__file__).parent / "src" / "html_guidelines.md"


class SlideInput(BaseModel):
    """One slide's worth of data."""
    html: str
    title: str | None = None
    notes: str | None = None


# ── Tools ──────────────────────────────────────────────────────────────────────

@mcp.tool(
    name="check_slides",
    description=(
        "Validate a list of HTML slides for content overflow before generating the PPTX. "
        "Returns a report per slide: whether overflow was detected and which elements are responsible. "
        "No file is written. Call this before generate_pptx to catch layout problems early."
    ),
    annotations={
        "title": "Check slides for overflow",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def check_slides(slides: list[SlideInput]) -> list[dict[str, Any]]:
    """Run overflow detection on every slide and return structured results."""

    def _run() -> list[dict[str, Any]]:
        converter = HtmlToPptxConverter(on_overflow="ignore")
        for s in slides:
            converter.add_slide_from_string(s.html, title=s.title, notes=s.notes)
        results = converter.check_all_slides()
        return [
            {
                "slide_index": r.slide_index,
                "has_overflow": r.has_overflow,
                "violations": [
                    {
                        "tag": v.tag,
                        "text_snippet": v.text_snippet,
                        "overflow_edges": v.overflow_edges,
                        "overflow_px": v.overflow_px,
                    }
                    for v in r.violations
                ],
            }
            for r in results
        ]

    return await asyncio.to_thread(_run)


@mcp.tool(
    name="generate_pptx",
    description=(
        "Render a list of HTML slides to a .pptx file using headless Chromium. "
        "Each slide is screenshotted at 1280×720 px and embedded in PowerPoint. "
        "Returns the saved file path and slide count."
    ),
    annotations={
        "title": "Generate PowerPoint",
        "readOnlyHint": False,
        "destructiveHint": False,
        "idempotentHint": False,
        "openWorldHint": False,
    },
)
async def generate_pptx(
    slides: list[SlideInput],
    filename: str = "output.pptx",
    on_overflow: str = settings.on_overflow,
) -> dict[str, Any]:
    """Convert HTML slides to a .pptx file and save it to the output directory."""

    def _run() -> dict[str, Any]:
        settings.output_dir.mkdir(parents=True, exist_ok=True)
        output_path = settings.output_dir / filename

        converter = HtmlToPptxConverter(on_overflow=on_overflow)  # type: ignore[arg-type]
        for s in slides:
            converter.add_slide_from_string(s.html, title=s.title, notes=s.notes)

        saved = converter.save(output_path)
        return {"path": saved, "slide_count": len(slides)}

    return await asyncio.to_thread(_run)


@mcp.tool(
    name="get_pptx",
    description=(
        "Read a previously generated .pptx file and return it as a base64-encoded string. "
        "Use the path returned by generate_pptx."
    ),
    annotations={
        "title": "Download PPTX as base64",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def get_pptx(path: str) -> dict[str, str]:
    """Return a .pptx file as base64 so the client can download it."""

    def _read() -> dict[str, str]:
        data = Path(path).read_bytes()
        return {
            "path": path,
            "base64": base64.b64encode(data).decode("ascii"),
            "mime_type": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        }

    return await asyncio.to_thread(_read)


@mcp.tool(
    name="get_html_guidelines",
    description=(
        "Rules for generating self-contained 1280×720 px HTML slides that render correctly "
        "in the headless Chromium renderer. Read this before writing any slide HTML."
    ),
    annotations={
        "title": "Guideline to write a HTML",
        "readOnlyHint": True,
        "destructiveHint": False,
        "idempotentHint": True,
        "openWorldHint": False,
    },
)
async def get_html_guidelines():
    return _GUIDELINES_PATH.read_text(encoding="utf-8")


# ── Resource ───────────────────────────────────────────────────────────────────

@mcp.resource(
    uri="html://guidelines",
    name="Slide HTML guidelines",
    description=(
        "Rules for generating self-contained 1280×720 px HTML slides that render correctly "
        "in the headless Chromium renderer. Read this before writing any slide HTML."
    ),
    mime_type="text/markdown",
)
def html_guidelines() -> str:
    """Return the HTML slide authoring guidelines."""
    return _GUIDELINES_PATH.read_text(encoding="utf-8")


# ── Prompt ─────────────────────────────────────────────────────────────────────
# Call the `html://guidelines` resource
@mcp.prompt(
    name="create_presentation",
    description="Generate a complete PowerPoint presentation from a topic description.",
)
def create_presentation(topic: str, num_slides: str = "5", style: str = "professional") -> str:
    # MCP prompt arguments are always strings; cast here for arithmetic use
    _n = int(num_slides)
    return f"""You are creating a {style} PowerPoint presentation on: **{topic}**

## Steps

1. **Read constraints first**
   Call the `get_html_guidelines` tool. Follow every rule it describes — especially:
   - Canvas is exactly 1280 × 720 px with `overflow: hidden`
   - No external URLs; inline CSS/JS only; system fonts only; base64 images only

2. **Design {_n} slides**
   Plan a logical narrative arc:
   - Slide 1: Title / hero slide
   - Slides 2–{_n - 1}: Content slides (use varied layouts from the guidelines)
   - Slide {_n}: Summary or call-to-action

3. **Generate HTML for each slide**
   Produce complete, self-contained HTML per slide.
   Use the minimal template from the guidelines as your starting point.

4. **Validate (recommended)**
   Call `check_slides` with all your HTML strings.
   Fix any overflow violations before proceeding.

5. **Generate the PPTX**
   Call `generate_pptx` with:
   - `slides`: list of {{html, title, notes}} objects
   - `filename`: a descriptive name like "{topic.lower().replace(' ', '_')}.pptx"

6. **Report back**
   Tell the user the saved file path and offer to adjust any slides.
"""


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    settings.output_dir.mkdir(parents=True, exist_ok=True)
    mcp.run(transport="http", host=settings.host, port=settings.port)

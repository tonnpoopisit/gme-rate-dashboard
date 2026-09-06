#!/usr/bin/env python3
"""
render_table_png.py

Renders a JSON table dump (as produced by excel_range_to_json.ps1) as a PNG,
styled to match Excel's actual rendered appearance - the JSON's colors/bold
already account for conditional formatting (read via .DisplayFormat), so this
script just needs to lay them out faithfully.

Exists because Excel's own Range.CopyPicture + clipboard route, when driven via
COM automation, reliably produces corrupted captures (right image size, wrong
content) rather than an honest screenshot of the range. Rendering the already-
extracted data as HTML and screenshotting that with a real browser engine sidesteps
the problem entirely.

Usage:
    python render_table_png.py <table_data.json> <output.png> [--compact]
        [--font-size N] [--pad-v N] [--pad-h N] [--colors N]

--compact renders at device_scale_factor=1 (instead of 2) and re-encodes as an
adaptive-palette PNG via Pillow, cutting file size roughly 7-9x versus the
default with no visible quality loss for flat-color/text tables like this one.
Use it when the PNG needs to fit a hard payload-size limit - e.g. Power
Platform's "manual trigger" webhook URLs (the environment.api.powerplatform.com
kind, as opposed to classic Logic App HTTP triggers) reject request bodies over
28KB, and a base64-encoded default-quality render of an 11-row table alone runs
60-90KB, well over that even before any JSON wrapper overhead.

--font-size/--pad-v/--pad-h/--colors override the HTML table's font size (px),
cell padding (px, vertical/horizontal), and (with --compact) the adaptive
palette's color count - defaults (14/4/10/32) are unchanged from the original
Thailand table, which fits comfortably under 28KB with just --compact. A much
taller table (e.g. Laos's 5-block, 32-row ranking table) doesn't fit under that
same limit at the default font/padding even with --compact's scale=1 and
32-color quantization - shrinking the *render* itself (smaller font/padding,
so Playwright rasterizes fewer pixels to begin with) keeps text crisp, unlike
shrinking file size after the fact by resizing an already-rendered PNG, which
blurs it. Laos's report uses --font-size 11 --pad-v 2 --pad-h 5 --colors 16,
verified to land around 25KB total once wrapped in the Adaptive Card JSON.
"""
import json
import sys
from pathlib import Path

from playwright.sync_api import sync_playwright


# Distinct from every background color used across the Thailand/Laos tables
# (white, header gray, GME salmon, provider blue/yellow/light-blue/light-green)
# so a stale-row border reads as clearly different rather than blending in.
STALE_BORDER_COLOR = "#E06600"


def build_html(data: dict, font_size: int = 14, pad_v: int = 4, pad_h: int = 10) -> str:
    rows_html = []
    any_stale = False
    for row in data["rows"]:
        cells_html = []
        for cell in row:
            # skip=True cells are ones a merged cell above/beside them
            # already covers via rowspan/colspan (e.g. a "Country" label
            # merged down 9 rows) - an HTML table renders that the same way
            # Excel does, by the covered rows simply not emitting that
            # column, not by repeating or blanking it out.
            if cell.get("skip"):
                continue
            is_stale = cell.get("stale", False)
            any_stale = any_stale or is_stale
            border = f"2px solid {STALE_BORDER_COLOR}" if is_stale else "1px solid #BFBFBF"
            style = (
                f"background:{cell['bg']};color:{cell['fg']};"
                f"font-weight:{'bold' if cell['bold'] else 'normal'};"
                f"text-align:{cell.get('align', 'left')};"
                f"border:{border};"
            )
            rowspan = cell.get("rowspan", 1)
            colspan = cell.get("colspan", 1)
            span_attrs = ""
            if rowspan > 1:
                span_attrs += f' rowspan="{rowspan}"'
            if colspan > 1:
                span_attrs += f' colspan="{colspan}"'
            text = cell["text"] or "&nbsp;"
            cells_html.append(f'<td{span_attrs} style="{style}">{text}</td>')
        rows_html.append("<tr>" + "".join(cells_html) + "</tr>")

    caption = ""
    if any_stale:
        caption = (
            f'<div style="font-family:Calibri,Arial,sans-serif;font-size:{max(font_size - 3, 9)}px;'
            f'color:{STALE_BORDER_COLOR};padding-top:4px;">'
            f"&#9888; Amber border = this row's fetch failed this run - showing its last known value, not a fresh rate."
            f"</div>"
        )

    return f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"><style>
  body {{ margin: 0; padding: 0; background: white; }}
  /* display:inline-block so #capture shrinks to its content's actual width
     instead of a block div's default of stretching to fill the viewport -
     without this, the screenshot includes a big blank margin to the right
     of the table out to the viewport edge (this happened for real: the
     stale-row caption below the table was added by wrapping table+caption
     in a plain <div>, which silently introduced this). */
  #capture {{ display: inline-block; }}
  table {{ border-collapse: collapse; font-family: Calibri, Arial, sans-serif; font-size: {font_size}px; }}
  td {{ padding: {pad_v}px {pad_h}px; white-space: nowrap; vertical-align: middle; }}
</style></head><body>
<div id="capture">
<table>{"".join(rows_html)}</table>
{caption}
</div>
</body></html>"""


def main() -> None:
    args = sys.argv[1:]
    if len(args) < 2:
        sys.exit(
            "Usage: python render_table_png.py <table_data.json> <output.png> "
            "[--compact] [--font-size N] [--pad-v N] [--pad-h N] [--colors N]"
        )

    json_path, png_path = Path(args[0]), Path(args[1])
    rest = args[2:]

    compact = "--compact" in rest
    font_size, pad_v, pad_h, colors = 14, 4, 10, 32

    def _take_int(flag, default):
        if flag in rest:
            return int(rest[rest.index(flag) + 1])
        return default

    font_size = _take_int("--font-size", font_size)
    pad_v = _take_int("--pad-v", pad_v)
    pad_h = _take_int("--pad-h", pad_h)
    colors = _take_int("--colors", colors)

    # utf-8-sig: PowerShell's `Set-Content -Encoding UTF8` writes a BOM that
    # plain utf-8 decoding chokes on; utf-8-sig strips it if present and
    # behaves like plain utf-8 otherwise.
    data = json.loads(json_path.read_text(encoding="utf-8-sig"))

    html_path = json_path.with_suffix(".html")
    html_path.write_text(build_html(data, font_size, pad_v, pad_h), encoding="utf-8")

    scale = 1 if compact else 2
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width": 1200, "height": 900}, device_scale_factor=scale)
        page.goto(html_path.resolve().as_uri())
        # #capture wraps the table + the stale-row caption (when present) so
        # the caption isn't cropped out - locating just "table" would only
        # grab the table itself.
        page.locator("#capture").screenshot(path=str(png_path))
        browser.close()

    if compact:
        from PIL import Image

        img = Image.open(png_path)
        # Flat-color, mostly-text tables like this one use very few distinct
        # colors, so an adaptive palette loses no visible quality while
        # letting PNG's compression do far less work. Fewer colors (e.g. 16)
        # trades a bit of color-accuracy margin for a smaller file - safe as
        # long as the table doesn't use many more distinct colors than that.
        img.convert("P", palette=Image.ADAPTIVE, colors=colors).save(png_path, optimize=True)

    print(f"Saved: {png_path} ({png_path.stat().st_size:,} bytes)")


if __name__ == "__main__":
    main()

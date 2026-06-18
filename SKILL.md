---
name: html-deck-export
description: Convert local HTML slide decks or presentation preview pages into 16:9 PDF and PowerPoint PPTX outputs. Use when the user asks to export an HTML presentation preview, roadshow deck, brochure-style deck, web PPT, or front-end slide deck to PDF/PPTX, especially when the workflow needs high-fidelity PDF rendering, separate background and mask image layers, and editable text boxes in PPTX.
---

# HTML Deck Export

## Overview

Use this skill to export a local HTML slide deck into:

- A high-fidelity 16:9 PDF rendered by Playwright.
- A 16:9 PPTX with each slide built from separate background image, mask image, and editable text boxes.
- Layer assets for inspection or reuse.

The bundled script is tuned for brochure-style roadshow preview decks and should be the default path for HTML decks that use full-page background images and overlay text.

## Expected Input

Use `scripts/html_deck_hybrid_export.py` for HTML files that follow this structure:

- Slides are `section.slide` elements inside the HTML.
- Each slide has a background image marked as `img.slide-bg`.
- Optional mask/overlay styling is represented by page CSS, commonly `.wash`.
- The deck is designed for 16:9 display.

If the HTML differs, inspect the DOM first and update the text selectors in `extract_text_boxes()` before exporting the PPTX. The PDF path is usually more tolerant because it renders the page directly.

## Export Workflow

1. Locate the input `.html` file and decide the output directory.
2. Run the bundled script:

```powershell
python "C:\Users\Administrator\.codex\skills\html-deck-export\scripts\html_deck_hybrid_export.py" `
  "D:\path\to\deck.html" `
  -o "D:\path\to\export-output"
```

3. Review `hybrid-export-report.txt` for slide count, PDF page count, PPTX path, text box count, and Chrome/Edge path.
4. For visual verification, render or open at least one early page and one late page from the PDF. If the PDF looks right but the PPTX text is imperfect, treat the PDF as the visual source of truth and the PPTX as the editable working file.

## Outputs

The script writes these files into the output directory:

- `html-deck-16x9-playwright.pdf`: high-fidelity PDF export.
- `html-deck-16x9-editable-text.pptx`: PowerPoint file with editable top-layer text.
- `editable-text-boxes.json`: extracted text box geometry and styles.
- `hybrid-export-report.txt`: export summary and validation facts.
- `layers/background/`: cropped full-slide background images.
- `layers/mask/`: generated mask/overlay images.
- `layers/full/`: full rendered slide screenshots.
- `layers/source-background/`: original background image copies.

## Quality Rules

- Keep PDF output as the layout authority because it preserves the browser-rendered design most faithfully.
- Keep PPTX text editable whenever possible. Do not flatten all text into a single foreground image unless the user explicitly prefers visual fidelity over editability.
- Preserve 16:9 sizing. The default script uses `1600x900` raster layers and a `16 x 9 in` PPTX canvas.
- If title sizes or font placement look wrong in PPTX, first inspect `editable-text-boxes.json`; then adjust selector coverage or the font-size scale in `make_editable_pptx()`.
- If the exported PPTX misses text, extend the selector list inside `extract_text_boxes()`.

## Dependencies

The script expects Python packages already used by this workflow:

```powershell
pip install beautifulsoup4 pillow python-pptx pypdf playwright
python -m playwright install chromium
```

It also looks for Chrome or Edge in common Windows install paths. Pass `--chrome "path\to\chrome.exe"` if needed.

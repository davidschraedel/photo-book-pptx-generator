# Photo Book PPTX Generator

Turns a folder of photos into a print-ready photo book. Drop images in, run one script, get a PowerPoint with four photos per page—fitted and centered automatically. Export to PDF and submit to a photo book printing service.

Supports common image formats and Canon RAW (`.cr2`). Up to 280 photos (70 pages × 4).

## Setup

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
uv sync
```

## Usage

1. Add photos to `photos/current/` (sorted by filename).
2. Run from the project root:

```bash
uv run generate_book.py
```

3. Open `output/book_current.pptx`, export as PDF, and submit to your print service.
4. Move used photos from `photos/current/` to `photos/used/`.

The layout comes from `book_template.pptx` (included in the repo).

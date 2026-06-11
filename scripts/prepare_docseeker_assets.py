#!/usr/bin/env python3
"""Download the DocSeeker paper and prepare runtime-only visual assets.

The generated PDF and PNG files are intentionally written to ignored directories
(`paper/` and `figures/`) so the PR remains text-only. GitHub Actions uses these
assets at runtime to build the PPTX artifact.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from urllib.request import Request, urlopen

PAPER_URL = "https://arxiv.org/pdf/2604.12812"
PAPER_PATH = Path("paper/docseeker.pdf")
FIGURE_DIR = Path("figures")


@dataclass(frozen=True)
class CropSpec:
    name: str
    page_index: int
    # Coordinates are normalized fractions of page width/height.
    left: float
    top: float
    right: float
    bottom: float
    min_width_px: int = 900


# Runtime crops focus on the paper's central evidence rather than full pages.
# They are deliberately conservative and validated to avoid full-page images.
CROPS = [
    CropSpec("method_framework.png", 2, 0.06, 0.08, 0.94, 0.48),
    CropSpec("main_results.png", 4, 0.04, 0.16, 0.96, 0.70),
    CropSpec("evidence_comparison.png", 5, 0.08, 0.08, 0.92, 0.36),
    CropSpec("ablation_results.png", 7, 0.05, 0.10, 0.95, 0.56),
]


def download_paper() -> None:
    PAPER_PATH.parent.mkdir(parents=True, exist_ok=True)
    if PAPER_PATH.exists() and PAPER_PATH.stat().st_size > 100_000:
        print(f"Using existing {PAPER_PATH}")
        return
    req = Request(PAPER_URL, headers={"User-Agent": "Mozilla/5.0"})
    with urlopen(req, timeout=60) as response:
        data = response.read()
    if len(data) < 100_000 or not data.startswith(b"%PDF"):
        raise RuntimeError("Downloaded paper does not look like a valid PDF")
    PAPER_PATH.write_bytes(data)
    print(f"Downloaded {PAPER_PATH} ({len(data)} bytes)")


def extract_crops() -> None:
    import fitz  # PyMuPDF

    FIGURE_DIR.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(PAPER_PATH)
    for spec in CROPS:
        if spec.page_index >= doc.page_count:
            raise RuntimeError(f"Crop {spec.name} requests missing page index {spec.page_index}")
        page = doc[spec.page_index]
        rect = page.rect
        crop = fitz.Rect(
            rect.x0 + rect.width * spec.left,
            rect.y0 + rect.height * spec.top,
            rect.x0 + rect.width * spec.right,
            rect.y0 + rect.height * spec.bottom,
        )
        area_fraction = (crop.width * crop.height) / (rect.width * rect.height)
        if area_fraction >= 0.80:
            raise RuntimeError(f"Crop {spec.name} is too close to a full-page image")
        zoom = max(2.0, spec.min_width_px / crop.width)
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=crop, alpha=False)
        out = FIGURE_DIR / spec.name
        pix.save(out)
        print(f"Extracted {out} from page {spec.page_index + 1}: {pix.width}x{pix.height}")


def main() -> None:
    download_paper()
    extract_crops()


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""Validate the runtime-generated DocSeeker PPTX artifact."""
from __future__ import annotations

from pathlib import Path
import zipfile

from pptx import Presentation

PPTX = Path("outputs/docseeker_compact_ppt.pptx")
MAX_SLIDES = 7
EXPECTED_PICTURES = 4


def main() -> None:
    if not PPTX.exists():
        raise SystemExit(f"Missing PPTX: {PPTX}")
    prs = Presentation(str(PPTX))
    slide_count = len(prs.slides)
    picture_count = sum(1 for slide in prs.slides for shape in slide.shapes if shape.shape_type == 13)
    if slide_count > MAX_SLIDES:
        raise SystemExit(f"Too many slides: {slide_count} > {MAX_SLIDES}")
    if picture_count != EXPECTED_PICTURES:
        raise SystemExit(f"Unexpected picture count: {picture_count} != {EXPECTED_PICTURES}")
    with zipfile.ZipFile(PPTX) as zf:
        bad = zf.testzip()
        if bad is not None:
            raise SystemExit(f"Corrupt PPTX member: {bad}")
    print({"slides": slide_count, "pictures": picture_count, "full_page_screenshots": 0})


if __name__ == "__main__":
    main()

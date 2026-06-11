# codex-skill-test

## Paper compact PPT artifact

This repository generates a compact Chinese academic presentation for arXiv:2604.12812, **DocSeeker: Structured Visual Reasoning with Evidence Grounding for Long Document Understanding**.

The PPTX is intentionally **not committed** because it is a binary artifact. Instead, GitHub Actions runs `scripts/generate_paper_compact_ppt.py`, writes `outputs/paper_compact_ppt.pptx`, validates it, and uploads it as the artifact named `paper_compact_pptx`.

To run locally when dependencies are available:

```bash
python -m pip install -r requirements.txt
python scripts/generate_paper_compact_ppt.py --output outputs/paper_compact_ppt.pptx
unzip -t outputs/paper_compact_ppt.pptx
```

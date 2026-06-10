# DocSeeker compact PPT generation summary

- Output PPTX: `outputs/attention_is_all_you_need_compact.pptx`.
- Paper URL/PDF source: https://arxiv.org/pdf/2604.12812 (`DocSeeker: Structured Visual Reasoning with Evidence Grounding for Long Document Understanding`, arXiv v5, 20 pages).
- Deck slide count: 7.
- Embedded visual assets: 6 SVG figure/table crops or screenshot-style extracts in `outputs/docseeker_assets/`.
- Included original-paper visuals: Figure 1 overview/ALR, Figure 2 training framework/EGRA, Table 1 main results, Table 2 full-doc vs evidence-only, Figure 3 document-length analysis, Figure 4 RAG integration, Table 3 data ablation, Table 4 resolution/efficiency ablation, Table 5 EviGRPO reward ablation.
- Validation: PPTX zip package opened and parsed; 7 slide XML files found; 6 media assets embedded; each slide relationship XML is present.
- GitHub Actions artifact fallback: `.github/workflows/generate-docseeker-ppt.yml` regenerates this PPT and uploads artifact `attention_is_all_you_need_compact_pptx`.
- PR policy: `outputs/attention_is_all_you_need_compact.pptx` is intentionally ignored and not committed because GitHub PR creation/review does not support this binary PPTX.
- Download path: run the GitHub Actions workflow `Generate DocSeeker compact PPT`, then download artifact `attention_is_all_you_need_compact_pptx` from the workflow run page.

# codex-skill-test

## MARDoc compact PPT artifact

This repository generates a Chinese compact academic PPT for arXiv:2606.05749,
"MARDoc: A Memory-Aware Refinement Agent Framework for Multimodal Long Document QA".

The PPTX binary is intentionally not committed. Instead, the GitHub Actions
workflow builds `outputs/paper_compact_ppt.pptx` and uploads it as the artifact
named `paper_compact_pptx`.

To download the deck:

1. Open the **Generate paper compact PPT** workflow run in GitHub Actions.
2. Wait for the `build-pptx` job to finish successfully.
3. Download the `paper_compact_pptx` artifact from the run summary.

The generation script validates that the PPTX can be reopened with
`python-pptx`, has no more than seven slides, embeds paper screenshots, and
passes ZIP integrity checks.

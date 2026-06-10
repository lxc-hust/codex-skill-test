---
name: paper-compact-ppt
description: Create a concise, high-density Chinese presentation from an academic paper PDF or paper URL. Use when the user asks to read a paper and make a PPT/slides for presentation, especially when they want the deck to include background, motivation, method, experiments, ablations, analysis, conclusions, and original paper figure/table screenshots. Prefer 6-7 slides unless the user requests otherwise.
---

# Paper Compact PPT

## Goal

Turn an academic paper into a polished Chinese presentation deck that is concise, information-dense, and presentation-ready.

Default output:
- `.pptx`
- 6-7 slides unless the user specifies another length
- Chinese slide text
- Original paper screenshots for key figures/tables whenever useful

## Workflow

1. Download or locate the paper PDF.
2. Read the paper enough to identify:
   - background and task setting
   - motivation and pain points
   - proposed method
   - main framework figure
   - experimental setup
   - main result table
   - ablation studies
   - analysis experiments
   - limitations and conclusion
3. Extract or screenshot original key figures/tables from the PDF.
4. Build a compact PPT that combines:
   - dense Chinese summary bullets
   - original paper figure/table screenshots
   - short takeaways next to each screenshot
5. Validate that the PPT opens and has the intended slide count.

## Slide Structure

Prefer this 7-slide structure:

1. **Title / One-Sentence Contribution**
   - paper title
   - one-sentence contribution
   - 3-4 headline metrics or claims

2. **Background & Motivation**
   - task difficulty
   - limitations of prior methods
   - core problem the paper solves

3. **Method Overview**
   - include original framework figure if available
   - explain pipeline in 3-4 concise blocks

4. **Main Experiments**
   - include original main result table
   - summarize strongest results and comparisons

5. **Ablation / Fine-Grained Analysis I**
   - include component ablation or important diagnostic figure
   - state what each component contributes

6. **Ablation / Fine-Grained Analysis II**
   - include memory/tool/feedback/parameter analysis tables or figures
   - explain mechanism-level conclusions

7. **Conclusion / Limitations / Takeaways**
   - summarize why the method works
   - list limitations
   - give 2-3 reusable insights

If the paper has fewer experiments, merge slides 5-6. If it has many critical analyses, keep 7 slides but use multiple screenshots per slide.

## Design Style

- Keep content compact and information-dense.
- Avoid long prose paragraphs.
- Use short Chinese bullets beside figures/tables.
- Prefer original paper screenshots for:
  - method framework figure
  - main result table
  - ablation tables
  - analysis figures
  - cost/efficiency table
- Do not recreate paper tables manually unless screenshots are unreadable.
- Use a clean academic style:
  - light background
  - restrained colors
  - clear title hierarchy
  - figure/table on one side, takeaway bullets on the other
- Each slide should answer: “What should the audience remember?”

## Figure/Table Extraction

When working from PDF:

- Render pages to images with PyMuPDF or another PDF renderer.
- Crop figures/tables from the original PDF as PNG images.
- Use high enough resolution for readability, usually 2x-3x scale.
- Name crops descriptively, e.g.:
  - `fig1_framework.png`
  - `table1_main_results.png`
  - `table2_ablation.png`
  - `fig3_analysis.png`
- Insert screenshots into PPT with borders or frames.

## Content Density Rules

- Keep decks within 7 slides by default.
- Prefer fewer words plus stronger visual evidence.
- For each figure/table, add 2-4 takeaway bullets.
- Emphasize conclusions from ablations and analysis experiments, not only headline results.
- Include numeric results when they support the story.
- Mention limitations briefly on the final slide.

## Validation

Before final delivery:

- Confirm the `.pptx` opens using `python-pptx` or another available checker.
- Confirm slide count.
- Confirm key screenshots are embedded.
- Confirm output file is saved in the requested outputs directory.
- If possible, list slide titles and number of embedded pictures as a quick sanity check.

## Final Response

Return:
- link to the `.pptx`
- short summary of slide count and included original figures/tables
- mention any validation performed

---
name: paper-compact-ppt
description: Read an academic paper PDF or URL and create a concise, information-dense Chinese PPT for presentation. Use when the user asks to read a paper and generate slides, especially when they want a compact deck covering background, motivation, method, experiments, ablations, analysis, and conclusions, with original paper figure/table screenshots whenever useful.
---

# Paper Compact PPT

## Goal

Read an academic paper and generate a **Chinese presentation PPT** that is:

- concise
- information-dense
- presentation-ready
- visually grounded in the original paper

Default output requirements:

- output format: `.pptx`
- language: Chinese
- length: **no more than 7 slides**
- style: compact academic presentation
- include original paper screenshots for key figures/tables whenever useful

The deck should focus on the paper’s:

- background
- motivation
- method
- experiments
- conclusions

If the paper includes important ablations or analysis experiments, include their conclusions as well.

---

## Highest-Priority Rules

These rules have higher priority than all others:

1. **Keep the PPT within 7 slides unless the user explicitly asks otherwise.**
2. **Content must be concise and information-dense.**
3. **Do not turn the PPT into a long translation of the paper.**
4. **Prefer original paper screenshots for key figures/tables.**
5. **Prioritize the most important story of the paper, not exhaustive coverage.**
6. **If a section is weak or unimportant, compress it rather than forcing full coverage.**

---

## Output Style

The PPT should follow these style principles:

- Chinese slide text
- short bullets, not long paragraphs
- each slide should communicate a clear takeaway
- figure/table on one side, takeaway bullets on the other when possible
- clean academic layout
- restrained colors
- light background
- clear title hierarchy

Avoid:

- large blocks of prose
- overly verbose explanations
- recreating tables manually when screenshots are readable
- excessive decorative design

---

## Recommended Slide Structure

Prefer **6-7 slides**.

### Slide 1. Title / Paper Positioning
Include:

- paper title
- one-sentence summary of the contribution
- 2-4 key highlights or claims

### Slide 2. Background & Motivation
Include:

- task background
- why the problem matters
- limitations of prior methods
- the core motivation of this paper

### Slide 3. Method Overview
Include:

- the paper’s main framework figure if available
- 3-5 concise bullets explaining the pipeline / method idea
- emphasize what is new

### Slide 4. Main Experiments
Include:

- main result table screenshot
- concise comparison to baselines
- 2-4 takeaway bullets highlighting the strongest results

### Slide 5. Ablation or Key Analysis I
If the paper has ablations:

- include the most important ablation table/figure screenshot
- summarize what each component contributes

If the paper has no ablations but has important analysis:

- include the most important analysis figure/table
- summarize the conclusion

### Slide 6. Analysis II / Efficiency / More Evidence
Use this slide for one of the following if available:

- another important analysis experiment
- efficiency or cost comparison
- parameter / module / memory / scaling study
- qualitative examples

If the paper has limited analysis content, merge this material into Slide 5.

### Slide 7. Conclusion / Limitations
Include:

- why the method works
- main conclusion
- limitations
- 2-3 reusable insights or takeaways

---

## Figure/Table Usage Rules

When working from a PDF or paper URL:

1. Prefer original screenshots from the paper for:
   - method overview figure
   - main result table
   - ablation table
   - analysis figure
   - efficiency/cost table
2. Use only the most informative figures/tables.
3. Do not manually redraw tables unless screenshots are unreadable.
4. If a screenshot is used, place 2-4 short takeaway bullets next to it.
5. Crop figures/tables cleanly and keep them readable.

Recommended file naming:

- `fig1_method.png`
- `table1_main_results.png`
- `table2_ablation.png`
- `fig3_analysis.png`

---

## Workflow

1. Obtain the paper PDF from a user-provided PDF or URL.
2. Read the paper enough to identify:
   - background and task setting
   - motivation
   - method
   - main framework figure
   - main experiments
   - important ablations
   - important analysis experiments
   - conclusions and limitations
3. Extract or screenshot the most important original figures/tables from the paper.
4. Build a compact Chinese PPT.
5. Keep the deck within 7 slides by default.
6. Save the output as `.pptx`.
7. Validate the PPT.

---

## Content Compression Rules

Use these compression rules to keep the deck strong:

- Prefer **fewer slides with stronger evidence**
- Prefer **2-4 strong bullets** over many weak bullets
- Summarize each figure/table with short Chinese takeaways
- Focus on:
  - what problem is solved
  - what the core idea is
  - whether experiments support the claim
  - what ablations/analysis reveal
- If the paper has too many experiments, keep only the most decision-relevant ones
- If the paper has few ablations, do not force two separate ablation slides

---

## PPT Generation Requirements

When generating the PPT:

1. Use a reliable PPT generation method such as `python-pptx`.
2. Ensure the PPT is a valid `.pptx` file.
3. If images are inserted, ensure they are embedded correctly.
4. Prefer a side-by-side layout:
   - screenshot on one side
   - bullets on the other side

---

## Validation

Before final delivery, validate all of the following:

1. The `.pptx` file exists.
2. The PPT can be opened programmatically.
3. The slide count is within the intended range.
4. Key screenshots are embedded.
5. The PPT is not empty or obviously corrupted.

Recommended validation checks:

- open with `python-pptx`
- confirm slide count
- count embedded pictures
- confirm file size is reasonable

If possible, report:

- slide count
- slide titles
- number of embedded screenshots

---

## Delivery / Binary File Handling

If the environment supports direct repository commits for `.pptx`, save the file to the requested output path.

If binary `.pptx` files cannot be cleanly included in a PR:

1. keep scripts/workflows in the PR
2. generate the PPT in GitHub Actions
3. upload the `.pptx` as an artifact
4. optionally also export a PDF preview if helpful

Never silently return an invalid or corrupted `.pptx`.

---

## Final Response

Return:

- path to the generated `.pptx`
- slide count
- which original figures/tables were included
- validation status
- if artifact delivery is used, explain how to download it

---
name: paper-compact-ppt
description: "Read a research paper PDF or paper URL and create a concise, information-dense Chinese academic PPT. Use this skill when the user wants a presentation deck covering the core background, motivation, method, experiments, conclusion, and, when available, ablations or analysis. Prefer original paper figure/table screenshots when useful, but never use full-page paper screenshots."
---

# Paper Compact PPT

## Core Task

Read the paper and generate a **Chinese academic presentation PPT**.

The PPT should:

- be concise
- be information-dense
- be suitable for academic presentation
- focus on the paper’s core story
- use original paper figures/tables when they help explain the story

Default output:

- format: `.pptx`
- language: Chinese
- length: preferably within **7 slides**
- style: compact academic presentation

## Highest-Priority Instructions

These instructions are mandatory.

1. Create a PPT, not a poster, unless the user explicitly asks for a poster.
2. Keep the content concise and information-dense.
3. Prefer a deck within 7 slides unless the user asks otherwise.
4. Focus on the paper’s core background, motivation, method, experiments, and conclusion.
5. If the paper includes useful ablation or analysis experiments, include their key conclusions.
6. Do not translate the paper section by section.
7. Do not dump screenshots into the deck.
8. Use original paper screenshots only for important figures/tables, and crop them tightly.
9. Never use full-page PDF screenshots.
10. Every slide should communicate a clear takeaway.
11. Layout quality is as important as content quality.
12. Captions must describe the content, not the extraction process.

---

## What to Cover

The deck should usually cover:

* background
* motivation
* problem setting
* method
* main experiments
* conclusion

If useful and available, also include:

* ablation results
* analysis experiments
* qualitative results
* efficiency or cost comparison
* limitations

Do not force every possible section into the deck.
Prioritize the most important content.

---

## Default Slide Structure

Use this structure by default, but merge slides when appropriate.

### Slide 1. Title and Core Contribution

Include:

* paper title
* one-sentence summary
* 2-4 key highlights or claims

### Slide 2. Background and Motivation

Include:

* task background
* why the problem matters
* limitations of prior methods
* the key motivation of this paper

### Slide 3. Method

Include:

* the core method idea
* the method pipeline or framework
* what is new or important

If the paper has a clear method figure, use it.

### Slide 4. Main Experiments

Include:

* the main result table or figure
* strongest comparisons
* 2-4 concise takeaways

### Slide 5. Ablation or Analysis

If the paper has ablation experiments, include the most important ablation results.

If ablations are weak or absent, use this slide for:

* analysis experiments
* qualitative comparisons
* efficiency evidence

### Slide 6. More Evidence or Key Insights

Use this only if needed.

Possible content:

* another useful analysis
* qualitative results
* efficiency or robustness
* a second important result table/figure

If not needed, merge this content into Slide 5.

### Slide 7. Conclusion

Include:

* the main conclusion
* why the method works
* limitations
* 2-3 takeaways

---

## Screenshot and Figure/Table Policy

Use screenshots selectively.

### Use screenshots for:

* method overview figure
* main result table
* ablation table
* important analysis figure
* qualitative result figure
* efficiency/cost table

### Do not use:

* full-page screenshots
* introduction pages
* related work pages
* pages with lots of irrelevant text
* unreadable tiny tables
* screenshots used only to fill space

### Screenshot rules:

1. Crop tightly around the target figure or table.
2. Remove page margins and unrelated text.
3. Keep the crop readable.
4. Use only a small number of strong visuals.
5. For each inserted visual, add 2-4 Chinese takeaway bullets explaining what it shows.

If a figure or table cannot be cropped clearly and readably, summarize it in text instead.

---

## Caption Rules

Captions should describe the figure/table content, not the extraction process.

Never include words such as:

* 紧裁
* 裁剪
* 截图
* cropped
* tight crop
* extracted from page
* source page

Bad captions:

* 方法框架图（紧裁）
* 主实验结果表（裁剪）
* Cropped Table 1

Good captions:

* 方法整体框架
* 主实验结果对比
* 消融实验结果
* 定性对比示例
* 效率与性能对比

A caption should answer what the visual is about, not how it was obtained.

---

## Writing Style

Use Chinese academic presentation language.

Write in a way that is:

* compact
* clear
* conclusion-oriented
* easy to present aloud

Prefer:

* short bullets
* explicit takeaways
* result-oriented wording
* key numbers only when helpful

Avoid:

* long paragraphs
* copying the abstract
* copying the paper wording
* overly detailed implementation descriptions
* vague bullets without conclusions

Recommended density:

* 3-5 main bullets per slide
* 2-4 takeaway bullets for each figure/table
* no large text blocks
* no more than 2 screenshots on one slide unless necessary

---

## Text Wrapping and Readability Rules

Text readability is a hard requirement.

Rules:

* Avoid long lines of text.
* Keep each bullet to one clear idea.
* Prefer one-line bullets when possible.
* Avoid bullets longer than two lines.
* Manually shorten text before reducing font size.
* Do not allow title text to wrap into two lines unless unavoidable.
* Do not use font sizes that are too small to read in presentation mode.
* Leave enough line spacing between bullets.
* Avoid dense Chinese paragraphs.
* Use concise academic Chinese, not verbose explanation.

If text overflows or wraps poorly:

1. shorten the wording;
2. split the sentence into two bullets;
3. remove secondary details;
4. move details to speaker notes only if needed;
5. never simply shrink the font until unreadable.

---

## Layout Style

Use a clean, stable academic layout.

The layout must be planned before writing elements to slides.

Preferred patterns:

1. Figure left, takeaways right.
2. Table top, takeaways bottom.
3. Text-only slide for background or conclusion.
4. Two-panel comparison slide only when necessary.

Rules:

* Use 16:9 widescreen slides.
* Use consistent margins on all slides.
* Keep all titles aligned.
* Keep text boxes aligned to a grid.
* Preserve image aspect ratios.
* Do not stretch or squeeze figures/tables.
* Do not overlap text, images, captions, or page elements.
* Do not place captions too close to figures.
* Do not place content too close to slide edges.
* Do not use arbitrary floating text boxes.
* Keep enough whitespace between title, visual, caption, and bullets.
* Prefer one dominant visual per slide.
* Avoid more than two visuals on one slide.
* If a table is wide, place it across the top or center, with takeaways below.
* If a figure is vertical or narrow, place it on the left with takeaways on the right.
* The deck should look like a polished academic presentation, not a screenshot collage.

---

## Layout Implementation Guidelines

When using `python-pptx`, define stable layout constants before creating slides.

Recommended slide structure:

* slide width: 13.333 inches
* slide height: 7.5 inches
* top title region: about 0.5-0.8 inches
* outer margin: about 0.4-0.6 inches
* gap between visual and text: about 0.25-0.4 inches

Recommended font hierarchy:

* slide title: 26-32 pt
* section label or key message: 18-22 pt
* main bullets: 15-18 pt
* captions and notes: 10-13 pt

Do not reduce body text below a readable size just to fit more content.
Shorten the content instead.

---

## Layout Templates

Use one of these layout templates for each slide.

### Template A: Text-Only Concept Slide

Use for:

* background
* motivation
* problem setting
* conclusion
* limitations

Layout:

* title at top
* key message block in upper-middle
* 3-5 bullets below
* optional small summary box

### Template B: Figure Left, Takeaways Right

Use for:

* method overview
* architecture diagram
* qualitative example
* analysis figure

Layout:

* title at top
* figure on the left, about 55-65% of slide width
* takeaway bullets on the right, about 30-40% of slide width
* caption below the figure

### Template C: Table Top, Takeaways Bottom

Use for:

* main results
* ablation results
* efficiency comparison
* benchmark comparison

Layout:

* title at top
* table screenshot in the upper or central region
* 2-4 takeaway bullets below the table
* optional small conclusion box

### Template D: Two-Panel Evidence Slide

Use only when two visuals are both necessary.

Layout:

* title at top
* left visual and right visual with similar size
* comparison takeaways at bottom

Do not use more than two visuals.

---

## Content Selection Principles

When the paper contains too much material, compress aggressively.

Prioritize:

1. the core problem
2. the key motivation
3. the main method idea
4. the strongest experimental evidence
5. the most informative ablation or analysis
6. the main conclusion and limitations

Deprioritize:

* excessive related work
* minor implementation details
* redundant tables
* secondary experiments
* weak ablations

The goal is to help the audience understand the paper quickly.

---

## Internal Planning Requirement

Before generating the PPT, internally decide:

* what the main story is
* which slides are necessary
* which figures/tables are worth using
* which layout each slide should use
* what the main takeaway of each slide is
* whether each slide risks overcrowding

Do not expose this internal planning unless the user asks.

---

## Generation Requirements

When generating the PPT:

1. Use a reliable method such as `python-pptx`.
2. Create a real, valid `.pptx` file.
3. Keep the slide count within the intended limit.
4. Insert visuals only when they improve the presentation.
5. Make sure the deck is readable and presentation-ready.
6. Do not expose implementation details such as cropping or extraction in slide captions.
7. Make sure all visible captions and text are polished Chinese.

---

## Final Layout Inspection

Before delivery, inspect the generated slides and fix layout problems.

Check:

1. no caption contains words such as “紧裁”, “裁剪”, “截图”, “cropped”, or “tight crop”;
2. slide titles do not wrap awkwardly;
3. bullets are not overly long;
4. text boxes do not overflow;
5. figures and tables are not distorted;
6. captions are clean and content-based;
7. table screenshots are readable;
8. every slide has consistent margins;
9. no element is too close to slide edges;
10. no text/image/caption overlaps;
11. no slide looks crowded or broken.

If any issue is found, revise the PPT before final delivery.

---

## Validation

Before final delivery, validate the PPT.

Required checks:

1. The `.pptx` file exists.
2. It can be reopened programmatically.
3. Slide count is reasonable and preferably within 7.
4. The file is not corrupted.
5. Inserted figures/tables are readable.
6. No full-page paper screenshots are used.
7. The layout is not visibly broken.
8. Captions do not mention cropping, screenshots, or extraction.
9. Text wrapping is acceptable.

If validation fails, fix the deck before delivery.

Never return a corrupted PPT.

---

## Binary File Delivery

If `.pptx` files cannot be committed directly into a PR because they are binary files:

1. Do not force the `.pptx` into the PR.
2. Commit only text files such as scripts and workflow files.
3. Generate the PPT in GitHub Actions.
4. Upload the `.pptx` as an artifact.
5. Ensure the artifact has been validated before upload.

If useful, also generate a PDF preview artifact.

---

## Final Response

In the final response, report:

* the PPT path or artifact name
* slide count
* which original figures/tables were used
* whether ablation/analysis content was included
* whether captions were cleaned
* validation status
* how to download the file if artifact delivery is used

---


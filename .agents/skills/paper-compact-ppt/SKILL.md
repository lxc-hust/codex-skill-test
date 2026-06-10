---

name: paper-compact-ppt
description: "Create a high-quality, concise, information-dense academic presentation deck from a research paper PDF or paper URL. Use this skill when the user asks to read a paper and generate a PPT/slides for academic presentation. The deck should cover background, motivation, method, experiments, ablations, analysis, conclusions, and limitations. Use carefully cropped original paper figures/tables when they directly support the presentation story, and never use full-page paper screenshots."
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

# Paper Compact PPT

## Goal

Create a high-quality academic presentation PPT from a research paper.

The output should be:

* presentation-ready
* concise but information-dense
* organized around the paper’s core story
* visually clean and readable
* grounded in the paper’s actual figures, tables, experiments, and claims

Default output:

* format: `.pptx`
* slide language: Chinese, unless the user requests another language
* slide count: no more than 7 slides, unless the user explicitly asks otherwise
* style: compact academic presentation, not a poster and not a paper summary document

The deck should explain the paper’s:

* background
* motivation
* core problem
* method
* experiments
* ablations, if available
* analysis experiments, if available
* conclusions and limitations

The final PPT should look like a polished academic talk deck, not a collection of PDF screenshots.

---

## Highest-Priority Principles

These principles override all lower-level instructions.

1. Quality of the academic story is more important than mechanically covering every paper section.
2. Do not translate the paper section by section.
3. Do not create a screenshot dump.
4. Do not insert full-page PDF screenshots into the PPT.
5. Only use tightly cropped local screenshots of important figures or tables.
6. Each slide must have a clear message and a clear reason to exist.
7. Every figure/table screenshot must directly support the slide’s main takeaway.
8. If a figure/table cannot be cropped cleanly and readably, summarize it in text instead.
9. Keep the deck within 7 slides by default.
10. The final `.pptx` must be valid, openable, and not corrupted.
11. Slide layout quality matters as much as content quality.
12. Do not place text boxes, figures, and tables arbitrarily. Use stable academic layout templates.

---

## Task Interpretation

When the user provides a paper URL or PDF and asks for a PPT, interpret the task as:

> Read the paper, identify its core academic contribution, and create a compact presentation deck that helps an audience quickly understand the paper’s background, motivation, method, experimental evidence, and conclusions.

If the user gives additional requirements, follow them over the default structure.

Typical user requirements may include:

* make it concise
* keep it within 7 slides
* include background, motivation, method, experiments, and conclusion
* include ablation or analysis if available
* use original paper figures/tables
* do not use full-page screenshots
* improve slide layout quality
* make the deck suitable for an academic presentation

Treat these as strong constraints.

---

## What Makes a Good Deck

A good deck should answer these questions clearly:

1. What problem does the paper solve?
2. Why is the problem important or difficult?
3. What is the key idea of the method?
4. What evidence shows that the method works?
5. What do the ablations or analysis experiments reveal?
6. What are the main conclusions and limitations?

The PPT should prioritize:

* clear academic narrative
* important experimental evidence
* compact wording
* readable visual layout
* precise takeaways
* consistent slide structure
* readable cropped paper figures and tables

Avoid:

* long prose paragraphs
* exhaustive paper summaries
* copying the abstract into slides
* listing too many details without hierarchy
* using screenshots as decoration
* putting raw full-page paper images into slides
* random or unaligned placement of text boxes and images
* unreadable tables squeezed into small areas

---

## Recommended Workflow

Follow this workflow.

### Step 1. Obtain and inspect the paper

Download or locate the paper PDF.

Read enough of the paper to identify:

* title and core contribution
* abstract and introduction
* problem setting
* motivation and pain points
* method overview
* main experiments
* main result table
* ablation studies
* analysis experiments
* qualitative results, if important
* conclusion and limitations

### Step 2. Build the presentation story

Before generating the PPT, decide the core story:

* What is the paper trying to fix?
* Why do previous methods fall short?
* What is the main technical idea?
* What is the strongest evidence?
* What should the audience remember after the talk?

Do not start by selecting screenshots. Start by selecting the story.

### Step 3. Select only the most useful visual evidence

Choose at most 4-6 important figures/tables from the paper.

Prioritize:

* method framework figure
* main result table
* important ablation table
* important analysis figure
* qualitative comparison figure
* efficiency/cost table

Do not use a figure/table if it is not essential to the presentation story.

### Step 4. Create an internal slide plan

Before creating the PPT file, create an internal slide plan.

For each slide, decide:

* slide title
* slide purpose
* selected layout template
* main visual, if any
* text region position
* figure/table region position
* 2-4 key takeaways
* whether the slide risks overcrowding

Do not start generating the `.pptx` until this layout plan is clear.

The layout plan should guide all element placement in `python-pptx`.

### Step 5. Crop figures and tables carefully

When using original paper visuals:

1. Render the relevant PDF page.
2. Locate the local figure/table region.
3. Crop tightly around the figure/table.
4. Remove page margins, headers, footers, unrelated paragraphs, references, and unrelated neighboring figures.
5. Save the crop as PNG.
6. Insert only the cropped PNG into the PPT.

Recommended filenames:

* `fig_method_overview.png`
* `table_main_results.png`
* `table_ablation.png`
* `fig_analysis.png`
* `fig_qualitative.png`
* `table_efficiency.png`

### Step 6. Generate the PPT

Create a compact academic PPT with no more than 7 slides by default.

Use a reliable PPT generation method such as `python-pptx`.

### Step 7. Validate the PPT

Before delivery, verify that the file is valid and the content meets the quality requirements.

---

## Strict Screenshot Policy

Original paper screenshots are allowed only when they improve the deck.

### Allowed screenshots

Use tightly cropped screenshots of:

* method framework figures
* architecture diagrams
* main result tables
* important ablation tables
* important analysis figures
* qualitative comparison figures
* efficiency or cost tables

### Forbidden screenshots

Never insert:

* full PDF pages
* full-page screenshots
* screenshots with large paper margins
* screenshots dominated by introduction text
* screenshots of related work paragraphs
* screenshots where the target figure/table occupies only a small part of the image
* screenshots containing headers, footers, page numbers, references, or unrelated text
* unreadable tiny tables
* random paper pages used as visual filler

### Screenshot quality criteria

A screenshot is acceptable only if:

* the relevant figure/table occupies most of the image area
* the crop is readable in PowerPoint
* the crop is directly related to the slide’s message
* the slide explains what the audience should learn from it

If these criteria are not satisfied, do not use the screenshot.

---

## Slide Structure

Use this structure by default. Merge or adapt slides when the paper does not contain enough material for a separate slide.

### Slide 1. Title and Core Contribution

Purpose:

* position the paper
* explain its one-sentence contribution
* preview the most important results or claims

Include:

* paper title
* one-sentence contribution
* 2-4 key highlights
* optional small teaser visual only if it is useful and clean

Do not insert a full-page screenshot.

### Slide 2. Background and Motivation

Purpose:

* explain why the task matters
* explain why prior methods are insufficient
* identify the core pain point

Include:

* task background
* practical or research motivation
* limitations of existing approaches
* the problem this paper targets

Usually this slide should be mostly text-based.

Do not use introduction page screenshots.

### Slide 3. Method Overview

Purpose:

* explain the core method clearly

Include:

* tightly cropped method framework figure if available
* 3-5 concise bullets explaining the pipeline
* what is technically new or different

Preferred layout:

* cropped method figure on the left
* takeaway bullets on the right

If the paper has no clear method figure, create a clean text-based method summary instead of using a full-page screenshot.

### Slide 4. Main Experimental Results

Purpose:

* show the strongest empirical evidence

Include:

* tightly cropped main result table or figure
* strongest comparisons against baselines
* 2-4 concise takeaways

Focus on what the results prove, not on copying every number.

### Slide 5. Ablation Study

Purpose:

* explain which components matter

If ablations exist, include:

* tightly cropped ablation table or figure
* what each important component contributes
* the mechanism-level conclusion

If ablations are weak or absent, merge this slide with analysis or qualitative evidence.

### Slide 6. Analysis, Qualitative Results, or Efficiency

Purpose:

* provide deeper evidence beyond headline performance

Use this slide for one of:

* analysis experiments
* qualitative comparisons
* scaling behavior
* robustness analysis
* efficiency or cost comparison
* failure cases

Include only visuals that are tightly cropped and readable.

### Slide 7. Conclusion, Limitations, and Takeaways

Purpose:

* summarize what the audience should remember

Include:

* main conclusion
* why the method works
* limitations
* 2-3 reusable insights

Avoid introducing large new figures on the final slide unless they are essential.

---

## Writing Style

Use compact Chinese academic presentation language by default.

Good slide writing:

* use short bullets
* each bullet should express one clear claim
* emphasize conclusions, not paper section summaries
* use numbers only when they support a point
* explain what each experiment demonstrates
* use reasoning-oriented language such as “shows”, “indicates”, “suggests”, and “therefore”

Bad slide writing:

* long paragraphs
* direct abstract translation
* copying paper wording
* listing many details without hierarchy
* using vague bullets such as “the method is effective”
* putting a figure on a slide without explaining it

Recommended density:

* 3-5 main bullets per slide
* 2-4 takeaway bullets for each inserted figure/table
* no large paragraphs
* no slide filled only with screenshots
* no more than 2 screenshots on one slide unless absolutely necessary

---

## Layout Principles

The PPT must use a stable academic layout system instead of placing elements arbitrarily.

General layout rules:

* Use a 16:9 widescreen slide size.
* Use consistent margins across all slides.
* Keep a clear title area at the top of every slide.
* Divide the remaining content area into structured regions.
* Align all visual and text elements to a consistent grid.
* Avoid overlapping text boxes, images, tables, and captions.
* Avoid placing elements too close to slide edges.
* Avoid filling the entire slide with dense text or oversized screenshots.
* Prefer balanced layouts with clear reading order from left to right and top to bottom.

Every slide should have:

1. a clear title region;
2. one main content region;
3. one supporting explanation region;
4. sufficient whitespace between regions.

Do not let figures, tables, or text boxes float freely without alignment.

---

## Slide Layout Templates

Use one of the following layout templates for each slide.

### Template A: Text-Only Concept Slide

Use for motivation, problem setting, conclusions, and limitations.

Layout:

* Title: top 10-15% of the slide.
* Main message block: upper-middle area.
* Supporting bullets: below the main message.
* Optional small summary box: bottom-right.

Rules:

* Use 3-5 concise bullets.
* Do not use more than two text columns.
* Keep line length moderate.
* Use bold emphasis only for key terms.

Suitable for:

* background
* motivation
* problem definition
* conclusion
* limitations

### Template B: Figure Left, Takeaways Right

Use for method overview, framework figures, qualitative examples, and analysis figures.

Layout:

* Title: top region.
* Left region: cropped figure, occupying about 55-65% of slide width.
* Right region: 2-4 takeaway bullets, occupying about 30-40% of slide width.
* Optional caption: below the figure.

Rules:

* The figure must be large enough to read.
* The right-side bullets must explain what the audience should notice in the figure.
* Do not place long paragraphs next to the figure.
* Do not put more than one large figure in this layout.

Suitable for:

* method overview
* architecture diagram
* qualitative comparison
* analysis figure

### Template C: Table Top, Takeaways Bottom

Use for main results, ablation results, and efficiency comparisons.

Layout:

* Title: top region.
* Table screenshot: center-top region, occupying about 70-80% of slide width.
* Takeaway bullets: bottom region, below the table.
* Optional highlighted observation box: bottom-right.

Rules:

* The table must be tightly cropped and readable.
* Do not stretch the table non-proportionally.
* Do not include a full-page screenshot.
* Use 2-4 bullets below the table to explain the main result.
* If the table is too wide or too dense, crop only the most relevant rows/columns if possible.
* If the table remains unreadable after cropping, replace it with a textual summary.

Suitable for:

* main benchmark table
* ablation table
* comparison table
* cost or efficiency table

### Template D: Two-Panel Evidence Slide

Use when two cropped visuals need to be compared.

Layout:

* Title: top region.
* Left visual: one cropped figure/table.
* Right visual: another cropped figure/table.
* Bottom region: 2-3 comparison bullets.

Rules:

* Use this template only when both visuals are necessary.
* Both visuals should have similar visual weight.
* Do not use more than two visuals.
* The bottom bullets must explain the comparison.

Suitable for:

* before/after comparison
* two analysis figures
* qualitative vs quantitative evidence
* baseline vs proposed method

### Template E: Method Pipeline Slide

Use when the paper has no clean method figure or when the method needs simplification.

Layout:

* Title: top region.
* Middle region: 3-5 step pipeline drawn with simple boxes and arrows.
* Bottom region: concise explanation of each step.

Rules:

* Use this only when the original method figure is unavailable, unreadable, or too complex.
* Keep the pipeline abstract and easy to present.
* Do not overdraw complicated architectures manually.

Suitable for:

* method summary
* algorithm flow
* training pipeline
* inference pipeline

---

## Table Layout Rules

Tables are often the most difficult part of the deck. Follow these rules strictly.

1. Prefer cropped original paper tables only when they remain readable.
2. A table screenshot should usually occupy the central visual area, not the entire slide.
3. Keep the table aspect ratio unchanged.
4. Do not squeeze a wide table into a narrow column.
5. Do not place long text beside a very wide table.
6. For wide tables, use the “Table Top, Takeaways Bottom” layout.
7. For narrow tables, use the “Figure/Table Left, Takeaways Right” layout.
8. If a table has many columns, crop to the most relevant columns when possible.
9. If a table has many rows, crop to the most relevant baselines and the proposed method when possible.
10. If the original table is unreadable after cropping, summarize the key numbers in text instead of inserting the screenshot.

Every table slide must answer:

* What is being compared?
* Which method performs best?
* What is the main conclusion?
* Why does this result matter?

---

## Text Box Rules

Text boxes must be compact and aligned.

Rules:

* Use short bullets instead of paragraphs.
* Use consistent font sizes across slides.
* Use larger font for slide titles.
* Use medium font for key bullets.
* Use smaller font only for captions or secondary notes.
* Avoid more than 5 bullets on one slide.
* Avoid more than 2 lines per bullet when possible.
* Keep text boxes aligned to the same left edge within a slide.
* Leave clear spacing between bullets.
* Do not place text over images or tables.

Recommended hierarchy:

* Slide title: large and bold.
* Section label or key message: medium-large and bold.
* Main bullets: medium.
* Captions or notes: small but readable.

---

## Figure and Screenshot Layout Rules

For every inserted figure or table screenshot:

* Preserve aspect ratio.
* Add a thin border or subtle frame if it improves readability.
* Keep enough whitespace around the image.
* Do not crop off important labels, legends, or axis titles.
* Do not insert images that become unreadable after resizing.
* Do not use full-slide images unless the user explicitly requests a visual-only slide.
* Do not place multiple unrelated screenshots on the same slide.

If the figure is too complex:

* crop the most relevant sub-region;
* or use a simplified text explanation;
* or split evidence across two slides only if the slide count limit still allows it.

---

## Content Selection Rules

When the paper is long, compress aggressively.

Prioritize:

1. core problem and motivation
2. central method idea
3. strongest result
4. most informative ablation
5. most useful analysis
6. key limitations

Deprioritize:

* extensive related work
* implementation details not needed for understanding
* secondary metrics
* repeated experimental tables
* minor ablations
* long benchmark descriptions

The PPT should help someone present the paper, not reproduce the paper.

---

## Figure/Table Selection Plan

Before generating the final PPT, create an internal figure/table selection plan.

For each candidate visual, evaluate:

* page number
* figure/table name
* why it matters
* which slide it supports
* whether it can be cropped cleanly
* whether it is readable after cropping

Only use visuals that pass this check.

Do not include this internal plan in the final PPT unless the user asks for it.

---

## Layout Quality Check

Before final delivery, inspect the generated PPT for layout quality.

Check every slide for:

1. title is visible and aligned;
2. text boxes do not overlap;
3. images and tables do not overlap text;
4. figures/tables are readable;
5. no full-page PDF screenshots are used;
6. table screenshots are not excessively squeezed;
7. each visual has nearby explanatory bullets;
8. slide has sufficient margins and whitespace;
9. content follows a clear reading order;
10. the deck looks consistent across slides.

If a slide looks crowded, fix it by:

* reducing bullet count;
* enlarging the main visual;
* removing secondary details;
* switching to a better layout template;
* replacing an unreadable screenshot with text summary.

Never deliver a deck with visibly broken layout.

---

## Validation Requirements

Before final delivery, validate the `.pptx`.

Required checks:

1. The `.pptx` file exists.
2. It can be opened with `python-pptx`.
3. Slide count is no more than 7 by default.
4. The file size is reasonable.
5. Embedded pictures exist when useful paper figures/tables are available.
6. No full-page PDF screenshots are inserted.
7. Each inserted screenshot is paired with explanatory takeaways.
8. The deck is not just a paper screenshot collection.
9. The PPT can be delivered as a valid file or artifact.
10. No slide has visibly broken layout.

Recommended technical checks:

* open with `python-pptx`
* count slides
* count embedded pictures
* check image dimensions
* inspect whether any inserted image looks like a full PDF page
* run `unzip -t` on the `.pptx` if available

If validation fails, fix the PPT before delivery.

Never deliver a corrupted or invalid `.pptx`.

---

## Binary Delivery Rule

If `.pptx` files cannot be committed directly into a PR because they are binary files:

1. Do not force the `.pptx` into the PR.
2. Commit the generation script, helper assets, and GitHub Actions workflow instead.
3. Generate the PPTX in GitHub Actions.
4. Upload the final `.pptx` as a GitHub Actions artifact.
5. Ensure the artifact passes validation.

If possible, also provide a PDF preview artifact for quick inspection.

---

## Final Response

In the final response, report:

* generated PPT path or artifact name
* slide count
* list of original paper figures/tables used
* confirmation that no full-page screenshots were used
* validation result
* layout quality check result
* download instructions if artifact delivery is used

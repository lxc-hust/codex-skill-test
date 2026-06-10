#!/usr/bin/env python3
"""Generate a compact Chinese academic PPT for DocSeeker.

This script is intentionally text-only in git. At runtime it:
1. downloads the arXiv PDF;
2. renders and crops selected figure/table regions into ignored local folders;
3. generates ``outputs/docseeker_compact_ppt.pptx`` with python-pptx;
4. validates the generated deck.

The crop boxes are normalized page coordinates selected for the arXiv version of
"DocSeeker: Structured Visual Reasoning with Evidence Grounding for Long
Document Understanding" (arXiv:2604.12812). They are deliberately tight enough
to avoid full-page screenshots and are used only as runtime artifacts.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Iterable
import argparse
import zipfile

import fitz  # PyMuPDF
import requests
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

PAPER_URL = "https://arxiv.org/pdf/2604.12812"
PAPER_PATH = Path("paper/docseeker_2604.12812.pdf")
CROP_DIR = Path("crops/docseeker")
OUTPUT = Path("outputs/docseeker_compact_ppt.pptx")
MAX_SLIDES = 7

# Internal slide plan required by the paper-compact-ppt workflow.
SLIDE_PLAN = [
    ("DocSeeker：面向长文档理解的证据定位式结构化视觉推理", "A/Text", None),
    ("为什么长文档更难：低信噪比 + 监督稀缺", "A/Text", None),
    ("核心方法：Analysis–Localization–Reasoning + 两阶段训练", "B/Figure-left", "fig_method_overview.png"),
    ("主结果：ALR 监督带来跨数据集泛化提升", "C/Table-top", "table_main_results.png"),
    ("长度分析：文档越长，普通 MLLM 越容易被噪声淹没", "B/Figure-left", "fig_length_analysis.png"),
    ("消融与效率：ALR、Page ID、EviGRPO、EGRA 均有贡献", "C/Table-top", "table_ablation.png"),
    ("结论与局限：让模型先找证据，再进行可核验推理", "A/Text", None),
]

# Runtime figure/table selection plan. Page numbers are 1-indexed; boxes are
# normalized (x0, y0, x1, y1) fractions of the page. The script also performs a
# safety check that every crop is far smaller than a full page.
@dataclass(frozen=True)
class CropSpec:
    name: str
    page: int
    box: tuple[float, float, float, float]
    filename: str
    purpose: str


CROP_PLAN = [
    CropSpec("Figure 2 / method framework", 4, (0.07, 0.08, 0.93, 0.46), "fig_method_overview.png", "method overview"),
    CropSpec("Table 1 / main benchmark results", 6, (0.05, 0.10, 0.95, 0.42), "table_main_results.png", "main results"),
    CropSpec("Figure 3 / length analysis", 7, (0.08, 0.08, 0.92, 0.45), "fig_length_analysis.png", "length robustness"),
    CropSpec("Tables 3-5 / ablation and efficiency", 8, (0.05, 0.08, 0.95, 0.54), "table_ablation.png", "ablation/efficiency"),
]

BLUE = RGBColor(30, 58, 138)
DARK = RGBColor(15, 23, 42)
TEXT = RGBColor(31, 41, 55)
MUTED = RGBColor(100, 116, 139)
LIGHT_BLUE = RGBColor(239, 246, 255)
LIGHT_GRAY = RGBColor(248, 250, 252)
LIGHT_GREEN = RGBColor(236, 253, 245)
LIGHT_ORANGE = RGBColor(255, 247, 237)
BORDER = RGBColor(203, 213, 225)


def download_pdf(url: str = PAPER_URL, dest: Path = PAPER_PATH) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 100_000:
        return dest
    response = requests.get(url, timeout=90, headers={"User-Agent": "codex-docseeker-ppt/1.0"})
    response.raise_for_status()
    if not response.content.startswith(b"%PDF"):
        raise RuntimeError(f"Downloaded file is not a PDF: {response.headers.get('content-type')}")
    dest.write_bytes(response.content)
    return dest


def crop_pdf_figures(pdf_path: Path, crop_dir: Path = CROP_DIR) -> dict[str, Path]:
    crop_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    paths: dict[str, Path] = {}
    for spec in CROP_PLAN:
        if spec.page < 1 or spec.page > doc.page_count:
            raise RuntimeError(f"Crop page out of range for {spec.name}: {spec.page}/{doc.page_count}")
        page = doc[spec.page - 1]
        rect = page.rect
        x0, y0, x1, y1 = spec.box
        crop = fitz.Rect(rect.x0 + x0 * rect.width, rect.y0 + y0 * rect.height, rect.x0 + x1 * rect.width, rect.y0 + y1 * rect.height)
        area_ratio = (crop.width * crop.height) / (rect.width * rect.height)
        if area_ratio >= 0.60:
            raise RuntimeError(f"Crop too close to full-page screenshot for {spec.name}: {area_ratio:.2%}")
        pix = page.get_pixmap(matrix=fitz.Matrix(2.4, 2.4), clip=crop, alpha=False)
        out = crop_dir / spec.filename
        pix.save(out)
        paths[spec.filename] = out
    doc.close()
    return paths


def set_text(frame, paragraphs: Iterable[str], size: int = 18, color: RGBColor = TEXT, bold_first: bool = False) -> None:
    frame.clear()
    for i, line in enumerate(paragraphs):
        p = frame.paragraphs[0] if i == 0 else frame.add_paragraph()
        p.text = line
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.font.bold = bool(bold_first and i == 0)
        p.space_after = Pt(6)


def add_title(slide, title: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.45), Inches(0.24), Inches(12.4), Inches(0.55))
    set_text(box.text_frame, [title], 25, DARK, True)


def add_footer(slide) -> None:
    box = slide.shapes.add_textbox(Inches(0.45), Inches(7.12), Inches(12.4), Inches(0.22))
    set_text(box.text_frame, ["DocSeeker, arXiv:2604.12812 / compact Chinese academic deck"], 8, MUTED)


def add_panel(slide, x, y, w, h, lines, fill=LIGHT_GRAY, border=BORDER, size=18, bold_first=True) -> None:
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = border
    shape.line.width = Pt(1)
    set_text(shape.text_frame, lines, size, TEXT, bold_first)


def add_picture_fit(slide, image_path: Path, x, y, w, h) -> None:
    pic = slide.shapes.add_picture(str(image_path), Inches(x), Inches(y), width=Inches(w))
    if pic.height > Inches(h):
        ratio = Inches(h) / pic.height
        pic.width = int(pic.width * ratio)
        pic.height = Inches(h)
    pic.left = Inches(x) + int((Inches(w) - pic.width) / 2)
    pic.top = Inches(y) + int((Inches(h) - pic.height) / 2)
    pic.line.color.rgb = BORDER
    pic.line.width = Pt(0.75)


def blank_slide(prs: Presentation):
    return prs.slides.add_slide(prs.slide_layouts[6])


def build_deck(crops: dict[str, Path], output: Path = OUTPUT) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(13.333333)
    prs.slide_height = Inches(7.5)

    # Slide 1
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[0][0])
    add_panel(slide, 0.65, 1.20, 7.7, 0.9, ["一句话：把长文档 VQA 从“直接回答”改造成“先分析问题、定位证据页、再基于证据推理”的可核验流程。"], LIGHT_BLUE, RGBColor(147, 197, 253), 21, True)
    add_panel(slide, 0.85, 2.35, 5.8, 3.2, ["核心贡献", "• 低信噪比：关键证据淹没在大量无关页面中", "• 监督稀缺：短答案/证据页标签缺少推理过程", "• ALR：Analysis–Localization–Reasoning 结构化工作流", "• EGRA：证据页高分辨率，非证据页降采样以节省显存"], LIGHT_GRAY, BORDER, 18)
    add_panel(slide, 7.25, 2.35, 5.0, 3.2, ["听众应记住", "1. 难点不是“看不到”，而是“噪声太多且缺少细粒度监督”。", "2. Evidence pages 让推理可定位、可检查。", "3. SFT + EviGRPO 分别解决范式注入和结果驱动优化。"], LIGHT_GRAY, BORDER, 18)
    add_footer(slide)

    # Slide 2
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[1][0])
    add_panel(slide, 0.7, 1.2, 3.9, 4.9, ["背景：纯视觉长文档理解", "• 直接输入多页图像，保留版面/视觉信息", "• 避免 OCR 管线的级联误差", "• 但上下文越长，视觉 token 越冗余"], LIGHT_GRAY, BORDER, 18)
    add_panel(slide, 4.9, 1.2, 3.9, 4.9, ["挑战 1：低信噪比", "• 关键证据稀疏，干扰页面密集", "• RAG Top-K 存在 recall/noise 两难", "• 文档增长会放大页面间干扰"], LIGHT_ORANGE, RGBColor(253, 186, 116), 18)
    add_panel(slide, 9.1, 1.2, 3.5, 4.9, ["挑战 2：监督稀缺", "• 数据集通常只有短答案 + 证据页", "• 缺少定位与综合证据的过程监督", "• 短答案 SFT 易记忆，OOD 泛化弱"], RGBColor(254, 242, 242), RGBColor(254, 202, 202), 18)
    add_footer(slide)

    # Slide 3
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[2][0])
    add_picture_fit(slide, crops["fig_method_overview.png"], 0.55, 1.15, 7.2, 3.95)
    add_panel(slide, 8.05, 1.15, 4.6, 4.55, ["图中应关注", "• ALR 把输出拆成分析、定位、推理、答案四段", "• SFT 用蒸馏 ALR CoT 建立基本格式与能力", "• EviGRPO 奖励同时覆盖格式、证据页、答案", "• EGRA 让长文档训练在显存上可承受"], LIGHT_GRAY, BORDER, 18)
    add_footer(slide)

    # Slide 4
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[3][0])
    add_picture_fit(slide, crops["table_main_results.png"], 0.65, 1.0, 12.0, 3.9)
    add_panel(slide, 0.9, 5.15, 11.5, 1.35, ["• DocSeeker 在多个长文档/多页基准上优于同规模直接回答模型。", "• 短答案 SFT 主要提升域内指标；ALR CoT 对 OOD 长文档更关键。", "• GRPO 阶段在 SFT 基础上继续带来增益，说明证据奖励能优化定位-回答闭环。"], LIGHT_GRAY, BORDER, 16, False)
    add_footer(slide)

    # Slide 5
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[4][0])
    add_picture_fit(slide, crops["fig_length_analysis.png"], 0.7, 1.2, 7.0, 4.3)
    add_panel(slide, 8.05, 1.25, 4.55, 4.35, ["图中应关注", "• Baseline 随页面数增长明显下降", "• DocSeeker 曲线更平稳，说明能在噪声中保持证据定位", "• 长度越大，二者差距越明显", "• 证据定位能力是长文档鲁棒性的关键中介变量"], LIGHT_GRAY, BORDER, 18)
    add_footer(slide)

    # Slide 6
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[5][0])
    add_picture_fit(slide, crops["table_ablation.png"], 0.65, 1.0, 12.0, 4.25)
    add_panel(slide, 0.9, 5.35, 11.5, 1.05, ["• ALR CoT > Vanilla CoT > raw short answer：结构化证据定位比普通推理链更可迁移。", "• Page ID、EGRA、EviGRPO 分别对应定位锚点、训练效率/信噪比、奖励层面的证据约束。", "• 组件贡献互补：围绕“证据可定位”统一设计，而非单纯扩大数据或分辨率。"], LIGHT_GRAY, BORDER, 15, False)
    add_footer(slide)

    # Slide 7
    slide = blank_slide(prs); add_title(slide, SLIDE_PLAN[6][0])
    add_panel(slide, 0.85, 1.2, 5.8, 4.9, ["主要结论", "• 长文档理解瓶颈：证据稀疏、噪声密集、监督粗粒度", "• DocSeeker 用 ALR 将“找证据”显式化，提升可解释性与鲁棒性", "• SFT + EviGRPO 分别解决范式注入与证据/答案联合优化", "• EGRA 在多页训练中平衡细节、上下文长度与显存"], LIGHT_GREEN, RGBColor(134, 239, 172), 18)
    add_panel(slide, 7.05, 1.2, 5.4, 4.9, ["局限与讨论", "• 依赖高质量教师模型与验证机制，蒸馏成本仍不可忽视", "• 极端版式/跨文档任务仍需更多验证", "• 推理时全高分辨率处理长文档仍有计算成本", "• 证据页正确不等于最终答案必然正确，仍需更细粒度 grounding"], LIGHT_ORANGE, RGBColor(253, 186, 116), 18)
    add_footer(slide)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    return output


def validate_pptx(path: Path = OUTPUT) -> dict[str, int]:
    prs = Presentation(str(path))
    slide_count = len(prs.slides)
    if slide_count > MAX_SLIDES:
        raise RuntimeError(f"Slide count exceeds limit: {slide_count} > {MAX_SLIDES}")
    picture_count = sum(1 for slide in prs.slides for shape in slide.shapes if shape.shape_type == 13)
    if picture_count != len(CROP_PLAN):
        raise RuntimeError(f"Expected {len(CROP_PLAN)} pictures, found {picture_count}")
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad:
            raise RuntimeError(f"Corrupt PPTX member: {bad}")
    return {"slides": slide_count, "pictures": picture_count, "full_page_screenshots": 0}


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--skip-download", action="store_true", help="Use an existing paper PDF under paper/.")
    args = parser.parse_args()
    pdf = PAPER_PATH if args.skip_download else download_pdf()
    crops = crop_pdf_figures(pdf)
    pptx = build_deck(crops)
    result = validate_pptx(pptx)
    print(f"wrote {pptx}")
    print(result)


if __name__ == "__main__":
    main()

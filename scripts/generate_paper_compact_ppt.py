#!/usr/bin/env python3
"""Generate a compact Chinese academic PPT for arXiv:2604.12812 (DocSeeker).

The deck is intentionally generated in CI rather than committed as a binary.
It uses python-pptx and embeds tightly cropped original paper figures from the
arXiv HTML rendering, plus compact editable tables reconstructed from the paper.
"""
from __future__ import annotations

import argparse
import io
import os
import zipfile
from pathlib import Path
from typing import Iterable

import requests
from PIL import Image, ImageOps
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.util import Cm, Pt

PAPER_URL = "https://arxiv.org/abs/2604.12812"
FIGURE_URLS = {
    "fig1_overview": "https://arxiv.org/html/2604.12812v5/x1.png",
    "fig3_length": "https://arxiv.org/html/2604.12812v5/x3.png",
    "fig4_rag": "https://arxiv.org/html/2604.12812v5/x4.png",
}

OUT = Path("outputs/paper_compact_ppt.pptx")
ASSET_DIR = Path("assets/paper_2604_12812")
WIDE_W, WIDE_H = Cm(33.867), Cm(19.05)
COLORS = {
    "navy": RGBColor(27, 55, 100),
    "blue": RGBColor(47, 103, 177),
    "green": RGBColor(91, 155, 77),
    "light_blue": RGBColor(235, 243, 252),
    "light_green": RGBColor(237, 247, 235),
    "gray": RGBColor(94, 104, 121),
    "dark": RGBColor(30, 36, 48),
    "orange": RGBColor(224, 132, 44),
    "red": RGBColor(192, 64, 64),
}


def download_figures(asset_dir: Path) -> dict[str, Path]:
    asset_dir.mkdir(parents=True, exist_ok=True)
    paths: dict[str, Path] = {}
    session = requests.Session()
    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; paper-compact-ppt/1.0)",
        "Referer": "https://arxiv.org/html/2604.12812v5",
    }
    for name, url in FIGURE_URLS.items():
        path = asset_dir / f"{name}.png"
        if not path.exists():
            r = session.get(url, headers=headers, timeout=45)
            r.raise_for_status()
            img = Image.open(io.BytesIO(r.content)).convert("RGB")
            img.save(path)
        paths[name] = path

    # Figure 1 contains two useful panels. Crop tightly to avoid slide clutter.
    fig1 = Image.open(paths["fig1_overview"]).convert("RGB")
    w, h = fig1.size
    # left: performance plots; right: ALR workflow. Keep only panel areas.
    crops = {
        "fig1_results_panel": (0, 0, int(w * 0.49), h),
        "fig1_alr_panel": (int(w * 0.50), 0, w, h),
    }
    for name, box in crops.items():
        crop_path = asset_dir / f"{name}.png"
        if not crop_path.exists():
            crop = fig1.crop(box)
            crop = ImageOps.expand(crop, border=2, fill="white")
            crop.save(crop_path)
        paths[name] = crop_path
    return paths


def add_title(slide, text: str, subtitle: str | None = None):
    box = slide.shapes.add_textbox(Cm(1.0), Cm(0.45), Cm(31.8), Cm(1.25))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = text
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(25)
    run.font.bold = True
    run.font.color.rgb = COLORS["navy"]
    if subtitle:
        sub = slide.shapes.add_textbox(Cm(1.05), Cm(1.55), Cm(31.6), Cm(0.55))
        st = sub.text_frame
        st.clear()
        p2 = st.paragraphs[0]
        r2 = p2.add_run()
        r2.text = subtitle
        r2.font.name = "Microsoft YaHei"
        r2.font.size = Pt(10.5)
        r2.font.color.rgb = COLORS["gray"]


def add_footer(slide, idx: int):
    line = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Cm(1.0), Cm(18.3), Cm(31.8), Cm(0.03))
    line.fill.solid(); line.fill.fore_color.rgb = RGBColor(218, 224, 235)
    line.line.fill.background()
    box = slide.shapes.add_textbox(Cm(1.0), Cm(18.38), Cm(31.8), Cm(0.35))
    tf = box.text_frame; tf.clear()
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.RIGHT
    r = p.add_run(); r.text = f"DocSeeker · arXiv:2604.12812 · {idx}/7"
    r.font.name = "Aptos"; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(130, 139, 153)


def set_bg(slide, color=RGBColor(248, 250, 253)):
    fill = slide.background.fill
    fill.solid(); fill.fore_color.rgb = color


def add_bullets(slide, x, y, w, h, bullets: Iterable[str], font_size=15, color=None, gap=0.11):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.clear(); tf.word_wrap = True; tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    for i, text in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = text
        p.level = 0
        p.space_after = Pt(font_size * gap)
        p.font.name = "Microsoft YaHei"
        p.font.size = Pt(font_size)
        p.font.color.rgb = color or COLORS["dark"]
    return box


def add_card(slide, x, y, w, h, title: str, body: list[str], accent="blue"):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, w, h)
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
    shape.line.color.rgb = RGBColor(221, 228, 239)
    bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, x, y, Cm(0.12), h)
    bar.fill.solid(); bar.fill.fore_color.rgb = COLORS[accent]
    bar.line.fill.background()
    title_box = slide.shapes.add_textbox(x + Cm(0.35), y + Cm(0.25), w - Cm(0.55), Cm(0.55))
    tf = title_box.text_frame; tf.clear()
    p = tf.paragraphs[0]; r = p.add_run(); r.text = title
    r.font.name = "Microsoft YaHei"; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = COLORS[accent]
    add_bullets(slide, x + Cm(0.35), y + Cm(0.9), w - Cm(0.55), h - Cm(1.0), body, font_size=11.5)


def fit_picture(slide, img_path: Path, x, y, w, h):
    with Image.open(img_path) as im:
        iw, ih = im.size
    ratio = min(w / iw, h / ih)
    nw, nh = int(iw * ratio), int(ih * ratio)
    return slide.shapes.add_picture(str(img_path), x + (w - nw) / 2, y + (h - nh) / 2, width=nw, height=nh)


def make_table(slide, x, y, w, h, rows, col_widths=None, font_size=8.8, header_fill=COLORS["navy"]):
    table_shape = slide.shapes.add_table(len(rows), len(rows[0]), x, y, w, h)
    table = table_shape.table
    if col_widths:
        for i, cw in enumerate(col_widths):
            table.columns[i].width = cw
    for r, row in enumerate(rows):
        for c, val in enumerate(row):
            cell = table.cell(r, c)
            cell.text = str(val)
            cell.margin_left = Cm(0.04); cell.margin_right = Cm(0.04)
            cell.margin_top = Cm(0.03); cell.margin_bottom = Cm(0.03)
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for run in p.runs:
                    run.font.name = "Aptos"
                    run.font.size = Pt(font_size)
                    if r == 0:
                        run.font.bold = True; run.font.color.rgb = RGBColor(255, 255, 255)
                    else:
                        run.font.color.rgb = COLORS["dark"]
            if r == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = header_fill
            elif r % 2 == 0:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(242, 246, 251)
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb = RGBColor(255, 255, 255)
    return table_shape


def add_badge(slide, x, y, text, color="green"):
    shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, x, y, Cm(4.8), Cm(0.68))
    shape.fill.solid(); shape.fill.fore_color.rgb = COLORS[color]
    shape.line.fill.background()
    tf = shape.text_frame; tf.clear(); tf.margin_left = Cm(0.12); tf.margin_right = Cm(0.12)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    r.font.name = "Microsoft YaHei"; r.font.size = Pt(11); r.font.bold = True; r.font.color.rgb = RGBColor(255, 255, 255)


def build_deck(output: Path, assets: dict[str, Path]) -> None:
    prs = Presentation()
    prs.slide_width, prs.slide_height = WIDE_W, WIDE_H
    blank = prs.slide_layouts[6]

    # 1 Title
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "DocSeeker：面向长文档理解的结构化视觉推理与证据定位",
              "Hao Yan et al., CVPR 2026 Highlight · arXiv:2604.12812v5")
    fit_picture(s, assets["fig1_results_panel"], Cm(1.1), Cm(2.35), Cm(14.9), Cm(10.4))
    add_card(s, Cm(17.0), Cm(2.4), Cm(15.2), Cm(4.0), "一句话贡献", [
        "把多页文档 VQA 从“直接回答”改造成 Analysis → Localization → Reasoning 的显式证据驱动流程。",
        "两阶段训练：ALR CoT 蒸馏 SFT + EviGRPO，同时优化格式、证据页定位与答案正确性。",
    ], "blue")
    add_card(s, Cm(17.0), Cm(7.0), Cm(15.2), Cm(4.8), "核心结果", [
        "同 7B backbone 下，相比 Baseline 在 5 个 benchmark 上提升约 30–60%。",
        "在 MMLongBench-doc 等 OOD 长文档场景中明显更稳健；与视觉 RAG 结合时能缓解 top-k 噪声困境。",
        "训练只依赖相对短的多页文档，却能泛化到超长文档。",
    ], "green")
    add_badge(s, Cm(17.0), Cm(12.45), "关键词：证据定位 / 长上下文抗噪 / 可解释推理", "orange")
    add_footer(s, 1)

    # 2 Motivation
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "动机：长文档 VQA 的瓶颈不是“看不见”，而是“找不到 + 学不会”")
    add_card(s, Cm(1.0), Cm(2.25), Cm(9.8), Cm(5.3), "挑战 1：低信噪比", [
        "关键证据常埋在数十/数百页中；无关页面会淹没视觉 token。",
        "RAG 的 top-k 两难：k 小易漏证据，k 大引入噪声。",
        "纯视觉 MLLM 长上下文中性能随页数增长快速下降。",
    ], "red")
    add_card(s, Cm(12.0), Cm(2.25), Cm(9.8), Cm(5.3), "挑战 2：监督稀缺", [
        "现有多页 DocVQA 多只有短答案与证据页，不含中间推理链。",
        "短答案 SFT 容易学到数据集记忆/捷径，OOD 泛化弱。",
        "缺少显式证据归因，难解释、难验证。",
    ], "orange")
    add_card(s, Cm(23.0), Cm(2.25), Cm(9.8), Cm(5.3), "目标", [
        "让模型先分析问题，再定位证据页，最后基于证据推理。",
        "把证据页定位变成可监督、可奖励、可输出的能力。",
        "在长文档与 RAG 噪声下维持稳定性能。",
    ], "blue")
    add_title(s, "", "研究假设：显式页面级证据 grounding 能提升长上下文抗噪性，并形成可迁移推理能力。")
    # visual equation/story
    for i, (label, desc, color) in enumerate([
        ("Analyze", "分解问题意图\n确定所需证据", "blue"),
        ("Locate", "扫描并引用页 ID\n筛出关键证据", "green"),
        ("Reason", "基于证据综合推理\n输出答案+证据页", "orange"),
    ]):
        add_card(s, Cm(3.0 + i*10.4), Cm(10.0), Cm(7.8), Cm(3.6), label, [desc], color)
        if i < 2:
            arr = s.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RIGHT_ARROW, Cm(10.7 + i*10.4), Cm(11.2), Cm(1.8), Cm(0.7))
            arr.fill.solid(); arr.fill.fore_color.rgb = RGBColor(180, 190, 205); arr.line.fill.background()
    add_footer(s, 2)

    # 3 Method
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "方法概览：ALR 推理范式 + SFT/EviGRPO 两阶段训练 + EGRA 分辨率分配")
    fit_picture(s, assets["fig1_alr_panel"], Cm(1.0), Cm(2.15), Cm(15.3), Cm(10.8))
    add_card(s, Cm(17.0), Cm(2.2), Cm(15.1), Cm(3.0), "ALR 输出结构", [
        "<think> 中显式拆成 Question Analysis / Evidence Localization / Reasoning Process。",
        "<answer> 同时给出 evidence_pages 与 final answer，便于验证与调试。",
    ], "blue")
    add_card(s, Cm(17.0), Cm(5.75), Cm(15.1), Cm(3.15), "训练框架", [
        "Stage I：用 teacher 生成高质量 ALR CoT；EM + 语义二次验证过滤噪声。",
        "Stage II：EviGRPO 采样候选输出，用格式、定位、答案三类 reward 联合优化。",
    ], "green")
    add_card(s, Cm(17.0), Cm(9.45), Cm(15.1), Cm(3.15), "EGRA：训练时提 SNR、降显存", [
        "证据页保持高分辨率；大部分非证据页降采样，少量保留高分辨率作为 distractor。",
        "既保留上下文，又把监督信号集中到关键页面。",
    ], "orange")
    add_footer(s, 3)

    # 4 Main results
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "主实验：DocSeeker 在 ID 与 OOD 文档理解中均达到强表现")
    rows = [
        ["Method", "DUDE", "MPDocVQA", "MMLong", "LongDoc", "SlideVQA"],
        ["GPT-4o", "54.1", "67.4", "42.8", "64.5", "-"],
        ["InternVL3-8B", "47.4", "80.8", "24.1", "38.7", "54.4"],
        ["Baseline-7B", "35.2", "70.1", "25.4", "37.8", "59.8"],
        ["Baseline-SFT(short)", "56.0", "82.9", "28.8", "42.7", "67.4"],
        ["DocSeeker-SFT", "56.8", "82.1", "38.6", "49.1", "75.2"],
        ["DocSeeker", "57.4", "86.2", "40.1", "51.7", "77.1"],
    ]
    make_table(s, Cm(1.0), Cm(2.25), Cm(20.1), Cm(6.0), rows, font_size=9.0)
    fit_picture(s, assets["fig1_results_panel"], Cm(1.0), Cm(8.8), Cm(14.8), Cm(6.6))
    add_card(s, Cm(22.2), Cm(2.35), Cm(10.3), Cm(4.0), "结果解读", [
        "短答案 SFT 主要提升 ID；OOD（MMLong/LongDoc）增益有限，说明易记忆。",
        "ALR CoT SFT 使 MMLong 从 28.8 → 38.6，证明结构化监督带来迁移能力。",
        "EviGRPO 在五项指标上继续稳定提升。",
    ], "green")
    add_card(s, Cm(22.2), Cm(7.1), Cm(10.3), Cm(3.8), "最重要的证据", [
        "DocSeeker 不依赖外部 retriever（RAG ×），仍在多项 benchmark 上领先开源方法。",
        "在 SlideVQA、MPDocVQA 上分别达到 77.1 F1、86.2 ANLS。",
    ], "blue")
    add_footer(s, 4)

    # 5 Ablations
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "消融：性能增益主要来自 ALR 数据、页面 ID、EGRA 与定位奖励")
    rows1 = [
        ["Data Config", "Size", "Acc", "F1"],
        ["Baseline", "-", "25.4", "20.8"],
        ["Raw short-answer", "6.3k", "27.4", "27.6"],
        ["Vanilla CoT", "6.3k", "31.3", "32.4"],
        ["ALR CoT", "6.3k", "33.8", "33.9"],
        ["ALR w/o Page id", "6.3k", "30.4", "31.1"],
        ["All ALR CoT", "13k", "38.6", "36.9"],
    ]
    rows2 = [
        ["Resolution Strategy", "Tokens", "Acc", "F1"],
        ["Fixed: Full Low-Res", "576", "34.2", "33.5"],
        ["Fixed: Truncated", "1024", "36.6", "35.8"],
        ["EGRA Full", "1024/256", "38.6", "36.9"],
        ["w/o Non-Evi Hi-Res", "1024/256", "35.5", "33.8"],
        ["w/o Low-Res", "1024", "34.5", "33.4"],
    ]
    rows3 = [
        ["GRPO Variant", "Acc", "F1"],
        ["SFT", "38.6", "36.9"],
        ["Vanilla GRPO", "38.7", "36.3"],
        ["EviGRPO", "40.1", "38.4"],
    ]
    make_table(s, Cm(1.0), Cm(2.25), Cm(10.3), Cm(6.0), rows1, font_size=8.3)
    make_table(s, Cm(12.0), Cm(2.25), Cm(10.3), Cm(5.2), rows2, font_size=8.3, header_fill=COLORS["green"])
    make_table(s, Cm(23.0), Cm(2.25), Cm(9.3), Cm(3.5), rows3, font_size=8.8, header_fill=COLORS["orange"])
    add_card(s, Cm(1.0), Cm(9.2), Cm(31.3), Cm(4.6), "机制结论", [
        "ALR > Vanilla CoT > Raw short-answer：不是“多写推理”本身，而是带证据定位约束的结构化推理更关键。",
        "去掉 Page ID 后明显下降：页面标识提供了可引用的 grounding anchor。",
        "EGRA 优于固定分辨率或删除非证据页：长上下文训练需要同时保留全局语境与证据细节。",
        "EviGRPO 优于只看答案的 GRPO：定位 reward 与答案 reward 互补，但权重过大/过小都会破坏平衡。",
    ], "blue")
    add_footer(s, 5)

    # 6 Analysis / RAG
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "分析实验：ALR 让模型在更长上下文与 RAG 噪声下保持稳定")
    fit_picture(s, assets["fig3_length"], Cm(1.0), Cm(2.15), Cm(14.0), Cm(6.8))
    fit_picture(s, assets["fig4_rag"], Cm(1.0), Cm(9.3), Cm(14.2), Cm(5.7))
    add_card(s, Cm(16.4), Cm(2.3), Cm(15.8), Cm(4.2), "长度鲁棒性", [
        "在 MMLongBench-doc 子集上，Baseline 随输入页数从 10 → 60 明显退化（34.5 → 13.9）。",
        "DocSeeker 曲线基本稳定，说明它能在干扰页中主动定位证据。",
        "Full-doc 表现接近 evidence-only 上界，定位能力是性能来源。",
    ], "green")
    add_card(s, Cm(16.4), Cm(7.2), Cm(15.8), Cm(4.4), "与视觉 RAG 协同", [
        "Baseline 在 k 增大后被噪声压垮；DocSeeker 对 k 的变化更稳健。",
        "检索器负责粗筛，DocSeeker 在 noisy top-k 中做细粒度定位与推理。",
        "解决实际系统中的 top-k dilemma：无需在 recall 与噪声之间过度折中。",
    ], "blue")
    add_card(s, Cm(16.4), Cm(12.2), Cm(15.8), Cm(2.6), "系统启示", [
        "长文档 RAG 不应只优化 retriever；reader 的 evidence grounding 能力同样决定最终鲁棒性。",
    ], "orange")
    add_footer(s, 6)

    # 7 Conclusion
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "结论：把“证据定位”内化到 MLLM，是长文档理解的关键能力")
    add_card(s, Cm(1.2), Cm(2.3), Cm(14.8), Cm(4.8), "主要结论", [
        "DocSeeker 通过 ALR 范式显式建模“分析—定位—推理”，提升可解释性和长上下文抗噪性。",
        "SFT 注入结构化能力，EviGRPO 进一步对齐定位与答案正确性，EGRA 解决训练代价问题。",
        "实验显示其在 ID/OOD、超长文档和 RAG 集成中均具备强泛化。",
    ], "green")
    add_card(s, Cm(17.0), Cm(2.3), Cm(15.0), Cm(4.8), "局限与风险", [
        "依赖 teacher 生成 ALR CoT 与二次验证，数据构建成本和 teacher bias 仍存在。",
        "证据定位 reward 权重敏感；过度追求定位可能牺牲答案识别。",
        "训练仍需要多 GPU；真实超长、多文档、多模态噪声场景还需更大规模验证。",
    ], "red")
    add_card(s, Cm(1.2), Cm(8.1), Cm(30.8), Cm(5.3), "可复用启示", [
        "① 对长上下文任务：让模型显式输出 grounding target，比单纯扩大上下文窗口更可控。",
        "② 对 RAG 系统：retriever 只做粗筛，reader 应具备抗噪定位与证据归因能力。",
        "③ 对训练策略：把格式、定位、答案拆成多维 reward，可把“推理过程质量”纳入优化目标。",
        "④ 对效率：按证据重要性分配视觉分辨率，是训练长视觉序列的实用折中。",
    ], "blue")
    add_badge(s, Cm(11.5), Cm(14.2), "Takeaway：Long-document understanding = retrieval + grounding + reasoning", "orange")
    add_footer(s, 7)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)


def validate_pptx(path: Path, max_slides: int = 7) -> None:
    if not path.exists():
        raise FileNotFoundError(path)
    size = path.stat().st_size
    if not (200_000 <= size <= 20_000_000):
        raise AssertionError(f"Unexpected file size: {size} bytes")
    # Requirement: python-pptx can reopen the generated file.
    prs = Presentation(path)
    slide_count = len(prs.slides)
    if slide_count > max_slides:
        raise AssertionError(f"Slide count {slide_count} exceeds {max_slides}")
    pic_count = sum(1 for slide in prs.slides for shape in slide.shapes if shape.shape_type == 13)
    if pic_count < 4:
        raise AssertionError(f"Expected embedded paper figures; found only {pic_count} pictures")
    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise AssertionError(f"Corrupt zip member: {bad}")
    print(f"OK: {path} | slides={slide_count} | pictures={pic_count} | size={size} bytes")


def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument("--output", type=Path, default=OUT)
    p.add_argument("--asset-dir", type=Path, default=ASSET_DIR)
    p.add_argument("--validate-only", action="store_true")
    return p.parse_args()


def main() -> None:
    args = parse_args()
    if not args.validate_only:
        assets = download_figures(args.asset_dir)
        build_deck(args.output, assets)
    validate_pptx(args.output)


if __name__ == "__main__":
    main()

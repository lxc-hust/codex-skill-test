#!/usr/bin/env python3
"""Generate a compact Chinese academic PPT for arXiv:2606.05749.

The script downloads the paper PDF, extracts original paper page/figure/table
screenshots with PyMuPDF, builds a <=7 slide deck with python-pptx, and validates
that the resulting PPTX can be reopened and contains embedded pictures.
"""
from __future__ import annotations

import argparse
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable

import fitz  # PyMuPDF
import requests
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

PAPER_ID = "2606.05749"
PAPER_URLS = [
    f"https://arxiv.org/pdf/{PAPER_ID}",
    f"https://arxiv.org/pdf/{PAPER_ID}.pdf",
    f"https://export.arxiv.org/pdf/{PAPER_ID}",
]
TITLE = "MARDoc: A Memory-Aware Refinement Agent Framework for Multimodal Long Document QA"
AUTHORS = "Kaifeng Chen, Hongtao Liu, Qiyao Peng, Jian Yang, Yongqiang Liu, Xiaochen Zhang, Qing Yang"

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)
BG = RGBColor(248, 250, 252)
NAVY = RGBColor(24, 45, 78)
BLUE = RGBColor(39, 97, 170)
TEAL = RGBColor(24, 132, 126)
GRAY = RGBColor(79, 89, 105)
LIGHT = RGBColor(229, 235, 244)
ACCENT = RGBColor(236, 129, 49)


@dataclass
class Asset:
    key: str
    path: Path
    page: int
    note: str


def download_pdf(dest: Path) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and dest.stat().st_size > 100_000:
        return dest
    headers = {"User-Agent": "Mozilla/5.0 (compatible; paper-compact-ppt/1.0)"}
    last_error: Exception | None = None
    for url in PAPER_URLS:
        try:
            with requests.get(url, headers=headers, timeout=60, stream=True) as resp:
                resp.raise_for_status()
                content_type = resp.headers.get("content-type", "")
                tmp = dest.with_suffix(".tmp")
                with tmp.open("wb") as fh:
                    for chunk in resp.iter_content(1024 * 128):
                        if chunk:
                            fh.write(chunk)
                if tmp.stat().st_size < 100_000:
                    raise RuntimeError(f"downloaded file too small from {url}: {tmp.stat().st_size} bytes")
                if "pdf" not in content_type.lower():
                    with tmp.open("rb") as fh:
                        if not fh.read(5).startswith(b"%PDF"):
                            raise RuntimeError(f"download from {url} does not look like a PDF: {content_type}")
                tmp.replace(dest)
                return dest
        except Exception as exc:  # keep trying mirrors in Actions
            last_error = exc
    raise RuntimeError(f"Could not download paper PDF; last error: {last_error}")


def page_text(doc: fitz.Document, idx: int) -> str:
    return doc[idx].get_text("text")


def find_page(doc: fitz.Document, required: Iterable[str], preferred: Iterable[str] = ()) -> int | None:
    required_l = [x.lower() for x in required]
    preferred_l = [x.lower() for x in preferred]
    best: tuple[int, int] | None = None
    for i in range(len(doc)):
        txt = page_text(doc, i).lower()
        if all(x in txt for x in required_l):
            score = sum(1 for x in preferred_l if x in txt)
            if best is None or score > best[0]:
                best = (score, i)
    return best[1] if best else None


def safe_render(page: fitz.Page, rect: fitz.Rect, out: Path) -> None:
    mat = fitz.Matrix(2.0, 2.0)
    pix = page.get_pixmap(matrix=mat, clip=rect, alpha=False)
    pix.save(out)


def placeholder_asset(out: Path, title: str, subtitle: str) -> None:
    out.parent.mkdir(parents=True, exist_ok=True)
    img = Image.new("RGB", (1600, 950), (247, 250, 253))
    draw = ImageDraw.Draw(img)
    try:
        font_big = ImageFont.truetype("DejaVuSans-Bold.ttf", 54)
        font_mid = ImageFont.truetype("DejaVuSans.ttf", 34)
    except Exception:
        font_big = font_mid = None
    draw.rounded_rectangle((40, 40, 1560, 910), radius=30, outline=(39, 97, 170), width=5)
    draw.text((90, 110), title, fill=(24, 45, 78), font=font_big)
    draw.text((90, 210), subtitle, fill=(79, 89, 105), font=font_mid)
    draw.text((90, 820), "注：若 PDF 图表定位失败，此占位图会让工作流显式暴露问题。", fill=(180, 80, 50), font=font_mid)
    img.save(out)


def extract_assets(pdf: Path, asset_dir: Path) -> list[Asset]:
    asset_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf)
    specs = [
        ("method", ["explorer", "refiner", "reflector"], ["overview", "framework", "mardoc"], "方法主图 / 框架页"),
        ("main_results", ["mmlongbench-doc", "docbench"], ["table", "overall", "qwen3", "docagent"], "主实验结果表"),
        ("ablation", ["ablation"], ["refiner", "reflector", "memory", "m_e", "m_r"], "消融实验表"),
        ("analysis", ["latency"], ["token", "cost", "evidence", "pages"], "效率/分析图表"),
    ]
    assets: list[Asset] = []
    used_pages: set[int] = set()
    for key, required, preferred, note in specs:
        page_idx = find_page(doc, required, preferred)
        if page_idx is None:
            # fallback: avoid duplicating pages when possible
            page_idx = next((i for i in range(min(len(doc), 10)) if i not in used_pages), 0)
        used_pages.add(page_idx)
        page = doc[page_idx]
        rect = page.rect
        # Full-page screenshots retain original table/figure context and are robust
        # across arXiv layout changes. PPT inserts them large enough for reading.
        crop = fitz.Rect(rect.x0 + rect.width * 0.02, rect.y0 + rect.height * 0.04, rect.x1 - rect.width * 0.02, rect.y1 - rect.height * 0.04)
        out = asset_dir / f"{key}_paper_page_{page_idx + 1}.png"
        try:
            safe_render(page, crop, out)
        except Exception:
            placeholder_asset(out, note, f"PDF page extraction failed for {key}")
        assets.append(Asset(key=key, path=out, page=page_idx + 1, note=note))
    doc.close()
    return assets


def set_bg(slide) -> None:
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = BG


def add_title(slide, title: str, subtitle: str | None = None) -> None:
    box = slide.shapes.add_textbox(Inches(0.35), Inches(0.16), Inches(12.55), Inches(0.55))
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = title
    p.font.size = Pt(24)
    p.font.bold = True
    p.font.color.rgb = NAVY
    if subtitle:
        sub = slide.shapes.add_textbox(Inches(0.38), Inches(0.72), Inches(12.3), Inches(0.28))
        stf = sub.text_frame
        stf.text = subtitle
        stf.paragraphs[0].font.size = Pt(9.5)
        stf.paragraphs[0].font.color.rgb = GRAY
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.35), Inches(0.95), Inches(12.6), Inches(0.02))
    line.fill.solid(); line.fill.fore_color.rgb = LIGHT
    line.line.color.rgb = LIGHT


def add_bullets(slide, x, y, w, h, bullets: list[str], font_size: int = 15, color=RGBColor(39, 48, 64)):
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.08)
    tf.margin_right = Inches(0.04)
    tf.clear()
    for idx, text in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = text
        p.level = 0
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.space_after = Pt(7)
    return box


def add_tag(slide, x, y, text: str, color=BLUE):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(1.55), Inches(0.32))
    shape.fill.solid(); shape.fill.fore_color.rgb = color
    shape.line.color.rgb = color
    tf = shape.text_frame; tf.clear(); tf.text = text
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER; p.font.size = Pt(10); p.font.bold = True; p.font.color.rgb = RGBColor(255, 255, 255)


def add_image(slide, path: Path, x, y, w, h):
    pic = slide.shapes.add_picture(str(path), x, y, width=w)
    if pic.height > h:
        pic.height = h
    # center vertically in the allocated box
    pic.top = y + int((h - pic.height) / 2)
    return pic


def asset_map(assets: list[Asset]) -> dict[str, Asset]:
    return {a.key: a for a in assets}


def build_ppt(output: Path, assets: list[Asset]) -> None:
    am = asset_map(assets)
    prs = Presentation()
    prs.slide_width = SLIDE_W
    prs.slide_height = SLIDE_H
    blank = prs.slide_layouts[6]

    # 1 title
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "MARDoc：面向多模态长文档问答的记忆感知细化 Agent", "arXiv:2606.05749 · 中文精炼汇报 · 7页以内")
    box = s.shapes.add_textbox(Inches(0.65), Inches(1.25), Inches(7.15), Inches(1.35))
    tf = box.text_frame; tf.clear(); tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = TITLE; p.font.size = Pt(22); p.font.bold = True; p.font.color.rgb = NAVY
    p2 = tf.add_paragraph(); p2.text = AUTHORS; p2.font.size = Pt(12); p2.font.color.rgb = GRAY
    add_bullets(s, Inches(0.72), Inches(3.05), Inches(6.95), Inches(2.35), [
        "一句话：用结构化记忆替代不断膨胀的交互历史，缓解长文档多跳 QA 的证据稀释与噪声累积。",
        "框架：Explorer 负责多粒度检索，Refiner 将轨迹压缩为证据/推理记忆，Reflector 判断充分性并反馈下一轮搜索。",
        "效果：在 MMLongBench-Doc 与 DocBench 上优于同骨干基线；Qwen3-30B 在 MMLongBench-Doc 达到 57.1% overall。",
        "价值：以适度 token/延迟开销换取更稳定的复杂文档理解。",
    ], 15)
    add_image(s, am["method"].path, Inches(8.05), Inches(1.25), Inches(4.7), Inches(5.65))
    add_tag(s, Inches(8.15), Inches(6.75), "原文框架截图", TEAL)

    # 2 motivation
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "背景与动机：长文档 QA 的核心瓶颈是“上下文污染”", "从一次性长上下文，转向可迭代、可检查、可压缩的证据管理")
    add_bullets(s, Inches(0.65), Inches(1.25), Inches(5.85), Inches(5.6), [
        "任务：多模态长文档 QA 需要跨页面、跨文本/表格/图像聚合证据。",
        "现有 agent 常把检索轨迹、观察、推理全部追加到同一上下文；轮次越多，关键信息越分散。",
        "多跳问题受影响最大：模型既要找证据，也要记住证据间依赖，噪声会放大错误推理。",
        "MARDoc 的假设：把“原始轨迹”压缩为结构化证据与推理记忆，再由反思器检查缺口，可降低噪声并保留关键依赖。",
    ], 17)
    # simple process diagram
    labels = [("长文档", BLUE), ("检索轨迹增长", ACCENT), ("证据稀释", ACCENT), ("结构化记忆", TEAL), ("更稳推理", TEAL)]
    x = Inches(7.0)
    for i, (lab, col) in enumerate(labels):
        y = Inches(1.35 + i * 1.02)
        shp = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y, Inches(4.9), Inches(0.62))
        shp.fill.solid(); shp.fill.fore_color.rgb = col; shp.line.color.rgb = col
        shp.text_frame.text = lab
        pp = shp.text_frame.paragraphs[0]; pp.font.size = Pt(18); pp.font.bold = True; pp.font.color.rgb = RGBColor(255, 255, 255); pp.alignment = PP_ALIGN.CENTER
        if i < len(labels) - 1:
            arr = s.shapes.add_textbox(Inches(9.25), y + Inches(0.62), Inches(0.5), Inches(0.28))
            arr.text_frame.text = "↓"; arr.text_frame.paragraphs[0].font.size = Pt(18); arr.text_frame.paragraphs[0].font.color.rgb = GRAY
    add_bullets(s, Inches(7.1), Inches(6.45), Inches(5.2), Inches(0.6), ["关键转变：从“全量历史”到“任务相关记忆”。"], 14, NAVY)

    # 3 method
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "方法概览：Explore–Refine–Reflect 的闭环记忆机制", "每轮只携带压缩后的结构化记忆，而非完整历史")
    add_image(s, am["method"].path, Inches(0.45), Inches(1.18), Inches(7.0), Inches(5.85))
    add_bullets(s, Inches(7.75), Inches(1.2), Inches(5.15), Inches(5.65), [
        "Document Processing：MinerU2.5 解析版面，视觉元素用 Qwen3-VL 生成可检索描述。",
        "Explorer：基于 ReAct 使用 search / get_section / get_image / get_table_image 等多粒度工具。",
        "Refiner：把交互轨迹压缩为 Evidence Memory Mᴱ 与 Reasoning Memory Mᴿ。",
        "Reflector：判断证据是否足够；不足则输出定向反馈 Rₜ，引导下一轮检索。",
        "创新点：证据节点与推理链分离，便于定位“证据错”还是“逻辑错”。",
    ], 14)
    add_tag(s, Inches(0.55), Inches(6.85), f"截图页 p.{am['method'].page}", TEAL)

    # 4 main results
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "主实验：同骨干下稳定超过迭代检索/推理基线", "基准：MMLongBench-Doc、DocBench；骨干：Qwen3-VL 30B/8B 系列")
    add_image(s, am["main_results"].path, Inches(0.45), Inches(1.15), Inches(7.35), Inches(5.95))
    add_bullets(s, Inches(8.0), Inches(1.2), Inches(4.95), Inches(5.65), [
        "MARDoc 在两个多模态长文档 QA 基准上均展现强结果。",
        "Qwen3-30B 设置下，MMLongBench-Doc overall 为 57.1%，接近使用 Claude 3.5 Sonnet 的 DocAgent。",
        "优势主要来自对跨页证据的保留与去噪，而不是简单扩展上下文。",
        "说明：结构化记忆可以提升同一 MLLM 骨干的 agentic document QA 上限。",
    ], 15)
    add_tag(s, Inches(0.55), Inches(6.85), f"原文结果表 p.{am['main_results'].page}", BLUE)

    # 5 ablation
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "消融：Refiner/Reflector 与结构化记忆是核心增益来源", "不只是压缩文本，而是保留事实节点与推理依赖")
    add_image(s, am["ablation"].path, Inches(0.45), Inches(1.15), Inches(7.35), Inches(5.95))
    add_bullets(s, Inches(8.0), Inches(1.2), Inches(4.95), Inches(5.65), [
        "移除 Refiner 会显著退化：Reflector 面对噪声历史时难以准确判断证据缺口。",
        "Mᴱ 偏事实 grounding，对不可回答问题和幻觉抑制更关键。",
        "Mᴿ 偏逻辑依赖，对多跳问题的推理连贯性更关键。",
        "普通文本压缩效果较差，证明“结构化表示”本身是必要设计。",
        "Reflector 的反馈 Rₜ 能随迭代提升单跳/多跳准确率。",
    ], 14)
    add_tag(s, Inches(0.55), Inches(6.85), f"原文消融页 p.{am['ablation'].page}", ACCENT)

    # 6 analysis/cost
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "分析与代价：多跳/跨页场景收益更明显，开销可控", "适合证据分散、需要迭代检索与校验的复杂文档理解")
    add_image(s, am["analysis"].path, Inches(0.45), Inches(1.15), Inches(7.35), Inches(5.95))
    add_bullets(s, Inches(8.0), Inches(1.2), Inches(4.95), Inches(5.65), [
        "细粒度性能分析：证据页数增加时，MARDoc 的性能下降小于基线，体现更强多跳鲁棒性。",
        "文档处理/工具消融：细粒度语义节点与多粒度工具显著改善检索有效性。",
        "计算成本：相比 DocAgent，约增加 22.5k tokens 与 15.9s/sample 延迟，但带来约 7.8% accuracy gain。",
        "取舍：更高推理可靠性换取适度推理开销；可通过早停或训练化优化进一步降低成本。",
    ], 14)
    add_tag(s, Inches(0.55), Inches(6.85), f"原文分析页 p.{am['analysis'].page}", TEAL)

    # 7 conclusion
    s = prs.slides.add_slide(blank); set_bg(s)
    add_title(s, "结论：把记忆从上下文中“外置并结构化”", "MARDoc 对长文档 Agent 设计的可复用启示")
    add_bullets(s, Inches(0.7), Inches(1.25), Inches(6.1), Inches(5.7), [
        "为什么有效：过滤无关轨迹、保留证据来源、显式维护推理依赖，并让反思器针对缺口发起下一轮检索。",
        "主要结论：结构化记忆比完整历史更适合迭代式多模态长文档 QA。",
        "局限：依赖 prompt 工程；主要验证 Qwen3-VL 家族；迭代闭环带来额外 latency/token。",
        "启示 1：长上下文不是万能，信息组织形式同样决定推理质量。",
        "启示 2：复杂 QA agent 应区分检索、记忆压缩与充分性检查三个职责。",
        "启示 3：未来可做训练化 Refiner/Reflector、跨骨干泛化与自适应早停。",
    ], 15)
    # mini architecture card
    card = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(7.25), Inches(1.35), Inches(5.25), Inches(4.85))
    card.fill.solid(); card.fill.fore_color.rgb = RGBColor(255,255,255); card.line.color.rgb = LIGHT
    add_tag(s, Inches(7.55), Inches(1.65), "Takeaway", NAVY)
    add_bullets(s, Inches(7.55), Inches(2.25), Inches(4.6), Inches(3.4), [
        "Agent 不是只需要更多上下文，\n而是需要更好的工作记忆。",
        "Mᴱ = 我知道什么事实；\nMᴿ = 这些事实如何支撑答案；\nRₜ = 下一步该补什么证据。",
    ], 18, NAVY)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)


def validate_pptx(path: Path, max_slides: int = 7) -> dict[str, object]:
    if not path.exists():
        raise AssertionError(f"PPTX not found: {path}")
    size = path.stat().st_size
    if size < 100_000:
        raise AssertionError(f"PPTX file is suspiciously small: {size} bytes")
    prs = Presentation(path)
    slide_count = len(prs.slides)
    if slide_count > max_slides:
        raise AssertionError(f"slide count {slide_count} exceeds {max_slides}")
    titles = []
    picture_count = 0
    for slide in prs.slides:
        title = ""
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.text.strip() and not title:
                title = shape.text.strip().split("\n", 1)[0]
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                picture_count += 1
        titles.append(title)
    if picture_count < 4:
        raise AssertionError(f"expected at least 4 embedded pictures, found {picture_count}")
    with zipfile.ZipFile(path) as zf:
        bad = zf.testzip()
        if bad is not None:
            raise AssertionError(f"zip integrity check failed at {bad}")
    return {"path": str(path), "size": size, "slides": slide_count, "pictures": picture_count, "titles": titles}


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--output", default="outputs/paper_compact_ppt.pptx", type=Path)
    ap.add_argument("--workdir", default="outputs/_paper_assets", type=Path)
    ap.add_argument("--pdf", default=None, type=Path)
    args = ap.parse_args(argv)

    pdf = args.pdf or (args.workdir / f"arxiv_{PAPER_ID}.pdf")
    if args.pdf is None:
        pdf = download_pdf(pdf)
    elif not pdf.exists():
        raise FileNotFoundError(pdf)

    assets = extract_assets(pdf, args.workdir)
    build_ppt(args.output, assets)
    meta = validate_pptx(args.output)
    print("Generated PPTX:", meta["path"])
    print("File size:", meta["size"])
    print("Slide count:", meta["slides"])
    print("Embedded pictures:", meta["pictures"])
    print("Slide titles:")
    for i, title in enumerate(meta["titles"], 1):
        print(f"  {i}. {title}")
    print("Original screenshot assets:")
    for a in assets:
        print(f"  - {a.key}: {a.path} (paper page {a.page}; {a.note})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

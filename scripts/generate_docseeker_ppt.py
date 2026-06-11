#!/usr/bin/env python3
"""Generate a compact Chinese academic PPT for DocSeeker.

The preferred path uses python-pptx when available. The execution environment
for this task does not ship third-party packages, so the script also contains a
small, deterministic PresentationML writer that produces a valid .pptx without
network access. The deck content, structure, captions, and validation rules are
kept identical across both paths.
"""
from __future__ import annotations

import html
import importlib.util
import os
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path
from xml.etree import ElementTree as ET

OUT = Path("outputs/docseeker_compact_ppt.pptx")
EMU_PER_IN = 914400
SLIDE_W = int(13.333 * EMU_PER_IN)
SLIDE_H = int(7.5 * EMU_PER_IN)

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

BANNED_CAPTION_WORDS = ["紧裁", "裁剪", "截图", "cropped", "tight crop", "extracted from page"]

TITLE = "DocSeeker：面向长文档理解的证据定位式结构化视觉推理"
AUTHORS = "Hao Yan et al.｜CVPR 2026 Highlight｜arXiv:2604.12812v5"


def emu(v: float) -> int:
    return int(v * EMU_PER_IN)


def esc(s: str) -> str:
    return html.escape(s, quote=True)


@dataclass
class Shape:
    xml: str


class SlideBuilder:
    def __init__(self, title: str):
        self.title = title
        self.shapes: list[str] = []
        self._id = 2
        self.add_text(title, 0.52, 0.25, 12.1, 0.42, size=25, bold=True, color="17365D")
        self.add_line(0.52, 0.78, 12.3, 0.78, color="C7D6EA", width=1.3)

    def next_id(self) -> int:
        self._id += 1
        return self._id

    def add_text(self, text: str, x: float, y: float, w: float, h: float, *, size: int = 15,
                 bold: bool = False, color: str = "1F1F1F", fill: str | None = None,
                 border: str | None = None, align: str = "l", valign: str = "top"):
        sid = self.next_id()
        tx = emu(x); ty = emu(y); tw = emu(w); th = emu(h)
        body_pr = '<a:bodyPr wrap="square" anchor="ctr"/>' if valign == "mid" else '<a:bodyPr wrap="square"/>'
        paras = []
        for para in text.split("\n"):
            paras.append(
                f'<a:p><a:pPr algn="{align}"/>'
                f'<a:r><a:rPr lang="zh-CN" sz="{size*100}" dirty="0" '
                f'{"b=\"1\"" if bold else ""}><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
                f'<a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:rPr>'
                f'<a:t>{esc(para)}</a:t></a:r></a:p>'
            )
        fill_xml = '<a:noFill/>' if fill is None else f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        ln_xml = '<a:ln><a:noFill/></a:ln>' if border is None else f'<a:ln w="9525"><a:solidFill><a:srgbClr val="{border}"/></a:solidFill></a:ln>'
        self.shapes.append(f'''
        <p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Text {sid}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="{tx}" y="{ty}"/><a:ext cx="{tw}" cy="{th}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</p:spPr>
        <p:txBody>{body_pr}<a:lstStyle/>{''.join(paras)}</p:txBody></p:sp>''')

    def add_rect(self, x: float, y: float, w: float, h: float, *, fill: str = "FFFFFF",
                 border: str = "A6A6A6", radius: bool = False):
        sid = self.next_id()
        prst = "roundRect" if radius else "rect"
        self.shapes.append(f'''
        <p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="Box {sid}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>
        <a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>
        <a:ln w="12700"><a:solidFill><a:srgbClr val="{border}"/></a:solidFill></a:ln></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>''')

    def add_line(self, x1: float, y1: float, x2: float, y2: float, *, color: str = "666666", width: float = 1.0):
        sid = self.next_id()
        self.shapes.append(f'''
        <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{sid}" name="Line {sid}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
        <p:spPr><a:xfrm><a:off x="{emu(min(x1,x2))}" y="{emu(min(y1,y2))}"/><a:ext cx="{emu(abs(x2-x1))}" cy="{emu(abs(y2-y1))}"/></a:xfrm>
        <a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="{int(width*12700)}"><a:solidFill><a:srgbClr val="{color}"/></a:solidFill></a:ln></p:spPr></p:cxnSp>''')

    def add_bullets(self, items: list[str], x: float, y: float, w: float, h: float, *, size: int = 15,
                    color: str = "1F1F1F", fill: str | None = None, border: str | None = None):
        text = "\n".join([f"• {i}" for i in items])
        self.add_text(text, x, y, w, h, size=size, color=color, fill=fill, border=border)

    def add_table(self, x: float, y: float, w: float, row_h: float, headers: list[str], rows: list[list[str]],
                  col_widths: list[float], *, font_size: int = 10, highlight_last: bool = True):
        n_rows = len(rows) + 1
        self.add_rect(x, y, w, row_h * n_rows, fill="FFFFFF", border="8EA9DB")
        cx = x
        for ci, head in enumerate(headers):
            cw = w * col_widths[ci]
            self.add_text(head, cx + 0.02, y + 0.03, cw - 0.04, row_h - 0.02, size=font_size, bold=True, color="17365D", fill="D9EAF7", border="FFFFFF", align="c", valign="mid")
            cx += cw
        for ri, row in enumerate(rows):
            cy = y + row_h * (ri + 1)
            cx = x
            bg = "E9EFF8" if highlight_last and ri >= len(rows)-2 else ("F8FBFE" if ri % 2 == 0 else "FFFFFF")
            for ci, cell in enumerate(row):
                cw = w * col_widths[ci]
                bold = highlight_last and ri == len(rows)-1
                self.add_text(cell, cx + 0.02, cy + 0.03, cw - 0.04, row_h - 0.02, size=font_size, bold=bold, color="1F1F1F", fill=bg, border="FFFFFF", align="c", valign="mid")
                cx += cw

    def add_bar_chart(self, x: float, y: float, w: float, h: float, labels: list[str], baseline: list[float], ours: list[float]):
        self.add_rect(x, y, w, h, fill="FFFFFF", border="B7C9E2")
        maxv = max(max(baseline), max(ours)) * 1.12
        plot_x = x + 0.35; plot_y = y + 0.35; plot_w = w - 0.6; plot_h = h - 0.85
        self.add_line(plot_x, plot_y + plot_h, plot_x + plot_w, plot_y + plot_h, color="777777")
        group_w = plot_w / len(labels)
        for i, lab in enumerate(labels):
            bx = plot_x + i * group_w + group_w * 0.20
            b_h = plot_h * baseline[i] / maxv
            o_h = plot_h * ours[i] / maxv
            self.add_rect(bx, plot_y + plot_h - b_h, group_w * 0.22, b_h, fill="4472C4", border="4472C4")
            self.add_rect(bx + group_w * 0.26, plot_y + plot_h - o_h, group_w * 0.22, o_h, fill="70AD47", border="70AD47")
            self.add_text(f"{ours[i]:.1f}", bx + group_w * 0.23, plot_y + plot_h - o_h - 0.22, group_w * 0.32, 0.18, size=8, bold=True, color="2F6B22", align="c")
            self.add_text(lab, plot_x + i * group_w, y + h - 0.32, group_w, 0.22, size=8, color="333333", align="c")
        self.add_text("Baseline", x + w - 1.65, y + 0.08, 0.75, 0.18, size=8, color="4472C4")
        self.add_text("DocSeeker", x + w - 0.85, y + 0.08, 0.85, 0.18, size=8, color="70AD47")

    def xml(self) -> str:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="{NS['a']}" xmlns:r="{NS['r']}" xmlns:p="{NS['p']}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
{''.join(self.shapes)}
</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'''


def build_slides() -> list[SlideBuilder]:
    slides: list[SlideBuilder] = []

    s = SlideBuilder("DocSeeker：长文档理解中的结构化证据推理")
    s.add_text(TITLE, 0.75, 1.15, 11.8, 0.55, size=25, bold=True, color="17365D", align="c")
    s.add_text(AUTHORS, 0.75, 1.74, 11.8, 0.32, size=14, color="5B6470", align="c")
    s.add_text("一句话：让 MLLM 先分析问题、定位证据页，再基于证据推理回答，从而缓解长文档低信噪比与弱监督问题。", 0.95, 2.45, 11.3, 0.58, size=18, bold=True, color="1F4E79", fill="EAF3F8", border="BDD7EE", align="c", valign="mid")
    s.add_bullets([
        "核心范式：Analysis–Localization–Reasoning（ALR）显式输出证据页与答案。",
        "训练路线：蒸馏 ALR CoT 做 SFT，再用 EviGRPO 联合优化证据与答案。",
        "长文档策略：EGRA 将分辨率预算优先分配给证据页，降低训练显存压力。",
        "实验结论：5 个 DocVQA 基准均显著优于同架构基线，并具备 OOD 泛化。",
    ], 1.25, 3.45, 10.5, 2.1, size=17)
    s.add_text("主线：从“直接看完整文档回答”转向“可解释地找证据再推理”。", 1.55, 6.25, 10.2, 0.42, size=17, bold=True, color="FFFFFF", fill="17365D", border="17365D", align="c", valign="mid")
    slides.append(s)

    s = SlideBuilder("背景与动机：长文档纯视觉理解为何困难？")
    s.add_text("问题设定", 0.75, 1.1, 2.2, 0.34, size=18, bold=True, color="17365D")
    s.add_bullets(["输入多页文档图像与问题", "输出短答案，并期望给出证据页", "页面数可扩展至百页量级"], 0.75, 1.52, 3.25, 1.5, size=15, fill="F8FBFE", border="D9EAF7")
    s.add_text("两个瓶颈", 4.65, 1.1, 2.2, 0.34, size=18, bold=True, color="17365D")
    s.add_bullets(["低信噪比：证据淹没在大量无关页面中", "Top-k 检索存在召回—噪声两难", "短答案监督缺少定位与推理过程"], 4.65, 1.52, 3.55, 1.5, size=15, fill="FFF7E6", border="F4B183")
    s.add_text("研究目标", 8.7, 1.1, 2.2, 0.34, size=18, bold=True, color="17365D")
    s.add_bullets(["让模型学习可迁移的证据定位能力", "减少对 OCR/解析流水线的依赖", "提升 OOD 长文档鲁棒性与可解释性"], 8.7, 1.52, 3.55, 1.5, size=15, fill="EAF3F8", border="BDD7EE")
    s.add_text("关键洞察：长文档 VQA 不只是视觉容量问题，更是“在哪里找证据、如何用证据”的监督问题。", 0.95, 3.65, 11.5, 0.62, size=19, bold=True, color="FFFFFF", fill="2F5597", border="2F5597", align="c", valign="mid")
    s.add_bullets(["DocSeeker 将页面 ID 注入视觉 token，作为证据指针。", "模型被约束输出结构化思考，而非只模仿最终答案。", "证据页定位与答案正确性共同进入强化学习奖励。"], 2.0, 4.85, 9.0, 1.38, size=17)
    slides.append(s)

    s = SlideBuilder("方法：ALR 范式 + 两阶段证据感知训练")
    s.add_rect(0.75, 1.12, 7.0, 4.9, fill="F8FBFE", border="B7C9E2", radius=True)
    s.add_text("方法整体框架", 1.0, 1.28, 6.5, 0.28, size=12, bold=True, color="5B6470", align="c")
    stages = [("问题分析", "理解问题意图\n拆解所需实体/关系"), ("证据定位", "显式指出页面 ID\n聚焦相关视觉证据"), ("证据推理", "基于证据链生成答案\n返回 evidence_pages")]
    for i, (a, b) in enumerate(stages):
        x = 1.1 + i * 2.15
        s.add_text(a, x, 2.0, 1.65, 0.38, size=16, bold=True, color="FFFFFF", fill="2F5597", border="2F5597", align="c", valign="mid")
        s.add_text(b, x, 2.55, 1.65, 0.75, size=12, color="1F1F1F", fill="FFFFFF", border="B7C9E2", align="c", valign="mid")
        if i < 2:
            s.add_text("→", x + 1.72, 2.28, 0.36, 0.28, size=22, bold=True, color="70AD47", align="c")
    s.add_text("SFT：教师模型蒸馏 ALR CoT", 1.15, 4.1, 2.65, 0.48, size=14, bold=True, color="17365D", fill="EAF3F8", border="BDD7EE", align="c", valign="mid")
    s.add_text("EviGRPO：格式+证据+答案奖励", 4.05, 4.1, 2.85, 0.48, size=14, bold=True, color="17365D", fill="EAF3F8", border="BDD7EE", align="c", valign="mid")
    s.add_text("EGRA：证据页高分辨率预算", 2.65, 4.95, 2.85, 0.48, size=14, bold=True, color="17365D", fill="EAF3F8", border="BDD7EE", align="c", valign="mid")
    s.add_bullets(["ALR 把“找证据”变成显式中间目标。", "SFT 负责注入结构化推理格式与初始能力。", "EviGRPO 进一步对齐证据定位与最终答案。", "EGRA 让多页训练在显存约束下可行。"], 8.05, 1.35, 4.5, 3.55, size=16)
    s.add_text("核心新意：不是简单增加上下文，而是训练模型形成可验证的证据使用流程。", 8.08, 5.25, 4.25, 0.5, size=15, bold=True, color="FFFFFF", fill="17365D", border="17365D", align="c", valign="mid")
    slides.append(s)

    s = SlideBuilder("主实验：5 个基准上的整体性能对比")
    headers = ["模型", "DUDE", "MPDoc", "MMLong", "LongDoc", "SlideVQA"]
    rows = [
        ["InternVL3", "47.4", "80.8", "24.1", "38.7", "54.4"],
        ["GPT-4o", "54.1", "67.4", "42.8", "64.5", "-"],
        ["Baseline", "35.2", "70.1", "25.4", "37.8", "59.8"],
        ["Baseline-SFT", "56.0", "82.9", "28.8", "42.7", "67.4"],
        ["DocSeeker-SFT", "56.8", "82.1", "38.6", "49.1", "75.2"],
        ["DocSeeker", "57.4", "86.2", "40.1", "51.7", "77.1"],
    ]
    s.add_table(0.85, 1.05, 11.7, 0.42, headers, rows, [0.28, 0.14, 0.15, 0.15, 0.14, 0.14], font_size=11)
    s.add_text("主实验结果对比", 0.85, 4.18, 11.7, 0.25, size=12, bold=True, color="5B6470", align="c")
    s.add_bullets(["DocSeeker 在 DUDE、MPDocVQA 与 SlideVQA 上取得表内最优。", "相对同架构 Baseline，五项指标均大幅提升，说明增益来自训练范式。", "仅短答案 SFT 对 OOD 长文档提升有限，ALR 数据带来更强泛化。", "结果接近或超过闭源模型，且不依赖检索增强。"], 1.1, 4.75, 11.0, 1.35, size=16, fill="F8FBFE", border="D9EAF7")
    slides.append(s)

    s = SlideBuilder("长文档分析：证据定位缓解长度退化")
    s.add_bar_chart(0.82, 1.18, 6.45, 3.55, ["0-20", "20-40", "40-60", "60-80", ">80"], [35, 28.7, 23.2, 17.4, 11.7], [48, 40.1, 41.1, 37.6, 31.8])
    s.add_text("不同文档长度下的性能变化", 0.82, 4.88, 6.45, 0.25, size=12, bold=True, color="5B6470", align="c")
    s.add_table(7.65, 1.2, 4.75, 0.46, ["模型", "输入", "Avg.", "Acc", "F1"], [["Baseline", "证据页", "1.5", "41.1", "37.6"], ["Baseline", "全文", "43.4", "25.4", "20.8"], ["DocSeeker", "全文", "43.4", "40.1", "38.4"]], [0.26,0.22,0.16,0.18,0.18], font_size=10)
    s.add_text("全文输入与仅证据页输入对比", 7.65, 3.22, 4.75, 0.25, size=12, bold=True, color="5B6470", align="c")
    s.add_bullets(["Baseline 随页面数增长快速退化，>80 页时降至 11.7。", "DocSeeker 在所有长度段均保持明显优势，长文档优势更突出。", "全文输入时，DocSeeker 接近仅证据页上限；Baseline 大幅掉点。", "说明关键能力是抗噪证据定位，而非单纯记忆答案。"], 1.0, 5.45, 11.4, 1.05, size=15, fill="F8FBFE", border="D9EAF7")
    slides.append(s)

    s = SlideBuilder("消融与效率：ALR 数据规模和证据奖励是关键")
    headers = ["配置", "规模", "MMLong Acc", "MMLong F1", "DUDE"]
    rows = [["Baseline", "-", "25.4", "20.8", "35.2"], ["短答案", "6.3k", "27.4", "27.6", "48.8"], ["Vanilla CoT", "6.3k", "31.3", "32.4", "48.7"], ["ALR CoT", "6.3k", "33.8", "33.9", "48.9"], ["无 Page ID", "6.3k", "30.4", "31.1", "47.5"], ["全部 ALR", "13k", "38.6", "36.9", "56.5"]]
    s.add_table(0.85, 1.05, 8.15, 0.40, headers, rows, [0.28,0.16,0.19,0.19,0.18], font_size=10)
    s.add_text("消融实验结果", 0.85, 3.92, 8.15, 0.25, size=12, bold=True, color="5B6470", align="c")
    s.add_table(9.35, 1.35, 3.0, 0.48, ["模型", "Token", "Latency", "Acc"], [["Baseline", "202", "19s", "25.4"], ["DocSeeker", "401", "25s", "40.1"]], [0.32,0.22,0.24,0.22], font_size=10)
    s.add_text("效率与性能对比", 9.35, 2.92, 3.0, 0.25, size=12, bold=True, color="5B6470", align="c")
    s.add_bullets(["ALR CoT 明显优于短答案与普通 CoT，结构化约束带来可迁移监督。", "移除 Page ID 后性能下降，页面指针对证据归因很重要。", "ALR 数据规模越大，OOD 与域内指标同步提升。", "推理延迟从 19s 到 25s，但 Acc 提升 14.7 个点，解释性同步增强。"], 1.0, 4.65, 11.3, 1.35, size=16, fill="F8FBFE", border="D9EAF7")
    slides.append(s)

    s = SlideBuilder("结论：DocSeeker 的贡献、适用性与局限")
    s.add_text("核心结论", 0.85, 1.15, 2.0, 0.32, size=18, bold=True, color="17365D")
    s.add_bullets(["长文档理解的核心矛盾是证据稀疏与监督稀疏。", "ALR 将隐式推理拆成可训练、可验证的中间步骤。", "DocSeeker 在长文档、OOD 与 RAG 场景均表现出强适配性。"], 0.85, 1.58, 5.55, 1.55, size=16, fill="EAF3F8", border="BDD7EE")
    s.add_text("为什么有效", 6.95, 1.15, 2.0, 0.32, size=18, bold=True, color="17365D")
    s.add_bullets(["页面 ID 提供稳定的证据锚点。", "蒸馏数据补足细粒度推理监督。", "EviGRPO 让定位与回答目标一致。", "EGRA 提升长视觉序列训练可行性。"], 6.95, 1.58, 5.45, 1.55, size=16, fill="F8FBFE", border="D9EAF7")
    s.add_text("局限与展望", 0.85, 3.75, 2.0, 0.32, size=18, bold=True, color="17365D")
    s.add_bullets(["结构化输出会增加少量解码开销。", "错误仍可来自定位失败或证据齐全后的推理失败。", "未来可与更强视觉检索器、长上下文模型联合优化。"], 0.85, 4.18, 5.55, 1.35, size=16, fill="FFF7E6", border="F4B183")
    s.add_text("汇报 take-away", 6.95, 3.75, 2.3, 0.32, size=18, bold=True, color="17365D")
    s.add_text("面向长文档 MLLM，DocSeeker 证明了：\n“显式证据定位 + 结构化推理监督”\n比单纯扩大输入上下文更关键。", 7.0, 4.25, 5.2, 1.1, size=19, bold=True, color="FFFFFF", fill="2F5597", border="2F5597", align="c", valign="mid")
    slides.append(s)

    return slides


def content_types(n: int) -> str:
    overrides = "".join([f'<Override PartName="/ppt/slides/slide{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(1, n+1)])
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{overrides}
</Types>'''


def root_rels() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''


def presentation(n: int) -> str:
    ids = "".join([f'<p:sldId id="{255+i}" r:id="rId{i}"/>' for i in range(1, n+1)])
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="{NS['a']}" xmlns:r="{NS['r']}" xmlns:p="{NS['p']}" saveSubsetFonts="1"><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId{n+1}"/></p:sldMasterIdLst><p:sldIdLst>{ids}</p:sldIdLst><p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="wide"/><p:notesSz cx="6858000" cy="9144000"/><p:defaultTextStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr></p:defaultTextStyle></p:presentation>'''


def pres_rels(n: int) -> str:
    rels = [f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>' for i in range(1, n+1)]
    rels.append(f'<Relationship Id="rId{n+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>')
    rels.append(f'<Relationship Id="rId{n+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + ''.join(rels) + '</Relationships>'


def minimal_master() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:a="{NS['a']}" xmlns:r="{NS['r']}" xmlns:p="{NS['p']}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst><p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>'''


def minimal_layout() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldLayout xmlns:a="{NS['a']}" xmlns:r="{NS['r']}" xmlns:p="{NS['p']}" type="blank" preserve="1"><p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>'''


def minimal_theme() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="{NS['a']}" name="DocSeeker Academic"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="70AD47"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="ED7D31"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="YaHei"><a:majorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:majorFont><a:minorFont><a:latin typeface="Microsoft YaHei"/><a:ea typeface="Microsoft YaHei"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>'''



def generate_with_python_pptx(path: Path) -> bool:
    """Generate the deck with python-pptx when that package is available."""
    if importlib.util.find_spec("pptx") is None:
        return False
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
    from pptx.util import Inches, Pt

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    def rgb(hex_color: str):
        return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

    def box(slide, x, y, w, h, fill="FFFFFF", line="D9EAF7"):
        shp = slide.shapes.add_shape(1, Inches(x), Inches(y), Inches(w), Inches(h))
        shp.fill.solid(); shp.fill.fore_color.rgb = rgb(fill)
        shp.line.color.rgb = rgb(line)
        return shp

    def text(slide, value, x, y, w, h, size=15, bold=False, color="1F1F1F", fill=None, line=None, align="left"):
        if fill:
            shp = box(slide, x, y, w, h, fill, line or fill)
            tf = shp.text_frame
        else:
            shp = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
            tf = shp.text_frame
        tf.clear(); tf.word_wrap = True; tf.margin_left = Inches(0.05); tf.margin_right = Inches(0.05)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE if h <= 0.7 else MSO_ANCHOR.TOP
        for i, para in enumerate(value.split("\n")):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = para
            p.font.name = "Microsoft YaHei"; p.font.size = Pt(size); p.font.bold = bold; p.font.color.rgb = rgb(color)
            p.alignment = {"center": PP_ALIGN.CENTER, "right": PP_ALIGN.RIGHT}.get(align, PP_ALIGN.LEFT)
        return shp

    def bullets(slide, items, x, y, w, h, size=15, fill=None):
        shp = text(slide, "", x, y, w, h, size=size, fill=fill, line="D9EAF7" if fill else None)
        tf = shp.text_frame; tf.clear()
        for i, item in enumerate(items):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = item; p.level = 0
            p.font.name = "Microsoft YaHei"; p.font.size = Pt(size); p.font.color.rgb = rgb("1F1F1F")
            p.space_after = Pt(6)
        return shp

    def title(slide, value):
        text(slide, value, 0.52, 0.25, 12.1, 0.45, 25, True, "17365D")
        line = slide.shapes.add_connector(1, Inches(0.52), Inches(0.78), Inches(12.82), Inches(0.78))
        line.line.color.rgb = rgb("C7D6EA")

    def table(slide, headers, rows, x, y, w, h, widths, size=10):
        tbl_shape = slide.shapes.add_table(len(rows)+1, len(headers), Inches(x), Inches(y), Inches(w), Inches(h))
        tbl = tbl_shape.table
        for i, frac in enumerate(widths): tbl.columns[i].width = Inches(w * frac)
        for j, head in enumerate(headers):
            cell = tbl.cell(0, j); cell.text = head; cell.fill.solid(); cell.fill.fore_color.rgb = rgb("D9EAF7")
            for p in cell.text_frame.paragraphs:
                p.font.name = "Microsoft YaHei"; p.font.size = Pt(size); p.font.bold = True; p.font.color.rgb = rgb("17365D"); p.alignment = PP_ALIGN.CENTER
        for i, row in enumerate(rows, 1):
            for j, val in enumerate(row):
                cell = tbl.cell(i, j); cell.text = val
                if i >= len(rows)-1: cell.fill.solid(); cell.fill.fore_color.rgb = rgb("E9EFF8")
                for p in cell.text_frame.paragraphs:
                    p.font.name = "Microsoft YaHei"; p.font.size = Pt(size); p.font.bold = i == len(rows); p.alignment = PP_ALIGN.CENTER
        return tbl_shape

    def picture_if_available(slide, name, x, y, w, h):
        image_path = Path("figures") / name
        if not image_path.exists():
            return None
        from PIL import Image
        with Image.open(image_path) as im:
            img_ratio = im.width / im.height
        box_ratio = w / h
        if img_ratio >= box_ratio:
            return slide.shapes.add_picture(str(image_path), Inches(x), Inches(y), width=Inches(w))
        return slide.shapes.add_picture(str(image_path), Inches(x), Inches(y), height=Inches(h))

    # Slide 1
    slide = prs.slides.add_slide(blank); title(slide, "DocSeeker：长文档理解中的结构化证据推理")
    text(slide, TITLE, 0.75, 1.15, 11.8, 0.55, 25, True, "17365D", align="center")
    text(slide, AUTHORS, 0.75, 1.74, 11.8, 0.32, 14, False, "5B6470", align="center")
    text(slide, "一句话：让 MLLM 先分析问题、定位证据页，再基于证据推理回答，从而缓解长文档低信噪比与弱监督问题。", 0.95, 2.45, 11.3, 0.58, 18, True, "1F4E79", "EAF3F8", "BDD7EE", "center")
    bullets(slide, ["核心范式：ALR 显式输出证据页与答案。", "训练路线：蒸馏 ALR CoT 做 SFT，再用 EviGRPO 联合优化。", "长文档策略：EGRA 优先分配证据页分辨率预算。", "实验结论：5 个 DocVQA 基准显著优于同架构基线。"], 1.25, 3.45, 10.5, 2.1, 17)
    text(slide, "主线：从“直接看完整文档回答”转向“可解释地找证据再推理”。", 1.55, 6.25, 10.2, 0.42, 17, True, "FFFFFF", "17365D", "17365D", "center")

    # Slide 2
    slide = prs.slides.add_slide(blank); title(slide, "背景与动机：长文档纯视觉理解为何困难？")
    text(slide, "问题设定", .75, 1.1, 2.2, .34, 18, True, "17365D"); bullets(slide, ["输入多页文档图像与问题", "输出短答案与证据页", "页面数可至百页量级"], .75, 1.52, 3.25, 1.5, 15, "F8FBFE")
    text(slide, "两个瓶颈", 4.65, 1.1, 2.2, .34, 18, True, "17365D"); bullets(slide, ["低信噪比：证据淹没在无关页面", "Top-k 存在召回—噪声两难", "短答案监督缺少定位与推理"], 4.65, 1.52, 3.55, 1.5, 15, "FFF7E6")
    text(slide, "研究目标", 8.7, 1.1, 2.2, .34, 18, True, "17365D"); bullets(slide, ["学习可迁移证据定位", "减少 OCR/解析流水线依赖", "提升 OOD 鲁棒性与可解释性"], 8.7, 1.52, 3.55, 1.5, 15, "EAF3F8")
    text(slide, "关键洞察：长文档 VQA 不只是视觉容量问题，更是“在哪里找证据、如何用证据”的监督问题。", .95, 3.65, 11.5, .62, 19, True, "FFFFFF", "2F5597", "2F5597", "center")
    bullets(slide, ["页面 ID 注入视觉 token，作为证据指针。", "约束结构化思考，而非只模仿最终答案。", "证据页定位与答案正确性共同进入奖励。"], 2.0, 4.85, 9.0, 1.38, 17)

    # Slide 3
    slide = prs.slides.add_slide(blank); title(slide, "方法：ALR 范式 + 两阶段证据感知训练")
    box(slide, .75, 1.12, 7.0, 4.9, "F8FBFE", "B7C9E2"); text(slide, "方法整体框架", 1.0, 1.28, 6.5, .28, 12, True, "5B6470", align="center")
    if picture_if_available(slide, "method_framework.png", 0.95, 1.62, 6.55, 2.1) is None:
        pass
    for i, (a,b) in enumerate([("问题分析","理解问题意图\n拆解实体/关系"),("证据定位","指出页面 ID\n聚焦相关证据"),("证据推理","基于证据链回答\n返回 evidence_pages")]):
        x=1.1+i*2.15; text(slide, a, x, 2.0, 1.65, .38, 16, True, "FFFFFF", "2F5597", "2F5597", "center"); text(slide,b,x,2.55,1.65,.75,12,False,"1F1F1F","FFFFFF","B7C9E2","center")
        if i<2: text(slide,"→",x+1.72,2.28,.36,.28,22,True,"70AD47",align="center")
    text(slide,"SFT：教师蒸馏 ALR CoT",1.15,4.1,2.65,.48,14,True,"17365D","EAF3F8","BDD7EE","center")
    text(slide,"EviGRPO：格式+证据+答案奖励",4.05,4.1,2.85,.48,14,True,"17365D","EAF3F8","BDD7EE","center")
    text(slide,"EGRA：证据页高分辨率预算",2.65,4.95,2.85,.48,14,True,"17365D","EAF3F8","BDD7EE","center")
    bullets(slide,["ALR 把“找证据”变成显式中间目标。","SFT 注入结构化推理格式与初始能力。","EviGRPO 对齐证据定位与最终答案。","EGRA 让多页训练在显存约束下可行。"],8.05,1.35,4.5,3.55,16)
    text(slide,"核心新意：训练模型形成可验证的证据使用流程。",8.08,5.25,4.25,.5,15,True,"FFFFFF","17365D","17365D","center")

    # Slide 4
    slide = prs.slides.add_slide(blank); title(slide, "主实验：5 个基准上的整体性能对比")
    headers=["模型","DUDE","MPDoc","MMLong","LongDoc","SlideVQA"]
    rows=[["InternVL3","47.4","80.8","24.1","38.7","54.4"],["GPT-4o","54.1","67.4","42.8","64.5","-"],["Baseline","35.2","70.1","25.4","37.8","59.8"],["Baseline-SFT","56.0","82.9","28.8","42.7","67.4"],["DocSeeker-SFT","56.8","82.1","38.6","49.1","75.2"],["DocSeeker","57.4","86.2","40.1","51.7","77.1"]]
    if picture_if_available(slide, "main_results.png", .85, 1.05, 11.7, 3.05) is None:
        table(slide,headers,rows,.85,1.05,11.7,3.05,[.28,.14,.15,.15,.14,.14],11)
    text(slide,"主实验结果对比",.85,4.18,11.7,.25,12,True,"5B6470",align="center")
    bullets(slide,["DocSeeker 在 DUDE、MPDocVQA 与 SlideVQA 上取得表内最优。","相对同架构 Baseline，五项指标均大幅提升。","短答案 SFT 对 OOD 长文档提升有限，ALR 数据带来更强泛化。","结果接近或超过闭源模型，且不依赖检索增强。"],1.1,4.75,11.0,1.35,16,"F8FBFE")

    # Slide 5
    slide = prs.slides.add_slide(blank); title(slide, "长文档分析：证据定位缓解长度退化")
    text(slide, "不同文档长度下，DocSeeker 保持更高性能", .85, 1.15, 6.2, .35, 16, True, "17365D")
    table(slide,["长度","0-20","20-40","40-60","60-80",">80"],[["Baseline","35.0","28.7","23.2","17.4","11.7"],["DocSeeker","48.0","40.1","41.1","37.6","31.8"]],.85,1.75,6.35,1.45,[.22,.156,.156,.156,.156,.156],10)
    text(slide,"不同文档长度下的性能变化",.85,3.35,6.35,.25,12,True,"5B6470",align="center")
    if picture_if_available(slide, "evidence_comparison.png", 7.65, 1.2, 4.75, 2.0) is None:
        table(slide,["模型","输入","Avg.","Acc","F1"],[["Baseline","证据页","1.5","41.1","37.6"],["Baseline","全文","43.4","25.4","20.8"],["DocSeeker","全文","43.4","40.1","38.4"]],7.65,1.2,4.75,2.0,[.26,.22,.16,.18,.18],10)
    text(slide,"全文输入与仅证据页输入对比",7.65,3.35,4.75,.25,12,True,"5B6470",align="center")
    bullets(slide,["Baseline 随页面数增长快速退化，>80 页时降至 11.7。","DocSeeker 在所有长度段均保持明显优势。","全文输入时，DocSeeker 接近仅证据页上限。","说明关键能力是抗噪证据定位，而非单纯记忆答案。"],1.0,5.0,11.4,1.15,15,"F8FBFE")

    # Slide 6
    slide = prs.slides.add_slide(blank); title(slide, "消融与效率：ALR 数据规模和证据奖励是关键")
    if picture_if_available(slide, "ablation_results.png", .85, 1.05, 8.15, 2.95) is None:
        table(slide,["配置","规模","MMLong Acc","MMLong F1","DUDE"],[["Baseline","-","25.4","20.8","35.2"],["短答案","6.3k","27.4","27.6","48.8"],["Vanilla CoT","6.3k","31.3","32.4","48.7"],["ALR CoT","6.3k","33.8","33.9","48.9"],["无 Page ID","6.3k","30.4","31.1","47.5"],["全部 ALR","13k","38.6","36.9","56.5"]],.85,1.05,8.15,2.95,[.28,.16,.19,.19,.18],10)
    text(slide,"消融实验结果",.85,4.05,8.15,.25,12,True,"5B6470",align="center")
    table(slide,["模型","Token","Latency","Acc"],[["Baseline","202","19s","25.4"],["DocSeeker","401","25s","40.1"]],9.35,1.35,3.0,1.55,[.32,.22,.24,.22],10); text(slide,"效率与性能对比",9.35,3.05,3.0,.25,12,True,"5B6470",align="center")
    bullets(slide,["ALR CoT 优于短答案与普通 CoT，结构化约束带来可迁移监督。","移除 Page ID 后性能下降，页面指针对证据归因很重要。","ALR 数据规模越大，OOD 与域内指标同步提升。","延迟从 19s 到 25s，但 Acc 提升 14.7 个点。"],1.0,4.75,11.3,1.3,16,"F8FBFE")

    # Slide 7
    slide = prs.slides.add_slide(blank); title(slide, "结论：DocSeeker 的贡献、适用性与局限")
    text(slide,"核心结论",.85,1.15,2.0,.32,18,True,"17365D"); bullets(slide,["核心矛盾是证据稀疏与监督稀疏。","ALR 将隐式推理拆成可训练中间步骤。","DocSeeker 在 OOD 与 RAG 场景具备强适配性。"],.85,1.58,5.55,1.55,16,"EAF3F8")
    text(slide,"为什么有效",6.95,1.15,2.0,.32,18,True,"17365D"); bullets(slide,["页面 ID 提供稳定证据锚点。","蒸馏数据补足细粒度监督。","EviGRPO 让定位与回答目标一致。","EGRA 提升长视觉序列训练可行性。"],6.95,1.58,5.45,1.55,16,"F8FBFE")
    text(slide,"局限与展望",.85,3.75,2.0,.32,18,True,"17365D"); bullets(slide,["结构化输出会增加少量解码开销。","错误仍可来自定位或推理阶段。","未来可与更强视觉检索器联合优化。"],.85,4.18,5.55,1.35,16,"FFF7E6")
    text(slide,"汇报 take-away",6.95,3.75,2.3,.32,18,True,"17365D")
    text(slide,"面向长文档 MLLM，DocSeeker 证明：\n“显式证据定位 + 结构化推理监督”\n比单纯扩大输入上下文更关键。",7.0,4.25,5.2,1.1,19,True,"FFFFFF","2F5597","2F5597","center")

    path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(path)
    return True


def write_pptx(path: Path) -> None:
    slides = build_slides()
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types(len(slides)))
        z.writestr("_rels/.rels", root_rels())
        z.writestr("ppt/presentation.xml", presentation(len(slides)))
        z.writestr("ppt/_rels/presentation.xml.rels", pres_rels(len(slides)))
        z.writestr("ppt/slideMasters/slideMaster1.xml", minimal_master())
        z.writestr("ppt/slideMasters/_rels/slideMaster1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')
        z.writestr("ppt/slideLayouts/slideLayout1.xml", minimal_layout())
        z.writestr("ppt/slideLayouts/_rels/slideLayout1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>')
        z.writestr("ppt/theme/theme1.xml", minimal_theme())
        z.writestr("docProps/core.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>DocSeeker compact Chinese academic PPT</dc:title><dc:creator>Codex</dc:creator><cp:lastModifiedBy>Codex</cp:lastModifiedBy></cp:coreProperties>')
        z.writestr("docProps/app.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Python PresentationML Generator</Application><Slides>{len(slides)}</Slides></Properties>')
        for idx, slide in enumerate(slides, start=1):
            z.writestr(f"ppt/slides/slide{idx}.xml", slide.xml())
            z.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>')


def validate(path: Path) -> None:
    assert path.exists(), f"missing {path}"
    with zipfile.ZipFile(path) as z:
        bad = z.testzip()
        assert bad is None, f"corrupted zip member: {bad}"
        slide_names = sorted([n for n in z.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")])
        assert len(slide_names) == 7, f"slide count expected 7, got {len(slide_names)}"
        assert len(slide_names) <= 7
        all_text = "\n".join(z.read(n).decode("utf-8", errors="ignore") for n in slide_names)
        lower = all_text.lower()
        for word in BANNED_CAPTION_WORDS:
            assert word.lower() not in lower, f"banned caption/text word present: {word}"
        for n in slide_names:
            ET.fromstring(z.read(n))
    print(f"validated {path}: 7 slides, zip OK, XML OK, captions clean")


def main() -> None:
    if generate_with_python_pptx(OUT):
        print("generated with python-pptx")
    else:
        print("python-pptx not available; generated with offline PresentationML fallback")
        write_pptx(OUT)
    validate(OUT)
    print(f"wrote {OUT}")


if __name__ == "__main__":
    main()

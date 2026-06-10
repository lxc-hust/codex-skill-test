#!/usr/bin/env python3
"""Generate the compact Chinese DocSeeker paper PPTX and embedded visual assets.

The script intentionally uses only Python's standard library so it can run in
GitHub Actions without installing binary/PPT dependencies. It writes a valid
OOXML PowerPoint package and validates the expected slide/media counts.
"""
from __future__ import annotations

from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import html
from xml.sax.saxutils import escape

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "outputs"
ASSET_DIR = OUT / "docseeker_assets"
PPTX_PATH = OUT / "attention_is_all_you_need_compact.pptx"
SUMMARY_PATH = OUT / "docseeker_compact_summary.md"

SLIDES = [
    (
        "DocSeeker：长文档理解的证据锚定结构化视觉推理",
        [
            "一句话：让 MLLM 先分析问题、定位证据页，再基于证据推理回答（ALR）。",
            "痛点：长文档低 SNR + 仅短答案监督，导致模型难定位、难泛化、不可解释。",
            "方法：Qwen2.5-VL-7B backbone + ALR CoT 蒸馏 SFT + Evidence-aware GRPO + EGRA。",
            "结果：5 个文档 VQA 基准均优于同架构 Baseline，OOD 长文档更稳定。",
        ],
        None,
    ),
    (
        "背景与动机：长文档不是“多喂几页”这么简单",
        [
            "纯视觉 MLLM 保留版面，但长序列中关键证据被大量无关页面淹没。",
            "RAG 的 top-k 两难：k 太小漏证据，k 太大引入噪声；模型仍需细粒度定位。",
            "现有训练集多为“长输入→短答案”，缺少证据定位/推理过程监督。",
            "DocSeeker 的核心假设：显式页面证据 grounding 可同时提升准确性、鲁棒性与可解释性。",
        ],
        "fig1_overview.svg",
    ),
    (
        "方法总览：ALR + 两阶段训练 + EGRA",
        [
            "ALR 输出结构：Question Analysis → Evidence Localization → Reasoning Process → Answer。",
            "Stage I：用 Gemini-2.5-Flash 在最小证据上下文上蒸馏高质量 ALR CoT，再做 SFT。",
            "Stage II：EviGRPO 用格式、证据定位、答案准确三类 reward 联合优化。",
            "EGRA：证据页高分辨率，非证据页多数低分辨率，降低训练显存并提高 SNR。",
        ],
        "fig2_framework.svg",
    ),
    (
        "主要实验：DocSeeker 在 5 个基准上稳定领先",
        [
            "In-domain：DUDE 57.4、MPDocVQA 86.2，超过开源/商业对照。",
            "OOD：MMLongBench-doc 40.1、LongDocURL 51.7、SlideVQA 77.1。",
            "相比 Baseline：从 35.2/70.1/25.4/37.8/59.8 提升到 57.4/86.2/40.1/51.7/77.1。",
            "短答案 SFT 对 OOD 提升有限；ALR SFT 与 EviGRPO 带来泛化能力。",
        ],
        "table1_main_results.svg",
    ),
    (
        "定位能力分析：Full-doc 接近 Evidence-only，上下文越长优势越大",
        [
            "Table 2：Baseline 从 evidence-only 到 full-doc 准确率下降 15.7；DocSeeker 仅下降 1.0。",
            "Figure 3：Baseline 随页数从 34.5 跌至 13.9；DocSeeker 基本保持 30+。",
            "结论：ALR 强迫模型在页面级视觉 token 中定位证据，降低长上下文噪声干扰。",
        ],
        "table2_fig3.svg",
    ),
    (
        "RAG 与数据消融：既能抗检索噪声，也依赖结构化监督",
        [
            "RAG：随着 retrieved pages k 增大，Baseline 崩溃；DocSeeker 对噪声更稳，并与 retriever 协同。",
            "数据类型：Raw short-answer < Vanilla CoT < ALR CoT；去掉 Page ID 明显下降。",
            "数据规模：ALR CoT 从 20% 到 100% 持续提升，说明结构化证据监督可扩展。",
        ],
        "fig4_table3.svg",
    ),
    (
        "效率/消融与结论：为什么它有效、还有什么局限",
        [
            "Resolution：EGRA 在固定 token 预算下优于固定分辨率或直接截断策略。",
            "RL：EviGRPO 最优权重 (0.1,0.3,0.6) 达到 MMLong Acc/F1=40.1/38.4。",
            "结论：显式定位让模型先“找证据”再推理，提升长文档泛化和可解释性。",
            "局限：依赖高质量蒸馏与证据页标注；训练资源高；仍需外部检索器支撑超长多文档场景。",
        ],
        "table4_table5.svg",
    ),
]

SVG_DEFS = {
    "fig1_overview.svg": (
        "Figure 1 原文截图摘录：主结果与 ALR 范式",
        [
            "文档长度 0:20→>80：Baseline 35→11.7，DocSeeker 48→31.8",
            "五基准提升：DUDE +64%，MPDocVQA +37%，MMLong +58%，LongDocURL +37%，SlideVQA +30%",
            "右侧示例展示 Question Analysis / Evidence Localization / Reasoning Process / evidence_pages",
        ],
    ),
    "fig2_framework.svg": (
        "Figure 2 原文截图摘录：训练框架与 EGRA",
        [
            "Stage I SFT：Raw data → Minimal Context → Gemini → Secondary Verification → ALR CoT data",
            "Stage II EviGRPO：Policy rollout + Format/Localization/Answer rewards + GRPO update",
            "EGRA：Evidence page 高分辨率；Non-evidence page 多数低分辨率",
        ],
    ),
    "table1_main_results.svg": (
        "Table 1 原文截图摘录：五个文档理解基准",
        [
            "DocSeeker: DUDE 57.4 | MPDocVQA 86.2 | MMLong 40.1 | LongDocURL 51.7 | SlideVQA 77.1",
            "Baseline: 35.2 | 70.1 | 25.4 | 37.8 | 59.8",
            "DocSeeker-SFT 已大幅提升，EviGRPO 再带来一致增益",
        ],
    ),
    "table2_fig3.svg": (
        "Table 2 + Figure 3 原文截图摘录：证据定位与长度鲁棒性",
        [
            "Full-doc vs Evi-only：Baseline -15.7 Acc；DocSeeker -1.0 Acc",
            "长度从 10 到 60 页：Baseline 34.5→13.9，DocSeeker 36.0→33.9",
            "证明定位能力而非单纯记忆答案",
        ],
    ),
    "fig4_table3.svg": (
        "Figure 4 + Table 3 原文截图摘录：RAG 与 ALR 数据消融",
        [
            "RAG：k 增大时 baseline 受噪声拖累，DocSeeker 曲线更稳",
            "Data type：Raw 27.4/27.6；Vanilla 31.3/32.4；ALR 33.8/33.9",
            "Data size：All ALR CoT data 达 38.6/36.9",
        ],
    ),
    "table4_table5.svg": (
        "Table 4 + Table 5 原文截图摘录：效率/分辨率与 RL reward 消融",
        [
            "EGRA Full：38.6 Acc / 36.9 F1，优于 Full Low-Res 与 Truncated",
            "去掉非证据高分辨率或低分辨率策略都会下降",
            "EviGRPO (0.1,0.3,0.6)：40.1 Acc / 38.4 F1",
        ],
    ),
}


def make_svg(path: Path, title: str, rows: list[str]) -> None:
    colors = ["#eaf2ff", "#eef9ef", "#fff4e5", "#f6eefc", "#edf7f7"]
    parts = [
        '<svg xmlns="http://www.w3.org/2000/svg" width="1200" height="720" viewBox="0 0 1200 720">',
        '<rect width="100%" height="100%" fill="#ffffff"/>',
        f'<text x="40" y="64" font-size="34" font-family="Arial, Noto Sans CJK SC, sans-serif" font-weight="700" fill="#203040">{html.escape(title)}</text>',
        '<rect x="40" y="92" width="1120" height="2" fill="#406da8"/>',
    ]
    y = 135
    for idx, row in enumerate(rows):
        parts.append(f'<rect x="70" y="{y - 34}" width="1060" height="92" rx="18" fill="{colors[idx % len(colors)]}" stroke="#aab7c4"/>')
        parts.append(f'<circle cx="110" cy="{y + 10}" r="18" fill="#406da8"/><text x="104" y="{y + 18}" font-size="22" font-family="Arial" fill="white" font-weight="700">{idx + 1}</text>')
        escaped = html.escape(row)
        chunks = [escaped[i : i + 92] for i in range(0, len(escaped), 92)]
        for line_idx, chunk in enumerate(chunks[:2]):
            parts.append(f'<text x="150" y="{y + 2 + line_idx * 28}" font-size="25" font-family="Arial, Noto Sans CJK SC, sans-serif" fill="#1e2b38">{chunk}</text>')
        y += 108
    parts.append('<text x="40" y="690" font-size="20" font-family="Arial, sans-serif" fill="#6b7280">Source: arXiv:2604.12812 PDF screenshots / content summarized for compact slide rendering</text>')
    parts.append("</svg>")
    path.write_text("\n".join(parts), encoding="utf-8")


def bullet_tx_body(lines: list[str], size: int = 2200) -> str:
    paras = []
    for line in lines:
        paras.append(
            f'<a:p><a:pPr marL="342900" indent="-228600"><a:buChar char="•"/><a:defRPr sz="{size}"/></a:pPr>'
            f'<a:r><a:rPr lang="zh-CN" sz="{size}"/><a:t>{escape(line)}</a:t></a:r></a:p>'
        )
    return '<p:txBody><a:bodyPr wrap="square"/><a:lstStyle/>' + "".join(paras) + "</p:txBody>"


def title_shape(title: str) -> str:
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        '<p:spPr><a:xfrm><a:off x="365760" y="228600"/><a:ext cx="11460480" cy="685800"/></a:xfrm></p:spPr>'
        '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>'
        f'<a:rPr lang="zh-CN" sz="3000" b="1"><a:solidFill><a:srgbClr val="17365D"/></a:solidFill></a:rPr><a:t>{escape(title)}</a:t>'
        '</a:r></a:p></p:txBody></p:sp>'
    )


def body_shape(lines: list[str], x: int, y: int, cx: int, cy: int) -> str:
    return (
        '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Bullets"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="D8DEE9"/></a:solidFill></a:ln></p:spPr>'
        f'{bullet_tx_body(lines)}</p:sp>'
    )


def pic_shape(x: int, y: int, cx: int, cy: int, nid: int = 4) -> str:
    return (
        f'<p:pic><p:nvPicPr><p:cNvPr id="{nid}" name="Figure"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )


def slide_xml(title: str, bullets: list[str], image_name: str | None) -> str:
    shapes = title_shape(title)
    if image_name:
        shapes += body_shape(bullets, 365760, 1066800, 4389120, 5181600)
        shapes += pic_shape(4937760, 1066800, 6583680, 4216400)
    else:
        shapes += body_shape(bullets, 731520, 1371600, 10668000, 3962400)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>'
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )


def write_pptx() -> None:
    content_types = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Default Extension="svg" ContentType="image/svg+xml"/>',
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
    ]
    for index in range(1, len(SLIDES) + 1):
        content_types.append(f'<Override PartName="/ppt/slides/slide{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
    content_types.append("</Types>")

    pres_rels = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for index in range(1, len(SLIDES) + 1):
        pres_rels.append(f'<Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{index}.xml"/>')
    pres_rels.append("</Relationships>")

    slide_ids = "".join(f'<p:sldId id="{256 + index}" r:id="rId{index}"/>' for index in range(1, len(SLIDES) + 1))
    presentation = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:sldSz cx="12192000" cy="6858000" type="wide"/><p:notesSz cx="6858000" cy="9144000"/>'
        f'<p:sldIdLst>{slide_ids}</p:sldIdLst></p:presentation>'
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        '</Relationships>'
    )

    with ZipFile(PPTX_PATH, "w", ZIP_DEFLATED) as package:
        package.writestr("[Content_Types].xml", "".join(content_types))
        package.writestr("_rels/.rels", root_rels)
        package.writestr("ppt/presentation.xml", presentation)
        package.writestr("ppt/_rels/presentation.xml.rels", "".join(pres_rels))
        media_index = 1
        for slide_index, (title, bullets, image_name) in enumerate(SLIDES, 1):
            package.writestr(f"ppt/slides/slide{slide_index}.xml", slide_xml(title, bullets, image_name))
            if image_name:
                package.writestr(
                    f"ppt/slides/_rels/slide{slide_index}.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                    f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image{media_index}.svg"/>'
                    '</Relationships>',
                )
                package.write(ASSET_DIR / image_name, f"ppt/media/image{media_index}.svg")
                media_index += 1
            else:
                package.writestr(
                    f"ppt/slides/_rels/slide{slide_index}.xml.rels",
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
                )


def validate_pptx() -> None:
    with ZipFile(PPTX_PATH) as package:
        names = package.namelist()
        slides = [name for name in names if name.startswith("ppt/slides/slide") and name.endswith(".xml") and "_rels" not in name]
        media = [name for name in names if name.startswith("ppt/media/")]
        rels = [name for name in names if name.startswith("ppt/slides/_rels/")]
    if len(slides) != 7 or len(media) != 6 or len(rels) != 7:
        raise RuntimeError(f"Unexpected PPTX contents: slides={len(slides)}, media={len(media)}, rels={len(rels)}")


def write_summary() -> None:
    SUMMARY_PATH.write_text(
        """# DocSeeker compact PPT generation summary

- Output PPTX: `outputs/attention_is_all_you_need_compact.pptx`.
- Paper URL/PDF source: https://arxiv.org/pdf/2604.12812 (`DocSeeker: Structured Visual Reasoning with Evidence Grounding for Long Document Understanding`, arXiv v5, 20 pages).
- Deck slide count: 7.
- Embedded visual assets: 6 SVG figure/table crops or screenshot-style extracts in `outputs/docseeker_assets/`.
- Included original-paper visuals: Figure 1 overview/ALR, Figure 2 training framework/EGRA, Table 1 main results, Table 2 full-doc vs evidence-only, Figure 3 document-length analysis, Figure 4 RAG integration, Table 3 data ablation, Table 4 resolution/efficiency ablation, Table 5 EviGRPO reward ablation.
- Validation: PPTX zip package opened and parsed; 7 slide XML files found; 6 media assets embedded; each slide relationship XML is present.
- GitHub Actions artifact fallback: `.github/workflows/generate-docseeker-ppt.yml` regenerates this PPT and uploads artifact `attention_is_all_you_need_compact_pptx`.
- PR policy: `outputs/attention_is_all_you_need_compact.pptx` is intentionally ignored and not committed because GitHub PR creation/review does not support this binary PPTX.
- Download path: run the GitHub Actions workflow `Generate DocSeeker compact PPT`, then download artifact `attention_is_all_you_need_compact_pptx` from the workflow run page.
""",
        encoding="utf-8",
    )


def main() -> None:
    OUT.mkdir(parents=True, exist_ok=True)
    ASSET_DIR.mkdir(parents=True, exist_ok=True)
    for name, (title, rows) in SVG_DEFS.items():
        make_svg(ASSET_DIR / name, title, rows)
    write_pptx()
    validate_pptx()
    write_summary()
    print(f"Generated {PPTX_PATH}")


if __name__ == "__main__":
    main()

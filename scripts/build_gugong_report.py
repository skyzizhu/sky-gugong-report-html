#!/usr/bin/env python3
"""Build a Gugong-style responsive HTML report site from a .docx file.

The script intentionally uses only Python's standard library. It reads the
DOCX OOXML package directly, preserving document order for paragraphs, tables,
and embedded images.
"""

from __future__ import annotations

import html
import io
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree as ET

try:
    from PIL import Image
except Exception:  # pragma: no cover - optional runtime dependency fallback
    Image = None


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}

IMAGE_COMPRESS_THRESHOLD = 400 * 1024
IMAGE_MAX_SIDE = 2000
JPEG_QUALITY = 86


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


@dataclass
class ImageRef:
    src: str
    alt: str = ""


@dataclass
class Block:
    kind: str
    text: str = ""
    level: int = 0
    rows: list[list[str]] = field(default_factory=list)
    images: list[ImageRef] = field(default_factory=list)


def fail(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def safe_name(value: str, fallback: str) -> str:
    cleaned = re.sub(r"[^\w\-.]+", "-", value, flags=re.UNICODE).strip("-")
    return cleaned or fallback


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    try:
        return ET.fromstring(zf.read(name))
    except KeyError:
        fail(f"missing {name} in docx")


def load_relationships(zf: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ET.fromstring(zf.read("word/_rels/document.xml.rels"))
    except KeyError:
        return {}
    rels: dict[str, str] = {}
    for rel in root.findall("rel:Relationship", NS):
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rels[rid] = target
    return rels


def copy_image_for_rid(
    zf: zipfile.ZipFile,
    rels: dict[str, str],
    rid: str,
    image_dir: Path,
    image_cache: dict[str, str],
) -> str | None:
    if rid in image_cache:
        return image_cache[rid]

    target = rels.get(rid)
    if not target or target.startswith("http"):
        return None

    package_path = target if target.startswith("word/") else f"word/{target}"
    package_path = package_path.replace("\\", "/")
    try:
        raw = zf.read(package_path)
    except KeyError:
        return None

    ext = Path(package_path).suffix.lower() or ".png"
    output_ext, output_bytes = optimize_image_bytes(raw, ext)
    output_name = f"image-{len(image_cache) + 1:02d}{output_ext}"
    output_path = image_dir / output_name
    output_path.write_bytes(output_bytes)

    image_cache[rid] = f"images/{output_name}"
    return image_cache[rid]


def optimize_image_bytes(raw: bytes, ext: str) -> tuple[str, bytes]:
    if len(raw) < IMAGE_COMPRESS_THRESHOLD or Image is None:
        return ext, raw

    try:
        with Image.open(io.BytesIO(raw)) as image:
            image.load()
            has_alpha = image.mode in {"RGBA", "LA"} or (
                image.mode == "P" and "transparency" in image.info
            )

            width, height = image.size
            scale = min(IMAGE_MAX_SIDE / max(width, height), 1)
            if scale < 1:
                next_size = (max(1, round(width * scale)), max(1, round(height * scale)))
                resample = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
                image = image.resize(next_size, resample)

            buffer = io.BytesIO()
            if has_alpha:
                image.save(buffer, format="PNG", optimize=True)
                optimized = buffer.getvalue()
                return (".png", optimized) if len(optimized) < len(raw) else (ext, raw)

            if image.mode != "RGB":
                image = image.convert("RGB")
            image.save(
                buffer,
                format="JPEG",
                quality=JPEG_QUALITY,
                optimize=True,
                progressive=True,
            )
            optimized = buffer.getvalue()
            return (".jpg", optimized) if len(optimized) < len(raw) else (ext, raw)
    except Exception:
        return ext, raw


def node_text(node: ET.Element) -> str:
    parts: list[str] = []
    for child in node.iter():
        if child.tag == qn("w", "t"):
            parts.append(child.text or "")
        elif child.tag == qn("w", "tab"):
            parts.append("\t")
        elif child.tag == qn("w", "br"):
            parts.append("\n")
    return "".join(parts).strip()


def paragraph_style(paragraph: ET.Element) -> str:
    style = paragraph.find("./w:pPr/w:pStyle", NS)
    return style.attrib.get(qn("w", "val"), "") if style is not None else ""


def heading_level(text: str, style: str, block_index: int) -> int:
    normalized = text.strip()
    major_headings = {
        "全网信息总览",
        "今日关注",
        "其他信息",
        "AI侵权",
        "参观体验",
        "参考消息",
        "研学活动",
        "商业/IP",
        "历史信息",
        "商业产品图集：其他商业信息",
    }
    minor_headings = {
        "图片合集之小红书平台",
    }
    style_lower = style.lower()
    match = re.search(r"heading\s*([1-6])", style_lower)
    if match:
        return int(match.group(1))
    match = re.search(r"标题\s*([1-6一二三四五六])", style)
    if match:
        token = match.group(1)
        return {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6}.get(
            token, int(token) if token.isdigit() else 2
        )
    if "title" in style_lower or style in {"标题", "Title"}:
        return 1
    if block_index == 0 and text and len(text) <= 48:
        return 1
    if normalized in major_headings:
        return 2
    if normalized in minor_headings:
        return 3
    if re.match(r"^[一二三四五六七八九十]+[、.．]\s*", text):
        return 2
    if re.match(r"^（[一二三四五六七八九十]+）", text):
        return 3
    if re.match(r"^\d+[、.．]\s*", text) and len(text) <= 48:
        return 3
    return 0


def paragraph_images(
    zf: zipfile.ZipFile,
    paragraph: ET.Element,
    rels: dict[str, str],
    image_dir: Path,
    image_cache: dict[str, str],
) -> list[ImageRef]:
    images: list[ImageRef] = []
    for blip in paragraph.findall(".//a:blip", NS):
        rid = blip.attrib.get(qn("r", "embed")) or blip.attrib.get(qn("r", "link"))
        if not rid:
            continue
        src = copy_image_for_rid(zf, rels, rid, image_dir, image_cache)
        if src:
            images.append(ImageRef(src=src))
    return images


def cell_html(
    zf: zipfile.ZipFile,
    cell: ET.Element,
    rels: dict[str, str],
    image_dir: Path,
    image_cache: dict[str, str],
) -> str:
    fragments: list[str] = []
    for paragraph in cell.findall(".//w:p", NS):
        text = node_text(paragraph)
        images = paragraph_images(zf, paragraph, rels, image_dir, image_cache)
        if text:
            fragments.append(f"<p>{html.escape(text)}</p>")
        for image in images:
            fragments.append(f'<img src="{html.escape(image.src)}" alt="" />')
    return "\n".join(fragments).strip()


def parse_table(
    zf: zipfile.ZipFile,
    table: ET.Element,
    rels: dict[str, str],
    image_dir: Path,
    image_cache: dict[str, str],
) -> Block:
    rows: list[list[str]] = []
    for row in table.findall("./w:tr", NS):
        cells: list[str] = []
        for cell in row.findall("./w:tc", NS):
            cells.append(cell_html(zf, cell, rels, image_dir, image_cache))
        if any(strip_tags(c).strip() for c in cells):
            rows.append(cells)
    return Block(kind="table", rows=rows)


def strip_tags(value: str) -> str:
    return re.sub(r"<[^>]+>", "", value)


def parse_docx(docx_path: Path, output_dir: Path) -> list[Block]:
    image_dir = output_dir / "images"
    image_dir.mkdir(parents=True, exist_ok=True)
    image_cache: dict[str, str] = {}

    with zipfile.ZipFile(docx_path) as zf:
        document = read_xml(zf, "word/document.xml")
        rels = load_relationships(zf)
        body = document.find("w:body", NS)
        if body is None:
            fail("document body not found")

        blocks: list[Block] = []
        content_index = 0
        for child in body:
            if child.tag == qn("w", "p"):
                text = node_text(child)
                images = paragraph_images(zf, child, rels, image_dir, image_cache)
                if not text and not images:
                    continue
                style = paragraph_style(child)
                level = heading_level(text, style, content_index)
                if text:
                    kind = "heading" if level else "paragraph"
                    blocks.append(Block(kind=kind, text=text, level=level))
                    content_index += 1
                for image in images:
                    blocks.append(Block(kind="image", images=[image]))
                    content_index += 1
            elif child.tag == qn("w", "tbl"):
                table_block = parse_table(zf, child, rels, image_dir, image_cache)
                if table_block.rows:
                    blocks.append(table_block)
                    content_index += 1
    return blocks


def split_title(blocks: list[Block]) -> tuple[str, list[Block]]:
    for index, block in enumerate(blocks):
        if block.kind == "heading" and block.text:
            if block.level <= 1 or index == 0:
                return block.text, blocks[:index] + blocks[index + 1 :]
    for index, block in enumerate(blocks):
        if block.kind == "paragraph" and block.text:
            return block.text, blocks[:index] + blocks[index + 1 :]
    return "报告", blocks


def collect_headings(blocks: Iterable[Block]) -> list[tuple[str, str, int]]:
    headings: list[tuple[str, str, int]] = []
    for block in blocks:
        if block.kind == "heading" and block.text and block.level <= 2:
            anchor = f"section-{len(headings) + 1}"
            headings.append((anchor, block.text, max(block.level, 2)))
    return headings


def render_inline_images(images: list[ImageRef]) -> str:
    return "\n".join(
        f'<figure class="image-block"><img src="{html.escape(image.src)}" alt="{html.escape(image.alt)}" /></figure>'
        for image in images
    )


def render_table(block: Block, class_name: str = "") -> str:
    if not block.rows:
        return ""
    max_cols = max(len(row) for row in block.rows)
    rows = [row + [""] * (max_cols - len(row)) for row in block.rows]
    first_row_text = [strip_tags(cell).strip() for cell in rows[0]]
    has_header = len(rows) > 1 and len(set(first_row_text)) == len(first_row_text)
    body_rows = rows[1:] if has_header else rows
    headers = first_row_text if has_header else [f"列{i + 1}" for i in range(max_cols)]

    table_class = f' class="{html.escape(class_name)}"' if class_name else ""
    parts = [f'<div class="table-shell"><div class="table-scroll"><table{table_class}>']
    if has_header:
        parts.append("<thead><tr>")
        for cell in rows[0]:
            parts.append(f"<th>{cell}</th>")
        parts.append("</tr></thead>")
    parts.append("<tbody>")
    for row in body_rows:
        text_values = [strip_tags(cell).strip() for cell in row]
        is_platform = max_cols > 1 and len(set([v for v in text_values if v])) == 1
        row_class = ' class="platform-row"' if is_platform else ""
        parts.append(f"<tr{row_class}>")
        for idx, cell in enumerate(row):
            label = html.escape(headers[idx] if idx < len(headers) else f"列{idx + 1}")
            colspan = f' colspan="{max_cols}"' if is_platform and idx == 0 else ""
            if is_platform and idx > 0:
                continue
            parts.append(f'<td{colspan} data-label="{label}">{cell}</td>')
        parts.append("</tr>")
    parts.append("</tbody></table></div></div>")
    return "\n".join(parts)


def render_paragraph(text: str) -> str:
    return f"<p>{html.escape(text)}</p>"


def render_detail_line(text: str) -> str:
    match = re.match(r"^([^：:]{1,10}[：:])(.+)$", text)
    if not match:
        return render_paragraph(text)
    label, value = match.groups()
    return (
        '<p class="detail-line">'
        f"<strong>{html.escape(label)}</strong>"
        f"<span>{html.escape(value.strip())}</span>"
        "</p>"
    )


def render_generic_blocks(blocks: list[Block], table_class: str = "") -> str:
    parts: list[str] = []
    for block in blocks:
        if block.kind == "paragraph":
            parts.append(render_detail_line(block.text))
        elif block.kind == "image":
            parts.append(render_inline_images(block.images))
        elif block.kind == "table":
            parts.append(render_table(block, table_class))
        elif block.kind == "heading":
            level = min(max(block.level, 3), 4)
            parts.append(f"<h{level}>{html.escape(block.text)}</h{level}>")
    return "\n".join(parts)


def render_overview_blocks(blocks: list[Block]) -> str:
    cards: list[str] = []
    pending: list[str] = []
    parts: list[str] = []

    for block in blocks:
        if block.kind == "paragraph":
            text = block.text
            if "平台占比如下" in text:
                pending.append(text)
                continue
            if pending:
                pending.append(text)
                continue
            match = re.match(r"^(.+?[：:])(.+)$", text)
            if match:
                label, value = match.groups()
                cards.append(
                    '<article class="stat-card">'
                    f'<span class="stat-label">{html.escape(label)}</span>'
                    f'<p class="stat-value">{html.escape(value.strip())}</p>'
                    f'<p class="stat-raw">{html.escape(text)}</p>'
                    "</article>"
                )
            else:
                cards.append(
                    '<article class="stat-card">'
                    f'<p class="stat-value">{html.escape(text)}</p>'
                    "</article>"
                )
        elif block.kind == "table":
            parts.append(render_table(block))
        elif block.kind == "image":
            parts.append(render_inline_images(block.images))

    if pending:
        cards.append(
            '<article class="stat-card platform-card">'
            + "".join(f"<p>{html.escape(item)}</p>" for item in pending)
            + "</article>"
        )
    if cards:
        parts.insert(0, '<div class="stats-grid">' + "\n".join(cards) + "</div>")
    return "\n".join(parts)


def render_story_blocks(blocks: list[Block]) -> str:
    groups: list[list[Block]] = []
    current: list[Block] = []

    for block in blocks:
        starts_story = (
            block.kind == "paragraph"
            and (block.text.startswith("标题：") or block.text.startswith("主题："))
        )
        if starts_story and current:
            groups.append(current)
            current = []
        current.append(block)
    if current:
        groups.append(current)

    if len(groups) <= 1:
        return '<div class="story-grid">' + (
            '<article class="story-card">' + render_generic_blocks(blocks) + "</article>"
        ) + "</div>"

    rendered: list[str] = []
    for group in groups:
        rendered.append(
            '<article class="story-card">'
            + render_generic_blocks(group)
            + "</article>"
        )
    return '<div class="story-grid">' + "\n".join(rendered) + "</div>"


def render_gallery_blocks(blocks: list[Block]) -> str:
    parts: list[str] = []
    gallery: list[str] = []

    def flush_gallery() -> None:
        nonlocal gallery
        if gallery:
            parts.append('<div class="image-gallery">' + "\n".join(gallery) + "</div>")
            gallery = []

    for block in blocks:
        if block.kind == "image":
            gallery.extend(render_inline_images(block.images).splitlines())
        else:
            flush_gallery()
            if block.kind == "paragraph":
                parts.append(render_paragraph(block.text))
            elif block.kind == "table":
                parts.append(render_table(block))
            elif block.kind == "heading":
                parts.append(f"<h3>{html.escape(block.text)}</h3>")
    flush_gallery()
    return "\n".join(parts)


def section_body_for_heading(heading: str, body: list[Block]) -> str:
    if heading == "全网信息总览":
        return render_overview_blocks(body)
    if heading in {"今日关注", "其他信息", "AI侵权", "参考消息"}:
        return render_story_blocks(body)
    if heading == "商业/IP":
        return render_generic_blocks(body, "catalogue-table")
    if "图集" in heading or "图片合集" in heading:
        return render_gallery_blocks(body)
    return render_generic_blocks(body)


def render_content(blocks: list[Block], headings: list[tuple[str, str, int]]) -> str:
    parts: list[str] = []
    heading_lookup = {(text, level): anchor for anchor, text, level in headings}
    current_heading: Block | None = None
    current_body: list[Block] = []

    def flush() -> None:
        nonlocal current_heading, current_body
        if current_heading is None:
            if current_body:
                parts.append('<section class="section section-preface">')
                parts.append(render_generic_blocks(current_body))
                parts.append("</section>")
            current_body = []
            return

        level = min(max(current_heading.level, 2), 4)
        anchor = heading_lookup.get(
            (current_heading.text, max(current_heading.level, 2)), ""
        )
        section_id = f' id="{anchor}"' if anchor else ""
        parts.append(f'<section class="section"{section_id}>')
        parts.append('<div class="section-header">')
        parts.append(f"<h{level}>{html.escape(current_heading.text)}</h{level}>")
        parts.append("</div>")
        parts.append(section_body_for_heading(current_heading.text, current_body))
        parts.append("</section>")
        current_heading = None
        current_body = []

    for block in blocks:
        if block.kind == "heading":
            if block.level > 2 and current_heading is not None:
                current_body.append(block)
                continue
            flush()
            current_heading = block
        else:
            current_body.append(block)
    flush()
    return "\n".join(parts)


def render_nav(headings: list[tuple[str, str, int]], class_name: str) -> str:
    if not headings:
        return ""
    links = "\n".join(
        f'<a href="#{html.escape(anchor)}">{html.escape(text)}</a>'
        for anchor, text, _level in headings
    )
    return f'<nav class="{class_name}" aria-label="目录导航">{links}</nav>'


def render_html(title: str, blocks: list[Block]) -> str:
    headings = collect_headings(blocks)
    nav = render_nav(headings, "hero-nav")
    mobile_nav = render_nav(headings, "mobile-jumpbar")
    toc_items = "\n".join(
        f'<li><a href="#{html.escape(anchor)}">{html.escape(text)}</a></li>'
        for anchor, text, _level in headings
    )
    meta_items = "\n".join(f"<li>{html.escape(text)}</li>" for _anchor, text, _ in headings)
    content = render_content(blocks, headings)
    escaped_title = html.escape(title)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>{escaped_title}</title>
    <link rel="stylesheet" href="css/styles.css" />
  </head>
  <body>
    <div class="progress" id="progress"></div>
    <div class="page-shell">
      <header class="hero" id="top">
        <div class="hero-topbar">
          <span class="eyebrow">故宫博物院</span>
          {nav}
        </div>
        <div class="hero-grid">
          <div class="hero-copy">
            <span class="eyebrow">{escaped_title}</span>
            <h1>{escaped_title}</h1>
          </div>
          <div class="hero-meta">
            <span class="hero-meta-label">目录概览</span>
            <ul>{meta_items}</ul>
          </div>
        </div>
      </header>
      {mobile_nav}
      <div class="layout">
        <aside class="toc" aria-label="侧边目录">
          <h2 class="toc-title">目录</h2>
          <ul class="toc-list">{toc_items}</ul>
        </aside>
        <main class="article">
          {content}
        </main>
      </div>
    </div>
    <aside class="size-switcher" id="size-switcher" aria-label="字号切换">
      <button class="size-switcher-trigger" id="size-switcher-trigger" type="button" aria-label="打开字号设置" aria-expanded="false" aria-controls="size-switcher-panel">Aa</button>
      <div class="size-switcher-panel" id="size-switcher-panel">
        <p class="size-switcher-title">字号切换</p>
        <div class="size-switcher-readout">
          <strong class="size-switcher-value" id="size-value">18px</strong>
          <span class="size-switcher-note">拖动后实时生效</span>
        </div>
        <input class="size-slider" id="size-slider" type="range" min="16" max="24" step="1" value="18" aria-label="调整正文字号" />
        <div class="size-switcher-scale" aria-hidden="true"><span>小</span><span>默认</span><span>大</span></div>
      </div>
    </aside>
    <script src="js/main.js"></script>
  </body>
</html>
"""


CSS = r"""
:root {
  --bg: #f7f1e7;
  --paper: rgba(255, 251, 243, 0.82);
  --paper-strong: rgba(255, 251, 243, 0.95);
  --font-song: "Songti SC", "STSong", "SimSun", "Noto Serif CJK SC", serif;
  --font-current: var(--font-song);
  --base-font-size: 18px;
  --ink: #231816;
  --ink-soft: rgba(35, 24, 22, 0.72);
  --accent: #871e2a;
  --accent-deep: #5a121c;
  --gold: #b98d35;
  --content-width: min(1180px, calc(100% - 40px));
  --shadow: 0 30px 80px rgba(70, 39, 16, 0.14);
}
* { box-sizing: border-box; }
html { font-size: var(--base-font-size); scroll-behavior: smooth; }
body {
  margin: 0;
  color: var(--ink);
  font-family: var(--font-current);
  line-height: 1.75;
  min-height: 100vh;
  overflow-x: hidden;
  background:
    radial-gradient(circle at 15% 20%, rgba(135, 30, 42, 0.18), transparent 32%),
    radial-gradient(circle at 85% 12%, rgba(185, 141, 53, 0.2), transparent 28%),
    linear-gradient(180deg, #f3ebdf 0%, #f7f1e7 36%, #efe2ca 100%);
}
a { color: inherit; text-decoration: none; }
img { display: block; max-width: 100%; }
figure { margin: 0; }
.progress {
  position: fixed; inset: 0 auto auto 0; z-index: 50; width: 100%; height: 3px;
  transform-origin: left center; transform: scaleX(0);
  background: linear-gradient(90deg, var(--accent), #d2a039, #54756c);
}
.page-shell { width: var(--content-width); margin: 0 auto; padding: 24px 0 72px; }
.hero {
  position: relative; overflow: hidden; min-height: calc(100svh - 48px);
  padding: 28px; border-radius: 40px; isolation: isolate; box-shadow: var(--shadow);
  border: 1px solid rgba(255,255,255,.4);
  background: linear-gradient(130deg, rgba(82,19,30,.94), rgba(122,36,50,.84) 35%, rgba(178,127,44,.72));
}
.hero::after {
  content: ""; position: absolute; inset: 0; z-index: -1;
  background: linear-gradient(180deg, rgba(16,10,8,.12), rgba(16,10,8,.48));
}
.hero-topbar { display: flex; align-items: center; justify-content: space-between; gap: 16px; min-width: 0; }
.eyebrow {
  display: inline-flex; align-items: center; gap: 12px;
  color: rgba(255,243,228,.8); font-size: .9rem; letter-spacing: .28em;
}
.eyebrow::before { content: ""; width: 44px; height: 1px; background: rgba(255,243,228,.58); }
.hero-nav { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 10px; }
.hero-nav a, .mobile-jumpbar a {
  padding: 8px 14px; border-radius: 999px; color: rgba(255,243,228,.9);
  border: 1px solid rgba(255,243,228,.22); background: rgba(255,251,243,.08);
  font-size: .88rem;
}
.hero-grid {
  display: grid; grid-template-columns: minmax(0, 1.1fr) minmax(280px, .74fr);
  gap: 36px; align-items: end; min-height: calc(100svh - 190px); padding-top: 68px;
}
.hero-copy h1 {
  margin: 0; font-size: clamp(3.35rem, 7.8vw, 6.75rem);
  line-height: .96; letter-spacing: -.04em; color: #fff8ef;
}
.hero-meta {
  display: grid; gap: 18px; width: 100%; min-width: 0; align-self: stretch;
  padding: 26px; border-radius: 28px; color: #fff7ea;
  border: 1px solid rgba(255,243,228,.18);
  background: linear-gradient(180deg, rgba(255,251,243,.16), rgba(255,251,243,.06));
  backdrop-filter: blur(14px);
}
.hero-meta-label { font-size: .9rem; letter-spacing: .24em; color: rgba(255,243,228,.68); }
.hero-meta ul { display: grid; gap: 14px; margin: 0; padding: 0; list-style: none; }
.hero-meta li { padding-bottom: 12px; border-bottom: 1px solid rgba(255,243,228,.12); overflow-wrap: anywhere; }
.hero-meta li:last-child { border-bottom: 0; }
.mobile-jumpbar { display: none; }
.layout { display: grid; grid-template-columns: 270px minmax(0, 1fr); gap: 28px; margin-top: 28px; align-items: start; }
.toc {
  position: sticky; top: 24px; padding: 22px 20px; border-radius: 28px;
  background: rgba(255,251,243,.68); border: 1px solid rgba(255,255,255,.54);
  box-shadow: 0 16px 36px rgba(88,54,29,.08); backdrop-filter: blur(16px);
}
.toc-title { margin: 0 0 14px; font-size: 1.3rem; letter-spacing: .08em; }
.toc-list { display: grid; gap: 8px; margin: 0; padding: 0; list-style: none; }
.toc-list a { display: block; padding: 10px 12px; border-radius: 14px; color: var(--ink-soft); }
.toc-list a:hover, .toc-list a.is-active { color: var(--accent); background: rgba(135,30,42,.08); }
.article { display: grid; gap: 28px; }
.section {
  position: relative; overflow: hidden; padding: 28px; border-radius: 32px;
  background: linear-gradient(180deg, var(--paper-strong), var(--paper));
  border: 1px solid rgba(255,255,255,.72); box-shadow: 0 24px 50px rgba(90,56,31,.08);
  opacity: 0; transform: translateY(24px); transition: opacity 720ms ease, transform 720ms ease;
  scroll-margin-top: 24px;
}
.section.is-visible { opacity: 1; transform: translateY(0); }
.section h2 { margin: 0; font-size: clamp(2.1rem, 3.2vw, 3.15rem); line-height: 1.08; }
.section h3 { margin: 0; font-size: clamp(1.45rem, 2.2vw, 1.95rem); line-height: 1.18; }
.section h4 { margin: 0; font-size: 1.22rem; }
.section-header { display: grid; gap: 10px; margin-bottom: 22px; }
.section p { margin: 0 0 1em; color: var(--ink-soft); }
.detail-line {
  display: grid; grid-template-columns: max-content minmax(0, 1fr); gap: 10px;
  align-items: start; margin: 0 0 12px;
}
.detail-line strong { color: var(--accent-deep); font-weight: 700; }
.detail-line span { color: var(--ink-soft); }
.stats-grid {
  display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px;
}
.stat-card {
  min-height: 148px; padding: 22px; border-radius: 24px;
  background: linear-gradient(160deg, rgba(255,255,255,.7), rgba(246,238,224,.92));
  border: 1px solid rgba(185,141,53,.18);
  display: grid; align-content: space-between; gap: 14px;
}
.stat-card:nth-child(1) {
  background: linear-gradient(160deg, rgba(125,24,39,.95), rgba(96,17,29,.88));
  color: #fff6ea;
}
.stat-card:nth-child(2) {
  background: linear-gradient(160deg, rgba(191,148,57,.28), rgba(255,250,240,.92));
}
.stat-card:nth-child(3) {
  background: linear-gradient(160deg, rgba(79,114,104,.18), rgba(255,252,246,.92));
}
.stat-card:nth-child(4) {
  background: linear-gradient(160deg, rgba(255,245,228,.94), rgba(236,220,184,.8));
}
.stat-label { color: inherit; font-size: .85rem; letter-spacing: .12em; opacity: .8; }
.stat-value {
  margin: 0; color: inherit; font-size: clamp(1.35rem, 2.3vw, 2.35rem);
  line-height: 1.16; font-weight: 700;
}
.stat-raw { margin: 0; color: inherit; opacity: .78; font-size: .95rem; }
.platform-card p { margin: 0; color: inherit; }
.story-grid {
  display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px;
}
.story-card {
  display: grid; align-content: start; gap: 12px;
  padding: 22px; border-radius: 26px;
  background: linear-gradient(160deg, rgba(255,255,255,.86), rgba(245,236,221,.9));
  border: 1px solid rgba(111,74,43,.12);
}
.story-card p { margin: 0; }
.story-card .image-block { margin: 4px 0 0; }
.story-card .image-block img { width: 100%; height: auto; object-fit: contain; }
.image-gallery {
  display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 16px;
  align-items: start;
}
.image-gallery .image-block { margin: 0; }
.image-gallery img { display: block; width: 100%; max-width: 100%; height: auto; object-fit: contain; }
.image-block { margin: 18px 0; }
.image-block img, td img {
  width: min(100%, 520px); height: auto; border-radius: 20px; object-fit: contain;
  box-shadow: 0 14px 30px rgba(72,44,22,.12);
}
.image-gallery .image-block img { width: 100%; }
.table-shell { overflow: hidden; border-radius: 24px; border: 1px solid rgba(111,74,43,.12); background: rgba(255,255,255,.76); margin: 18px 0; }
.table-scroll { overflow-x: auto; -webkit-overflow-scrolling: touch; }
table { width: 100%; min-width: 720px; border-collapse: collapse; }
th, td { padding: 16px 18px; text-align: left; vertical-align: top; border-bottom: 1px solid rgba(111,74,43,.12); }
th { background: rgba(135,30,42,.94); color: #fff7ea; font-weight: 600; }
td p { margin: 0; color: inherit; }
.platform-row td { background: rgba(239,226,201,.78); color: var(--accent-deep); font-weight: 600; }
.size-switcher { position: fixed; right: 20px; bottom: 20px; z-index: 60; }
.size-switcher-trigger {
  display: grid; place-items: center; width: 56px; height: 56px; border-radius: 999px;
  border: 1px solid rgba(111,74,43,.16); background: rgba(255,251,243,.92);
  color: var(--accent-deep); font: inherit; font-size: 22px; cursor: pointer;
  box-shadow: 0 18px 42px rgba(79,47,21,.14); backdrop-filter: blur(18px);
}
.size-switcher-panel {
  position: absolute; right: 0; bottom: 68px; display: grid; gap: 10px;
  width: min(220px, calc(100vw - 32px)); padding: 14px; border-radius: 22px;
  background: rgba(255,251,243,.96); border: 1px solid rgba(111,74,43,.14);
  box-shadow: 0 18px 42px rgba(79,47,21,.14); backdrop-filter: blur(18px);
  opacity: 0; transform: translateY(10px) scale(.98); transform-origin: right bottom;
  pointer-events: none; transition: opacity 180ms ease, transform 180ms ease;
}
.size-switcher.is-open .size-switcher-panel { opacity: 1; transform: translateY(0) scale(1); pointer-events: auto; }
.size-switcher-title { margin: 0; color: var(--accent-deep); font-size: 14px; letter-spacing: .18em; }
.size-switcher-readout { display: flex; justify-content: space-between; gap: 12px; align-items: baseline; }
.size-switcher-value { color: var(--accent-deep); font-size: 24px; line-height: 1; }
.size-switcher-note, .size-switcher-scale { color: rgba(35,24,22,.58); font-size: 13px; }
.size-slider { width: 100%; margin: 0; accent-color: var(--accent); cursor: pointer; }
.size-switcher-scale { display: flex; justify-content: space-between; gap: 8px; font-size: 12px; }
@media (max-width: 1120px) {
  .hero { min-height: auto; }
  .hero-grid, .layout, .story-grid { grid-template-columns: 1fr; }
  .toc { display: none; }
  .mobile-jumpbar {
    position: sticky; top: 12px; z-index: 30; display: flex; gap: 10px; overflow-x: auto;
    margin: 16px 0 4px; padding: 10px; border-radius: 999px;
    background: rgba(255,251,243,.82); border: 1px solid rgba(255,255,255,.72);
    box-shadow: 0 14px 30px rgba(80,47,23,.08); scrollbar-width: none;
  }
  .mobile-jumpbar a { flex: 0 0 auto; color: var(--accent-deep); background: rgba(135,30,42,.06); }
}
@media (max-width: 760px) {
  :root { --content-width: calc(100% - 20px); }
  .page-shell { padding: 10px 0 48px; }
  .hero { padding: 20px; border-radius: 28px; }
  .hero-nav { display: none; }
  .hero-grid { min-height: auto; gap: 22px; padding-top: 28px; }
  .hero-copy h1 { font-size: clamp(2.6rem, 13.2vw, 3.35rem); }
  .hero-meta { padding: 16px; border-radius: 22px; }
  .hero-meta-label { font-size: .84rem; }
  .hero-meta li { font-size: .95rem; }
  .section {
    overflow: visible; padding: 22px 18px; border-radius: 24px;
    opacity: 1; transform: none;
  }
  .stats-grid, .story-grid, .image-gallery { grid-template-columns: 1fr; }
  .detail-line { grid-template-columns: 1fr; gap: 4px; }
  .stat-card { min-height: 0; padding: 18px; border-radius: 22px; }
  .story-card { padding: 18px; border-radius: 22px; }
  .image-gallery { gap: 14px; }
  .image-gallery img { height: auto; max-height: none; }
  .table-shell { overflow: visible; background: transparent; border: 0; }
  .table-scroll { overflow: visible; }
  table { min-width: 0; border-collapse: separate; border-spacing: 0; }
  thead { position: absolute; width: 1px; height: 1px; margin: -1px; overflow: hidden; clip: rect(0,0,0,0); white-space: nowrap; border: 0; }
  tbody, tr, td { display: block; width: 100%; }
  tbody { display: grid; gap: 14px; }
  tr { overflow: hidden; border-radius: 20px; background: rgba(255,255,255,.86); border: 1px solid rgba(111,74,43,.12); box-shadow: 0 14px 28px rgba(81,48,24,.06); }
  td { padding: 12px 14px; border-bottom: 1px solid rgba(111,74,43,.1); }
  td:last-child { border-bottom: 0; }
  td::before { content: attr(data-label); display: block; margin-bottom: 6px; color: var(--accent); font-size: 12px; letter-spacing: .12em; }
  .platform-row td { border-bottom: 0; background: rgba(239,226,201,.88); }
  td img { width: min(100%, 260px); border-radius: 18px; }
  .catalogue-table tbody { gap: 16px; }
  .catalogue-table tr:not(.platform-row) {
    display: grid; grid-template-columns: minmax(0, .9fr) minmax(0, 1.1fr);
    gap: 10px 14px; padding: 16px; border-radius: 24px;
    background: linear-gradient(160deg, rgba(255,255,255,.94), rgba(247,239,226,.88));
  }
  .catalogue-table tr:not(.platform-row) td {
    min-width: 0; padding: 0; border-bottom: 0;
  }
  .catalogue-table tr:not(.platform-row) td::before {
    margin-bottom: 4px; font-size: 12px; letter-spacing: .14em;
  }
  .catalogue-table tr:not(.platform-row) td:first-child {
    color: var(--ink); font-size: 1.12rem; font-weight: 700; line-height: 1.35;
  }
  .catalogue-table tr:not(.platform-row) td:nth-child(2) {
    color: var(--ink-soft); line-height: 1.45;
  }
  .size-switcher { right: 12px; bottom: 12px; }
  .size-switcher-trigger { width: 52px; height: 52px; font-size: 20px; }
  .size-switcher-panel { bottom: 62px; width: min(220px, calc(100vw - 24px)); padding: 12px; border-radius: 18px; }
}
@media (prefers-reduced-motion: reduce) {
  html { scroll-behavior: auto; }
  *, *::before, *::after { animation: none !important; transition: none !important; }
  .section { opacity: 1; transform: none; }
}
"""


JS = r"""
const progress = document.getElementById("progress");
const sections = Array.from(document.querySelectorAll(".section"));
const navLinks = Array.from(document.querySelectorAll(".toc-list a"));
const sizeSwitcher = document.getElementById("size-switcher");
const sizeSwitcherTrigger = document.getElementById("size-switcher-trigger");
const sizeSlider = document.getElementById("size-slider");
const sizeValue = document.getElementById("size-value");
const sizeStorageKey = "sky-gugong-report-font-size";

const applySizeChoice = (sizeValuePx) => {
  const numericValue = Number(sizeValuePx);
  const nextValue = Number.isFinite(numericValue)
    ? Math.min(Math.max(numericValue, 16), 24)
    : 18;
  document.documentElement.style.setProperty("--base-font-size", `${nextValue}px`);
  sizeSlider.value = String(nextValue);
  sizeValue.textContent = `${nextValue}px`;
  window.localStorage.setItem(sizeStorageKey, String(nextValue));
};

const setSizeSwitcherOpen = (isOpen) => {
  sizeSwitcher.classList.toggle("is-open", isOpen);
  sizeSwitcherTrigger.setAttribute("aria-expanded", String(isOpen));
};

const updateProgress = () => {
  const total = document.documentElement.scrollHeight - window.innerHeight;
  const ratio = total > 0 ? window.scrollY / total : 0;
  progress.style.transform = `scaleX(${Math.min(Math.max(ratio, 0), 1)})`;
};

const visibilityObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (entry.isIntersecting) entry.target.classList.add("is-visible");
  });
}, { threshold: 0.16, rootMargin: "0px 0px -8% 0px" });

sections.forEach((section) => visibilityObserver.observe(section));

const navObserver = new IntersectionObserver((entries) => {
  entries.forEach((entry) => {
    if (!entry.isIntersecting) return;
    const id = entry.target.id;
    navLinks.forEach((link) => {
      link.classList.toggle("is-active", link.getAttribute("href") === `#${id}`);
    });
  });
}, { threshold: 0.5, rootMargin: "-10% 0px -35% 0px" });

sections.forEach((section) => navObserver.observe(section));
sizeSwitcherTrigger.addEventListener("click", () => {
  setSizeSwitcherOpen(!sizeSwitcher.classList.contains("is-open"));
});
document.addEventListener("click", (event) => {
  if (!sizeSwitcher.contains(event.target)) setSizeSwitcherOpen(false);
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") setSizeSwitcherOpen(false);
});
sizeSlider.addEventListener("input", (event) => applySizeChoice(event.target.value));
applySizeChoice(window.localStorage.getItem(sizeStorageKey) || "18");
setSizeSwitcherOpen(false);
updateProgress();
window.addEventListener("scroll", updateProgress, { passive: true });
window.addEventListener("resize", updateProgress);
"""


def prepare_output(output_dir: Path) -> None:
    if output_dir.exists():
        shutil.rmtree(output_dir)
    (output_dir / "css").mkdir(parents=True, exist_ok=True)
    (output_dir / "js").mkdir(parents=True, exist_ok=True)
    (output_dir / "images").mkdir(parents=True, exist_ok=True)


def build(docx_path: Path, output_dir: Path) -> None:
    if not docx_path.exists():
        fail(f"input docx not found: {docx_path}")
    if docx_path.suffix.lower() != ".docx":
        fail("input must be a .docx file")
    prepare_output(output_dir)
    blocks = parse_docx(docx_path, output_dir)
    if not blocks:
        fail("no content found in docx")
    title, content_blocks = split_title(blocks)
    (output_dir / "index.html").write_text(
        render_html(title, content_blocks), encoding="utf-8"
    )
    (output_dir / "css" / "styles.css").write_text(CSS.strip() + "\n", encoding="utf-8")
    (output_dir / "js" / "main.js").write_text(JS.strip() + "\n", encoding="utf-8")
    print(f"created {output_dir}")
    print(f"blocks: {len(blocks)}")
    print(f"images: {len(list((output_dir / 'images').glob('*')))}")


def main(argv: list[str]) -> None:
    if len(argv) != 3:
        fail("usage: build_gugong_report.py input.docx output-folder")
    build(Path(argv[1]).expanduser().resolve(), Path(argv[2]).expanduser().resolve())


if __name__ == "__main__":
    main(sys.argv)

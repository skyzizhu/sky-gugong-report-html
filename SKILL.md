---
name: sky-gugong-report-html
description: Convert a Word .docx report into a Gugong Museum editorial-style responsive HTML site using a fixed visual template, preserving the Word outline order, all text, tables, and embedded images.
---

# Sky Gugong Report HTML

Use this skill when the user provides a Word `.docx` report and asks to generate an HTML page/site in the established Gugong Museum report style.

## Output Contract

Generate a folder containing:

- `index.html`
- `css/styles.css`
- `js/main.js`
- `images/` with images extracted from the Word file

The HTML must be a standalone report page using the Gugong visual language: warm paper background, deep palace red accents, gold/jade secondary tones, large Songti typography, rounded editorial cards, mobile-first responsive tables/cards, and a collapsible bottom-right font-size slider.

## Required Layout Components

- `全网信息总览`: render as compact statistic cards, not plain paragraphs.
- `今日关注`, `其他信息`, `AI侵权`, `参考消息`: render title/source/summary/image groups as event cards.
- `商业/IP`: render Word tables as compact commercial tables that become card-like rows on mobile.
- `图片合集` or image-heavy sections: render extracted images as a responsive gallery grid.
- Keep the visual rhythm close to the standard template: rounded cards, warm translucent surfaces, deep red labels, and tight mobile spacing.
- Images must show in full by default. Do not crop screenshots or product images with `object-fit: cover`; use full-image display rules unless the user explicitly asks for cropping.
- Subheadings under a section must stay visually grouped with their parent section instead of creating empty standalone sections.
- On mobile, report content must be visible without relying on scroll-triggered reveal classes. Decorative reveal animations may be disabled on mobile to avoid hidden late-page content.

## Non-Negotiable Content Rules

- Preserve the Word document outline order exactly.
- Do not drop, summarize, rewrite, invent, or add report content.
- Do not add explanatory UI copy as report content. Navigation labels may be derived only from Word headings.
- Preserve all non-empty paragraphs, headings, tables, and embedded images.
- Extract Word images into an `images/` directory and link them with relative paths.
- If the source structure is ambiguous, prefer including content over omitting it.
- After generation, inspect the output for missing image links and obvious lost sections.

## Recommended Workflow

1. Run the bundled converter:

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py input.docx output-folder
```

2. Open `output-folder/index.html` locally and verify:

- The section order matches the Word outline.
- All extracted images display.
- Tables are readable on desktop and collapse into card-like rows on mobile.
- The bottom-right `Aa` button opens the font-size slider and adjusts body text live.

3. If the Word file has unusual formatting that the script cannot infer cleanly, manually adjust only structure/styling while preserving source content verbatim.

## Implementation Notes

- The converter uses only Python standard library modules and reads `.docx` as OOXML, so it does not require `python-docx`.
- Heading detection uses Word paragraph styles first, then conservative Chinese outline heuristics.
- The first strong title-like paragraph becomes the hero title. Remaining headings and content render in document order.
- Tables are rendered as semantic HTML tables and become mobile cards through CSS.
- Embedded images are copied from the Word package into the generated `images/` directory.
- Image extraction optimizes large images for the web: files under 400KB are preserved as-is; larger images are resized only if their longest side exceeds 2000px and are saved with JPEG quality 86 when that reduces file size. Image proportions must not be changed or cropped.

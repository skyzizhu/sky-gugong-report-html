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

After generating the folder, do not upload automatically. First show the local `index.html` path and ask the user to confirm whether to upload the complete output directory to Aliyun OSS. Upload only after the user explicitly confirms, then return the browser-viewable `final_url`.

By default, when no output folder is passed, the converter creates the site next to the input Word file using the current date format `YYYYMMDD_report`, for example `20260509_report`. If that directory already exists, it automatically creates `YYYYMMDD_report_2`, `YYYYMMDD_report_3`, and so on instead of overwriting an older report.

The HTML must be a standalone report page using the Gugong visual language: warm paper background, deep palace red accents, gold/jade secondary tones, large Songti typography, rounded editorial cards, mobile-first responsive tables/cards, and a collapsible bottom-right font-size slider.

## Required Layout Components

- `全网信息总览`: render as compact statistic cards, not plain paragraphs.
- Overview statistic cards should show each metric once. Do not repeat the original raw sentence under the formatted value.
- `今日关注`, `其他信息`, `AI侵权`, `参考消息`: render title/source/summary/image groups as event cards.
- `商业/IP` sections and sections containing `商业产品图集`: render Word tables as compact commercial tables that become card-like rows on mobile.
- On mobile commercial table cards, text cells should keep the current typography but use a stable two-line summary: shorter metadata fields on the first line separated by small dots (`·`), and the final product/title field on the second line. Image cells stay visually separate.
- `图片合集` or image-heavy sections: render extracted images as a responsive gallery grid.
- Keep the visual rhythm close to the standard template: rounded cards, warm translucent surfaces, deep red labels, and tight mobile spacing.
- Hero report titles should stay expressive but not oversized; keep the main hero heading about 16px smaller than the earlier template scale on both desktop and mobile.
- First-level in-report section titles should stay restrained; avoid oversized section headings and keep them about 6px smaller than the earlier template scale.
- Second-level and third-level in-report headings should use body-sized typography with bold weight, not enlarged heading typography.
- Overview statistic cards should use compact vertical spacing: values should sit close to their labels, and card height plus bottom padding should avoid excessive blank space below the content across all quick overview cards.
- Overview statistic card labels are second-level titles and should be visibly larger and bolder than body text. The statistic values and topic lines underneath are body content, not headings, and should use restrained body-sized typography.
- Overview statistic card text should be smaller than the original oversized template scale, and the primary red card should use a lighter red gradient with forced high-contrast warm-white label and value text for readability.
- In overview statistic cards, values separated by `|` or `｜` should render as separate lines instead of keeping the divider character inline.
- In overview statistic cards, a standalone label ending with `：` or `:` must stay grouped with the following content paragraphs until the next label/table/image. Do not split a label such as `核心热词：` and its following value lines into separate cards.
- Avoid duplicate title presentation in the hero area. The main report title should not be echoed twice as both a label and the primary heading.
- Images must show in full by default. Do not crop screenshots or product images with `object-fit: cover`; use full-image display rules unless the user explicitly asks for cropping.
- Subheadings under a section must stay visually grouped with their parent section instead of creating empty standalone sections.
- On mobile, report content must be visible without relying on scroll-triggered reveal classes. Decorative reveal animations may be disabled on mobile to avoid hidden late-page content.
- When a Word table has no real header row, do not invent generic mobile labels such as `列1` or `列2`. Show only labels that genuinely exist in the source content.
- On mobile, table rows rendered as nested cards inside a section/card should sit 3px closer to the outer card edges on both sides, and table-cell vertical padding should stay compact so row cards do not have excessive empty space at the bottom.
- Detail cards should use compact internal rhythm: reduce spacing between the card title and its content, keep labeled fields such as `平台TOP3` / `正面` / `吐槽` / `质疑` close together inside the same card, and keep adjacent detail cards slightly closer together.
- Tables with headers `媒体 / 标题 / 数据` should render as polished media report tables with forced warm-white header text on a deep red background, never inherited black text, subtle row striping, compact mobile typography, and horizontal scrolling on mobile instead of converting rows into stacked cards.
- Keep the visual finish refined without changing layout: use warm gray body text, softened borders and shadows on cards/tables, and thin warm-toned horizontal scrollbars for overflow tables.

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
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py input.docx
```

Optional: if you want to specify a folder manually, you can still pass it as the second argument. If that folder name already exists, the converter will create a suffixed directory such as `_2` instead of deleting the existing one.

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/build_gugong_report.py input.docx output-folder
```

2. Open the generated `index.html` locally and verify:

- The section order matches the Word outline.
- All extracted images display.
- Tables are readable on desktop and collapse into card-like rows on mobile.
- The bottom-right `Aa` button opens the font-size slider and adjusts body text live.

3. Ask the user whether to upload the generated folder to Aliyun OSS.

Do not run the upload command until the user explicitly confirms. The confirmation should happen after local generation and verification, so the user can decide whether the report is ready to publish.

4. After confirmation, upload the generated folder to Aliyun OSS:

```bash
python3 /Users/fushan/.codex/skills/sky-gugong-report-html/scripts/upload_to_oss.py output-folder
```

The upload script reads `config/oss_config.json` by default. Keep the real config local and uncommitted; commit only `config/oss_config.example.json`.

- The OSS upload prints `public_url` and `final_url`. Return only `final_url` as the final deliverable link. `final_url` must be a browser-viewable signed URL with `response-content-disposition=inline`, not a link that downloads `index.html`. It should use the configured `public_base_url` such as `http://report.blynkai.com`, and `signed_url_expires_days` should usually be `3` so the link is available for three days after upload.

5. If the Word file has unusual formatting that the script cannot infer cleanly, manually adjust only structure/styling while preserving source content verbatim.

## Implementation Notes

- The converter uses only Python standard library modules and reads `.docx` as OOXML, so it does not require `python-docx`.
- Heading detection uses Word paragraph styles first, then conservative Chinese outline heuristics.
- The first strong title-like paragraph becomes the hero title. Remaining headings and content render in document order.
- Tables are rendered as semantic HTML tables and become mobile cards through CSS.
- Embedded images are copied from the Word package into the generated `images/` directory.
- Image extraction optimizes large images for the web: files under 400KB are preserved as-is; larger images are resized only if their longest side exceeds 2000px and are saved with JPEG quality 86 when that reduces file size. Image proportions must not be changed or cropped.

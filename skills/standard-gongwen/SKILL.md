---
name: standard-gongwen
description: Use when Codex needs to draft, format, review, or correct Chinese official documents and project-report 公文 according to fixed layout rules, including Word/DOCX page setup, cover pages, fonts, headings, body text, table text, page numbers, attachments, signature blocks, and related 公文格式要求.
---

# 标准公文

## Overview

Apply this skill when creating, editing, or auditing Chinese official documents. Preserve the user's substantive content while enforcing the formatting rules below.

When working with `.docx` or Word-targeted output, use the available document tooling for rendering and visual verification when possible. After formatting, check page layout, heading hierarchy, paragraph spacing, page numbers, attachments, and signature/date placement.

## Workflow

1. Identify document parts: title, subtitle, issuing unit, body headings, body paragraphs, attachments, signature, date, page numbers, and版记/抄送 if present.
2. Apply page setup first, then paragraph defaults, then part-specific typography.
3. Preserve official wording unless the user asks for rewriting.
4. Flag unclear or missing required content instead of inventing issuing units, dates, attachment names, or official abbreviations.
5. For Word/DOCX results, prefer `scripts/format_docx.py` for repeatable formatting, choose the profile before editing, then run `scripts/verify_docx.py` and render or inspect the output before delivery.
6. If LibreOffice/`soffice` is available, render with the Documents skill `render_docx.py`; if it is unavailable, disclose that visual QA was not completed and rely on structural DOCX inspection.

## Bundled Tools

- `scripts/generate_docx.py`: generate a new standard-gongwen DOCX from structured JSON. Usage: `python scripts/generate_docx.py 通知 content.json output.docx`; supported types are `通知`, `报告`, `请示`, `函`, and `纪要`. Use `--example` to generate a sample notice.
- `scripts/format_docx.py`: apply this standard to an existing DOCX. Usage: `python scripts/format_docx.py input.docx --profile standard --backup` for ordinary 公文, or `python scripts/format_docx.py input.docx --profile project-report --output output.docx` for project technical reports.
- `scripts/verify_docx.py`: inspect key DOCX settings after formatting, including per-section page geometry, margins, header/footer distances, odd/even page settings, paragraph samples, and table-text font buckets.
- Documents skill `render_docx.py`: use after formatting when available to convert the DOCX to PNG pages for visual QA. This renderer depends on LibreOffice/`soffice`.

When using the bundled scripts, run them with the Codex bundled Python environment if available because it includes `python-docx` in many Codex document workflows. If imports fail, resolve dependencies through the Documents skill/runtime rather than rewriting the script ad hoc.

Use `generate_docx.py` for new structured 公文 when the user provides fields such as issuer, doc number, title, recipient, body, attachments, issuer signature, date, copy-to, and print information. Use `format_docx.py` for already-existing DOCX files.

## Project Report Profile

Use `--profile project-report` for technical justification reports and similar procurement project reports with a cover page, report body, market/risk tables, and a signature page.

1. Preserve the original report or write a copy before formatting.
2. Treat the report cover as its own layout block when it appears before the first `一、` heading. Keep the report title centered in 小标宋/小标宋类 font at 二号 with fixed `33磅` line spacing; keep cover metadata such as 项目名称、项目类型、项目级别、项目主管部门 in 仿宋 `14pt` where the template uses it.
3. Treat body `一、` paragraphs as一级标题, `（一）` paragraphs as二级标题, and dot-number headings such as `1. 基本功能需求` as project-report third-level headings. Do not misclassify `1、合规目标` and similar body lists as headings.
4. Use the project-report body geometry from the standard template: A4 body sections with top/bottom `2.54cm`, left/right `3.17cm`, header/footer `1.27cm`. If the document has a separate cover section, format that cover section as A4 with top/bottom `2.54cm`, left/right `2.5cm`, header `1.5cm`, footer `1.75cm`.
5. Preserve table layout. Do not rewrite table styles, borders, widths, row heights, cell margins, cell alignment, or table paragraph geometry for project reports. Only normalize table run fonts so Chinese, digits, English, and symbols use `仿宋_GB2312`, preserving the table's existing font size and emphasis unless the user asks otherwise.

## Page Setup

- Set paper to A4: width `21cm`, height `29.7cm`.
- Set margins: top `3.7cm`, bottom `3.5cm`, left `2.8cm`, right `2.6cm`.
- Set header distance to `1.5cm`; set footer distance to `2.45cm`.
- Enable different odd and even pages.
- If configuring Word manually, use 页面布局 -> 页边距 -> 自定义边距, then set as default after all data is complete.

## Paragraph Defaults

- Use justified alignment for body text.
- Set paragraph style to 正文文本.
- Set spacing before `0` lines and after `0` lines.
- Set fixed line spacing to `29磅`.
- Set first-line indent to `2字符`.
- In正文 and all 仿宋 roles, set Chinese, digits, English letters, and punctuation/symbols to `仿宋_GB2312`; do not allow digits or Latin text to fall back to `Times New Roman`.
- Keep numbering markers and their following text at the same font size. For example, `1.` and the heading/body text after it must both be 三号 when the surrounding paragraph is 三号.

## Title Block

- Article title: use 小标宋, 二号, not bold, centered, fixed line spacing `33磅`.
- Subtitle: if present, start a new line after the dash/破折号; use 楷体_GB2312, 三号, not bold, centered.
- Unit name: place one blank line below title unless the surrounding正文 layout requires otherwise; use 仿宋体_GB2312, 三号, flush with the text area. Add a full-width colon after the name when appropriate.

## Heading Rules

- 一级标题: prefix with `一、`; use 黑体, 三号, not bold, single line spacing, first-line indent `2字符`; end with no punctuation.
- 二级标题: prefix with `（一）`; use 楷体_GB2312, 三号, bold, fixed line spacing `29磅`, first-line indent `2字符`.
- 三级标题: prefix with half-width Arabic numbering such as `1.`; use 仿宋体_GB2312, 三号, bold, fixed line spacing `29磅`, first-line indent. Add punctuation at the sentence end. Insert a half-width space after the `.`.
- 四级标题: prefix with `（1）`; use 仿宋体_GB2312, 三号, not bold, fixed line spacing `29磅`, first-line indent. Add punctuation at the sentence end.
- Apply the heading font and size to the whole heading paragraph, including Chinese sequence markers, Arabic numerals, English letters, parentheses, periods, colons, and other symbols.

## Body Text

- Use 三号 仿宋_GB2312.
- Use justified alignment.
- Set fixed line spacing to `29磅`.
- Set first-line indent to `2字符`.
- Ensure Chinese, numbers, English, and symbols all use 三号 `仿宋_GB2312` in body paragraphs and table body text.

## Page Numbers

- Use 四号 half-width Arabic numerals in 宋体.
- Place one em dash/一字线 on each side of the page number.
- Position page numbers at the bottom of the page.
- For odd pages, place the page number on the right with one-character spacing.
- For even pages, place the page number on the left with one-character spacing.
- Set footer boundary to `2.45cm`.
- In the footer editor, enable different first page when required by the document.

## Attachments

- In the main document, place the attachment notice one blank line below the body text; indent left by two characters.
- Write `附件：` with a full-width colon, followed by the attachment number/name, for example `附件：1.xxx`.
- In each attachment, place `附件` and the attachment sequence number in 三号 黑体 at the first line of the upper-left text area. Do not insert a space between them and do not add a colon after them.
- Format attachment body text the same as the main body text.
- Format 版记/抄送 in 四号 仿宋.

## Signature And Date

- Place the issuing authority signature above the document date.
- Center the issuing authority signature based on the document date; use the full name or standardized abbreviation.
- Write the document date with Arabic numerals and include full year, month, and day.

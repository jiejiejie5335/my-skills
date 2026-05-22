# My Skills

This repository collects reusable Codex skills.

## Included Skills

- `skills/standard-gongwen`: Draft, format, review, and verify Chinese official-document DOCX files using standard 公文 layout rules.
- `技术论证报告` Use --profile project-report for technical justification reports and similar procurement project reports with a cover page, report body, market/risk tables, and a signature page.
## Repository Layout

```text
my skills/
├── README.md
├── requirements.txt
└── skills/
    └── standard-gongwen/
        ├── SKILL.md
        ├── agents/
        └── scripts/
```

## Install

Install the Python dependency used by the bundled DOCX scripts:

```bash
python -m pip install -r requirements.txt
```

To make a skill discoverable by Codex, copy the selected skill directory into your Codex skills directory, for example:

```bash
cp -R skills/standard-gongwen ~/.codex/skills/
```

## Standard Gongwen Notes

The `standard-gongwen` scripts use `python-docx` and write Chinese DOCX font names such as `仿宋_GB2312`, `楷体_GB2312`, and 方正小标宋 variants into the document. Final rendering still depends on the fonts available in Word, LibreOffice, or the system used to open the DOCX.

Before publishing new sample documents or templates in this repository, remove real单位、联系人、项目数据、密级信息和其他不适合公开的内容。

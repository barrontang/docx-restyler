---
name: docx-restyler
description: Restyle and finalize Word documents by applying the visual and formatting rules of a reference .docx template to a source .docx file. Use when the user wants to standardize, reformat, or finalize a Word document according to a template, especially for headings, paragraphs, spacing, margins, tables, and quote-like blocks.
---

Use this skill when the user provides:
- a template `.docx` file
- a source `.docx` file
- and wants a finalized `.docx` output that follows the template's visual rules

## First version scope

Focus on **style transfer**, not content rewriting.

Support:
- page margins and page setup basics
- heading styles
- normal paragraph styles
- quote-like paragraph styles when detectable
- basic table reconstruction with template-facing style hints

Do not promise in v1:
- tracked changes
- comments
- footnotes/endnotes
- headers/footers parity
- perfect preservation of complex merged tables
- structural rewriting of the source document

## Workflow

1. Read the source and template files.
2. Read `references/design.md` and `references/mapping-rules.md` if implementation choices matter.
3. Run `scripts/restyle_docx.py` with:
   - `--template <template.docx>`
   - `--source <source.docx>`
   - `--output <final.docx>`
4. Inspect the output and report any uncertain mappings.

## Mapping strategy

Use a mixed strategy:
1. Prefer Word style-name mapping when styles are usable.
2. Fall back to content-feature heuristics when the source styles are messy.

Heuristics may use:
- paragraph length
- boldness
- font size
- heading-like numbering
- quote indentation signals

## Output expectations

Always report:
- input files used
- output file path
- style mapping assumptions
- anything that was preserved conservatively because the structure was too complex

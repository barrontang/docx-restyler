# docx-restyler summaries

## Chinese summary

`docx-restyler` is an OpenClaw skill for Word document restyling and finalization.

It takes:
1. a standard template `.docx`
2. a source `.docx`

and outputs a new `.docx` whose content comes from the source document while the formatting is aligned to the template.

It is designed for scenarios such as:
- official reports
- internal working materials
- meeting summaries
- Chinese formal/government-style documents

Current direction:
- preserve content
- restyle presentation
- prioritize Chinese heading hierarchy rules
- use template formatting as the standard

Current strengths:
- supports real template + source Word workflow
- focuses on headings, body paragraphs, basic tables, and tail sections
- tuned for Chinese official-document numbering such as `一、` / `（一）` / `1.` / `（1）`

Current limitations:
- complex tables still need work
- advanced Word features are not handled yet
- exact typography fidelity may still vary in difficult files

One-sentence summary:
**`docx-restyler` is an OpenClaw skill prototype that reformats a source Word document into the style of a reference template, especially for Chinese formal-document structures.**

---

## GitHub README-style summary

### docx-restyler

`docx-restyler` is an OpenClaw skill prototype for **restyling Word documents** using a reference `.docx` template.

#### What it does
Given:
- a **template Word document**
- a **source Word document**

it produces:
- a **new finalized `.docx`** whose content comes from the source file while formatting is aligned to the template.

#### Intended use cases
- official reports
- internal working documents
- structured enterprise materials
- Chinese government-style or formal business documents

#### Current approach
The current version focuses on:
- style transfer, not content rewriting
- template-driven formatting
- Chinese official heading hierarchy detection
- conservative body-text handling
- basic table carry-over

#### Current supported elements
- document title
- heading hierarchy
- body paragraphs
- basic quote-like blocks
- simple tables
- signature/date tail sections

#### Summary
**`docx-restyler` is an OpenClaw skill prototype for reformatting a source Word document into the style of a reference template, with extra attention to Chinese formal-document heading rules.**

---

## Product-spec summary

### Product name
**docx-restyler**

### Product type
OpenClaw skill / document finalization prototype

### Product goal
Enable users to generate a finalized `.docx` by applying the formatting rules of a standard template Word document to a source Word document.

### Primary workflow
**Input**
1. Template `.docx`
2. Source `.docx`

**Output**
- Finalized `.docx`

### Scope of v1
Included:
- title detection
- heading hierarchy mapping
- body text restyling
- page margin transfer
- paragraph spacing transfer
- basic table reproduction
- Chinese official-document heading recognition

Excluded:
- content rewriting
- structure redesign
- tracked changes/comments
- advanced header/footer parity
- complex table fidelity
- OCR/PDF input

### Summary
**docx-restyler is a practical prototype for Word document restyling and finalization, optimized for template-based formatting transfer and Chinese formal-document heading recognition.**

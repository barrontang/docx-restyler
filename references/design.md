# docx-restyler design

## Goal

Take two Word files:
- a template `.docx`
- a source `.docx`

Produce a new `.docx` whose content comes from the source document while visual formatting is aligned to the template.

## v1 philosophy

- Preserve content
- Rebuild presentation
- Avoid overclaiming structural intelligence
- Be conservative when uncertain

## Recommended technical stack

- Python
- python-docx

## Core pipeline

1. Parse template document styles and section geometry.
2. Parse source document blocks in order.
3. Classify paragraphs into rough semantic buckets:
   - title
   - heading1
   - heading2
   - heading3
   - body
   - quote
4. Rebuild a fresh output document.
5. Apply page setup and paragraph styles derived from the template.
6. Recreate tables conservatively.

## Important note

The output document should generally be a newly built file, not an in-place mutation of the source file. This makes the behavior more predictable and easier to debug.

# Mapping rules

## Priority order

1. Source paragraph style name
2. Visible formatting hints
3. Text-shape heuristics

## Suggested paragraph classification heuristics

### Title
- first meaningful paragraph in document
- short length
- larger font or bold

### Heading 1
- short paragraph
- bold or numbered section pattern
- larger than body text

### Heading 2 / Heading 3
- similar to heading 1 but lower apparent emphasis

### Body
- default fallback for longer paragraphs

### Quote
- indentation or explicit quote style name
- may be shorter, emphasized, or offset from body

## Table handling

- preserve row/column counts
- preserve cell text
- do not aggressively preserve exotic merged layouts in v1
- prefer readable output over exact low-level fidelity

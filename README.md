# docx-restyler

[中文文档](README_CN.md)

Restyle and finalize Word documents by applying the formatting rules of a reference `.docx` template to a source `.docx` file.

## Quick summary

- **Input**: a template `.docx` + a source `.docx`
- **Output**: a newly finalized `.docx`
- **Focus**: formatting transfer, not content rewriting
- **Best fit**: official reports, internal documents, Chinese formal materials

## Why it matters

A lot of document work is not writing — it is reformatting.

`docx-restyler` exists to reduce the mechanical part: take a draft Word file, apply the formatting logic of a reference template, and output a cleaner final document.

## What it does

`docx-restyler` is an OpenClaw skill prototype for cases where you already have:
- a **template Word document** that represents the target format
- a **source Word document** that contains the content you want to finalize

The skill produces a new `.docx` whose content comes from the source file while the formatting is aligned to the template.

## Typical use cases

- official reports
- enterprise/internal reporting
- Chinese formal business or government-style materials
- documents that must be standardized before final delivery

## Current v1 focus

The current prototype focuses on:
- style transfer, not content rewriting
- heading hierarchy detection
- body paragraph normalization
- basic table carry-over
- signature/date tail handling

## Special handling

The skill now prioritizes Chinese formal heading rules such as:
- `一、`
- `（一）`
- `1.` / `1、`
- `（1）`

This is important because many real source Word files do not use reliable built-in Word styles.

## Current limitations

- complex table fidelity is still limited
- advanced Word features are not fully supported
- exact typography matching may still vary in difficult documents

## One-sentence summary

**`docx-restyler` is an OpenClaw skill prototype that reformats a source Word document into the style of a reference template, especially for Chinese formal-document structures.**

## Current status

This is still a prototype. It already works on real `.docx` template + source flows, but it should be treated as an iterative document-engine skill, not a finished production formatter.

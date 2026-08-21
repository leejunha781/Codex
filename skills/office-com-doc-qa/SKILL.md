---
name: office-com-doc-qa
description: Create, edit, inspect, export, and visually QA Microsoft Office documents on Windows using PowerShell 5.1 and Office COM automation. Use when Codex needs to work on resumes, proposals, application documents, Word DOC/DOCX files, PowerPoint PPT/PPTX decks, PDF export, PDF text checks, or PDF-to-PNG visual QA without Python, Node.js, LibreOffice, or ad hoc Office file manipulation.
---

# Office COM Document QA

Use PowerShell 5.1 and Microsoft Office COM automation for Office document work. Do not use Python, Node.js, LibreOffice, or non-COM Office editors in this environment.

## Workflow

1. Read local project instructions first, especially `AGENTS.md`, `CLAUDE.md`, and project memory files when present.
2. Confirm the canonical source file and output folder before editing. If the user did not explicitly ask to overwrite the canonical file, edit a timestamped copy or a clearly named working copy.
3. Inspect the document through Office COM before changing it. For Word/PDF text, use Word COM; for PowerPoint slide structure, use PowerPoint COM.
4. Make scoped edits with PowerShell scripts or short COM snippets. Preserve user formatting unless the request requires layout changes.
5. Export the edited file to PDF.
6. Render the PDF to PNG pages and inspect every changed page visually. Use `view_image` for local PNG review when available.
7. Report exact output paths, pages checked, and any remaining visual or content risks.

## Bundled Resources

- `scripts/Export-OfficePdfPng.ps1`: Export `.doc`, `.docx`, `.ppt`, `.pptx`, or copy `.pdf`, then render PDF pages to PNG files for visual QA.
- `references/office-com-patterns.md`: Word and PowerPoint COM gotchas, reliable edit patterns, and QA snippets.

Run scripts in Windows PowerShell 5.1 without changing execution policy. If the current shell is PowerShell 7, launch `powershell.exe` explicitly:

```powershell
$scriptPath = "C:\Users\namma\.codex\skills\office-com-doc-qa\scripts\Export-OfficePdfPng.ps1"
$inputPath = "D:\path\file.docx"
$renderDir = "C:\Users\namma\.claude\qa-renders"
$cmd = "`$scriptBlock = [scriptblock]::Create((Get-Content -LiteralPath '$scriptPath' -Raw)); & `$scriptBlock -InputPath '$inputPath' -RenderDir '$renderDir'"
powershell.exe -NoProfile -Command $cmd
```

## Editing Rules

- Wrap every Office COM session in `try/finally`; close documents and quit the application when no documents remain.
- Avoid Word `Find.Execute` from PowerShell. Iterate paragraphs and normalize whitespace for matching.
- When replacing Word paragraph text, duplicate the range and set `End = End - 1` to keep the paragraph mark.
- After replacing Word text, reset inherited formatting and reapply only intentional bold or other lead-in formatting.
- Iterate Word table cells through `$table.Range.Cells`; avoid `$table.Cell(row,col)` when merged cells may exist.
- After deleting or moving Word tables, verify nearby headings because table deletion can truncate adjacent text.
- For PowerPoint export, open presentations with `WithWindow=$false` and use `SaveAs($pdfPath, 32)` to PDF.
- Kill only orphaned hidden Office processes when necessary, and only after checking `MainWindowTitle` is empty.

## Visual QA Checklist

- Check page count and expected output files.
- Inspect first page, every changed page, and the final page at minimum; inspect all pages for resumes and short proposals.
- Verify no text is clipped, overlapped, unexpectedly bold, or merged into adjacent paragraphs.
- Verify tables, bullets, slide titles, footers, page numbers, and section order.
- If a PDF text check is needed, open the PDF read-only with Word COM and compare key phrases from the source/edit request.

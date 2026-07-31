# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Environment constraints

- **Python and Node.js are installed** on this machine. **LibreOffice is not available**. All Office file editing must use **PowerShell + Office COM automation** (Microsoft Office 2016+ is installed at `C:\Program Files\Microsoft Office\root\Office16\`).
- Default shell is **PowerShell 7.6.2**. Use **Windows PowerShell 5.1** only when a task specifically requires it, such as Office COM compatibility. No `&&` pipeline chaining — use `; if ($?) { ... }` or separate statements.
- To run a `.ps1` script without changing execution policy: `Get-Content script.ps1 -Raw | Invoke-Expression` (using `-ExecutionPolicy Bypass` is blocked by the security hook).

## Office COM patterns (Word / PowerShell)

```powershell
# Attach to Word (creates a new hidden instance)
$word = New-Object -ComObject Word.Application
$word.Visible = $false

# Always wrap in try/finally to prevent orphaned WINWORD processes
try {
    $doc = $word.Documents.Open($path, $false, $false)   # ReadOnly=$false for editing
    # ... edits ...
    $doc.Save()
    $doc.ExportAsFixedFormat("out.pdf", 17)  # wdExportFormatPDF = 17
    $doc.Close([ref]$true)
} finally {
    if ($word.Documents.Count -eq 0) { $word.Quit() }
}

# Kill an orphaned hidden Word process (titleless window)
Get-Process WINWORD | Where-Object { [string]::IsNullOrEmpty($_.MainWindowTitle) } | Stop-Process -Force
```

**Key COM gotchas:**
- Cast numeric assignments: `Font.Size=[single]10`, `Font.Bold=[int]1`, `LineWidth=[int]6`
- **Never use `Find.Execute`** from PowerShell — throws `DISP_E_TYPEMISMATCH`. Instead iterate `$doc.Paragraphs` and match by `($p.Range.Text -replace "[\r\n\a\x07]","" -replace "\s+"," ").Trim()` (normalize whitespace — headings may contain non-breaking spaces).
- Replace paragraph text: `$r=$p.Range.Duplicate; $r.End=$r.End-1; $r.Text=$newText` (End-1 keeps the trailing ¶ mark to avoid merging with the next paragraph).
- After replacing text with `End-1`, the new run may inherit the first character's bold formatting — explicitly reset `$p.Range.Font.Bold=[int]0` then re-apply bold only to the lead-in.
- Table cell iteration (merge-safe): iterate `$tb.Range.Cells` not `$tb.Cell(r,c)` — `.Cell(r,c)` breaks after vertical merges.
- To move a block (heading + table): capture `Range.Start` of start and next heading, `$doc.Range(s,e).Cut()`, then re-find the target anchor by text (positions shift after Cut), and `$doc.Range(t,t).Collapse(1).Paste()`.
- `Tables.Delete()` can silently truncate adjacent heading text — verify heading text after any table deletion.
- When inserting before an anchor: `InsertParagraphBefore` first, then add content, to prevent accidental table auto-merge.

## Rendering PDFs to PNG (for visual QA, no poppler)

```powershell
Add-Type -AssemblyName System.Runtime.WindowsRuntime
# Load types
[Windows.Storage.StorageFile,Windows.Storage,ContentType=WindowsRuntime] | Out-Null
[Windows.Data.Pdf.PdfDocument,Windows.Data.Pdf,ContentType=WindowsRuntime] | Out-Null
[Windows.Storage.Streams.InMemoryRandomAccessStream,Windows.Storage.Streams,ContentType=WindowsRuntime] | Out-Null
# Use AsTask helpers to await IAsyncOperation/IAsyncAction
# Then: PdfDocument.LoadFromFileAsync -> GetPage(i) -> RenderToStreamAsync with DestinationWidth=[uint32]1240
# Dump bytes to .png via [System.IO.File]::WriteAllBytes()
```

Reading a PDF's text: `$word.Documents.Open($pdfPath, $false, $true)` then `$doc.Content.Text`.

## Work folders

| Folder | Contents |
|---|---|
| `C:\Users\namma\.Codex\itt_work\` | PowerShell build/edit scripts and PNG renders for the ITT Cannon FAE resume |
| `C:\Users\namma\.Codex\plm_slide_work\` | PowerShell build scripts and PNG renders for the AVEVA PLM presentation deck |
| `D:\이력서\ITT Cannon\` | ITT Cannon FAE resume files (canonical: `..._Final_Integrated_v2.docx` + `.pdf`) |
| `D:\이력서\` | Job application documents |

## Active projects

Use `C:\Users\namma\.codex\memories\MEMORY.md` as the Codex-native durable registry and quick-pass index. `C:\Users\namma\.codex\memory\` is the cross-tool Claude-memory mirror, while `C:\Users\namma\.cursor\memory\` is the synchronization hub. For substantive work, search the native registry first, then open only the task-specific mirrored file it identifies; do not load the full corpus by default. Key mirrored files:
- `office-docs-com-automation.md` — full COM gotcha reference
- `itt-cannon-fae-resume.md` — resume edit history, salary facts, JD alignment notes
- `aveva-plm-application-deck.md` — presentation deck edit history

## Codex Pro operating mode

When the user asks to act as "Codex pro", "evolve", or absorb Claude/Cursor strengths, treat it as a higher-initiative execution mode:

- Invoke the personal `$codex-pro` skill for the reusable cowork workflow.
- **Claude strengths to absorb:** selective durable-context boot, long-horizon continuity, explicit task framing, careful tone/structure review, and honest uncertainty.
- **Codex strengths to use:** `AGENTS.md`, matching skills, native subagents, connectors/MCP, automations, browser/computer use, precise local edits, and command/test/visual-QA loops.
- **Default loop:** boot relevant context -> define goal/constraints/exit criteria -> plan and parallelize independent reads when useful -> execute the smallest correct change -> verify -> independent review -> report artifacts and open risks.
- **Outward-facing work:** resumes, cover letters, LinkedIn posts, decks, proposals, and interview material must receive an accuracy/tone/structure/confidentiality review pass and evidence-backed wording; verify current facts from authoritative sources.
- **Autonomy:** prefer safe end-to-end execution over advice-only responses. Ask only when the missing answer materially changes the result, requires new authority, or makes a reasonable assumption risky.

### Capability routing

- Put durable environment and repository constraints in `AGENTS.md`; put repeatable procedures in focused skills; use MCP/connectors for live external or private data; use subagents for independent exploration, testing, or review.
- For Word, PowerPoint, or Office-targeted PDF work on this machine, prefer the personal `office-com-doc-qa` skill and the Windows Office COM rules above. Do not route production Office edits through LibreOffice- or Python-dependent document workflows.
- Do not weaken sandbox, trust, or approval settings merely to make Codex seem more capable. Permission breadth is not model intelligence.
- Do not create stop hooks that automatically re-enter the same agent task; they can cause infinite execution loops. Use explicit automations only for genuinely scheduled work.
- Keep memory retrieval selective and write durable memory only when the user explicitly asks to learn, remember, register, or persist it.

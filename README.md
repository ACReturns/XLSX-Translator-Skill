# XLSX-Translator-Skill

A Claude skill that translates an Excel workbook (`.xlsx`) into another language **without breaking it**. Formatting, formulas, merged cells, tables, comments, charts, images and column widths all survive. Sheet tab names get translated too, and every reference to them is rewritten so formulas keep working.

You hand Claude a workbook and a target language. You get back:

- **`<name>_<lang>.xlsx`**: the same workbook, text translated, layout kept.
- **`<name>_<lang>_review.xlsx`**: a side-by-side review sheet listing every translated segment (source, translation, fit status, notes) plus a Summary sheet of anything that needs a human eye.

No API key is needed by default. Claude does the translating itself, in the session, and the scripts handle all the file surgery and checking.

## How it works

An `.xlsx` is a zip of XML files. Instead of rebuilding the workbook with a library (which loses styles and quirks), the scripts patch only the text nodes and leave every other byte alone.

| Step | Script | What it does |
|---|---|---|
| 1. Scan | `xlsx_scan.py` | Validates the file, unzips it, inventories every translatable segment and skips numbers, codes, URLs and formulas |
| 2. Protect | `protect.py` | Locks URLs, e-mails, cell references, placeholders (`{0}`, `%s`) and rich-text runs behind tokens so they can't be mangled |
| 3. Translate | `run_translate.py` + `session_io.py` | Claude translates in rounds of ~150–250 items; a validator rejects anything that drops a token, breaks line counts or exceeds a hard length limit |
| 4. Apply | `xlsx_apply.py`, `xlsx_sheetnames.py` | Writes translations into a copy of the workbook and renames tabs, following each rename through formulas, defined names, charts, validations and conditional formats |
| 5. Fit | `fit.py`, `xlsx_fit.py` | Measures whether new text fits its cell using real font metrics. If not: wrap, grow the row (max 3×), then shrink the font (min 80%). Still too long → flagged, never cut |
| 6. Verify | `xlsx_verify.py` | Confirms formulas, values, merges, widths and untouched parts are identical to the source, re-checks fit, and opens/recalculates the output in LibreOffice. **If verify fails, nothing is delivered.** |
| 7. Review | `build_review.py` | Builds the color-coded review workbook and summary |

## Using it

Once installed, just ask Claude something like *"Translate this spreadsheet from Korean to English"* and attach the file. You don't need to say "xlsx". Claude asks for the target language if you didn't give one, runs the pipeline, and returns both files with a short summary of what was translated, adjusted or flagged.

Optional controls you can ask for:

- **Glossary**: force specific term translations (`{"source term": "required translation"}`)
- **Keep tab names**: leave sheet tabs in the original language
- **Fit limits**: turn off wrapping or row growth, or change the minimum font scale

### Large workbooks

Session translation is fine for a few hundred segments. For 1,500+ segments Claude will offer an API route instead: `--engine claude` (Anthropic key) or `--engine google` (Google Cloud Translation key). Keys are set with `configure.py` through a hidden prompt, stored in a local `.env`, and never pasted in chat.

## What it doesn't do

- `.xlsx` only. No `.xls`, `.xlsm` (macros are refused on purpose), `.csv`, Word, PowerPoint or PDF.
- Never changes column widths.
- Doesn't translate headers/footers, text boxes and shapes, chart text, dropdown list values, number-format text, pivot tables or document properties. These are **listed** on the Summary sheet, not silently skipped.
- A tab referenced inside a text formula like `INDIRECT("Sheet!A1")` is kept in the original language and flagged rather than risk breaking it.
- Formulas or dropdowns that compare against translated text get a warning on the Summary sheet; they're never edited.
- Fit for Arabic, Hebrew and Thai is an estimate (substitute fonts). Eyeball those in Excel.

## Requirements

Python 3.9+, `lxml`, `openpyxl`, `Pillow`, LibreOffice (`soffice`) and fontconfig (`fc-match`). The skill runs `preflight.py` first and stops with a clear message if anything is missing.

## Repo contents

- `xlsx-translator.skill`: the packaged skill (a zip). Upload it in Claude's Skills settings to install.
  - `SKILL.md`: the instructions Claude follows
  - `scripts/`: the pipeline above
  - `references/`: full spec (`sop.md`), session-engine format, troubleshooting by exit code
  - `assets/sample_tabs_ko.xlsx`: a tiny Korean workbook for a smoke test (`run_translate.py assets/sample_tabs_ko.xlsx en --mock --outdir /tmp/smoke`)

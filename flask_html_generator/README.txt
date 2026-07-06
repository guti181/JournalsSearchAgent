# HTML Generator App

## Install
pip install -r requirements.txt

## Run
python app.py

Default URL:
http://127.0.0.1:5050

## Workflow
1. Upload the raw keren Excel file.
2. Upload the raw agreement Excel file.
3. Optionally check "Include cleaned Excel files in the output".
4. Click Generate.

## Output
- If checkbox is OFF: download only the generated HTML.
- If checkbox is ON: download a ZIP file containing:
  - generated HTML
  - cleaned keren Excel
  - cleaned agreement Excel
  - processed keren Excel
  - processed agreement Excel

## Processing steps
1. Clean the raw workbooks.
2. Rebuild myName automatically.
3. Add name_changed.
4. Recalculate Sage "Open Access Fund Better".
5. Generate the final HTML.

## Canonical myName selection
The generator now chooses the first encountered non-empty name within each ISSN-connected cluster, to match the historical/manual processed outputs more closely.

The processor now applies fallback grouping by normalized name when rows have no usable ISSN, to better match the historical manual outputs.

The processor now normalizes each row name before grouping, so canonical myName selection happens after a normalize-first step instead of the older group-first step.

## Search behavior
The generated HTML now matches journal and book queries against displayName only, not all alternate name variants. ISSN search uses only valid 8-character normalized ISSNs.

Canonical myName selection now prefers a non-ALL-CAPS candidate when one exists; otherwise it falls back to the first encountered non-empty name.

Short text queries in the generated HTML now use prefix matching (startsWith) for 1-2 characters, and broader contains matching for longer queries.

Books-sheet parsing now detects the in-sheet Publisher header dynamically, and publisher deduplication treats forms like 'MIT Press' and 'MIT University press' as duplicates.

Publisher deduplication now prefers the fuller university-style form (for example, 'MIT University Press') over shorter forms like 'MIT Press', except forms starting with 'University of ...' still rank lower.


Embedding note:
Generated HTML styles and language direction are scoped to #appRoot. The generated app no longer sets CSS variables on :root, styles body, or changes document.documentElement/body for language switching, so it can be planted inside a website with its own headers, footers, colors, and direction.

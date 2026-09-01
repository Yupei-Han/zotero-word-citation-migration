---
name: zotero-word-citation-migration
description: Audit and convert bracketed static numeric or EndNote citations in complex Word DOCX files into Zotero-editable CSL fields, then refresh and verify them without disturbing scientific content or native Word structures. Use for Word citation migration; do not use for LaTeX/Markdown citation keys or Zotero library curation alone.
---

# Zotero Word Citation Migration

Produce a separate DOCX whose citations are real `ADDIN ZOTERO_ITEM CSL_CITATION` fields and whose reference list is one `ADDIN ZOTERO_BIBL ... CSL_BIBLIOGRAPHY` field. Preserve visible text and layout until Zotero performs the authoritative refresh.

## Required route

Before running a packaged Python script, call `load_workspace_dependencies` and use the returned bundled Python so `lxml` and the document runtime are known; do not assume the system Python is suitable.

1. Preserve the source and hash it. Never overwrite it.
2. Run `scripts/audit_docx_citations.py` on the source. Classify citations as existing Zotero, EndNote, or bracketed static numeric; inventory every Word story and the bibliography. Treat a mixed system or any bare-number superscript candidate as a blocker.
3. Read [references/workflow.md](references/workflow.md). Follow the matching and migration branch for the detected source type.
4. Create an auditable reference map. Read [references/mapping-schema.md](references/mapping-schema.md) when generating or reviewing the map.
5. Stop before writing if any used reference is missing, ambiguous, duplicated, or mapped only by an unreviewed fuzzy match.
6. Run `scripts/migrate_docx_citations.py` with the frozen source, reviewed manifest, and a new output path. It patches only necessary OOXML parts and refuses unsupported structures or overwrites.
7. Run `scripts/validate_zotero_docx.py` before opening Word, passing `--source` and `--expect-citations` from the manifest occurrence count. Any error is a blocker.
8. On Windows, run `scripts/refresh_zotero_word.ps1` on a copy. It must load `Zotero.dotm`, execute `ZoteroRefresh`, save successfully, and leave the source untouched. Prove editability on a disposable copy by changing one citation through Zotero, saving, reopening, and confirming the field remains valid; never deliver the smoke-test copy.
9. Run the validator again with the refreshed file. Then render and inspect every page using the documents workflow. Read [references/qa-gates.md](references/qa-gates.md) for the final gates.

## Non-negotiable distinctions

- A static number formatted as superscript is not a Zotero citation.
- An embedded CSL item can be editable inside the document but is not necessarily linked to a live library item. Report linked and embedded counts separately.
- A syntactically valid field is not proven usable until Zotero itself refreshes and saves it.
- Numeric styles renumber according to Zotero's actual Word field traversal, including table citations. Do not promise a body-only numbering order if table citations remain dynamic.
- Zotero library imports or metadata writes require explicit authorization. Read-only lookup and document-local embedded item data do not authorize library mutation.

## Default implementation choices

- Match by normalized DOI first, then PMID/ISBN when available, then exact normalized title plus compatible author/year. Use a reviewed override for exceptions.
- Use stable occurrence locators (`part`, paragraph index, field/marker index) rather than display text alone. A display-text queue is only a fallback for legacy EndNote files.
- Prefer localized OOXML over full `python-docx` reserialization for complex scientific documents.
- Use a temporary file and atomic replace for the new output. Copy original ZIP metadata and fail if unexpected package parts change.
- Reconstruct Zotero document preferences in `docProps/custom.xml`; add content-type and package relationships only when absent.

## Stop conditions

Do not deliver as converted when any of these remains: mixed citation systems, unreviewed bare-number superscript candidates, unresolved mapping, duplicate target identity, malformed CSL JSON, duplicate `citationID`, missing URI or `itemData`, unbalanced fields, residual `ADDIN EN.*`, missing/multiple bibliography fields, unexpected OOXML/package changes, failed Zotero refresh or edit/save smoke test, or incomplete visual QA.

After maintaining these packaged scripts, run `scripts/self_test.py` with the bundled Python.

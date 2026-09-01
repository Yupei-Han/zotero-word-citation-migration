# Workflow

## 1. Freeze and inventory

- Copy the source to a new working path and record its SHA-256.
- Inventory `word/document.xml`, footnotes, endnotes, headers, footers, comments, and text boxes contained in those parts.
- Parse both simple fields and complex `begin -> instrText -> separate -> result -> end` fields.
- Record non-citation Word fields such as `PAGE`, `REF`, and `TOC`; they must survive unchanged.
- Detect the reference heading and each visible bibliography entry. If the default heading, numbering, or bracketed-marker structure does not match uniquely, stop; the packaged workflow does not accept ad hoc regex overrides.

Use:

```powershell
python scripts/audit_docx_citations.py input.docx --out citation_inventory.json
```

The inventory is read-only. Review its `classification`, `issues`, field counts, citation occurrences, used/unused references, and out-of-range numbers before proceeding.

## 2. Choose the source branch

The packaged writer accepts only EndNote or bracketed static numeric sources. Existing Zotero fields or mixed citation systems are blockers; do not use this workflow as a repair path.

### EndNote fields

Extract metadata from both visible results and embedded `ADDIN EN.CITE` / `ADDIN EN.CITE.DATA` payloads. Preserve occurrence order and nested-field boundaries.

Build the reference identity table from the embedded EndNote record number, DOI/title metadata, and visible bibliography. Do not rely solely on the displayed citation number when the field payload identifies the source more precisely.

For migration:

- Replace only the outer EndNote field instruction with a Zotero CSL payload.
- Preserve the visible result runs and their original color; keep or restore native `w:vertAlign` superscript.
- Remove nested or standalone `ADDIN EN.CITE.DATA` fields only after their metadata has been captured.
- Replace `ADDIN EN.REFLIST` with one Zotero bibliography field.

### Bracketed static numeric citations

The packaged writer intentionally supports bracketed numeric clusters. Native superscript runs containing bare numbers are reported as blockers because years, quantities, footnote markers, and citations cannot be separated reliably. Do not silently treat every superscript number as a citation.

Parse each marker and expand ranges. Build the number-to-reference identity table from the visible bibliography, then bind each identity to Zotero.

- Preserve the cached marker text during initial OOXML insertion.
- Insert one Zotero field per citation cluster, not one field per number.
- Wrap the existing visible bibliography paragraphs in one Zotero bibliography field.
- Let Zotero Refresh perform the authoritative renumbering and bibliography reorder.

## 3. Resolve reference identities

Use this deterministic cascade:

1. Exact normalized DOI.
2. Exact PMID or ISBN.
3. Exact normalized title with compatible first author and year.
4. Reviewed override keyed by the source reference identity.

Normalization may lowercase, Unicode-normalize, remove DOI URL prefixes and terminal punctuation, decode HTML entities, and remove non-alphanumeric title characters. It must not merge substantively different works.

Do not auto-write on a title substring or fuzzy score alone. Put zero or multiple matches in `unresolved`/`ambiguous`, show the candidates, and require a reviewed override.

For live Zotero items store the library ID, item key, canonical URI, and current CSL JSON. If an item is absent:

- With explicit permission, import a verified RIS record, re-read the created item, and bind its real key/URI.
- Without permission, use a complete document-local embedded CSL item and label it `embedded`. Do not claim library linkage.

## 4. Construct fields

Let the packaged writer construct the Zotero fields and document preferences. Require one unique `citationID` per citation cluster, complete `citationItems` (`id`, URI, and `itemData`) in cluster order, exactly one bibliography field, and generated session/style preferences without hard-coded historical IDs or versions. Preserve cached marker text as native Word superscript until Zotero Refresh updates it.

## 5. Write locally and atomically

Run the fail-closed writer with a new output path:

```powershell
python scripts/migrate_docx_citations.py input.docx mapping.json migrated.docx --report migration_report.json
```

- Verify the source hash immediately before writing.
- Rewrite from a temporary DOCX in the destination directory, then atomically replace only the new output.
- Preserve original `ZipInfo` entries and order when practical.
- Allow only the story parts that contain migrated fields, `docProps/custom.xml`, and—only if newly required—`[Content_Types].xml` and `_rels/.rels` to change.
- Re-open the output ZIP, run CRC validation, and compare the package member set.

## 6. Refresh with Zotero

The pre-refresh validator proves structure, not integration behavior. On Windows, run the packaged refresh helper on another copy:

```powershell
powershell -ExecutionPolicy Bypass -File scripts/refresh_zotero_word.ps1 `
  -InputDocx migrated.docx `
  -OutputDocx refreshed.docx `
  -ReportPath zotero_refresh_report.json
```

The helper creates an isolated Word instance, proves ownership of its `WINWORD` process, checks that `Zotero.dotm` is loaded, runs `ZoteroRefresh`, waits for asynchronous updates (45 seconds by default), saves, and releases COM objects. Never kill unrelated Word processes. A successful macro return is not sufficient evidence: require a changed output on the first authoritative refresh, and record the saved state plus pre/post hashes.

Refresh proves that Zotero can parse and update the fields; an edit/save smoke test proves actual editability. Duplicate the refreshed candidate, use Zotero Edit Citation to change one citation with a known linked item, save, close, reopen, and validate the changed field. Record the tested occurrence and discard the smoke-test copy. If this interactive gate cannot be run, report it as untested rather than claiming full editability.

After refresh, citation numbers and bibliography order may change. Compare non-citation prose and object counts, not the old cached citation result text. Always compare the refreshed file with its exact pre-refresh candidate; an older manuscript that also differs in authorized prose will produce a misleading preservation failure.

If a dynamic bibliography creates a blank last page, inspect the field end marker and its trailing paragraph before changing layout. Compact only the Zotero-controlled bibliography tail, refresh again, and repeat structural and visual QA.

## 7. Deliver

Deliver only the separate refreshed DOCX after all gates pass. Retain the inventory, mapping, migration report, pre-refresh validation, refresh report, post-refresh validation, and render evidence as the audit trail unless the user asks for cleanup.

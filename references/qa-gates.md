# QA gates

## Gate A: reference map

- Zero unresolved references.
- Zero ambiguous matches.
- Zero conflicting duplicate DOI, title, or Zotero item identities.
- Every used reference is covered; out-of-range citation numbers are zero.
- Each exception is a reviewed override, not an unlogged code constant.

## Gate B: pre-refresh DOCX

- Source SHA-256 is unchanged and output is a separate file.
- ZIP CRC passes; no duplicate package members; required relationships and content types resolve.
- All complex fields have a valid begin/separate/end sequence in each story.
- Expected Zotero citation count and exactly one bibliography field are present.
- All CSL JSON parses; `citationID` values are unique.
- Every citation item has an `id`, URI, and `itemData`; linked and embedded counts match the map.
- No `ADDIN EN.CITE`, `ADDIN EN.CITE.DATA`, or `ADDIN EN.REFLIST` remains after an EndNote migration.
- Visible non-field text is byte-for-byte or normalization-equivalent to the frozen source, subject only to explicitly authorized changes.
- Paragraph, table/row/cell, drawing/image, OMML, section, comment, revision, footer-field, native super/subscript, and custom XML invariants hold.
- Unexpected package parts are byte-identical.

Run:

```powershell
python scripts/validate_zotero_docx.py migrated.docx --source input.docx --expect-citations 81 --out pre_refresh_validation.json
```

Replace `81` with the exact number of manifest occurrences.

## Gate C: live Zotero refresh

- The refresh helper created and owned a new `WINWORD` process.
- `Zotero.dotm` was loaded.
- `ZoteroRefresh` succeeded, the document saved, and field count did not unexpectedly fall.
- The first authoritative refresh changed the copied document hash; a macro return without a changed file is not accepted as proof.
- The refreshed document reopens read-only in Word.
- All citations remain editable fields; linked and embedded identities remain distinguishable.
- On a disposable copy, one citation was changed through Zotero, saved, reopened, and structurally revalidated; the disposable copy is not the delivery file.
- Zotero-controlled display renumbering and bibliography reordering are accepted only if coverage remains exact.

## Gate D: post-refresh structure

- Re-run `validate_zotero_docx.py` on the refreshed file with `--source` set to the exact pre-refresh candidate and `--post-refresh` enabled.
- Compare non-citation prose after excluding field results and the bibliography result.
- Use the exact pre-refresh candidate as the comparison source, not an earlier prose revision.
- Verify reference identity coverage rather than old display numbers.
- Recheck field balance, payload parsing, unique IDs, EndNote residue, footer fields, comments/revisions, OMML, media, and ZIP integrity.

## Gate E: visual QA

- Render the final refreshed DOCX with Microsoft Word or the documents renderer.
- Inspect every page at 100% zoom.
- Check citation placement and superscript, punctuation adjacency, bibliography indentation/wrapping, table citations, equations, figures/captions, headers/footers, missing glyphs, clipping, overlap, blank pages, and page-count changes.
- If visual rendering is unavailable, report that gate as not run; do not imply it passed.

## Evidence to retain

Keep the source/output hashes, inventory, mapping manifest, migration report, both validator reports, refresh report, page-count comparison, and renderer output directory. These are audit artifacts, not normal user deliverables unless requested.

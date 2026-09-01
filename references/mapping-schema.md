# Mapping schema

Use explicit, reviewable JSON rather than an in-memory-only map. The packaged writer consumes this manifest directly:

```json
{
  "schema_version": 1,
  "source": {
    "docx": "C:/absolute/input.docx",
    "sha256": "...",
    "citation_system": "endnote|static-numeric"
  },
  "preferences": {
    "style": "http://www.zotero.org/styles/journal-style-id",
    "zotero_version": "current-installed-version",
    "citation_format": "superscript",
    "prefs": {"fieldType": "Field", "delayCitationUpdates": "true"}
  },
  "references": [
    {
      "source_id": "ref-14",
      "number": 14,
      "visible_reference": "...",
      "identifiers": {"doi": "10.x/...", "isbn": null, "pmid": null},
      "match": {
        "status": "linked|embedded|unresolved|ambiguous",
        "method": "doi|isbn|pmid|title-author-year|reviewed-override",
        "reviewed": true,
        "reviewer": "name-or-agent-run-id",
        "reason": "exact DOI match",
        "evidence": ["visible-reference", "zotero-csl-json"],
        "candidates": []
      },
      "zotero": {
        "library_id": 123,
        "item_key": "ABCDEFGH",
        "uri": "http://zotero.org/users/123/items/ABCDEFGH"
      },
      "csl_item": {
        "id": "...",
        "type": "article-journal",
        "title": "...",
        "author": [{"family": "...", "given": "..."}],
        "issued": {"date-parts": [[2026]]}
      }
    }
  ],
  "occurrences": [
    {
      "occurrence_id": "word/document.xml:p42:m1",
      "part": "word/document.xml",
      "paragraph_index": 42,
      "kind": "static-marker|endnote-citation",
      "index_in_paragraph": 1,
      "source_text": "[14-16]",
      "reference_numbers": [14, 15, 16]
    }
  ],
  "bibliography": {
    "kind": "wrap-paragraph-range",
    "part": "word/document.xml",
    "start_paragraph_index": 300,
    "end_paragraph_index": 374,
    "start_text": "[1] ...",
    "end_text": "[75] ..."
  },
  "issues": {"unresolved": [], "ambiguous": [], "duplicates": []}
}
```

For EndNote input, replace the bibliography object with an exact locator:

```json
{
  "kind": "endnote-reflist",
  "part": "word/document.xml",
  "paragraph_index": 300,
  "index_in_paragraph": 1,
  "source_text": "..."
}
```

An occurrence must contain exactly one non-empty identity list: use `reference_numbers` when the frozen source has stable reference numbers, or `source_ids` when an EndNote field is identified directly. Use the canonical `zotero.uri` for linked items. For a document-local embedded item, omit `zotero` and provide one explicit `embedded_uri` such as `urn:document:reference:ref-14`.

## Invariants

- `source.sha256` must equal the file being migrated.
- Mixed citation systems are rejected. Do not include a historical Zotero session ID; the writer generates a fresh one.
- Every used reference number resolves to exactly one `source_id` and one CSL item.
- Every linked item has a real Zotero URI; every embedded item has complete CSL data and an explicit embedded status.
- The same DOI/item key must not map to conflicting source identities. Legitimate duplicate bibliography entries must be resolved deliberately before migration.
- `occurrence_id` is unique and stable for the frozen source. Prefer locator identity to display-text matching.
- Range expansion preserves citation-item order and removes accidental duplicates only when the source cluster itself repeats the same work.
- A reviewed override records the reason and reviewer; it is not a silent hard-coded exception.

## Minimal CSL data

At minimum retain type, title, creators, issued date, and the identifiers/containers needed by the selected style. Preserve verified volume, issue, pages/article number, publisher, DOI, URL, ISBN/ISSN, and language when available. Do not synthesize missing metadata merely to make a field look complete.

## Linked versus embedded

- `linked`: URI resolves to the intended item in the current Zotero library. The document field should survive Zotero Refresh as a library-linked citation.
- `embedded`: the field carries its own CSL item data and can remain editable in the document, but the item is not asserted to exist in the current library.
- `unresolved` or `ambiguous`: migration must stop.

# Zotero Word Citation Migration

A Codex skill for converting bracketed numeric or EndNote citations in complex Word documents into editable Zotero CSL citation and bibliography fields.

## What it does

- Audits citation systems across Word stories and blocks unsupported mixed systems.
- Builds an auditable, reviewed reference map before modifying OOXML.
- Creates a separate output DOCX with `ADDIN ZOTERO_ITEM CSL_CITATION` fields and one Zotero bibliography field.
- Validates the package, refreshes through Zotero in Word, proves edit/save behavior on a disposable copy, and performs page-level rendering checks.

## Install

```powershell
git clone https://github.com/Yupei-Han/zotero-word-citation-migration "$HOME/.agents/skills/zotero-word-citation-migration"
```

Restart Codex if needed, then invoke:

```text
$zotero-word-citation-migration convert this DOCX from static numeric citations to Zotero-editable fields without changing scientific content.
```

## Prerequisites

- Zotero Desktop with the relevant bibliographic records available for read-only lookup.
- Microsoft Word and the Zotero Word integration for authoritative refresh and edit/save verification on Windows.
- A source DOCX that must remain untouched.

## Important limits

- A superscript number is not a Zotero citation.
- Metadata-only records are not full-text evidence.
- The skill stops for unresolved or duplicate mappings, malformed CSL JSON, unsupported mixed citation systems, refresh failures, or incomplete visual QA.
- It does not authorize Zotero library writes, imports, or metadata edits.

## Maintainer check

```powershell
python scripts/self_test.py
```

## Repository layout

- `scripts/` — audit, migration, validation, and Word refresh helpers.
- `references/` — workflow, mapping schema, and QA gates.
- `agents/openai.yaml` — Codex presentation metadata.

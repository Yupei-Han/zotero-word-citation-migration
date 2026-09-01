#!/usr/bin/env python3
"""Minimal static-numeric and EndNote regression check for the packaged tools."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


SCRIPTS = Path(__file__).resolve().parent
MIGRATE = SCRIPTS / "migrate_docx_citations.py"
VALIDATE = SCRIPTS / "validate_zotero_docx.py"
CONTENT_TYPES = b"""<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"""
RELATIONSHIPS = b"""<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"""
STATIC_BODY = """<w:p><w:r><w:t>Synthetic prose [1].</w:t></w:r></w:p><w:p><w:r><w:t>References</w:t></w:r></w:p><w:p><w:r><w:t>[1] Tester, A. Synthetic work. 2026.</w:t></w:r></w:p>"""
ENDNOTE_BODY = """
<w:p><w:r><w:t>Synthetic prose </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ADDIN EN.CITE synthetic-record </w:instrText></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ADDIN EN.CITE.DATA captured-record </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>[1]</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t>References</w:t></w:r></w:p>
<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> ADDIN EN.REFLIST synthetic </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>[1] Tester, A. Synthetic work. 2026.</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>
"""


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def package(path: Path, body: str) -> None:
    document = f'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr/></w:body></w:document>'
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", RELATIONSHIPS)
        archive.writestr("word/document.xml", document)


def manifest(source: Path, system: str) -> dict:
    reference = {
        "source_id": "ref-1", "number": 1,
        "identifiers": {"doi": "10.0000/synthetic.1"},
        "match": {"status": "embedded", "method": "reviewed-override", "reviewed": True, "reason": "fixture", "reviewer": "self-test"},
        "embedded_uri": "urn:synthetic:reference:1",
        "csl_item": {"id": "synthetic-ref-1", "type": "article-journal", "title": "Synthetic work", "author": [{"family": "Tester"}], "issued": {"date-parts": [[2026]]}},
    }
    result = {
        "schema_version": 1,
        "source": {"sha256": sha256(source), "citation_system": system},
        "references": [reference],
        "issues": {"unresolved": [], "ambiguous": [], "duplicates": []},
        "preferences": {"style": "http://www.zotero.org/styles/nature", "zotero_version": "10.0.0", "citation_format": "superscript", "prefs": {"fieldType": "Field"}},
    }
    if system == "static-numeric":
        result["occurrences"] = [{"occurrence_id": "word/document.xml:p0:m1", "part": "word/document.xml", "paragraph_index": 0, "index_in_paragraph": 1, "kind": "static-marker", "source_text": "[1]", "reference_numbers": [1]}]
        result["bibliography"] = {"kind": "wrap-paragraph-range", "part": "word/document.xml", "start_paragraph_index": 2, "end_paragraph_index": 2, "start_text": "[1] Tester, A. Synthetic work. 2026.", "end_text": "[1] Tester, A. Synthetic work. 2026."}
    else:
        result["occurrences"] = [{"occurrence_id": "word/document.xml:p0:f1", "part": "word/document.xml", "paragraph_index": 0, "index_in_paragraph": 1, "kind": "endnote-citation", "source_text": "[1]", "reference_numbers": [1]}]
        result["bibliography"] = {"kind": "endnote-reflist", "part": "word/document.xml", "paragraph_index": 2, "index_in_paragraph": 1, "source_text": "[1] Tester, A. Synthetic work. 2026."}
    return result


def run(arguments: list[object], expected: int = 0) -> subprocess.CompletedProcess[str]:
    result = subprocess.run(
        [sys.executable, "-B", *map(str, arguments)],
        text=True, encoding="utf-8", capture_output=True,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
    )
    if result.returncode != expected:
        raise AssertionError(f"exit {result.returncode}:\n{result.stdout}\n{result.stderr}")
    return result


def main() -> None:
    with tempfile.TemporaryDirectory(prefix="zotero-word-citation-migration-") as temp:
        root = Path(temp)
        for system, body in (("static-numeric", STATIC_BODY), ("endnote", ENDNOTE_BODY)):
            source, mapping, output = (root / f"{system}{suffix}" for suffix in (".docx", ".json", "-migrated.docx"))
            package(source, body)
            source_hash = sha256(source)
            mapping.write_text(json.dumps(manifest(source, system)), encoding="utf-8")
            run([MIGRATE, source, mapping, output])
            validated = run([VALIDATE, output, "--source", source, "--expect-citations", 1])
            assert json.loads(validated.stdout)["status"] == "pass"
            assert sha256(source) == source_hash
            run([MIGRATE, source, mapping, output], expected=2)
        missing = run([VALIDATE, root / "static-numeric-migrated.docx", "--expect-citations", 1], expected=2)
        assert "--source" in missing.stderr
    print("PASS: static-numeric, EndNote, source-bound validation, no-overwrite")


if __name__ == "__main__":
    main()

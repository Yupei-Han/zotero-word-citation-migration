#!/usr/bin/env python3
"""Fail-closed structural validation for a Zotero-enabled DOCX."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path
from typing import Any

from lxml import etree

from audit_docx_citations import STORY_RE, audit, classify_code, qn


LINKED_URI_RE = re.compile(
    r"^https?://(?:www\.)?zotero\.org/(?:users|groups)/\d+/items/[A-Z0-9]{8}$",
    re.I,
)


ZOTERO_FIELD_KINDS = {"zotero-citation", "zotero-bibliography"}


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def _zotero_result_is_active(
    complex_fields: list[dict[str, Any]], simple_fields: list[dict[str, Any]]
) -> bool:
    return any(
        row["separated"] and row.get("kind") in ZOTERO_FIELD_KINDS
        for row in complex_fields
    ) or any(row["kind"] in ZOTERO_FIELD_KINDS for row in simple_fields)


def canonical_story_integrity(raw: bytes) -> dict[str, Any]:
    """Return non-Zotero visible text and non-Zotero field inventory.

    The visible stream ignores run boundaries and empty paragraphs because Word
    may split runs or add an empty bibliography-tail paragraph during refresh.
    Zotero citation and bibliography results are excluded. Text boxes are covered
    because their paragraphs are descendants of the containing story part.
    """

    root = etree.fromstring(
        raw, etree.XMLParser(remove_blank_text=False, huge_tree=True)
    )
    complex_fields: list[dict[str, Any]] = []
    simple_fields: list[dict[str, Any]] = []
    visible_chunks: list[str] = []
    inventory: Counter[tuple[str, str, str]] = Counter()

    for event, node in etree.iterwalk(root, events=("start", "end")):
        if event == "start":
            if node.tag == qn("fldSimple"):
                code = (node.get(qn("instr"), "") or "").strip()
                kind = classify_code(code)
                excluded_by_outer = _zotero_result_is_active(
                    complex_fields, simple_fields
                )
                simple_fields.append(
                    {"kind": kind, "excluded_by_outer": excluded_by_outer}
                )
                if kind not in ZOTERO_FIELD_KINDS and not excluded_by_outer:
                    inventory[("simple", kind, sha256_text(code))] += 1
                continue

            if node.tag == qn("fldChar"):
                field_type = node.get(qn("fldCharType"), "")
                if field_type == "begin":
                    complex_fields.append(
                        {
                            "code_chunks": [],
                            "separated": False,
                            "kind": None,
                            "excluded_by_outer": _zotero_result_is_active(
                                complex_fields, simple_fields
                            ),
                        }
                    )
                elif field_type == "separate" and complex_fields:
                    row = complex_fields[-1]
                    row["separated"] = True
                    row["kind"] = classify_code(
                        "".join(row["code_chunks"]).strip()
                    )
                elif field_type == "end" and complex_fields:
                    row = complex_fields.pop()
                    code = "".join(row["code_chunks"]).strip()
                    kind = row.get("kind") or classify_code(code)
                    if (
                        kind not in ZOTERO_FIELD_KINDS
                        and not row["excluded_by_outer"]
                    ):
                        inventory[("complex", kind, sha256_text(code))] += 1
                continue

            if (
                node.tag == qn("instrText")
                and complex_fields
                and not complex_fields[-1]["separated"]
            ):
                complex_fields[-1]["code_chunks"].append(node.text or "")
                continue

            if _zotero_result_is_active(complex_fields, simple_fields):
                continue
            if node.tag in {qn("t"), qn("delText")} and node.text:
                visible_chunks.append(node.text)
            elif node.tag == qn("tab"):
                visible_chunks.append("\t")
            elif node.tag in {qn("br"), qn("cr")}:
                visible_chunks.append("\n")
            elif node.tag == qn("noBreakHyphen"):
                visible_chunks.append("\u2011")
            elif node.tag == qn("softHyphen"):
                visible_chunks.append("\u00ad")
            elif node.tag == qn("sym"):
                visible_chunks.append(
                    "\ufff0SYM:"
                    + (node.get(qn("font"), "") or "")
                    + ":"
                    + (node.get(qn("char"), "") or "")
                    + "\ufff1"
                )
        elif node.tag == qn("fldSimple") and simple_fields:
            simple_fields.pop()

    visible_text = "".join(visible_chunks)
    inventory_rows = [
        {
            "field_form": field_form,
            "kind": kind,
            "code_sha256": code_hash,
            "count": count,
        }
        for (field_form, kind, code_hash), count in sorted(inventory.items())
    ]
    return {
        "non_zotero_visible_text_sha256": sha256_text(visible_text),
        "non_zotero_visible_text_characters": len(visible_text),
        "non_zotero_field_inventory": inventory_rows,
    }


def compare_post_refresh_stories(
    source_parts: dict[str, bytes], output_parts: dict[str, bytes]
) -> tuple[list[str], dict[str, Any]]:
    errors: list[str] = []
    details: dict[str, Any] = {}
    source_stories = {name for name in source_parts if STORY_RE.fullmatch(name)}
    output_stories = {name for name in output_parts if STORY_RE.fullmatch(name)}
    if source_stories != output_stories:
        errors.append(
            "post-refresh Word story set changed; "
            f"added={sorted(output_stories - source_stories)}, "
            f"removed={sorted(source_stories - output_stories)}"
        )

    for part in sorted(source_stories & output_stories):
        try:
            before = canonical_story_integrity(source_parts[part])
            after = canonical_story_integrity(output_parts[part])
        except etree.XMLSyntaxError as exc:
            errors.append(f"post-refresh story comparison failed for {part}: {exc}")
            continue
        text_match = (
            before["non_zotero_visible_text_sha256"]
            == after["non_zotero_visible_text_sha256"]
        )
        fields_match = (
            before["non_zotero_field_inventory"]
            == after["non_zotero_field_inventory"]
        )
        details[part] = {
            "source": before,
            "output": after,
            "non_zotero_visible_text_match": text_match,
            "non_zotero_field_inventory_match": fields_match,
        }
        if not text_match:
            errors.append(
                f"post-refresh non-Zotero visible text changed in {part}"
            )
        if not fields_match:
            errors.append(
                f"post-refresh non-Zotero field instructions changed in {part}"
            )
    return errors, details


def read_package(path: Path) -> tuple[list[str], dict[str, bytes]]:
    with zipfile.ZipFile(path) as archive:
        names = archive.namelist()
        return names, {name: archive.read(name) for name in names}


def zotero_preferences(parts: dict[str, bytes]) -> tuple[dict[str, Any] | None, str | None]:
    raw = parts.get("docProps/custom.xml")
    if raw is None:
        return None, "docProps/custom.xml is missing"
    try:
        root = etree.fromstring(raw)
    except etree.XMLSyntaxError as exc:
        return None, f"custom properties XML is invalid: {exc}"
    rows: list[tuple[int, str]] = []
    for prop in root:
        name = prop.get("name") or ""
        match = re.fullmatch(r"ZOTERO_PREF_(\d+)", name)
        if not match:
            continue
        value = "".join(prop.itertext())
        rows.append((int(match.group(1)), value))
    rows.sort()
    if not rows:
        return None, "ZOTERO_PREF_* properties are missing"
    expected = list(range(1, len(rows) + 1))
    actual = [number for number, _ in rows]
    if actual != expected:
        return None, f"ZOTERO_PREF_* chunks are non-sequential: {actual}"
    try:
        data = etree.fromstring("".join(value for _, value in rows).encode("utf-8"))
    except etree.XMLSyntaxError as exc:
        return None, f"Zotero document preferences are invalid XML: {exc}"
    session = data.find("session")
    style = data.find("style")
    prefs = {
        node.get("name"): node.get("value")
        for node in data.xpath("./prefs/pref")
        if node.get("name")
    }
    report = {
        "chunks": len(rows),
        "data_version": data.get("data-version"),
        "zotero_version": data.get("zotero-version"),
        "session_id": session.get("id") if session is not None else None,
        "style": style.get("id") if style is not None else None,
        "has_bibliography": style.get("hasBibliography") if style is not None else None,
        "preferences": prefs,
    }
    if not report["session_id"]:
        return report, "Zotero session ID is missing"
    if not report["style"]:
        return report, "Zotero CSL style is missing"
    if prefs.get("fieldType") != "Field":
        return report, "Zotero fieldType is not Field"
    return report, None


def package_diff(
    source_parts: dict[str, bytes],
    output_parts: dict[str, bytes],
    *,
    post_refresh: bool,
) -> tuple[list[str], list[str]]:
    source_names = set(source_parts)
    output_names = set(output_parts)
    errors: list[str] = []
    warnings: list[str] = []
    if source_names != output_names:
        allowed_added = {"docProps/custom.xml"}
        added = output_names - source_names
        removed = source_names - output_names
        if removed or not added.issubset(allowed_added):
            errors.append(
                f"package member set changed unexpectedly; added={sorted(added)}, "
                f"removed={sorted(removed)}"
            )
    changed = sorted(
        name
        for name in source_names & output_names
        if source_parts[name] != output_parts[name]
    )
    if not post_refresh:
        allowed = {
            "[Content_Types].xml",
            "_rels/.rels",
            "docProps/custom.xml",
        }
        allowed.update(
            name
            for name in source_names | output_names
            if STORY_RE.fullmatch(name)
        )
        unexpected = sorted(set(changed) - allowed)
        if unexpected:
            errors.append(f"unexpected pre-refresh changed parts: {unexpected}")
    else:
        warnings.append(f"post-refresh changed parts: {changed}")
    return errors, warnings


def validate(
    path: Path,
    *,
    source: Path,
    expected_citations: int,
    post_refresh: bool,
) -> dict[str, Any]:
    report = audit(path)
    errors: list[str] = []
    warnings: list[str] = []
    source_report = audit(source)
    if source.resolve() == path.resolve():
        errors.append("source and output paths are the same")

    if report["zip"]["bad_member"]:
        errors.append(f"ZIP CRC failed at {report['zip']['bad_member']}")
    if report["zip"]["duplicate_names"]:
        errors.append(f"duplicate ZIP members: {report['zip']['duplicate_names']}")
    for category in ("xml", "field_sequence"):
        errors.extend(report["issues"][category])

    field_counts = report["field_counts"]
    citation_count = int(field_counts.get("zotero-citation", 0))
    bibliography_count = int(field_counts.get("zotero-bibliography", 0))
    if citation_count <= 0:
        errors.append("no Zotero citation fields were found")

    if expected_citations <= 0:
        errors.append("--expect-citations must be a positive integer")
    elif citation_count != expected_citations:
        errors.append(
            f"expected {expected_citations} Zotero citation fields, "
            f"found {citation_count}"
        )
    if bibliography_count != 1:
        errors.append(
            f"expected exactly one Zotero bibliography field, found {bibliography_count}"
        )
    endnote_count = sum(
        int(field_counts.get(name, 0))
        for name in ("endnote-citation", "endnote-data", "endnote-bibliography")
    )
    if endnote_count:
        errors.append(f"{endnote_count} EndNote fields remain")
    if report["citation_payload_errors"]:
        errors.append(
            f"{len(report['citation_payload_errors'])} Zotero citation payload(s) are invalid"
        )
    if report["duplicate_citation_ids"]:
        errors.append(
            f"duplicate citationID values: {report['duplicate_citation_ids']}"
        )

    linked_item_occurrences = 0
    embedded_item_occurrences = 0
    citation_ids: list[str] = []
    item_errors: list[str] = []
    item_warnings: list[str] = []
    citation_fields = [
        row for row in report["fields"] if row["kind"] == "zotero-citation"
    ]
    for field in citation_fields:
        occurrence_id = field["occurrence_id"]
        if not field["has_separator"]:
            errors.append(f"{occurrence_id} has no field separator")
        citation_id = field.get("citation_id")
        if not citation_id:
            errors.append(f"{occurrence_id} has no citationID")
        else:
            citation_ids.append(str(citation_id))
        items = field.get("citation_items")
        if not isinstance(items, list) or not items:
            item_errors.append(f"{occurrence_id} has no citationItems")
            continue
        for index, item in enumerate(items, 1):
            label = f"{occurrence_id}/item{index}"
            if item.get("id") in (None, ""):
                item_errors.append(f"{label} has no id")
            uris = item.get("uris")
            if (
                not isinstance(uris, list)
                or not uris
                or not all(isinstance(uri, str) and uri for uri in uris)
            ):
                item_errors.append(f"{label} has no valid URI")
                uris = []
            if not item.get("has_item_data"):
                item_errors.append(f"{label} has no itemData")
            if not item.get("type"):
                item_warnings.append(f"{label} itemData has no type")
            if not item.get("title"):
                item_warnings.append(f"{label} itemData has no title")
            if any(LINKED_URI_RE.fullmatch(uri) for uri in uris):
                linked_item_occurrences += 1
            else:
                embedded_item_occurrences += 1
    errors.extend(item_errors)
    warnings.extend(item_warnings)
    if len(citation_ids) != citation_count:
        errors.append(
            f"citationID coverage differs from citation field count: "
            f"{len(citation_ids)} vs {citation_count}"
        )

    names, parts = read_package(path)
    prefs, prefs_error = zotero_preferences(parts)
    if prefs_error:
        errors.append(prefs_error)

    changed_parts: list[str] = []
    post_refresh_story_integrity: dict[str, Any] | None = None
    source_names, source_parts = read_package(source)
    diff_errors, diff_warnings = package_diff(
        source_parts, parts, post_refresh=post_refresh
    )
    errors.extend(diff_errors)
    warnings.extend(diff_warnings)
    changed_parts = sorted(
        name
        for name in set(source_names) & set(names)
        if source_parts[name] != parts[name]
    )
    if (
        source_report["prose_noncitation_text_sha256"]
        != report["prose_noncitation_text_sha256"]
    ):
        errors.append("non-citation prose before References changed")
    structure_keys = (
        "tables",
        "table_rows",
        "table_cells",
        "drawings",
        "math_objects",
        "math_paragraphs",
        "sections",
        "revisions",
        "comment_anchors",
    )
    for key in structure_keys:
        before = source_report["structure"][key]
        after = report["structure"][key]
        if before != after:
            errors.append(f"structure count {key} changed: {before} -> {after}")
    if source_report["media_hashes"] != report["media_hashes"]:
        errors.append("media part names or bytes changed")
    if source_report["custom_xml_hashes"] != report["custom_xml_hashes"]:
        errors.append("customXml part names or bytes changed")
    if not post_refresh:
        source_story = {
            row["part"]: row["visible_text_sha256"]
            for row in source_report["story_parts"]
        }
        output_story = {
            row["part"]: row["visible_text_sha256"]
            for row in report["story_parts"]
        }
        if source_story != output_story:
            errors.append("pre-refresh visible story text changed")
        for key in ("document_paragraphs", "direct_body_paragraphs"):
            before = source_report["structure"][key]
            after = report["structure"][key]
            if before != after:
                errors.append(
                    f"pre-refresh structure count {key} changed: "
                    f"{before} -> {after}"
                )
    else:
        story_errors, post_refresh_story_integrity = compare_post_refresh_stories(
            source_parts, parts
        )
        errors.extend(story_errors)
        for key in ("document_paragraphs", "direct_body_paragraphs"):
            before = source_report["structure"][key]
            after = report["structure"][key]
            if before != after:
                warnings.append(
                    f"post-refresh {key} changed: {before} -> {after}; "
                    "verify bibliography tail and pagination"
                )

    result = {
        "schema_version": 1,
        "status": "pass" if not errors else "fail",
        "input": str(path.resolve()),
        "source": str(source.resolve()),
        "post_refresh": post_refresh,
        "expected_citation_fields": expected_citations,
        "expected_citation_count_basis": "explicit",
        "source_citation_classification": source_report["classification"],
        "citation_fields": citation_count,
        "bibliography_fields": bibliography_count,
        "citation_ids_unique": len(set(citation_ids)),
        "citation_item_occurrences": linked_item_occurrences
        + embedded_item_occurrences,
        "linked_item_occurrences": linked_item_occurrences,
        "embedded_item_occurrences": embedded_item_occurrences,
        "field_counts": field_counts,
        "field_characters_by_part": report["field_characters_by_part"],
        "references": {
            "visible_entries": report["reference_count"],
            "used_numbers": report["used_reference_numbers"],
            "unused_numbers": report["unused_reference_numbers"],
            "out_of_range_numbers": report["out_of_range_reference_numbers"],
        },
        "zotero_preferences": prefs,
        "structure": report["structure"],
        "source_sha256": source_report["sha256"],
        "output_sha256": report["sha256"],
        "changed_parts": changed_parts,
        "post_refresh_story_integrity": post_refresh_story_integrity,
        "errors": errors,
        "warnings": warnings,
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_docx", type=Path)
    parser.add_argument(
        "--source",
        type=Path,
        required=True,
        help="Exact frozen source/candidate for content and package fidelity comparison.",
    )
    parser.add_argument(
        "--expect-citations",
        type=int,
        required=True,
        help="Required positive citation-field count from the reviewed manifest.",
    )
    parser.add_argument(
        "--post-refresh",
        action="store_true",
        help=(
            "Treat --source as the exact pre-refresh candidate and compare every "
            "Word story outside Zotero-controlled field results."
        ),
    )
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()
    if args.expect_citations <= 0:
        parser.error("--expect-citations must be a positive integer")

    result = validate(
        args.input_docx,
        source=args.source,
        expected_citations=args.expect_citations,
        post_refresh=args.post_refresh,
    )
    rendered = json.dumps(result, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(rendered + "\n", encoding="utf-8")
    print(rendered)
    return 0 if result["status"] == "pass" else 2


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())

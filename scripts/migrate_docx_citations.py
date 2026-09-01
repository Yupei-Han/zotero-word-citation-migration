#!/usr/bin/env python3
"""Manifest-driven, fail-closed migration of Word citations to Zotero fields.

Supported source branches:

* ``static-numeric``: bracketed numeric markers such as ``[2,4-6]`` in direct
  paragraph runs. The manifest must identify every marker and a contiguous,
  direct-body bibliography paragraph range.
* ``endnote``: complex outer ``ADDIN EN.CITE`` fields. The manifest must
  identify every outer field. Captured ``ADDIN EN.CITE.DATA`` fields are
  removed and one explicitly located ``ADDIN EN.REFLIST`` field is converted.

The writer deliberately rejects unsupported structures instead of guessing.
It does not query, import, or modify a Zotero library. Its output is a
pre-refresh DOCX that still requires structural validation, Zotero Refresh,
post-refresh validation, and page-by-page visual QA.
"""

from __future__ import annotations

import argparse
import copy
import json
import os
import re
import sys
import tempfile
import uuid
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from lxml import etree

from audit_docx_citations import (
    FieldRecord,
    STORY_RE,
    W,
    ZOTERO_CITATION,
    paragraph_of,
    parse_field_records,
    sha256_file,
    sha256_text,
    text_outside_citation_fields,
    visible_text,
)


XML = "http://www.w3.org/XML/1998/namespace"
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CSL_SCHEMA = (
    "https://github.com/citation-style-language/schema/raw/master/"
    "csl-citation.json"
)
BIBLIOGRAPHY_CODE = (
    'ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]}'
    " CSL_BIBLIOGRAPHY"
)
STATIC_MARKER_RE = re.compile(
    r"\[(?P<body>\d+(?:(?:\s*[,;]\s*|\s*[-\u2013\u2014]\s*)\d+)*)\]"
)
FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
CUSTOM_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/custom-properties"
)
CUSTOM_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
)


class MigrationError(RuntimeError):
    """A fail-closed migration blocker."""


def qn(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def run_of(node: etree._Element) -> etree._Element | None:
    current: etree._Element | None = node
    while current is not None:
        if current.tag == qn(W, "r"):
            return current
        current = current.getparent()
    return None


def parse_numeric_cluster(marker: str) -> list[int]:
    match = STATIC_MARKER_RE.fullmatch(marker.strip())
    if not match:
        return []
    values: list[int] = []
    for token in re.split(r"\s*[,;]\s*", match.group("body")):
        range_match = re.fullmatch(r"(\d+)\s*[-\u2013\u2014]\s*(\d+)", token)
        if range_match:
            left, right = map(int, range_match.groups())
            if right < left:
                return []
            values.extend(range(left, right + 1))
        elif token.isdigit():
            values.append(int(token))
        else:
            return []
    return values


def normalize_doi(value: Any) -> str | None:
    if not isinstance(value, str) or not value.strip():
        return None
    result = value.strip().lower()
    result = re.sub(r"^(?:https?://(?:dx\.)?doi\.org/|doi:\s*)", "", result)
    return result.rstrip(".,;)")


def require_nonempty_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise MigrationError(f"{label} must be a non-empty string")
    return value.strip()


def reject_keys(value: dict[str, Any], label: str, *keys: str) -> None:
    found = sorted(set(value).intersection(keys))
    if found:
        raise MigrationError(f"{label} contains unsupported keys: {found}")


def validate_csl_item(item: Any, label: str) -> dict[str, Any]:
    if not isinstance(item, dict):
        raise MigrationError(f"{label}.csl_item must be an object")
    result = copy.deepcopy(item)
    require_nonempty_string(result.get("type"), f"{label}.csl_item.type")
    require_nonempty_string(result.get("title"), f"{label}.csl_item.title")
    creators = []
    for key in ("author", "editor", "container-author"):
        value = result.get(key)
        if isinstance(value, list):
            creators.extend(value)
    if not creators or not all(isinstance(value, dict) and value for value in creators):
        raise MigrationError(
            f"{label}.csl_item must contain at least one complete creator"
        )
    issued = result.get("issued")
    if not isinstance(issued, dict) or not isinstance(issued.get("date-parts"), list):
        raise MigrationError(f"{label}.csl_item.issued.date-parts is required")
    date_parts = issued["date-parts"]
    if not date_parts or not isinstance(date_parts[0], list) or not date_parts[0]:
        raise MigrationError(f"{label}.csl_item.issued.date-parts is empty")
    if not isinstance(date_parts[0][0], int):
        raise MigrationError(f"{label}.csl_item issued year must be an integer")
    item_id = result.get("id")
    if item_id in (None, "") or isinstance(item_id, (dict, list, bool)):
        raise MigrationError(f"{label}.csl_item.id is required")
    return result


@dataclass(frozen=True)
class PreparedReference:
    source_id: str
    number: int | None
    status: str
    item_id: str | int
    uris: tuple[str, ...]
    item_data: dict[str, Any]

    def citation_item(self) -> dict[str, Any]:
        return {
            "id": self.item_id,
            "uris": list(self.uris),
            "itemData": copy.deepcopy(self.item_data),
        }


def prepare_references(
    manifest: dict[str, Any],
) -> tuple[list[PreparedReference], dict[int, PreparedReference], dict[str, PreparedReference]]:
    rows = manifest.get("references")
    if not isinstance(rows, list) or not rows:
        raise MigrationError("manifest.references must be a non-empty array")
    prepared: list[PreparedReference] = []
    by_number: dict[int, PreparedReference] = {}
    by_source_id: dict[str, PreparedReference] = {}
    seen_targets: dict[str, str] = {}

    for index, row in enumerate(rows, 1):
        label = f"references[{index - 1}]"
        if not isinstance(row, dict):
            raise MigrationError(f"{label} must be an object")
        source_id = require_nonempty_string(row.get("source_id"), f"{label}.source_id")
        if source_id in by_source_id:
            raise MigrationError(f"duplicate source_id {source_id!r}")
        number = row.get("number")
        if number is not None and (not isinstance(number, int) or isinstance(number, bool) or number < 1):
            raise MigrationError(f"{label}.number must be a positive integer or null")
        if number is not None and number in by_number:
            raise MigrationError(f"duplicate reference number {number}")

        match = row.get("match")
        if not isinstance(match, dict):
            raise MigrationError(f"{label}.match must be an object")
        status = match.get("status")
        if status in {"unresolved", "ambiguous"}:
            raise MigrationError(f"{label} has blocking match status {status!r}")
        if status not in {"linked", "embedded"}:
            raise MigrationError(f"{label}.match.status must be linked or embedded")
        if match.get("reviewed") is not True:
            raise MigrationError(f"{label} mapping is not explicitly reviewed")
        method = require_nonempty_string(match.get("method"), f"{label}.match.method")
        if method == "reviewed-override":
            require_nonempty_string(match.get("reason"), f"{label}.match.reason")
            require_nonempty_string(match.get("reviewer"), f"{label}.match.reviewer")

        item_data = validate_csl_item(row.get("csl_item"), label)
        item_id = item_data["id"]
        zotero = row.get("zotero")
        if zotero is None:
            zotero = {}
        if not isinstance(zotero, dict):
            raise MigrationError(f"{label}.zotero must be an object")
        reject_keys(zotero, f"{label}.zotero", "uris", "item_id")
        uri_values = [
            value
            for value in (zotero.get("uri"), row.get("embedded_uri"))
            if value is not None
        ]
        uris: list[str] = []
        for uri in uri_values:
            value = require_nonempty_string(uri, f"{label} URI")
            if value not in uris:
                uris.append(value)
        if not uris:
            raise MigrationError(f"{label} must provide at least one explicit URI")
        item_key = zotero.get("item_key")
        if item_key is not None:
            item_key = require_nonempty_string(item_key, f"{label}.zotero.item_key")
        if status == "linked":
            if not item_key:
                raise MigrationError(f"{label} linked item lacks zotero.item_key")
            if not any(
                re.search(rf"/(?:items/)?{re.escape(item_key)}(?:$|[/?#])", uri, re.I)
                for uri in uris
            ):
                raise MigrationError(
                    f"{label} linked URI does not contain zotero.item_key"
                )

        identifiers = row.get("identifiers") or {}
        if not isinstance(identifiers, dict):
            raise MigrationError(f"{label}.identifiers must be an object")
        doi = normalize_doi(identifiers.get("doi") or item_data.get("DOI"))
        keys = [f"id:{item_id}"]
        keys.extend(f"uri:{uri.lower()}" for uri in uris)
        if item_key:
            keys.append(f"key:{item_key.upper()}")
        if doi:
            keys.append(f"doi:{doi}")
        for target in keys:
            prior = seen_targets.get(target)
            if prior is not None and prior != source_id:
                raise MigrationError(
                    f"duplicate target identity {target!r} for {prior!r} and {source_id!r}"
                )
            seen_targets[target] = source_id

        item_data["id"] = item_id
        reference = PreparedReference(
            source_id=source_id,
            number=number,
            status=status,
            item_id=item_id,
            uris=tuple(uris),
            item_data=item_data,
        )
        prepared.append(reference)
        by_source_id[source_id] = reference
        if number is not None:
            by_number[number] = reference

    return prepared, by_number, by_source_id


def validate_manifest_issues(manifest: dict[str, Any]) -> None:
    issues = manifest.get("issues", {})
    if not isinstance(issues, dict):
        raise MigrationError("manifest.issues must be an object")
    for key in ("unresolved", "ambiguous", "duplicates"):
        if issues.get(key):
            raise MigrationError(f"manifest.issues.{key} is not empty")


def resolve_occurrence_references(
    occurrence: dict[str, Any],
    by_number: dict[int, PreparedReference],
    by_source_id: dict[str, PreparedReference],
    label: str,
) -> list[PreparedReference]:
    has_numbers = "reference_numbers" in occurrence
    has_source_ids = "source_ids" in occurrence
    if has_numbers == has_source_ids:
        raise MigrationError(
            f"{label} must contain exactly one of reference_numbers or source_ids"
        )
    result: list[PreparedReference] = []
    if has_numbers:
        values = occurrence["reference_numbers"]
        if not isinstance(values, list) or not values:
            raise MigrationError(f"{label}.reference_numbers must be non-empty")
        for value in values:
            if not isinstance(value, int) or isinstance(value, bool):
                raise MigrationError(f"{label}.reference_numbers must be integers")
            if value not in by_number:
                raise MigrationError(f"{label} uses unmapped reference number {value}")
            result.append(by_number[value])
    else:
        values = occurrence["source_ids"]
        if not isinstance(values, list) or not values:
            raise MigrationError(f"{label}.source_ids must be non-empty")
        for value in values:
            if value not in by_source_id:
                raise MigrationError(f"{label} uses unmapped source_id {value!r}")
            result.append(by_source_id[value])
    target_ids = [str(value.item_id) for value in result]
    if len(target_ids) != len(set(target_ids)):
        raise MigrationError(f"{label} repeats the same target item in one cluster")
    return result


@dataclass
class PreparedOccurrence:
    source: dict[str, Any]
    occurrence_id: str
    part: str
    paragraph_index: int
    index_in_paragraph: int
    kind: str
    source_text: str
    references: list[PreparedReference]
    note_index: int
    citation_id: str = ""


def prepare_occurrences(
    manifest: dict[str, Any],
    source_system: str,
    by_number: dict[int, PreparedReference],
    by_source_id: dict[str, PreparedReference],
) -> list[PreparedOccurrence]:
    rows = manifest.get("occurrences")
    if not isinstance(rows, list) or not rows:
        raise MigrationError("manifest.occurrences must be a non-empty array")
    prepared: list[PreparedOccurrence] = []
    seen_ids: set[str] = set()
    seen_locators: set[tuple[str, int, int]] = set()
    allowed_kinds = (
        {"static-marker"}
        if source_system == "static-numeric"
        else {"endnote-citation"}
    )
    for index, row in enumerate(rows):
        label = f"occurrences[{index}]"
        if not isinstance(row, dict):
            raise MigrationError(f"{label} must be an object")
        reject_keys(row, label, "result")
        occurrence_id = require_nonempty_string(
            row.get("occurrence_id"), f"{label}.occurrence_id"
        )
        if occurrence_id in seen_ids:
            raise MigrationError(f"duplicate occurrence_id {occurrence_id!r}")
        seen_ids.add(occurrence_id)
        part = require_nonempty_string(row.get("part"), f"{label}.part")
        if not STORY_RE.fullmatch(part):
            raise MigrationError(f"{label}.part is not a supported Word story part")
        paragraph_index = row.get("paragraph_index")
        index_in_paragraph = row.get("index_in_paragraph")
        if not isinstance(paragraph_index, int) or isinstance(paragraph_index, bool) or paragraph_index < 0:
            raise MigrationError(f"{label}.paragraph_index must be a non-negative integer")
        if not isinstance(index_in_paragraph, int) or isinstance(index_in_paragraph, bool) or index_in_paragraph < 1:
            raise MigrationError(f"{label}.index_in_paragraph must be a 1-based integer")
        locator = (part, paragraph_index, index_in_paragraph)
        if locator in seen_locators:
            raise MigrationError(f"duplicate occurrence locator {locator!r}")
        seen_locators.add(locator)
        kind = require_nonempty_string(row.get("kind"), f"{label}.kind")
        if kind not in allowed_kinds:
            raise MigrationError(
                f"{label}.kind {kind!r} is incompatible with {source_system!r}"
            )
        source_text_value = row.get("source_text")
        if not isinstance(source_text_value, str) or not source_text_value:
            raise MigrationError(f"{label}.source_text must be a non-empty string")
        references = resolve_occurrence_references(
            row, by_number, by_source_id, label
        )
        note_index = row.get("note_index", 0)
        if not isinstance(note_index, int) or isinstance(note_index, bool) or note_index < 0:
            raise MigrationError(f"{label}.note_index must be a non-negative integer")
        prepared.append(
            PreparedOccurrence(
                source=row,
                occurrence_id=occurrence_id,
                part=part,
                paragraph_index=paragraph_index,
                index_in_paragraph=index_in_paragraph,
                kind=kind,
                source_text=source_text_value,
                references=references,
                note_index=note_index,
            )
        )
    return prepared


def rtf_escape(value: str) -> str:
    escaped: list[str] = []
    for char in value:
        codepoint = ord(char)
        if char in {"\\", "{", "}"}:
            escaped.append("\\" + char)
        elif codepoint < 128:
            escaped.append(char)
        else:
            if codepoint <= 0xFFFF:
                units = [codepoint]
            else:
                adjusted = codepoint - 0x10000
                units = [0xD800 + (adjusted >> 10), 0xDC00 + (adjusted & 0x3FF)]
            for unit in units:
                signed = unit if unit < 0x8000 else unit - 0x10000
                escaped.append(f"\\uc0\\u{signed}{{}}")
    return "".join(escaped)


def citation_payload(
    occurrence: PreparedOccurrence, citation_format: str
) -> dict[str, Any]:
    formatted = rtf_escape(occurrence.source_text)
    if citation_format == "superscript":
        formatted = "\\super " + formatted + "\\nosupersub{}"
    return {
        "citationID": occurrence.citation_id,
        "properties": {
            "formattedCitation": formatted,
            "plainCitation": occurrence.source_text,
            "noteIndex": occurrence.note_index,
        },
        "citationItems": [value.citation_item() for value in occurrence.references],
        "schema": CSL_SCHEMA,
    }


def citation_code(occurrence: PreparedOccurrence, citation_format: str) -> str:
    return ZOTERO_CITATION + " " + json.dumps(
        citation_payload(occurrence, citation_format),
        ensure_ascii=False,
        separators=(",", ":"),
    )


def sanitized_rpr(template: etree._Element) -> etree._Element | None:
    rpr = template.find(qn(W, "rPr"))
    if rpr is None:
        return None
    result = copy.deepcopy(rpr)
    for vert in result.findall(qn(W, "vertAlign")):
        result.remove(vert)
    return result


def new_run(template: etree._Element | None) -> etree._Element:
    return etree.Element(qn(W, "r"), attrib=dict(template.attrib) if template is not None else {})


def make_text_run(
    template: etree._Element, text: str, *, preserve_vertical: bool
) -> etree._Element:
    run = new_run(template)
    rpr = template.find(qn(W, "rPr")) if preserve_vertical else sanitized_rpr(template)
    if rpr is not None:
        run.append(copy.deepcopy(rpr))
    text_node = etree.SubElement(run, qn(W, "t"))
    if text.startswith(" ") or text.endswith(" ") or "  " in text:
        text_node.set(qn(XML, "space"), "preserve")
    text_node.text = text
    return run


def make_fldchar_run(template: etree._Element | None, kind: str) -> etree._Element:
    run = new_run(template)
    if template is not None:
        rpr = sanitized_rpr(template)
        if rpr is not None:
            run.append(rpr)
    node = etree.SubElement(run, qn(W, "fldChar"))
    node.set(qn(W, "fldCharType"), kind)
    return run


def make_instr_run(template: etree._Element | None, code: str) -> etree._Element:
    run = new_run(template)
    if template is not None:
        rpr = sanitized_rpr(template)
        if rpr is not None:
            run.append(rpr)
    node = etree.SubElement(run, qn(W, "instrText"))
    node.set(qn(XML, "space"), "preserve")
    node.text = f" {code} "
    return run


def replace_static_marker(
    paragraph: etree._Element,
    start: int,
    end: int,
    marker: str,
    code: str,
) -> None:
    runs = [child for child in paragraph if child.tag == qn(W, "r")]
    spans: list[tuple[etree._Element, int, int, str]] = []
    cursor = 0
    for run in runs:
        value = visible_text(run)
        spans.append((run, cursor, cursor + len(value), value))
        cursor += len(value)
    if "".join(value for _, _, _, value in spans) != visible_text(paragraph):
        raise MigrationError(
            "static marker paragraph contains text outside direct runs; unsupported"
        )
    affected = [row for row in spans if start < row[2] and end > row[1]]
    if not affected:
        raise MigrationError(f"static marker {marker!r} is not in direct runs")
    positions = [paragraph.index(row[0]) for row in affected]
    position = positions[0]
    interval = list(paragraph)[position : positions[-1] + 1]
    unsupported = [
        child.tag
        for child in interval
        if child.tag not in {qn(W, "r"), qn(W, "proofErr")}
    ]
    if unsupported:
        raise MigrationError(
            f"static marker {marker!r} crosses unsupported children {unsupported}"
        )

    result_by_run: dict[int, etree._Element] = {}
    reconstructed: list[str] = []
    for run, run_start, run_end, value in affected:
        run_unsupported = [
            child.tag
            for child in run
            if child.tag not in {qn(W, "rPr"), qn(W, "t")}
        ]
        if run_unsupported:
            raise MigrationError(
                f"static marker run has unsupported children {run_unsupported}"
            )
        local_start = max(0, start - run_start)
        local_end = min(len(value), end - run_start)
        segment = value[local_start:local_end]
        if segment:
            reconstructed.append(segment)
            result_by_run[id(run)] = make_text_run(
                run, segment, preserve_vertical=True
            )
    if "".join(reconstructed) != marker:
        raise MigrationError("run slices do not reconstruct the static marker")

    result_body: list[etree._Element] = []
    affected_ids = {id(row[0]) for row in affected}
    for child in interval:
        if id(child) in result_by_run:
            result_body.append(result_by_run[id(child)])
        elif child.tag == qn(W, "proofErr"):
            result_body.append(copy.deepcopy(child))
        elif id(child) not in affected_ids:
            raise MigrationError("unexpected run inside static marker interval")

    first_run, first_start, _, first_value = affected[0]
    last_run, last_start, _, last_value = affected[-1]
    prefix = first_value[: max(0, start - first_start)]
    suffix = last_value[min(len(last_value), end - last_start) :]
    replacement: list[etree._Element] = []
    if prefix:
        replacement.append(make_text_run(first_run, prefix, preserve_vertical=True))
    replacement.extend(
        [
            make_fldchar_run(first_run, "begin"),
            make_instr_run(first_run, code),
            make_fldchar_run(first_run, "separate"),
            *result_body,
            make_fldchar_run(first_run, "end"),
        ]
    )
    if suffix:
        replacement.append(make_text_run(last_run, suffix, preserve_vertical=True))
    for child in interval:
        paragraph.remove(child)
    for offset, child in enumerate(replacement):
        paragraph.insert(position + offset, child)


def set_field_code(record: FieldRecord, code: str) -> None:
    if record.simple:
        assert record.simple_node is not None
        record.simple_node.set(qn(W, "instr"), code)
        return
    if not record.closed or record.begin is None or record.end is None:
        raise MigrationError("cannot rewrite an unbalanced complex field")
    if record.separator is None:
        raise MigrationError("citation/bibliography field has no separator")
    if not record.instr_nodes:
        raise MigrationError("complex field has no instruction text")
    record.instr_nodes[0].text = f" {code} "
    record.instr_nodes[0].set(qn(XML, "space"), "preserve")
    for node in record.instr_nodes[1:]:
        node.text = ""


def remove_fld_data_nodes(record: FieldRecord) -> int:
    removed = 0
    for node in list(record.fld_data_nodes):
        parent = node.getparent()
        if parent is not None:
            parent.remove(node)
            removed += 1
    return removed


def remove_endnote_data_field(record: FieldRecord) -> None:
    if record.result.strip():
        raise MigrationError(
            "refusing to remove an EN.CITE.DATA field with visible result text"
        )
    if record.simple:
        assert record.simple_node is not None
        parent = record.simple_node.getparent()
        if parent is None:
            raise MigrationError("EN.CITE.DATA simple field is detached")
        parent.remove(record.simple_node)
        return
    if record.begin is None or record.end is None or not record.closed:
        raise MigrationError("EN.CITE.DATA field is unbalanced")
    begin_run = run_of(record.begin)
    end_run = run_of(record.end)
    begin_paragraph = paragraph_of(record.begin)
    end_paragraph = paragraph_of(record.end)
    if (
        begin_run is None
        or end_run is None
        or begin_paragraph is None
        or begin_paragraph is not end_paragraph
        or begin_run.getparent() is not begin_paragraph
        or end_run.getparent() is not begin_paragraph
    ):
        raise MigrationError(
            "EN.CITE.DATA removal supports only direct-run, single-paragraph fields"
        )
    children = list(begin_paragraph)
    start = children.index(begin_run)
    finish = children.index(end_run)
    if finish < start:
        raise MigrationError("EN.CITE.DATA field boundaries are reversed")
    for child in children[start : finish + 1]:
        begin_paragraph.remove(child)


def wrap_bibliography_range(
    root: etree._Element,
    bibliography: dict[str, Any],
) -> tuple[int, int]:
    if bibliography.get("part") != "word/document.xml":
        raise MigrationError(
            "static bibliography wrapping supports only word/document.xml"
        )
    paragraphs = list(root.iter(qn(W, "p")))
    start = bibliography.get("start_paragraph_index")
    end = bibliography.get("end_paragraph_index")
    if (
        not isinstance(start, int)
        or isinstance(start, bool)
        or not isinstance(end, int)
        or isinstance(end, bool)
        or start < 0
        or end < start
        or end >= len(paragraphs)
    ):
        raise MigrationError("invalid static bibliography paragraph range")
    first = paragraphs[start]
    last = paragraphs[end]
    body = root.find(f".//{{{W}}}body")
    if body is None or first.getparent() is not body or last.getparent() is not body:
        raise MigrationError(
            "static bibliography boundaries must be direct body paragraphs"
        )
    body_children = list(body)
    body_start = body_children.index(first)
    body_end = body_children.index(last)
    if any(child.tag != qn(W, "p") for child in body_children[body_start : body_end + 1]):
        raise MigrationError(
            "static bibliography range contains a non-paragraph body child"
        )
    expected_first = bibliography.get("start_text")
    expected_last = bibliography.get("end_text")
    if not isinstance(expected_first, str) or not isinstance(expected_last, str):
        raise MigrationError(
            "static bibliography requires exact start_text and end_text guards"
        )
    if visible_text(first) != expected_first or visible_text(last) != expected_last:
        raise MigrationError("static bibliography boundary text differs from manifest")
    if any(
        paragraph.xpath(".//w:fldChar|.//w:fldSimple", namespaces={"w": W})
        for paragraph in paragraphs[start : end + 1]
    ):
        raise MigrationError("static bibliography range already contains Word fields")

    first_template = next(first.iter(qn(W, "r")), None)
    last_template = next(last.iter(qn(W, "r")), None)
    insert_at = 1 if len(first) and first[0].tag == qn(W, "pPr") else 0
    for offset, run in enumerate(
        [
            make_fldchar_run(first_template, "begin"),
            make_instr_run(first_template, BIBLIOGRAPHY_CODE),
            make_fldchar_run(first_template, "separate"),
        ]
    ):
        first.insert(insert_at + offset, run)
    last.append(make_fldchar_run(last_template, "end"))
    return start, end


def validate_preferences(manifest: dict[str, Any]) -> tuple[dict[str, Any], str]:
    preferences = manifest.get("preferences")
    if not isinstance(preferences, dict):
        raise MigrationError("manifest.preferences must be an object")
    reject_keys(
        preferences,
        "preferences",
        "data_version",
        "bibliography_style_has_been_set",
    )
    for forbidden in ("session", "session_id", "user_id", "userID"):
        if forbidden in preferences:
            raise MigrationError(
                f"preferences.{forbidden} is forbidden; a fresh session is generated"
            )
    style = require_nonempty_string(preferences.get("style"), "preferences.style")
    zotero_version = require_nonempty_string(
        preferences.get("zotero_version"), "preferences.zotero_version"
    )
    if not re.fullmatch(r"\d+(?:\.\d+){0,3}(?:[-+][A-Za-z0-9.]+)?", zotero_version):
        raise MigrationError("preferences.zotero_version has an invalid format")
    citation_format = preferences.get("citation_format")
    if citation_format not in {"plain", "superscript"}:
        raise MigrationError(
            "preferences.citation_format must be plain or superscript"
        )
    prefs = preferences.get("prefs")
    if not isinstance(prefs, dict) or prefs.get("fieldType") != "Field":
        raise MigrationError('preferences.prefs must include "fieldType": "Field"')
    for name, value in prefs.items():
        require_nonempty_string(name, "preference name")
        if not isinstance(value, (str, bool, int, float)) or isinstance(value, (dict, list)):
            raise MigrationError(f"preference {name!r} has an unsupported value")
    result = copy.deepcopy(preferences)
    result["style"] = style
    result["zotero_version"] = zotero_version
    return result, citation_format


def zotero_preferences_xml(preferences: dict[str, Any], session_id: str) -> str:
    data = etree.Element(
        "data",
        {
            "data-version": "3",
            "zotero-version": preferences["zotero_version"],
        },
    )
    etree.SubElement(data, "session", {"id": session_id})
    etree.SubElement(
        data,
        "style",
        {
            "id": preferences["style"],
            "hasBibliography": "1",
            "bibliographyStyleHasBeenSet": "1",
        },
    )
    prefs_node = etree.SubElement(data, "prefs")
    for name, value in preferences["prefs"].items():
        if isinstance(value, bool):
            rendered = "true" if value else "false"
        else:
            rendered = str(value)
        etree.SubElement(prefs_node, "pref", {"name": name, "value": rendered})
    return etree.tostring(data, encoding="unicode", with_tail=False)


def patch_custom_properties(raw: bytes | None, preferences_xml: str) -> bytes:
    if raw is None:
        root = etree.Element(qn(CP, "Properties"), nsmap={None: CP, "vt": VT})
    else:
        root = etree.fromstring(
            raw, etree.XMLParser(remove_blank_text=False, huge_tree=True)
        )
        if root.tag != qn(CP, "Properties"):
            raise MigrationError("docProps/custom.xml has an unexpected root")
    for prop in list(root):
        if (prop.get("name") or "").startswith("ZOTERO_PREF_"):
            root.remove(prop)
    pids = [
        int(prop.get("pid"))
        for prop in root.findall(qn(CP, "property"))
        if (prop.get("pid") or "").isdigit()
    ]
    pid = max(pids, default=1) + 1
    chunks = [preferences_xml[index : index + 240] for index in range(0, len(preferences_xml), 240)]
    for index, value in enumerate(chunks, 1):
        prop = etree.SubElement(root, qn(CP, "property"))
        prop.set("fmtid", FMTID)
        prop.set("pid", str(pid))
        prop.set("name", f"ZOTERO_PREF_{index}")
        text = etree.SubElement(prop, qn(VT, "lpwstr"))
        text.text = value
        pid += 1
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def ensure_custom_content_type(raw: bytes) -> bytes:
    root = etree.fromstring(raw)
    overrides = root.findall(qn(CT, "Override"))
    if not any(node.get("PartName") == "/docProps/custom.xml" for node in overrides):
        node = etree.SubElement(root, qn(CT, "Override"))
        node.set("PartName", "/docProps/custom.xml")
        node.set("ContentType", CUSTOM_CONTENT_TYPE)
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def ensure_custom_relationship(raw: bytes) -> bytes:
    root = etree.fromstring(raw)
    relationships = root.findall(qn(REL, "Relationship"))
    if not any(node.get("Type") == CUSTOM_REL_TYPE for node in relationships):
        used = {node.get("Id") for node in relationships}
        number = 1
        while f"rId{number}" in used:
            number += 1
        node = etree.SubElement(root, qn(REL, "Relationship"))
        node.set("Id", f"rId{number}")
        node.set("Type", CUSTOM_REL_TYPE)
        node.set("Target", "docProps/custom.xml")
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def extract_csl_payload(code: str) -> dict[str, Any]:
    if not code.strip().startswith(ZOTERO_CITATION):
        raise MigrationError("not a Zotero citation instruction")
    remainder = code.strip()[len(ZOTERO_CITATION) :].lstrip()
    try:
        payload, end = json.JSONDecoder().raw_decode(remainder)
    except json.JSONDecodeError as exc:
        raise MigrationError(f"malformed Zotero CSL payload: {exc}") from exc
    if remainder[end:].strip():
        raise MigrationError("unexpected text after Zotero CSL payload")
    if not isinstance(payload, dict):
        raise MigrationError("Zotero CSL payload is not an object")
    return payload


def validate_final_fields(
    roots: dict[str, etree._Element], expected_citations: int
) -> dict[str, Any]:
    citation_ids: list[str] = []
    counts: Counter[str] = Counter()
    for part, root in roots.items():
        records, errors, _ = parse_field_records(root, part)
        if errors:
            raise MigrationError("; ".join(errors))
        for record in records:
            kind = record.kind
            counts[kind] += 1
            if kind.startswith("endnote-"):
                raise MigrationError(f"residual EndNote field in {part}")
            if kind == "zotero-citation":
                if record.simple or record.separator is None or not record.closed:
                    raise MigrationError("Zotero citation is not a balanced complex field")
                payload = extract_csl_payload(record.code)
                citation_id = payload.get("citationID")
                if not isinstance(citation_id, str) or not citation_id:
                    raise MigrationError("Zotero citation lacks citationID")
                citation_ids.append(citation_id)
                properties = payload.get("properties")
                if not isinstance(properties, dict):
                    raise MigrationError("Zotero citation lacks properties")
                if not isinstance(properties.get("formattedCitation"), str):
                    raise MigrationError("formattedCitation is missing")
                if not isinstance(properties.get("plainCitation"), str):
                    raise MigrationError("plainCitation is missing")
                items = payload.get("citationItems")
                if not isinstance(items, list) or not items:
                    raise MigrationError("Zotero citation has no citationItems")
                for item in items:
                    if not isinstance(item, dict) or item.get("id") in (None, ""):
                        raise MigrationError("citation item lacks id")
                    uris = item.get("uris")
                    if not isinstance(uris, list) or not uris or not all(
                        isinstance(uri, str) and uri for uri in uris
                    ):
                        raise MigrationError("citation item lacks URI")
                    if not isinstance(item.get("itemData"), dict):
                        raise MigrationError("citation item lacks itemData")
            elif kind == "zotero-bibliography":
                if record.simple or record.separator is None or not record.closed:
                    raise MigrationError(
                        "Zotero bibliography is not a balanced complex field"
                    )
    duplicates = sorted(
        value for value, count in Counter(citation_ids).items() if count > 1
    )
    if duplicates:
        raise MigrationError(f"duplicate citationID values: {duplicates}")
    if counts["zotero-citation"] != expected_citations:
        raise MigrationError(
            f"expected {expected_citations} Zotero citations, found "
            f"{counts['zotero-citation']}"
        )
    if counts["zotero-bibliography"] != 1:
        raise MigrationError(
            f"expected one Zotero bibliography, found {counts['zotero-bibliography']}"
        )
    return {
        "zotero_citations": counts["zotero-citation"],
        "zotero_bibliographies": counts["zotero-bibliography"],
        "citation_ids_unique": True,
    }


def structure_signature(root: etree._Element) -> dict[str, int]:
    return {
        "paragraphs": len(root.xpath(".//w:p", namespaces={"w": W})),
        "tables": len(root.xpath(".//w:tbl", namespaces={"w": W})),
        "rows": len(root.xpath(".//w:tr", namespaces={"w": W})),
        "cells": len(root.xpath(".//w:tc", namespaces={"w": W})),
        "drawings": len(root.xpath(".//w:drawing", namespaces={"w": W})),
        "revisions": len(
            root.xpath(
                ".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo",
                namespaces={"w": W},
            )
        ),
        "comments": len(
            root.xpath(
                ".//w:commentRangeStart|.//w:commentRangeEnd|.//w:commentReference",
                namespaces={"w": W},
            )
        ),
    }


def publish_temp_no_overwrite(temporary: Path, destination: Path) -> None:
    if destination.exists():
        raise MigrationError(f"refusing to overwrite {destination}")
    try:
        os.link(temporary, destination)
    except FileExistsError as exc:
        raise MigrationError(f"refusing to overwrite {destination}") from exc
    except OSError as exc:
        raise MigrationError(
            f"atomic no-overwrite publication failed for {destination}: {exc}"
        ) from exc


def atomic_write_new(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=path.parent
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
        publish_temp_no_overwrite(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def serialize_root(root: etree._Element) -> bytes:
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("manifest_json", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument(
        "--report",
        type=Path,
        help="Migration report path (default: OUTPUT.migration.json)",
    )
    args = parser.parse_args()

    input_docx = args.input_docx.resolve()
    manifest_path = args.manifest_json.resolve()
    output_docx = args.output_docx.resolve()
    report_path = (
        args.report.resolve()
        if args.report is not None
        else output_docx.with_suffix(".migration.json")
    )
    if not input_docx.is_file():
        raise MigrationError(f"input DOCX does not exist: {input_docx}")
    if not manifest_path.is_file():
        raise MigrationError(f"manifest does not exist: {manifest_path}")
    if input_docx == output_docx:
        raise MigrationError("output DOCX must be separate from the source")
    if output_docx.exists():
        raise MigrationError(f"refusing to overwrite {output_docx}")
    if report_path.exists():
        raise MigrationError(f"refusing to overwrite report {report_path}")
    if output_docx.suffix.lower() != ".docx":
        raise MigrationError("output path must end in .docx")
    output_docx.parent.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise MigrationError(f"cannot read manifest: {exc}") from exc
    if not isinstance(manifest, dict) or manifest.get("schema_version") != 1:
        raise MigrationError("manifest schema_version must be 1")
    validate_manifest_issues(manifest)
    source = manifest.get("source")
    if not isinstance(source, dict):
        raise MigrationError("manifest.source must be an object")
    expected_sha = require_nonempty_string(source.get("sha256"), "source.sha256").lower()
    if not re.fullmatch(r"[0-9a-f]{64}", expected_sha):
        raise MigrationError("source.sha256 is not a SHA-256 digest")
    input_sha = sha256_file(input_docx)
    if input_sha != expected_sha:
        raise MigrationError(
            f"source SHA-256 mismatch: manifest {expected_sha}, input {input_sha}"
        )
    source_system = source.get("citation_system")
    if source_system not in {"static-numeric", "endnote"}:
        raise MigrationError(
            "source.citation_system must be static-numeric or endnote"
        )

    preferences, citation_format = validate_preferences(manifest)
    references, by_number, by_source_id = prepare_references(manifest)
    occurrences = prepare_occurrences(
        manifest, source_system, by_number, by_source_id
    )
    bibliography = manifest.get("bibliography")
    if not isinstance(bibliography, dict):
        raise MigrationError("manifest.bibliography must be an object")
    reject_keys(bibliography, "bibliography", "first_text", "last_text", "result")
    expected_bibliography_kind = (
        "wrap-paragraph-range"
        if source_system == "static-numeric"
        else "endnote-reflist"
    )
    if bibliography.get("kind") != expected_bibliography_kind:
        raise MigrationError(
            f"bibliography.kind must be {expected_bibliography_kind!r}"
        )

    try:
        with zipfile.ZipFile(input_docx, "r") as archive:
            infos = archive.infolist()
            names = [info.filename for info in infos]
            duplicates = sorted(
                name for name, count in Counter(names).items() if count > 1
            )
            if duplicates:
                raise MigrationError(f"source ZIP has duplicate members: {duplicates}")
            bad_member = archive.testzip()
            if bad_member is not None:
                raise MigrationError(f"source ZIP CRC failure: {bad_member}")
            original = {info.filename: archive.read(info.filename) for info in infos}
    except zipfile.BadZipFile as exc:
        raise MigrationError(f"input is not a valid DOCX ZIP: {exc}") from exc
    if "word/document.xml" not in original:
        raise MigrationError("word/document.xml is missing")
    if "[Content_Types].xml" not in original or "_rels/.rels" not in original:
        raise MigrationError("required package metadata is missing")

    story_parts = sorted(name for name in original if STORY_RE.fullmatch(name))
    roots: dict[str, etree._Element] = {}
    fields_by_part: dict[str, list[FieldRecord]] = {}
    before_visible: dict[str, str] = {}
    before_structure: dict[str, dict[str, int]] = {}
    for part in story_parts:
        try:
            root = etree.fromstring(
                original[part],
                etree.XMLParser(remove_blank_text=False, huge_tree=True),
            )
        except etree.XMLSyntaxError as exc:
            raise MigrationError(f"malformed story part {part}: {exc}") from exc
        records, errors, _ = parse_field_records(root, part)
        if errors:
            raise MigrationError("; ".join(errors))
        roots[part] = root
        fields_by_part[part] = records
        before_visible[part] = visible_text(root)
        before_structure[part] = structure_signature(root)

    for occurrence in occurrences:
        if occurrence.part not in roots:
            raise MigrationError(
                f"occurrence part is absent from source: {occurrence.part}"
            )

    existing_zotero_citations = 0
    existing_zotero_bibliographies = 0
    for records in fields_by_part.values():
        existing_zotero_citations += sum(
            record.kind == "zotero-citation" for record in records
        )
        existing_zotero_bibliographies += sum(
            record.kind == "zotero-bibliography" for record in records
        )
        for record in records:
            if record.kind == "zotero-citation":
                extract_csl_payload(record.code)
    if existing_zotero_citations:
        raise MigrationError(
            "source already contains Zotero citations; mixed-system migration is unsupported"
        )
    if existing_zotero_bibliographies:
        raise MigrationError(
            "source already contains a Zotero bibliography; mixed bibliography conversion is unsupported"
        )

    session_id = uuid.uuid4().hex
    for index, occurrence in enumerate(occurrences, 1):
        occurrence.citation_id = f"ZWCM-{session_id[:12]}-{index:06d}"

    changed_story_parts: set[str] = set()
    removed_endnote_data = 0
    removed_fld_data_nodes = 0
    converted_endnote = 0
    inserted_static = 0

    if source_system == "static-numeric":
        if any(
            record.kind.startswith("endnote-")
            for records in fields_by_part.values()
            for record in records
        ):
            raise MigrationError("static-numeric branch does not accept EndNote fields")
        bibliography_part = bibliography.get("part")
        excluded: set[tuple[str, int]] = set()
        if bibliography_part == "word/document.xml":
            start = bibliography.get("start_paragraph_index")
            end = bibliography.get("end_paragraph_index")
            if isinstance(start, int) and isinstance(end, int):
                excluded.update((bibliography_part, index) for index in range(start, end + 1))

        discovered: dict[tuple[str, int, int], tuple[str, int, int]] = {}
        for part, root in roots.items():
            for p_index, paragraph in enumerate(root.iter(qn(W, "p"))):
                if (part, p_index) in excluded:
                    continue
                outside = text_outside_citation_fields(paragraph)
                for marker_index, match in enumerate(STATIC_MARKER_RE.finditer(outside), 1):
                    discovered[(part, p_index, marker_index)] = (
                        match.group(0),
                        match.start(),
                        match.end(),
                    )
        selected_keys = {
            (value.part, value.paragraph_index, value.index_in_paragraph)
            for value in occurrences
        }
        if set(discovered) != selected_keys:
            missing = sorted(set(discovered) - selected_keys)
            extra = sorted(selected_keys - set(discovered))
            raise MigrationError(
                f"static occurrence coverage mismatch; unmanifested={missing}, absent={extra}"
            )

        grouped: dict[tuple[str, int], list[tuple[PreparedOccurrence, int, int]]] = defaultdict(list)
        for occurrence in occurrences:
            key = (
                occurrence.part,
                occurrence.paragraph_index,
                occurrence.index_in_paragraph,
            )
            actual, start, end = discovered[key]
            if actual != occurrence.source_text:
                raise MigrationError(
                    f"{occurrence.occurrence_id} source_text differs: {actual!r}"
                )
            expected_numbers = [reference.number for reference in occurrence.references]
            if any(number is None for number in expected_numbers):
                raise MigrationError("static markers require numbered references")
            if parse_numeric_cluster(actual) != expected_numbers:
                raise MigrationError(
                    f"{occurrence.occurrence_id} marker expansion differs from references"
                )
            if "start" in occurrence.source and occurrence.source["start"] != start:
                raise MigrationError(f"{occurrence.occurrence_id} start offset differs")
            if "end" in occurrence.source and occurrence.source["end"] != end:
                raise MigrationError(f"{occurrence.occurrence_id} end offset differs")
            paragraph = list(roots[occurrence.part].iter(qn(W, "p")))[
                occurrence.paragraph_index
            ]
            if paragraph.xpath(".//w:fldChar|.//w:fldSimple", namespaces={"w": W}):
                raise MigrationError(
                    f"{occurrence.occurrence_id} paragraph already contains a Word field"
                )
            grouped[(occurrence.part, occurrence.paragraph_index)].append(
                (occurrence, start, end)
            )

        for (part, p_index), values in grouped.items():
            paragraph = list(roots[part].iter(qn(W, "p")))[p_index]
            for occurrence, start, end in sorted(values, key=lambda row: row[1], reverse=True):
                replace_static_marker(
                    paragraph,
                    start,
                    end,
                    occurrence.source_text,
                    citation_code(occurrence, citation_format),
                )
                inserted_static += 1
            changed_story_parts.add(part)
        wrap_bibliography_range(roots["word/document.xml"], bibliography)
        changed_story_parts.add("word/document.xml")
    else:
        all_endnote = {
            record.occurrence_key: record
            for records in fields_by_part.values()
            for record in records
            if record.kind == "endnote-citation"
        }
        selected_keys = {
            (value.part, value.paragraph_index, value.index_in_paragraph)
            for value in occurrences
        }
        if set(all_endnote) != selected_keys:
            missing = sorted(set(all_endnote) - selected_keys)
            extra = sorted(selected_keys - set(all_endnote))
            raise MigrationError(
                f"EndNote occurrence coverage mismatch; unmanifested={missing}, absent={extra}"
            )
        selected_records: list[tuple[PreparedOccurrence, FieldRecord]] = []
        for occurrence in occurrences:
            key = (
                occurrence.part,
                occurrence.paragraph_index,
                occurrence.index_in_paragraph,
            )
            record = all_endnote[key]
            if record.simple:
                raise MigrationError("simple EndNote citation fields are unsupported")
            if record.result != occurrence.source_text:
                raise MigrationError(
                    f"{occurrence.occurrence_id} cached result differs from manifest"
                )
            if "code_sha256" in occurrence.source:
                if occurrence.source["code_sha256"] != sha256_text(record.code):
                    raise MigrationError(
                        f"{occurrence.occurrence_id} EndNote code hash differs"
                    )
            selected_records.append((occurrence, record))

        data_records = [
            record
            for records in fields_by_part.values()
            for record in records
            if record.kind == "endnote-data"
        ]
        for record in data_records:
            # Preflight removability before mutating any XML.
            if record.result.strip():
                raise MigrationError("EN.CITE.DATA has visible result text")
            if not record.simple:
                if record.begin is None or record.end is None:
                    raise MigrationError("EN.CITE.DATA is unbalanced")
                begin_run = run_of(record.begin)
                end_run = run_of(record.end)
                paragraph = paragraph_of(record.begin)
                if (
                    begin_run is None
                    or end_run is None
                    or paragraph is None
                    or paragraph_of(record.end) is not paragraph
                    or begin_run.getparent() is not paragraph
                    or end_run.getparent() is not paragraph
                ):
                    raise MigrationError(
                        "EN.CITE.DATA is not a direct-run, single-paragraph field"
                    )

        bibliography_part = require_nonempty_string(
            bibliography.get("part"), "bibliography.part"
        )
        p_index = bibliography.get("paragraph_index")
        field_index = bibliography.get("index_in_paragraph")
        if not isinstance(p_index, int) or isinstance(p_index, bool) or p_index < 0:
            raise MigrationError("bibliography.paragraph_index is invalid")
        if not isinstance(field_index, int) or isinstance(field_index, bool) or field_index < 1:
            raise MigrationError("bibliography.index_in_paragraph must be 1-based")
        reflists = {
            record.occurrence_key: record
            for records in fields_by_part.values()
            for record in records
            if record.kind == "endnote-bibliography"
        }
        bibliography_key = (bibliography_part, p_index, field_index)
        if len(reflists) != 1 or bibliography_key not in reflists:
            raise MigrationError(
                "manifest must locate the sole ADDIN EN.REFLIST field exactly"
            )
        reflist = reflists[bibliography_key]
        expected_result = bibliography.get("source_text")
        if expected_result is not None and expected_result != reflist.result:
            raise MigrationError("EndNote bibliography cached result differs from manifest")

        for occurrence, record in selected_records:
            set_field_code(record, citation_code(occurrence, citation_format))
            removed_fld_data_nodes += remove_fld_data_nodes(record)
            changed_story_parts.add(occurrence.part)
            converted_endnote += 1
        for record in sorted(data_records, key=lambda value: value.order, reverse=True):
            remove_endnote_data_field(record)
            changed_story_parts.add(record.part)
            removed_endnote_data += 1
        set_field_code(reflist, BIBLIOGRAPHY_CODE)
        removed_fld_data_nodes += remove_fld_data_nodes(reflist)
        changed_story_parts.add(reflist.part)

    expected_citations = existing_zotero_citations + len(occurrences)
    final_field_report = validate_final_fields(roots, expected_citations)
    after_structure = {part: structure_signature(root) for part, root in roots.items()}
    if after_structure != before_structure:
        raise MigrationError("high-level Word story structure changed unexpectedly")
    for part, root in roots.items():
        if visible_text(root) != before_visible[part]:
            raise MigrationError(f"visible story text changed unexpectedly in {part}")

    preferences_xml = zotero_preferences_xml(preferences, session_id)
    replacements: dict[str, bytes] = {
        part: serialize_root(roots[part]) for part in changed_story_parts
    }
    custom_before = original.get("docProps/custom.xml")
    replacements["docProps/custom.xml"] = patch_custom_properties(
        custom_before, preferences_xml
    )
    replacements["[Content_Types].xml"] = ensure_custom_content_type(
        original["[Content_Types].xml"]
    )
    replacements["_rels/.rels"] = ensure_custom_relationship(original["_rels/.rels"])
    replacements = {
        name: value
        for name, value in replacements.items()
        if name not in original or original[name] != value
    }
    allowed_changes = set(changed_story_parts) | {
        "docProps/custom.xml",
        "[Content_Types].xml",
        "_rels/.rels",
    }
    if set(replacements) - allowed_changes:
        raise MigrationError(
            f"unexpected replacement parts: {sorted(set(replacements) - allowed_changes)}"
        )

    # Verify the frozen source again immediately before producing the output.
    second_sha = sha256_file(input_docx)
    if second_sha != expected_sha:
        raise MigrationError("source changed during migration; no output was written")

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=output_docx.stem + ".",
        suffix=".docx.tmp",
        dir=output_docx.parent,
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    output_published = False
    try:
        with zipfile.ZipFile(temporary, "w") as target:
            for info in infos:
                value = replacements.get(info.filename, original[info.filename])
                target.writestr(info, value)
            for name in sorted(set(replacements) - set(original)):
                target.writestr(name, replacements[name], compress_type=zipfile.ZIP_DEFLATED)
        with zipfile.ZipFile(temporary, "r") as check:
            if check.testzip() is not None:
                raise MigrationError("output ZIP CRC validation failed")
            expected_names = names + sorted(set(replacements) - set(original))
            if check.namelist() != expected_names:
                raise MigrationError("output package member set/order changed unexpectedly")
            changed_parts = [
                name
                for name in names
                if check.read(name) != original[name]
            ] + sorted(set(replacements) - set(original))
        output_sha = sha256_file(temporary)
        used_source_ids = {
            reference.source_id
            for occurrence in occurrences
            for reference in occurrence.references
        }
        report = {
            "schema_version": 1,
            "status": "migrated-pre-refresh",
            "input": str(input_docx),
            "input_sha256": input_sha,
            "manifest": str(manifest_path),
            "manifest_sha256": sha256_file(manifest_path),
            "output": str(output_docx),
            "output_sha256": output_sha,
            "source_citation_system": source_system,
            "style": preferences["style"],
            "zotero_version_from_manifest": preferences["zotero_version"],
            "fresh_session_id": session_id,
            "existing_zotero_citations_preserved": 0,
            "static_citations_inserted": inserted_static,
            "endnote_citations_converted": converted_endnote,
            "endnote_data_fields_removed": removed_endnote_data,
            "endnote_fld_data_nodes_removed": removed_fld_data_nodes,
            "bibliography_fields": 1,
            "references": {
                "manifest_total": len(references),
                "used": len(used_source_ids),
                "linked": sum(
                    reference.status == "linked" and reference.source_id in used_source_ids
                    for reference in references
                ),
                "embedded": sum(
                    reference.status == "embedded" and reference.source_id in used_source_ids
                    for reference in references
                ),
                "unused_source_ids": sorted(
                    reference.source_id
                    for reference in references
                    if reference.source_id not in used_source_ids
                ),
            },
            "field_validation": final_field_report,
            "visible_story_text_preserved": True,
            "high_level_structure_preserved": True,
            "changed_parts": changed_parts,
            "zotero_library_mutated": False,
            "next_required_gates": [
                "validate_zotero_docx.py",
                "ZoteroRefresh in Microsoft Word",
                "post-refresh structural validation",
                "page-by-page rendered visual QA",
            ],
            "support_limits": [
                "static markers must be bracketed numeric clusters in direct paragraph runs",
                "static marker paragraphs cannot already contain Word fields",
                "static bibliography must be a contiguous direct-body paragraph range",
                "EndNote citation fields must be balanced complex fields",
                "EN.CITE.DATA removal requires an empty-result direct-run single-paragraph field",
                "mixed pre-existing Zotero bibliography conversion is unsupported",
            ],
        }
        report_bytes = (
            json.dumps(report, ensure_ascii=False, indent=2) + "\n"
        ).encode("utf-8")
        publish_temp_no_overwrite(temporary, output_docx)
        output_published = True
        try:
            atomic_write_new(report_path, report_bytes)
        except Exception:
            if output_docx.is_file() and sha256_file(output_docx) == output_sha:
                output_docx.unlink()
                output_published = False
            raise
    finally:
        temporary.unlink(missing_ok=True)

    if not output_published:
        raise MigrationError("output publication did not complete")
    print(
        json.dumps(
            {
                "status": "migrated-pre-refresh",
                "output": str(output_docx),
                "report": str(report_path),
                "citation_fields": expected_citations,
                "bibliography_fields": 1,
                "changed_parts": changed_parts,
                "zotero_library_mutated": False,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
    try:
        raise SystemExit(main())
    except MigrationError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(2)

#!/usr/bin/env python3
"""Read-only inventory of citation systems and Word field structure in a DOCX."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
CP = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
NS = {"w": W, "m": M, "cp": CP}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
STORY_RE = re.compile(
    r"^word/(?:document|footnotes|endnotes|comments|header\d+|footer\d+)\.xml$"
)
DEFAULT_HEADING = r"^References$"
DEFAULT_REFERENCE = (
    r"^\s*(?:\[(?P<bracket>\d+)\]|(?P<plain>\d+)(?:[.)]|\t))"
    r"\s*(?P<body>.+?)\s*$"
)
DEFAULT_MARKER = r"\[(?P<body>\d+(?:(?:\s*[,;]\s*|\s*[-–—]\s*)\d+)*)\]"
BARE_NUMERIC_CLUSTER = re.compile(
    r"^\d+(?:(?:\s*[,;]\s*|\s*[-\u2013\u2014]\s*)\d+)*$"
)
DOI_RE = re.compile(r"10\.\d{4,9}/[-._;()/:A-Z0-9]+", re.I)
ZOTERO_CITATION = "ADDIN ZOTERO_ITEM CSL_CITATION"
ZOTERO_BIBLIOGRAPHY = "ADDIN ZOTERO_BIBL"
ENDNOTE_CITATION = "ADDIN EN.CITE"
ENDNOTE_DATA = "ADDIN EN.CITE.DATA"
ENDNOTE_BIBLIOGRAPHY = "ADDIN EN.REFLIST"


def qn(local: str) -> str:
    return f"{{{W}}}{local}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sha256_text(value: str) -> str:
    return hashlib.sha256(value.encode("utf-8")).hexdigest()


def visible_text(node: etree._Element) -> str:
    chunks: list[str] = []
    for element in node.iter():
        if element.tag in {qn("t"), qn("delText")} and element.text:
            chunks.append(element.text)
        elif element.tag == qn("tab"):
            chunks.append("\t")
        elif element.tag in {qn("br"), qn("cr")}:
            chunks.append("\n")
    return "".join(chunks)


def classify_code(code: str) -> str:
    value = code.strip()
    if value.startswith(ZOTERO_CITATION):
        return "zotero-citation"
    if value.startswith(ZOTERO_BIBLIOGRAPHY):
        return "zotero-bibliography"
    if value.startswith(ENDNOTE_DATA):
        return "endnote-data"
    if re.match(r"^ADDIN EN\.CITE(?:\s|$)", value):
        return "endnote-citation"
    if value.startswith(ENDNOTE_BIBLIOGRAPHY):
        return "endnote-bibliography"
    return "other"


def parse_numeric_cluster(value: str) -> list[int]:
    body = value.strip()
    body = re.sub(r"^[\[(]", "", body)
    body = re.sub(r"[\])]$", "", body)
    body = body.replace("\u00a0", " ").strip()
    if not body:
        return []
    numbers: list[int] = []
    for token in re.split(r"\s*[,;]\s*", body):
        token = token.strip()
        match = re.fullmatch(r"(\d+)\s*[-–—]\s*(\d+)", token)
        if match:
            left, right = map(int, match.groups())
            if right < left:
                return []
            numbers.extend(range(left, right + 1))
        elif re.fullmatch(r"\d+", token):
            numbers.append(int(token))
        else:
            return []
    return numbers


def extract_citation_payload(code: str) -> tuple[dict[str, Any] | None, str | None]:
    marker = "CSL_CITATION"
    position = code.find(marker)
    if position < 0:
        return None, "CSL_CITATION token is missing"
    source = code[position + len(marker) :].lstrip()
    try:
        payload, _ = json.JSONDecoder().raw_decode(source)
    except (json.JSONDecodeError, TypeError) as exc:
        return None, str(exc)
    if not isinstance(payload, dict):
        return None, "citation payload is not a JSON object"
    return payload, None


def paragraph_of(node: etree._Element) -> etree._Element | None:
    current: etree._Element | None = node
    while current is not None:
        if current.tag == qn("p"):
            return current
        current = current.getparent()
    return None


@dataclass
class FieldRecord:
    part: str
    order: int
    field_index: int | str
    paragraph_index: int | None
    index_in_paragraph: int
    simple: bool = False
    simple_node: etree._Element | None = None
    begin: etree._Element | None = None
    separator: etree._Element | None = None
    end: etree._Element | None = None
    end_paragraph_index: int | None = None
    instr_nodes: list[etree._Element] = field(default_factory=list)
    fld_data_nodes: list[etree._Element] = field(default_factory=list)
    result_chunks: list[str] = field(default_factory=list)
    closed: bool = False

    @property
    def code(self) -> str:
        if self.simple:
            assert self.simple_node is not None
            return (self.simple_node.get(qn("instr")) or "").strip()
        return "".join(node.text or "" for node in self.instr_nodes).strip()

    @property
    def kind(self) -> str:
        return classify_code(self.code)

    @property
    def result(self) -> str:
        if self.simple:
            assert self.simple_node is not None
            return visible_text(self.simple_node)
        return "".join(self.result_chunks)

    @property
    def occurrence_key(self) -> tuple[str, int | None, int]:
        return (self.part, self.paragraph_index, self.index_in_paragraph)


def parse_field_records(
    root: etree._Element, part: str
) -> tuple[list[FieldRecord], list[str], dict[str, int]]:
    paragraphs = list(root.iter(qn("p")))
    paragraph_ids = {id(node): index for index, node in enumerate(paragraphs)}
    per_paragraph: Counter[int | None] = Counter()
    stack: list[FieldRecord] = []
    records: list[FieldRecord] = []
    errors: list[str] = []
    character_counts: Counter[str] = Counter()
    order = 0
    complex_number = 0
    simple_number = 0

    def p_index(node: etree._Element) -> int | None:
        paragraph = paragraph_of(node)
        return paragraph_ids.get(id(paragraph)) if paragraph is not None else None

    for node in root.iter():
        if node.tag == qn("fldSimple"):
            order += 1
            simple_number += 1
            index = p_index(node)
            per_paragraph[index] += 1
            records.append(
                FieldRecord(
                    part=part,
                    order=order,
                    field_index=f"simple-{simple_number}",
                    paragraph_index=index,
                    index_in_paragraph=per_paragraph[index],
                    end_paragraph_index=index,
                    simple=True,
                    simple_node=node,
                    closed=True,
                )
            )
        elif node.tag == qn("fldChar"):
            kind = node.get(qn("fldCharType"), "")
            character_counts[kind] += 1
            if kind == "begin":
                order += 1
                complex_number += 1
                index = p_index(node)
                per_paragraph[index] += 1
                record = FieldRecord(
                    part=part,
                    order=order,
                    field_index=complex_number,
                    paragraph_index=index,
                    index_in_paragraph=per_paragraph[index],
                    begin=node,
                )
                records.append(record)
                stack.append(record)
            elif kind == "separate":
                if not stack:
                    errors.append(f"{part}: separator without begin")
                elif stack[-1].separator is not None:
                    errors.append(
                        f"{part}: duplicate separator in field {stack[-1].field_index}"
                    )
                else:
                    stack[-1].separator = node
            elif kind == "end":
                if not stack:
                    errors.append(f"{part}: end without begin")
                else:
                    record = stack.pop()
                    record.end = node
                    record.end_paragraph_index = p_index(node)
                    record.closed = True
        elif node.tag == qn("instrText") and stack and stack[-1].separator is None:
            stack[-1].instr_nodes.append(node)
        elif node.tag == qn("fldData") and stack:
            stack[-1].fld_data_nodes.append(node)
        elif node.tag in {qn("t"), qn("delText")} and node.text:
            for record in stack:
                if record.separator is not None:
                    record.result_chunks.append(node.text)
        elif node.tag == qn("tab"):
            for record in stack:
                if record.separator is not None:
                    record.result_chunks.append("\t")
        elif node.tag in {qn("br"), qn("cr")}:
            for record in stack:
                if record.separator is not None:
                    record.result_chunks.append("\n")

    errors.extend(f"{part}: unclosed field {record.field_index}" for record in stack)
    return records, errors, dict(character_counts)


def parse_fields(root: etree._Element, part: str) -> tuple[list[dict], list[str], dict]:
    records, errors, character_counts = parse_field_records(root, part)
    output: list[dict] = []
    for record in records:
        p_index = record.paragraph_index
        code = record.code
        result = record.result
        field_data = "".join(node.text or "" for node in record.fld_data_nodes)
        kind = record.kind
        row: dict[str, Any] = {
            "occurrence_id": (
                f"{part}:p{p_index}:f{record.index_in_paragraph}"
                if p_index is not None
                else f"{part}:field:{record.field_index}"
            ),
            "part": part,
            "paragraph_index": p_index,
            "index_in_paragraph": record.index_in_paragraph,
            "field_index": record.field_index,
            "kind": kind,
            "has_separator": record.simple or record.separator is not None,
            "end_paragraph_index": record.end_paragraph_index,
            "code_sha256": sha256_text(code),
            "code_preview": code[:240],
            "result": result,
            "field_data_present": bool(field_data),
            "field_data_sha256": sha256_text(field_data) if field_data else None,
            "simple": record.simple,
        }
        if kind == "endnote-citation":
            row["reference_numbers"] = parse_numeric_cluster(result)
        elif kind == "zotero-citation":
            payload, payload_error = extract_citation_payload(code)
            row["payload_error"] = payload_error
            if payload is not None:
                properties = payload.get("properties")
                plain = properties.get("plainCitation") if isinstance(properties, dict) else None
                items = payload.get("citationItems")
                item_rows = []
                if isinstance(items, list):
                    for item in items:
                        if not isinstance(item, dict):
                            item_rows.append({"malformed": True})
                            continue
                        data = item.get("itemData")
                        item_rows.append(
                            {
                                "id": item.get("id"),
                                "uris": item.get("uris"),
                                "type": data.get("type") if isinstance(data, dict) else None,
                                "title": data.get("title") if isinstance(data, dict) else None,
                                "DOI": data.get("DOI") if isinstance(data, dict) else None,
                                "has_item_data": isinstance(data, dict),
                            }
                        )
                row.update(
                    {
                        "citation_id": payload.get("citationID"),
                        "plain_citation": plain,
                        "reference_numbers": parse_numeric_cluster(str(plain or result)),
                        "citation_items": item_rows,
                    }
                )
        output.append(row)

    return output, errors, character_counts


def text_outside_citation_fields(paragraph: etree._Element) -> str:
    chunks: list[str] = []
    stack: list[dict[str, Any]] = []
    for node in paragraph.iter():
        if node.tag == qn("fldChar"):
            kind = node.get(qn("fldCharType"), "")
            if kind == "begin":
                stack.append({"code": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                stack.pop()
        elif node.tag == qn("instrText") and stack and not stack[-1]["separated"]:
            stack[-1]["code"].append(node.text or "")
        elif node.tag in {qn("t"), qn("delText")} and node.text:
            hidden = any(
                field["separated"]
                and classify_code("".join(field["code"]))
                in {
                    "zotero-citation",
                    "zotero-bibliography",
                    "endnote-citation",
                    "endnote-data",
                    "endnote-bibliography",
                }
                for field in stack
            )
            if not hidden:
                chunks.append(node.text)
        elif node.tag == qn("tab"):
            chunks.append("\t")
        elif node.tag in {qn("br"), qn("cr")}:
            chunks.append("\n")
    return "".join(chunks)

def unsupported_bare_superscripts(
    paragraph: etree._Element,
    *,
    marker_re: re.Pattern[str],
) -> list[dict[str, Any]]:
    """Return narrow native-superscript numeric candidates outside citation fields.

    Bare superscript numbers are intentionally not promoted to citation occurrences:
    they can also be exponents, charges, footnote markers, or other scientific
    notation. The caller emits a stop issue so they must be reviewed or handled by
    an explicit, task-specific marker rule.
    """
    chunks: list[str] = []
    superscript_spans: list[dict[str, Any]] = []
    stack: list[dict[str, Any]] = []
    offset = 0

    for node in paragraph.iter():
        if node.tag == qn("fldChar"):
            kind = node.get(qn("fldCharType"), "")
            if kind == "begin":
                stack.append({"code": [], "separated": False})
            elif kind == "separate" and stack:
                stack[-1]["separated"] = True
            elif kind == "end" and stack:
                stack.pop()
        elif node.tag == qn("instrText") and stack and not stack[-1]["separated"]:
            stack[-1]["code"].append(node.text or "")
        elif node.tag in {qn("t"), qn("delText")} and node.text:
            hidden = any(
                field["separated"]
                and classify_code("".join(field["code"]))
                in {
                    "zotero-citation",
                    "zotero-bibliography",
                    "endnote-citation",
                    "endnote-data",
                    "endnote-bibliography",
                }
                for field in stack
            )
            if hidden:
                continue
            value = node.text
            start = offset
            chunks.append(value)
            offset += len(value)
            run = node.getparent()
            while run is not None and run is not paragraph and run.tag != qn("r"):
                run = run.getparent()
            if run is not None and run.tag == qn("r"):
                is_superscript = bool(
                    run.xpath(
                        "./w:rPr/w:vertAlign[@w:val='superscript']",
                        namespaces=NS,
                    )
                )
                stripped = value.strip()
                if is_superscript and stripped and BARE_NUMERIC_CLUSTER.fullmatch(stripped):
                    left_trim = len(value) - len(value.lstrip())
                    superscript_spans.append(
                        {
                            "source_text": stripped,
                            "start": start + left_trim,
                            "end": start + left_trim + len(stripped),
                        }
                    )
        elif node.tag == qn("tab"):
            chunks.append("\t")
            offset += 1
        elif node.tag in {qn("br"), qn("cr")}:
            chunks.append("\n")
            offset += 1

    paragraph_text = "".join(chunks)
    supported_spans = [
        (match.start(), match.end()) for match in marker_re.finditer(paragraph_text)
    ]
    return [
        {
            **span,
            "context": paragraph_text[max(0, span["start"] - 60) : span["end"] + 60],
        }
        for span in superscript_spans
        if not any(
            span["start"] < supported_end and supported_start < span["end"]
            for supported_start, supported_end in supported_spans
        )
    ]



def package_structure(document_root: etree._Element, names: list[str]) -> dict[str, Any]:
    return {
        "document_paragraphs": len(document_root.xpath(".//w:p", namespaces=NS)),
        "direct_body_paragraphs": len(document_root.xpath("./w:body/w:p", namespaces=NS)),
        "tables": len(document_root.xpath(".//w:tbl", namespaces=NS)),
        "table_rows": len(document_root.xpath(".//w:tr", namespaces=NS)),
        "table_cells": len(document_root.xpath(".//w:tc", namespaces=NS)),
        "drawings": len(document_root.xpath(".//w:drawing", namespaces=NS)),
        "math_objects": len(document_root.xpath(".//m:oMath", namespaces=NS)),
        "math_paragraphs": len(document_root.xpath(".//m:oMathPara", namespaces=NS)),
        "sections": len(document_root.xpath(".//w:sectPr", namespaces=NS)),
        "revisions": len(
            document_root.xpath(
                ".//w:ins|.//w:del|.//w:moveFrom|.//w:moveTo", namespaces=NS
            )
        ),
        "comment_anchors": len(
            document_root.xpath(
                ".//w:commentRangeStart|.//w:commentRangeEnd|.//w:commentReference",
                namespaces=NS,
            )
        ),
        "native_superscript": len(
            document_root.xpath(".//w:vertAlign[@w:val='superscript']", namespaces=NS)
        ),
        "native_subscript": len(
            document_root.xpath(".//w:vertAlign[@w:val='subscript']", namespaces=NS)
        ),
        "media_parts": sorted(name for name in names if name.startswith("word/media/")),
        "custom_xml_parts": sorted(name for name in names if name.startswith("customXml/")),
    }


def audit(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(path)
    heading_re = re.compile(DEFAULT_HEADING, re.I)
    reference_re = re.compile(DEFAULT_REFERENCE)
    marker_re = re.compile(DEFAULT_MARKER)

    with zipfile.ZipFile(path) as archive:
        infos = archive.infolist()
        names = [info.filename for info in infos]
        duplicate_names = sorted(
            name for name, count in Counter(names).items() if count > 1
        )
        bad_member = archive.testzip()
        if "word/document.xml" not in names:
            raise RuntimeError("word/document.xml is missing")
        story_names = sorted(name for name in names if STORY_RE.fullmatch(name))
        roots: dict[str, etree._Element] = {}
        xml_errors: list[str] = []
        for name in story_names:
            try:
                roots[name] = etree.fromstring(
                    archive.read(name), etree.XMLParser(remove_blank_text=False, huge_tree=True)
                )
            except etree.XMLSyntaxError as exc:
                xml_errors.append(f"{name}: {exc}")
        document_root = roots["word/document.xml"]
        media_hashes = {
            name: hashlib.sha256(archive.read(name)).hexdigest()
            for name in names
            if name.startswith("word/media/")
        }
        custom_xml_hashes = {
            name: hashlib.sha256(archive.read(name)).hexdigest()
            for name in names
            if name.startswith("customXml/")
        }
        custom_properties: list[str] = []
        if "docProps/custom.xml" in names:
            try:
                custom_root = etree.fromstring(archive.read("docProps/custom.xml"))
                custom_properties = [
                    value
                    for value in (node.get("name") for node in custom_root)
                    if value
                ]
            except etree.XMLSyntaxError as exc:
                xml_errors.append(f"docProps/custom.xml: {exc}")

    all_fields: list[dict] = []
    field_errors: list[str] = []
    field_characters_by_part: dict[str, dict] = {}
    story_reports: list[dict] = []
    paragraph_elements: dict[str, list[etree._Element]] = {}
    for name, root in roots.items():
        fields, errors, counts = parse_fields(root, name)
        all_fields.extend(fields)
        field_errors.extend(errors)
        field_characters_by_part[name] = counts
        paragraphs = list(root.iter(qn("p")))
        paragraph_elements[name] = paragraphs
        outside_texts = [text_outside_citation_fields(p) for p in paragraphs]
        story_reports.append(
            {
                "part": name,
                "paragraphs": len(paragraphs),
                "visible_text_sha256": sha256_text("\n".join(visible_text(p) for p in paragraphs)),
                "noncitation_text_sha256": sha256_text("\n".join(outside_texts)),
            }
        )

    direct_paragraphs = document_root.xpath("./w:body/w:p", namespaces=NS)
    heading_matches = [
        index
        for index, paragraph in enumerate(direct_paragraphs)
        if heading_re.fullmatch(visible_text(paragraph).strip())
    ]
    references: list[dict] = []
    reference_elements: set[int] = set()
    reference_issues: list[str] = []
    heading_element: etree._Element | None = None
    if len(heading_matches) == 1:
        heading_index = heading_matches[0]
        heading_element = direct_paragraphs[heading_index]
        reference_elements.add(id(heading_element))
        for paragraph in direct_paragraphs[heading_index + 1 :]:
            value = visible_text(paragraph).strip()
            if not value:
                continue
            match = reference_re.fullmatch(value)
            if not match:
                reference_issues.append(
                    f"non-reference paragraph after heading: {value[:120]!r}"
                )
                break
            number = int(match.group("bracket") or match.group("plain"))
            doi_match = DOI_RE.search(value)
            references.append(
                {
                    "number": number,
                    "visible_text": value,
                    "doi": (
                        doi_match.group(0).lower().rstrip(".,;)")
                        if doi_match
                        else None
                    ),
                }
            )
            reference_elements.add(id(paragraph))
    elif not heading_matches:
        reference_issues.append("reference heading was not found")
    else:
        reference_issues.append(
            f"reference heading is ambiguous at direct paragraphs {heading_matches}"
        )

    expected_reference_numbers = list(range(1, len(references) + 1))
    actual_reference_numbers = [row["number"] for row in references]
    if references and actual_reference_numbers != expected_reference_numbers:
        reference_issues.append(
            f"reference numbering is not sequential: {actual_reference_numbers}"
        )

    static_occurrences: list[dict] = []
    bare_superscript_candidates: list[dict] = []
    for part, paragraphs in paragraph_elements.items():
        marker_counter: Counter[int] = Counter()
        bare_counter: Counter[int] = Counter()
        for index, paragraph in enumerate(paragraphs):
            if part == "word/document.xml" and id(paragraph) in reference_elements:
                continue
            filtered = text_outside_citation_fields(paragraph)
            for match in marker_re.finditer(filtered):
                numbers = parse_numeric_cluster(match.group(0))
                if not numbers:
                    continue
                marker_counter[index] += 1
                static_occurrences.append(
                    {
                        "occurrence_id": f"{part}:p{index}:m{marker_counter[index]}",
                        "part": part,
                        "paragraph_index": index,
                        "index_in_paragraph": marker_counter[index],
                        "kind": "static-marker",
                        "source_text": match.group(0),
                        "start": match.start(),
                        "end": match.end(),
                        "reference_numbers": numbers,
                        "context": filtered[max(0, match.start() - 60) : match.end() + 60],
                    }
                )
            for candidate in unsupported_bare_superscripts(
                paragraph,
                marker_re=marker_re,
            ):
                bare_counter[index] += 1
                bare_superscript_candidates.append(
                    {
                        "occurrence_id": f"{part}:p{index}:s{bare_counter[index]}",
                        "part": part,
                        "paragraph_index": index,
                        "index_in_paragraph": bare_counter[index],
                        "kind": "unsupported-bare-superscript-candidate",
                        **candidate,
                    }
                )

    counts = Counter(row["kind"] for row in all_fields)
    endnote_occurrences = [
        row for row in all_fields if row["kind"] == "endnote-citation"
    ]
    zotero_occurrences = [
        row for row in all_fields if row["kind"] == "zotero-citation"
    ]
    systems = []
    if zotero_occurrences:
        systems.append("zotero")
    if endnote_occurrences:
        systems.append("endnote")
    if static_occurrences:
        systems.append("static-numeric")
    classification = systems[0] if len(systems) == 1 else ("mixed" if systems else "none")

    # Count every detected citation system. A mixed document is blocked below,
    # but its inventory must expose all occurrences rather than silently
    # selecting one system as authoritative.
    active_occurrences = zotero_occurrences + endnote_occurrences + static_occurrences
    occurrence_counts_by_system = {
        "zotero": len(zotero_occurrences),
        "endnote": len(endnote_occurrences),
        "static-numeric": len(static_occurrences),
    }
    used_numbers = sorted(
        {
            number
            for occurrence in active_occurrences
            for number in occurrence.get("reference_numbers", [])
        }
    )
    reference_number_set = set(actual_reference_numbers)
    out_of_range = sorted(set(used_numbers) - reference_number_set)
    unused = sorted(reference_number_set - set(used_numbers))

    citation_ids = [
        row.get("citation_id")
        for row in zotero_occurrences
        if row.get("citation_id") not in (None, "")
    ]
    duplicate_citation_ids = sorted(
        value for value, count in Counter(citation_ids).items() if count > 1
    )
    payload_errors = [
        {
            "occurrence_id": row["occurrence_id"],
            "error": row.get("payload_error"),
        }
        for row in zotero_occurrences
        if row.get("payload_error")
    ]
    citation_system_issues: list[str] = []
    if len(systems) > 1:
        citation_system_issues.append(
            "mixed citation systems detected "
            f"({', '.join(systems)}); migration must stop until each occurrence "
            "is reviewed and one explicit source-system policy is chosen"
        )
    if bare_superscript_candidates:
        citation_system_issues.append(
            f"{len(bare_superscript_candidates)} native superscript numeric "
            "candidate(s) occur outside recognized citation fields/markers; "
            "bare superscripts are not auto-classified because they may be "
            "scientific notation or footnotes, so migration must stop for review"
        )


    heading_global_index = None
    prose_texts: list[str] = []
    document_paragraphs = paragraph_elements["word/document.xml"]
    if heading_element is not None:
        for index, paragraph in enumerate(document_paragraphs):
            if paragraph is heading_element:
                heading_global_index = index
                break
    for index, paragraph in enumerate(document_paragraphs):
        if heading_global_index is not None and index >= heading_global_index:
            continue
        prose = text_outside_citation_fields(paragraph)
        prose_texts.append(marker_re.sub("", prose))

    return {
        "schema_version": 1,
        "input": str(path.resolve()),
        "bytes": path.stat().st_size,
        "sha256": sha256_file(path),
        "zip": {
            "member_count": len(names),
            "duplicate_names": duplicate_names,
            "bad_member": bad_member,
        },
        "classification": classification,
        "systems_detected": systems,
        "story_parts": story_reports,
        "field_counts": dict(counts),
        "field_characters_by_part": field_characters_by_part,
        "fields": all_fields,
        "citation_payload_errors": payload_errors,
        "duplicate_citation_ids": duplicate_citation_ids,
        "references_heading_direct_paragraph": (
            heading_matches[0] if len(heading_matches) == 1 else None
        ),
        "references_heading_global_paragraph": heading_global_index,
        "reference_count": len(references),
        "references": references,
        "static_citation_occurrences": static_occurrences,
        "unsupported_bare_superscript_candidates": bare_superscript_candidates,
        "active_citation_occurrence_count": len(active_occurrences),
        "active_citation_occurrence_counts_by_system": occurrence_counts_by_system,
        "used_reference_numbers": used_numbers,
        "unused_reference_numbers": unused,
        "out_of_range_reference_numbers": out_of_range,
        "prose_noncitation_text_sha256": sha256_text("\n".join(prose_texts)),
        "structure": package_structure(document_root, names),
        "media_hashes": media_hashes,
        "custom_xml_hashes": custom_xml_hashes,
        "custom_properties": custom_properties,
        "issues": {
            "xml": xml_errors,
            "field_sequence": field_errors,
            "references": reference_issues,
            "citation_system": citation_system_issues,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("--out", type=Path)
    args = parser.parse_args()

    report = audit(args.input_docx)
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    if args.out:
        args.out.parent.mkdir(parents=True, exist_ok=True)
        args.out.write_text(rendered + "\n", encoding="utf-8")
    summary = {
        "input": report["input"],
        "classification": report["classification"],
        "field_counts": report["field_counts"],
        "reference_count": report["reference_count"],
        "active_citation_occurrence_count": report["active_citation_occurrence_count"],
        "active_citation_occurrence_counts_by_system": report[
            "active_citation_occurrence_counts_by_system"
        ],
        "unsupported_bare_superscript_candidate_count": len(
            report["unsupported_bare_superscript_candidates"]
        ),
        "used_reference_count": len(report["used_reference_numbers"]),
        "out_of_range_reference_numbers": report["out_of_range_reference_numbers"],
        "payload_errors": len(report["citation_payload_errors"]),
        "duplicate_citation_ids": report["duplicate_citation_ids"],
        "issues": report["issues"],
    }
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    raise SystemExit(main())

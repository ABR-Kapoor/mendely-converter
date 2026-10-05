import base64
import copy
import io
import json
import re
import uuid
import zipfile
from pathlib import Path

import bibtexparser
import streamlit as st
from docx import Document
from dotenv import load_dotenv
from lxml import etree
from rapidfuzz import fuzz


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
RUN_TAG = f"{{{W_NS}}}r"
TEXT_TAG = f"{{{W_NS}}}t"
PPR_TAG = f"{{{W_NS}}}pPr"


# ---------------------------------------------------------
# BIBTEX COMPATIBILITY
# ---------------------------------------------------------

def _normalize_bib_entry(entry):
    if hasattr(entry, "as_dict"):
        entry = entry.as_dict()

    if not isinstance(entry, dict):
        entry = dict(entry)

    normalized = {}

    for key in ("ID", "id"):
        if key in entry and entry[key] not in (None, ""):
            normalized["ID"] = str(entry[key])
            break

    for key in (
        "title",
        "author",
        "year",
        "journal",
        "booktitle",
        "doi",
        "url",
        "publisher",
        "pages",
        "volume",
        "number",
        "month",
        "editor",
    ):
        if key in entry and entry[key] not in (None, ""):
            normalized[key] = str(entry[key])

    if "ENTRYTYPE" in entry and entry["ENTRYTYPE"] not in (None, ""):
        normalized["ENTRYTYPE"] = str(entry["ENTRYTYPE"])

    if "ID" in normalized and "key" not in normalized:
        normalized["key"] = normalized["ID"]

    return normalized


def parse_bib(file):
    raw = file.read()

    if isinstance(raw, (bytes, bytearray)):
        raw = raw.decode("utf-8", errors="ignore")

    raw = str(raw).strip()

    if not raw:
        return []

    database = None

    if hasattr(bibtexparser, "parse_string"):
        try:
            database = bibtexparser.parse_string(raw)
        except Exception:
            database = None

    if database is None and hasattr(bibtexparser, "loads"):
        try:
            database = bibtexparser.loads(raw)
        except Exception:
            database = None

    if database is None:
        try:
            from bibtexparser import bparser

            parser = bparser.BibTexParser()
            parser.ignore_nonstandard_types = False
            database = parser.parse(raw)
        except Exception as exc:
            raise ValueError(
                "Unable to parse BibTeX content with the installed bibtexparser API. "
                f"Detected version: {getattr(bibtexparser, '__version__', 'unknown')}."
            ) from exc

    entries = getattr(database, "entries", database)

    if entries is None:
        return []

    return [_normalize_bib_entry(entry) for entry in entries]


load_dotenv()

st.set_page_config(
    page_title="Mendeley Converter",
    page_icon="📚",
    layout="wide",
)

st.title("📚 IEEE → Mendeley DOCX Converter")
st.caption("Convert manually typed IEEE [n] citations into real Mendeley Cite v3-compatible DOCX citation controls.")


# ---------------------------------------------------------
# DOCX
# ---------------------------------------------------------

def read_docx(file):
    doc = Document(file)
    paragraphs = []
    for p in doc.paragraphs:
        text = p.text.strip()
        if text:
            paragraphs.append(text)
    return paragraphs


# ---------------------------------------------------------
# FIND IN-TEXT CITATIONS
# ---------------------------------------------------------

def extract_citations(paragraphs):
    citations = []
    pattern = r"\[(\d+(?:\s*[-,]\s*\d+)*)\]"

    for paragraph in paragraphs:
        matches = re.findall(pattern, paragraph)
        for match in matches:
            numbers = re.split(r"\s*[, -]\s*", match)
            for number in numbers:
                if number.isdigit():
                    citations.append(int(number))
    return sorted(set(citations))


# ---------------------------------------------------------
# FIND REFERENCES SECTION
# ---------------------------------------------------------

def extract_references(paragraphs):
    references = {}
    started = False
    reference_pattern = re.compile(r"^\s*\[(\d+)\]\s*(.*)")

    for paragraph in paragraphs:
        lower = paragraph.lower()
        if not started and (lower == "references" or lower == "reference" or lower.startswith("references")):
            started = True
            continue
        if not started:
            continue
        match = reference_pattern.match(paragraph)
        if match:
            number = int(match.group(1))
            text = match.group(2).strip()
            references[number] = text
    return references


# ---------------------------------------------------------
# NORMALIZATION
# ---------------------------------------------------------

def normalize(text):
    text = text.lower()
    text = re.sub(r"https?://doi.org/", "", text)
    text = re.sub(r"doi:\s*", "", text)
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


# ---------------------------------------------------------
# BIB → SEARCHABLE TEXT
# ---------------------------------------------------------

def bib_text(entry):
    fields = []
    for key in ["title", "author", "year", "journal", "booktitle", "doi"]:
        if key in entry:
            fields.append(str(entry[key]))
    return " ".join(fields)


# ---------------------------------------------------------
# MATCHING
# ---------------------------------------------------------

def match_reference(word_reference, bib_entries):
    word_norm = normalize(word_reference)
    best_entry = None
    best_score = 0

    for entry in bib_entries:
        candidate = normalize(bib_text(entry))
        score = fuzz.token_set_ratio(word_norm, candidate)
        if score > best_score:
            best_score = score
            best_entry = entry
    return best_entry, best_score


# ---------------------------------------------------------
# CSL / MENDELEY CONVERSION
# ---------------------------------------------------------

def _parse_author_list(author_value):
    if not author_value:
        return []

    authors = []
    for part in str(author_value).split(" and "):
        part = part.strip()
        if not part:
            continue
        family = part
        given = ""
        if "," in part:
            family, given = [p.strip() for p in part.split(",", 1)]
        else:
            family = part
            given = ""
        authors.append({
            "family": family,
            "given": given,
        })
    return authors


def _entry_type_to_csl(entry_type):
    type_map = {
        "article": "article-journal",
        "inproceedings": "paper-conference",
        "conference": "paper-conference",
        "book": "book",
        "inbook": "chapter",
        "manual": "report",
        "misc": "article",
        "phdthesis": "thesis",
        "techreport": "report",
        "mastersthesis": "thesis",
    }
    return type_map.get(str(entry_type or "").lower(), "article")


def bib_entry_to_csl(entry):
    if not isinstance(entry, dict):
        return {}

    entry_type = entry.get("ENTRYTYPE", "article")
    item = {
        "id": entry.get("ID") or entry.get("key") or "unknown",
        "type": _entry_type_to_csl(entry_type),
    }

    if entry.get("title"):
        item["title"] = str(entry["title"]).strip()

    if entry.get("author"):
        authors = _parse_author_list(entry.get("author"))
        if authors:
            item["author"] = authors

    if entry.get("year"):
        year = str(entry["year"]).strip()
        if year:
            item["issued"] = {"date-parts": [[int(year)]]}

    if entry.get("journal"):
        item["container-title"] = str(entry["journal"]).strip()

    if entry.get("booktitle"):
        item["container-title"] = str(entry["booktitle"]).strip()

    if entry.get("doi"):
        item["DOI"] = str(entry["doi"]).strip()

    if entry.get("url"):
        item["URL"] = str(entry["url"]).strip()

    if entry.get("publisher"):
        item["publisher"] = str(entry["publisher"]).strip()

    if entry.get("volume"):
        item["volume"] = str(entry["volume"]).strip()

    if entry.get("number"):
        item["issue"] = str(entry["number"]).strip()

    if entry.get("pages"):
        item["page"] = str(entry["pages"]).strip()

    return item


def build_mendeley_citation_payload(citation_numbers, resolved_map):
    citation_items = []
    for number in citation_numbers:
        key = int(number)
        entry = resolved_map.get(key)
        if entry is None:
            entry = resolved_map.get(str(key))
        if entry is None:
            raise ValueError(f"Citation [{number}] could not be matched")
        entry_id = entry.get("ID") or entry.get("key") or str(key)
        citation_items.append({
            "id": entry_id,
            "itemData": bib_entry_to_csl(entry),
        })

    payload = {
        "citationID": str(uuid.uuid4()),
        "citationItems": citation_items,
        "properties": {"noteIndex": 0},
        "schema": "https://github.com/citation-style-language/schema/raw/master/csl-citation.json",
    }
    return payload


def create_mendeley_sdt(tag_value, cite_text):
    sdt = etree.Element(f"{{{W_NS}}}sdt")
    sdt_pr = etree.SubElement(sdt, f"{{{W_NS}}}sdtPr")
    tag = etree.SubElement(sdt_pr, f"{{{W_NS}}}tag")
    tag.set(f"{{{W_NS}}}val", tag_value)

    sdt_content = etree.SubElement(sdt, f"{{{W_NS}}}sdtContent")
    run = etree.SubElement(sdt_content, RUN_TAG)
    t = etree.SubElement(run, TEXT_TAG)
    t.text = cite_text
    if cite_text.startswith(" ") or cite_text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    return sdt


def extract_run_text(run_element):
    text_parts = []
    for node in run_element.iterfind(f".//{TEXT_TAG}"):
        if node.text:
            text_parts.append(node.text)
    return "".join(text_parts)


def find_citation_clusters(text):
    matches = list(re.finditer(r"\[(\d+)\]", text))
    if not matches:
        return []

    clusters = []
    i = 0
    while i < len(matches):
        start = matches[i].start()
        end = matches[i].end()
        numbers = [int(matches[i].group(1))]
        j = i + 1
        while j < len(matches):
            between = text[end:matches[j].start()]
            if re.fullmatch(r"\s*(?:,\s*|-\s*|–\s*|—\s*)?", between):
                numbers.append(int(matches[j].group(1)))
                end = matches[j].end()
                j += 1
            else:
                break
        clusters.append({
            "numbers": numbers,
            "start": start,
            "end": end,
            "label": text[start:end],
        })
        i = j
    return clusters


def replace_run_citations(run_element, resolved_map):
    run_text = extract_run_text(run_element)
    if "[" not in run_text:
        return [copy.deepcopy(run_element)]

    clusters = find_citation_clusters(run_text)
    if not clusters:
        return [copy.deepcopy(run_element)]

    fragments = []
    cursor = 0
    for cluster in clusters:
        start = cluster["start"]
        end = cluster["end"]

        if start > cursor:
            left_text = run_text[cursor:start]
            if left_text:
                left_run = copy.deepcopy(run_element)
                text_nodes = list(left_run.iterfind(f".//{TEXT_TAG}"))
                if text_nodes:
                    text_nodes[0].text = left_text
                    for node in text_nodes[1:]:
                        node.text = ""
                        node.tail = ""
                else:
                    t = etree.SubElement(left_run, TEXT_TAG)
                    t.text = left_text
                fragments.append(left_run)

        payload = build_mendeley_citation_payload(cluster["numbers"], resolved_map)
        json_bytes = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        encoded = base64.b64encode(json_bytes).decode("ascii")
        tag_value = f"MENDELEY_CITATION_v3_{encoded}"
        fragments.append(create_mendeley_sdt(tag_value, cluster["label"]))
        cursor = end

    if cursor < len(run_text):
        right_text = run_text[cursor:]
        if right_text:
            right_run = copy.deepcopy(run_element)
            text_nodes = list(right_run.iterfind(f".//{TEXT_TAG}"))
            if text_nodes:
                text_nodes[0].text = right_text
                for node in text_nodes[1:]:
                    node.text = ""
                    node.tail = ""
            else:
                t = etree.SubElement(right_run, TEXT_TAG)
                t.text = right_text
            fragments.append(right_run)

    return fragments


def paragraph_text(paragraph_elem):
    return "".join(paragraph_elem.itertext())


def convert_docx_to_mendeley(docx_bytes, resolved_map):
    output_buffer = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(docx_bytes)) as source_zip:
        with zipfile.ZipFile(output_buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as output_zip:
            in_references = False
            for item in source_zip.infolist():
                data = source_zip.read(item.filename)
                if item.filename == "word/document.xml":
                    root = etree.fromstring(data)
                    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
                    for para in root.iter(f"{W}p"):
                        para_text_value = paragraph_text(para)
                        trimmed = para_text_value.strip()
                        if re.match(r"^references?\b", trimmed, flags=re.IGNORECASE):
                            in_references = True
                            continue
                        if in_references:
                            continue

                        children = list(para)
                        new_children = []
                        for child in children:
                            if child.tag == RUN_TAG:
                                fragments = replace_run_citations(child, resolved_map)
                                new_children.extend(fragments)
                            else:
                                new_children.append(child)

                        para[:] = []
                        for child in new_children:
                            para.append(child)

                    data = etree.tostring(root, encoding="utf-8", xml_declaration=True, pretty_print=False)
                output_zip.writestr(item, data)
    return output_buffer.getvalue()


# ---------------------------------------------------------
# VALIDATION AND UI
# ---------------------------------------------------------

def analyze_uploaded_documents(docx_file, bib_file):
    paragraphs = read_docx(docx_file)
    citations = extract_citations(paragraphs)
    references = extract_references(paragraphs)
    bib_entries = parse_bib(bib_file)

    resolved_map = {}
    unresolved = []
    for number in citations:
        word_reference = references.get(number, "")
        if not word_reference:
            unresolved.append(number)
            continue
        entry, score = match_reference(word_reference, bib_entries)
        if entry is not None:
            resolved_map[number] = entry
        else:
            unresolved.append(number)

    results = []
    for number in citations:
        word_reference = references.get(number, "")
        if not word_reference:
            results.append({
                "Citation": f"[{number}]",
                "Word Reference": "NOT FOUND",
                "BibTeX Key": "—",
                "Score": 0,
                "Status": "❌ Missing",
            })
            continue
        entry, score = match_reference(word_reference, bib_entries)
        if entry is not None:
            status = "✓ Strong" if score >= 85 else "⚠ Review"
            results.append({
                "Citation": f"[{number}]",
                "Word Reference": word_reference,
                "BibTeX Key": entry.get("ID", "") or entry.get("key", ""),
                "Score": round(score),
                "Status": status,
            })
        else:
            results.append({
                "Citation": f"[{number}]",
                "Word Reference": word_reference,
                "BibTeX Key": "—",
                "Score": 0,
                "Status": "❌ No match",
            })

    return {
        "paragraphs": paragraphs,
        "citations": citations,
        "references": references,
        "bib_entries": bib_entries,
        "resolved_map": resolved_map,
        "unresolved": unresolved,
        "results": results,
    }


# ---------------------------------------------------------
# UI
# ---------------------------------------------------------

docx_file = st.file_uploader("Upload Word document", type=["docx"])
bib_file = st.file_uploader("Upload BibTeX file", type=["bib"])

if docx_file and bib_file:
    if st.button("Analyze citations", type="primary"):
        analysis = analyze_uploaded_documents(docx_file, bib_file)
        st.session_state["analysis"] = analysis
        st.success("Analysis complete.")

        col1, col2, col3 = st.columns(3)
        col1.metric("In-text citations", len(analysis["citations"]))
        col2.metric("Word references", len(analysis["references"]))
        col3.metric("BibTeX entries", len(analysis["bib_entries"]))

        st.divider()
        st.subheader("Citation mapping")
        st.dataframe(analysis["results"], use_container_width=True)

        unresolved = analysis["unresolved"]
        if unresolved:
            st.warning("Unresolved citations found before conversion:")
            for number in unresolved:
                st.write(f"❌ [{number}] could not be matched")

        if analysis["resolved_map"]:
            st.subheader("Citation conversion")
            for number, entry in sorted(analysis["resolved_map"].items()):
                st.write(f"[{number}] → {entry.get('ID') or entry.get('key')} ✓")
            st.success(f"{len(analysis['resolved_map'])}/{len(analysis['citations'])} citations resolved for DOCX conversion.")

    if "analysis" in st.session_state:
        analysis = st.session_state["analysis"]
        unresolved = analysis["unresolved"]
        if not unresolved and analysis["resolved_map"]:
            if st.button("Convert to Mendeley DOCX", type="primary"):
                with st.spinner("Generating Mendeley Cite v3 DOCX..."):
                    converted = convert_docx_to_mendeley(docx_file.getvalue(), analysis["resolved_map"])
                st.success("✓ Mendeley Cite v3 structures created")
                st.download_button(
                    label="Download generated DOCX",
                    data=converted,
                    file_name="paper_mendeley.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        else:
            st.info("Fix unresolved citations before generating the converted DOCX.")
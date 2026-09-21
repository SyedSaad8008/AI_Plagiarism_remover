
from __future__ import annotations

import argparse
import json
import re
import sys
import textwrap
from dataclasses import dataclass
from pathlib import Path

import requests
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree


@dataclass
class Flag:
    """One issue detected in a paragraph."""

    paragraph_index: int
    paragraph_text: str
    issue_type: str
    explanation: str
    suggested_rewrite: str


SYSTEM_PROMPT = textwrap.dedent("""\
    You are a writing-quality assistant reviewing a paragraph.

    Analyse the paragraph for **one specific category only**:

    1. FILLER_PHRASING -- Generic, vague, or padded phrasing that weakens the writing (e.g. "It is important to note that", "In today's rapidly changing
       world", "This paper aims to explore", "Needless to say", "It goes without
       saying", "It is widely accepted that", "In conclusion, this essay has
       shown", "Furthermore, it is evident that", "This highlights the importance
       of", "It is worth mentioning that", "It should be noted that",
       "In summary, this paper has demonstrated", "As we can see",
       "It is clear that").

    Respond with **valid JSON only** -- no prose, no markdown fences -- in this
    exact schema:

    {
      "flags": [
        {
          "issue_type": "filler_phrasing",
          "explanation": "<one concise sentence explaining the problem>",
          "suggested_rewrite": "<rewritten version of the flagged sentence/phrase only>"
        }
      ]
    }

    If the paragraph has NO issues, respond with {"flags": []}.
    Do NOT flag text that is clearly a heading, a figure caption, or a pure
    equation / formula.
""")


def call_lm_studio(
    paragraph: str,
    model: str,
    base_url: str,
    timeout: int = 60,
) -> list[dict]:
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": f"Paragraph:\n\n{paragraph}"},
        ],
        "temperature": 0.1,
        "max_tokens": 512,
        "response_format": {"type": "json_object"},
    }

    try:
        resp = requests.post(url, json=payload, timeout=timeout)
        resp.raise_for_status()
    except requests.exceptions.ConnectionError:
        print(
            f"  [ERROR] Cannot connect to server at {base_url}.\n"
            "          Is the server running?",
            file=sys.stderr,
        )
        sys.exit(1)
    except requests.exceptions.HTTPError as exc:
        print(f"  [WARN] HTTP {exc.response.status_code} for paragraph -- skipping.", file=sys.stderr)
        return []
    except requests.exceptions.Timeout:
        print("  [WARN] Request timed out -- skipping paragraph.", file=sys.stderr)
        return []

    try:
        content = resp.json()["choices"][0]["message"]["content"]
        data = json.loads(content)
        return data.get("flags", [])
    except (KeyError, json.JSONDecodeError, IndexError) as exc:
        print(f"  [WARN] Unexpected model response ({exc}) -- skipping paragraph.", file=sys.stderr)
        return []


_comment_id_counter: list[int] = [0]


def _next_comment_id() -> int:
    _comment_id_counter[0] += 1
    return _comment_id_counter[0]


def _build_comment_element(doc: Document, comment_id: int, author: str, text: str) -> etree._Element:
    comment = OxmlElement("w:comment")
    comment.set(qn("w:id"), str(comment_id))
    comment.set(qn("w:author"), author)
    comment.set(qn("w:date"), "2024-01-01T00:00:00Z")
    comment.set(qn("w:initials"), "DC")

    para = OxmlElement("w:p")
    run = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = text
    run.append(t)
    para.append(run)
    comment.append(para)
    return comment


def _ensure_comments_part(doc: Document) -> etree._Element:
    part = doc.part
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    try:
        comments_part = part.part_related_by(rel_type)
        return comments_part._element
    except KeyError:
        pass

    from docx.opc.part import Part
    from docx.opc.packuri import PackURI

    comments_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:comments'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
        ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '</w:comments>'
    )
    xml_bytes = comments_xml.encode("utf-8")
    comments_part_obj = Part(
        PackURI("/word/comments.xml"),
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml",
        xml_bytes,
        doc.part.package,
    )
    part.relate_to(comments_part_obj, rel_type)
    return etree.fromstring(xml_bytes)


def add_comment_to_paragraph(
    doc: Document,
    paragraph,
    comment_text: str,
    author: str = "Reviewer",
) -> None:
    comment_id = _next_comment_id()
    p_elem = paragraph._p

    comments_root = _ensure_comments_part(doc)
    comment_elem = _build_comment_element(doc, comment_id, author, comment_text)
    comments_root.append(comment_elem)

    range_start = OxmlElement("w:commentRangeStart")
    range_start.set(qn("w:id"), str(comment_id))
    p_elem.insert(0, range_start)

    range_end = OxmlElement("w:commentRangeEnd")
    range_end.set(qn("w:id"), str(comment_id))
    p_elem.append(range_end)

    ref_run = OxmlElement("w:r")
    comment_ref = OxmlElement("w:commentReference")
    comment_ref.set(qn("w:id"), str(comment_id))
    ref_run.append(comment_ref)
    p_elem.append(ref_run)


ISSUE_EMOJI = {
    "filler_phrasing": "NOTE: FILLER PHRASING",
}

ISSUE_EMOJI_TERMINAL = {
    "filler_phrasing": "FILLER PHRASING",
}


def format_comment(flags: list[dict]) -> str:
    parts: list[str] = []
    for flag in flags:
        label = ISSUE_EMOJI.get(flag["issue_type"], flag["issue_type"].upper())
        explanation = flag.get("explanation", "")
        rewrite = flag.get("suggested_rewrite", "")
        block = f"[{label}]\n{explanation}"
        if rewrite:
            block += f"\n\nSuggested rewrite:\n{rewrite}"
        parts.append(block)
    separator = "\n\n" + "-" * 40 + "\n\n"
    return "\n" + separator.join(parts) + "\n"


def scan_document(
    input_path: Path,
    output_path: Path,
    model: str,
    base_url: str,
    min_words: int,
    dry_run: bool,
) -> None:
    doc = Document(str(input_path))
    paragraphs = doc.paragraphs

    total = len(paragraphs)
    flagged_count = 0

    print(f"\n  Scanning '{input_path.name}' -- {total} paragraph(s) found.")
    print(f"  Model : {model}")
    print(f"  Server: {base_url}")
    if dry_run:
        print("  DRY RUN -- no file will be written.")
    print()

    all_flags: list[Flag] = []

    for idx, para in enumerate(paragraphs):
        text = para.text.strip()
        word_count = len(text.split())

        if word_count < min_words:
            continue

        preview = text[:80] + ("..." if len(text) > 80 else "")
        print(f"  [{idx + 1:>4}/{total}] {preview}")

        raw_flags = call_lm_studio(text, model, base_url)

        if not raw_flags:
            print("            -> clean")
            continue

        flagged_count += 1
        print(f"            -> {len(raw_flags)} flag(s)")
        for f in raw_flags:
            label = ISSUE_EMOJI_TERMINAL.get(f.get("issue_type", ""), f.get("issue_type", "").upper())
            print(f"               * [{label}] {f.get('explanation', '')}")

        flag_obj = Flag(
            paragraph_index=idx,
            paragraph_text=text,
            issue_type=raw_flags[0]["issue_type"],
            explanation="; ".join(f.get("explanation", "") for f in raw_flags),
            suggested_rewrite=raw_flags[0].get("suggested_rewrite", ""),
        )
        all_flags.append(flag_obj)

        if not dry_run:
            comment_text = format_comment(raw_flags)
            add_comment_to_paragraph(doc, para, comment_text)

    # Summary.
    print(f"\n{'=' * 60}")
    print(f"  Scan complete -- {flagged_count} / {total} paragraph(s) flagged.")
    print(f"{'=' * 60}\n")

    if dry_run:
        if not all_flags:
            print("  No issues found -- document looks clean!")
        else:
            print("  Issues found (dry-run summary):\n")
            for f in all_flags:
                print(f"  Paragraph {f.paragraph_index + 1}: [{f.issue_type.upper()}]")
                print(f"    Text   : {f.paragraph_text[:120]}")
                print(f"    Reason : {f.explanation}\n")
        return

    doc.save(str(output_path))
    print(f"  Saved -> {output_path}")
    print(
        "\n  Next steps:\n"
        "    1. Open the output file in Microsoft Word.\n"
        "    2. Go to Review -> Show Comments (or Next Comment).\n"
        "    3. For each comment:\n"
        "         - Filler phrase  -> rewrite in your own voice\n"
        "         - False positive -> just delete the comment\n"
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="doc_flagger",
        description="Scan a .docx file for filler phrasing and annotate it with Word review comments.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples:
              python ai_plagiarism_flagger.py paper.docx paper_flagged.docx \\
                  --model "qwen2.5-14b-instruct"

              python ai_plagiarism_flagger.py essay.docx essay_flagged.docx \\
                  --model "llama-3.1-8b-instruct" --dry-run --min-words 12
        """),
    )
    parser.add_argument("input", type=Path, help="Input .docx file")
    parser.add_argument("output", type=Path, help="Output .docx file (with comments added)")
    parser.add_argument(
        "--model",
        required=True,
        help='LM Studio model identifier, e.g. "qwen2.5-14b-instruct"',
    )
    parser.add_argument(
        "--base-url",
        default="http://localhost:1234/v1",
        help="LM Studio server base URL (default: http://localhost:1234/v1)",
    )
    parser.add_argument(
        "--min-words",
        type=int,
        default=8,
        help="Minimum word count to bother scanning a paragraph (default: 8)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Print flags to terminal only; do not write the output file",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path: Path = args.input
    output_path: Path = args.output

    if not input_path.exists():
        parser.error(f"Input file not found: {input_path}")
    if input_path.suffix.lower() != ".docx":
        parser.error(f"Input must be a .docx file, got: {input_path.suffix}")

    scan_document(
        input_path=input_path,
        output_path=output_path,
        model=args.model,
        base_url=args.base_url,
        min_words=args.min_words,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()

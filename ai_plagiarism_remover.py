
from __future__ import annotations

import argparse
import json
import sys
import textwrap
from dataclasses import dataclass, field
from pathlib import Path

import requests
from docx import Document
from docx.oxml.ns import qn
from lxml import etree


DETECT_PROMPT = textwrap.dedent("""\
You are a writing-quality assistant.

Analyse the paragraph below for EXACTLY this one category:

1. FILLER_PHRASING — generic, vague, or padded phrasing that weakens the writing
   (e.g. "It is important to note that", "In today's rapidly changing world",
   "This paper aims to explore", "Needless to say", "It goes without saying",
   "It is widely accepted that", "In conclusion, this essay has shown",
   "Furthermore, it is evident that", "This highlights the importance of",
   "It is worth mentioning that", "It should be noted that",
   "In summary, this paper has demonstrated", "As we can see",
   "It is clear that", "It goes without saying that").

Reply with VALID JSON ONLY — no prose, no markdown fences:

{
  "flags": [
    {
      "issue_type": "filler_phrasing",
      "flagged_phrase": "<exact phrase or sentence that is problematic>",
      "explanation": "<one concise sentence>"
    }
  ]
}

If the paragraph has NO issues, reply: {"flags": []}
Do NOT flag headings, figure captions, or pure equations/formulas.
""")

REWRITE_PROMPT = textwrap.dedent("""\
You are a writing editor.

Rewrite the paragraph below so that it:
  • Removes or replaces all vague filler phrases with direct, precise language.
  • Does NOT alter the factual content or the author's core argument.
  • Reads naturally and clearly.
  • Is approximately the same length as the original.

Reply with the rewritten paragraph ONLY — no extra commentary, no JSON,
no markdown fences.  Just the plain rewritten paragraph text.
""")


@dataclass
class Issue:
    issue_type: str
    flagged_phrase: str
    explanation: str


@dataclass
class ParagraphResult:
    index: int
    original: str
    issues: list[Issue] = field(default_factory=list)
    rewritten: str = ""

    @property
    def was_changed(self) -> bool:
        return bool(self.rewritten) and self.rewritten.strip() != self.original.strip()


def _post(url: str, payload: dict, timeout: int, base_url: str) -> dict | None:
    try:
        resp = requests.post(url, json=payload, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    except requests.exceptions.ConnectionError:
        print(
            f"\n[ERROR] Cannot reach server at {base_url}.\n"
            "        Make sure the server is running.",
            file=sys.stderr,
        )
        sys.exit(1)
    except requests.exceptions.HTTPError as exc:
        print(f"  [WARN] HTTP {exc.response.status_code} — skipping.", file=sys.stderr)
        return None
    except requests.exceptions.Timeout:
        print("  [WARN] Request timed out — skipping.", file=sys.stderr)
        return None


def detect_issues(paragraph: str, model: str, base_url: str) -> list[Issue]:
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": DETECT_PROMPT},
            {"role": "user",   "content": f"Paragraph:\n\n{paragraph}"},
        ],
        "temperature": 0.1,
        "max_tokens": 512,
        "response_format": {"type": "json_object"},
    }
    data = _post(url, payload, timeout=60, base_url=base_url)
    if data is None:
        return []
    try:
        content = data["choices"][0]["message"]["content"]
        flags = json.loads(content).get("flags", [])
        return [
            Issue(
                issue_type=f.get("issue_type", "unknown"),
                flagged_phrase=f.get("flagged_phrase", ""),
                explanation=f.get("explanation", ""),
            )
            for f in flags
        ]
    except (KeyError, json.JSONDecodeError, IndexError) as exc:
        print(f"  [WARN] Bad detect response ({exc}).", file=sys.stderr)
        return []


def rewrite_paragraph(paragraph: str, issues: list[Issue], model: str, base_url: str) -> str:
    issue_summary = "\n".join(
        f"  - [{i.issue_type.upper()}] {i.flagged_phrase!r}: {i.explanation}"
        for i in issues
    )
    user_msg = (
        f"Issues detected in this paragraph:\n{issue_summary}\n\n"
        f"Original paragraph:\n\n{paragraph}"
    )
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": REWRITE_PROMPT},
            {"role": "user",   "content": user_msg},
        ],
        "temperature": 0.3,
        "max_tokens": 1024,
    }
    data = _post(url, payload, timeout=90, base_url=base_url)
    if data is None:
        return paragraph
    try:
        return data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError):
        return paragraph


def _replace_paragraph_text(para, new_text: str) -> None:
    """
    Replace all text runs in a paragraph with a single run containing new_text.
    Paragraph-level formatting (style, spacing, indentation) is preserved.
    The first existing run's character formatting (font, bold, etc.) is kept
    for the replacement run; all other runs are removed.
    """
    p_elem = para._p
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Collect all <w:r> elements.
    runs = p_elem.findall(f"{{{W}}}r")

    if runs:
        first_run = runs[0]
        rpr = first_run.find(f"{{{W}}}rPr")

        new_run = etree.SubElement(p_elem, f"{{{W}}}r")
        if rpr is not None:
            import copy
            new_run.insert(0, copy.deepcopy(rpr))
        t_elem = etree.SubElement(new_run, f"{{{W}}}t")
        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_elem.text = new_text

        for r in runs:
            p_elem.remove(r)

        ppr = p_elem.find(f"{{{W}}}pPr")
        if ppr is not None:
            ppr.addnext(new_run)
        else:
            p_elem.insert(0, new_run)
    else:
        new_run = etree.SubElement(p_elem, f"{{{W}}}r")
        t_elem = etree.SubElement(new_run, f"{{{W}}}t")
        t_elem.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t_elem.text = new_text

LABEL = {
    "filler_phrasing": "FILLER PHRASING",
}


def process_file(
    input_path: Path,
    output_dir: Path,
    model: str,
    base_url: str,
    min_words: int,
    dry_run: bool,
) -> None:
    print(f"\n  {'─' * 60}")
    print(f"  Processing : {input_path.name}")
    print(f"  {'─' * 60}")

    doc = Document(str(input_path))
    paragraphs = doc.paragraphs
    total = len(paragraphs)
    results: list[ParagraphResult] = []

    for idx, para in enumerate(paragraphs):
        text = para.text.strip()
        if len(text.split()) < min_words:
            continue

        preview = (text[:75] + "...") if len(text) > 75 else text
        print(f"  [{idx + 1:>4}/{total}] {preview}")

        issues = detect_issues(text, model, base_url)

        if not issues:
            print("            -> clean")
            continue

        print(f"            -> {len(issues)} issue(s) detected:")
        for issue in issues:
            print(f"               * [{LABEL.get(issue.issue_type, issue.issue_type.upper())}]"
                  f" {issue.explanation}")

        print("            -> rewriting ...", end="", flush=True)
        rewritten = rewrite_paragraph(text, issues, model, base_url)
        print(" done")

        result = ParagraphResult(
            index=idx,
            original=text,
            issues=issues,
            rewritten=rewritten,
        )
        results.append(result)

        if not dry_run and result.was_changed:
            _replace_paragraph_text(para, rewritten)

    # Derive output filenames.
    stem = input_path.stem
    out_docx = output_dir / f"{stem}_cleaned.docx"

    changed_count = sum(1 for r in results if r.was_changed)
    flagged_count = sum(1 for r in results if r.issues)

    print(f"\n  Summary: {flagged_count} flagged, {changed_count} rewritten out of {total} paragraph(s).")

    if dry_run:
        print("\n  [DRY RUN] No files written.")
        return

    doc.save(str(out_docx))
    print(f"  Saved  -> {out_docx}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="doc_cleaner",
        description="Clean .docx files using a local LM Studio model.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=textwrap.dedent("""\
            Examples:
              python ai_plagiarism_remover.py --model "qwen2.5-14b-instruct"
              python ai_plagiarism_remover.py --model "llama-3.1-8b-instruct" --dry-run
              python ai_plagiarism_remover.py --model "qwen2.5-14b-instruct" \\
                  --input-dir ./my_papers --output-dir ./results
        """),
    )
    parser.add_argument(
        "--model", required=True,
        help='LM Studio model identifier, e.g. "qwen2.5-14b-instruct"',
    )
    parser.add_argument(
        "--base-url", default="http://localhost:1234/v1",
        help="LM Studio server base URL (default: http://localhost:1234/v1)",
    )
    parser.add_argument(
        "--min-words", type=int, default=8,
        help="Minimum word count to scan a paragraph (default: 8)",
    )
    parser.add_argument(
        "--input-dir", type=Path, default=Path("./input"),
        help="Folder containing input .docx files (default: ./input)",
    )
    parser.add_argument(
        "--output-dir", type=Path, default=Path("./output"),
        help="Folder for cleaned .docx (default: ./output)",
    )
    parser.add_argument(
        "--dry-run", action="store_true",
        help="Preview only; do not write any files",
    )
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_dir: Path  = args.input_dir
    output_dir: Path = args.output_dir

    if not input_dir.exists():
        input_dir.mkdir(parents=True)
        print(f"Created input directory: {input_dir.resolve()}")
        print("Place your .docx files in that folder and run again.")
        sys.exit(0)

    if not args.dry_run:
        output_dir.mkdir(parents=True, exist_ok=True)

    docx_files = sorted(input_dir.glob("*.docx"))
    if not docx_files:
        print(f"No .docx files found in {input_dir.resolve()}")
        sys.exit(0)

    print(f"\n{'=' * 62}")
    print(f"  DOCUMENT CLEANER")
    print(f"  Model    : {args.model}")
    print(f"  Server   : {args.base_url}")
    print(f"  Input    : {input_dir.resolve()}")
    print(f"  Output   : {output_dir.resolve()}")
    print(f"  Files    : {len(docx_files)} .docx file(s) found")
    if args.dry_run:
        print(f"  DRY RUN  : yes — no files will be written")
    print(f"{'=' * 62}")

    for docx_path in docx_files:
        process_file(
            input_path=docx_path,
            output_dir=output_dir,
            model=args.model,
            base_url=args.base_url,
            min_words=args.min_words,
            dry_run=args.dry_run,
        )

    print(f"\n{'=' * 62}")
    print("  All files processed.")
    if not args.dry_run:
        print(f"  Results are in: {output_dir.resolve()}")
    print(f"{'=' * 62}\n")


if __name__ == "__main__":
    main()

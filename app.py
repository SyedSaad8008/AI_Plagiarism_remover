

from __future__ import annotations

import copy
import json
import queue
import re
import sys
import textwrap
import threading
from dataclasses import dataclass, field

from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog

import requests
from docx import Document
from docx.oxml.ns import qn
from lxml import etree


# ─────────────────────────────────────────────────────────────────────────────
# Colour palette  (dark violet/slate theme)
C = {
    "bg":        "#0d0d1a",
    "panel":     "#13132b",
    "border":    "#2a2a4a",
    "accent":    "#7c3aed",
    "accent_h":  "#9d5cf6",
    "btn_start": "#059669",
    "btn_stop":  "#dc2626",
    "btn_open":  "#1d4ed8",
    "fg":        "#e2e8f0",
    "fg_dim":    "#64748b",
    "entry_bg":  "#1e1e38",
    "log_header":  "#a78bfa",
    "log_info":    "#94a3b8",
    "log_clean":   "#34d399",
    "log_flag":    "#fbbf24",
    "log_rewrite": "#60a5fa",
    "log_error":   "#f87171",
    "log_dim":     "#475569",
    "log_label":   "#f472b6",
}

FONT_UI    = ("Segoe UI",       10)
FONT_BOLD  = ("Segoe UI",       10, "bold")
FONT_HEAD  = ("Segoe UI",       13, "bold")
FONT_TITLE = ("Segoe UI",       16, "bold")
FONT_MONO  = ("Consolas",       10)
FONT_MONO_S= ("Consolas",        9)


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
      "flagged_phrase": "<exact phrase that is problematic>",
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
no markdown fences. Just the plain rewritten paragraph text.
""")

LABEL_MAP = {
    "filler_phrasing": "FILLER PHRASING",
}


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


def _extract_json(text: str) -> dict:
    text = text.strip()
    m = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    if m:
        text = m.group(1)
    m = re.search(r"\{.*\}", text, re.DOTALL)
    if m:
        text = m.group(0)
    return json.loads(text)


def _post(url: str, payload: dict, timeout: int) -> dict | None:
    resp = requests.post(url, json=payload, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def detect_issues(
    paragraph: str,
    model: str,
    base_url: str,
    q: queue.Queue,
    stop: threading.Event,
) -> list[Issue] | None:
    if stop.is_set():
        return []
    url = base_url.rstrip("/") + "/chat/completions"
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": DETECT_PROMPT},
            {"role": "user",   "content": f"Paragraph:\n\n{paragraph}"},
        ],
        "temperature": 0.1,
        "max_tokens": 512,
    }
    try:
        data = _post(url, payload, timeout=60)
        content = data["choices"][0]["message"]["content"]
        flags = _extract_json(content).get("flags", [])
        return [
            Issue(
                issue_type=f.get("issue_type", "unknown"),
                flagged_phrase=f.get("flagged_phrase", ""),
                explanation=f.get("explanation", ""),
            )
            for f in flags
        ]
    except requests.exceptions.ConnectionError:
        q.put({"type": "fatal",
               "text": f"Cannot connect to server at {base_url}.\n"
                       "Make sure the server is running."})
        return None
    except requests.exceptions.Timeout:
        q.put({"type": "log", "text": "  [timeout] — skipping paragraph.", "tag": "error"})
        return []
    except Exception as exc:
        q.put({"type": "log", "text": f"  [detect error] {exc}", "tag": "error"})
        return []


def rewrite_paragraph(
    paragraph: str,
    issues: list[Issue],
    model: str,
    base_url: str,
    q: queue.Queue,
    stop: threading.Event,
) -> str:
    if stop.is_set():
        return paragraph
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
    try:
        data = _post(url, payload, timeout=90)
        return data["choices"][0]["message"]["content"].strip()
    except requests.exceptions.Timeout:
        q.put({"type": "log", "text": "  [timeout on rewrite] — keeping original.", "tag": "error"})
        return paragraph
    except Exception as exc:
        q.put({"type": "log", "text": f"  [rewrite error] {exc}", "tag": "error"})
        return paragraph


def _replace_paragraph_text(para, new_text: str) -> None:
    """Replace all text runs, preserving paragraph-level style."""
    p_elem = para._p
    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    runs = p_elem.findall(f"{{{W}}}r")

    if runs:
        first_run = runs[0]
        rpr = first_run.find(f"{{{W}}}rPr")
        new_run = etree.SubElement(p_elem, f"{{{W}}}r")
        if rpr is not None:
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



def worker(
    input_dir: Path,
    output_dir: Path,
    model: str,
    base_url: str,
    min_words: int,
    dry_run: bool,
    q: queue.Queue,
    stop: threading.Event,
) -> None:

    def log(text: str, tag: str = "info") -> None:
        q.put({"type": "log", "text": text, "tag": tag})

    def progress(value: float, label: str = "") -> None:
        q.put({"type": "progress", "value": value, "label": label})

    def stats(flagged: int, rewritten: int, clean: int, total: int) -> None:
        q.put({"type": "stats",
               "flagged": flagged, "rewritten": rewritten,
               "clean": clean, "total": total})

    try:
        output_dir.mkdir(parents=True, exist_ok=True)
        docx_files = sorted(input_dir.glob("*.docx"))

        if not docx_files:
            log(f"No .docx files found in: {input_dir.resolve()}", "error")
            q.put({"type": "done", "success": False})
            return

        total_files = len(docx_files)
        log(f"Found {total_files} file(s) in {input_dir.resolve()}", "info")

        for file_idx, docx_path in enumerate(docx_files):
            if stop.is_set():
                log("Stopped by user.", "error")
                break

            log("", "dim")
            log(f"{'─' * 58}", "dim")
            log(f"  FILE {file_idx + 1}/{total_files}:  {docx_path.name}", "header")
            log(f"{'─' * 58}", "dim")

            doc = Document(str(docx_path))
            paragraphs = doc.paragraphs
            total_paras = len(paragraphs)
            results: list[ParagraphResult] = []
            scanned = 0
            g_flagged = g_rewritten = g_clean = 0

            for idx, para in enumerate(paragraphs):
                if stop.is_set():
                    break

                text = para.text.strip()
                if len(text.split()) < min_words:
                    continue

                scanned += 1
                pct = (file_idx + (idx / max(total_paras, 1))) / total_files
                label = f"{docx_path.name}  ·  Para {idx + 1}/{total_paras}"
                progress(pct, label)

                preview = (text[:72] + "…") if len(text) > 72 else text
                log(f"  [{idx + 1:>4}/{total_paras}]  {preview}", "info")

                issues = detect_issues(text, model, base_url, q, stop)
                if issues is None:
                    q.put({"type": "done", "success": False})
                    return
                if not issues:
                    log("             ✓  clean", "clean")
                    g_clean += 1
                    continue

                g_flagged += 1
                log(f"             ⚑  {len(issues)} issue(s):", "flag")
                for issue in issues:
                    lbl = LABEL_MAP.get(issue.issue_type, issue.issue_type.upper())
                    log(f"                [{lbl}]  {issue.explanation}", "label")
                    if issue.flagged_phrase:
                        log(f'                  phrase: "{issue.flagged_phrase}"', "dim")

                log("             ✏  rewriting …", "rewrite")
                rewritten = rewrite_paragraph(text, issues, model, base_url, q, stop)
                if rewritten != text:
                    log(f"             ✔  done  →  {(rewritten[:65] + '…') if len(rewritten) > 65 else rewritten}",
                        "rewrite")
                    g_rewritten += 1
                else:
                    log("             !  rewrite unchanged (original kept)", "error")

                result = ParagraphResult(index=idx, original=text,
                                         issues=issues, rewritten=rewritten)
                results.append(result)

                if not dry_run and result.was_changed:
                    _replace_paragraph_text(para, rewritten)

                stats(g_flagged, g_rewritten, g_clean, scanned)

            stem = docx_path.stem
            out_docx = output_dir / f"{stem}_cleaned.docx"

            log("", "dim")
            log(f"  Summary:  {g_flagged} flagged  |  {g_rewritten} rewritten  |  {g_clean} clean  (of {scanned} scanned)", "header")

            if not dry_run:
                doc.save(str(out_docx))
                log(f"  Saved →  {out_docx.name}", "clean")
                q.put({"type": "file_done",
                       "docx": str(out_docx)})
            else:
                log("  [DRY RUN] — no files written.", "flag")

        progress(1.0, "Done")
        log("", "dim")
        log("═" * 58, "dim")
        log("  ALL FILES PROCESSED", "header")
        log("═" * 58, "dim")
        q.put({"type": "done", "success": True, "output_dir": str(output_dir)})

    except Exception as exc:
        import traceback
        q.put({"type": "log", "text": f"[UNEXPECTED ERROR] {exc}", "tag": "error"})
        q.put({"type": "log", "text": traceback.format_exc(), "tag": "error"})
        q.put({"type": "done", "success": False})


def _btn(parent, text, command, bg, fg="#ffffff", width=None, **kw):
    cfg = dict(
        text=text, command=command,
        bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
        relief="flat", bd=0, cursor="hand2",
        font=FONT_BOLD, padx=14, pady=6,
    )
    if width:
        cfg["width"] = width
    cfg.update(kw)
    b = tk.Button(parent, **cfg)
    def _on(e):  b.config(bg=_lighten(bg))
    def _off(e): b.config(bg=bg)
    b.bind("<Enter>", _on)
    b.bind("<Leave>", _off)
    return b


def _lighten(hex_color: str, amount: int = 20) -> str:
    hex_color = hex_color.lstrip("#")
    r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
    r, g, b = min(255, r + amount), min(255, g + amount), min(255, b + amount)
    return f"#{r:02x}{g:02x}{b:02x}"


def _label(parent, text, font=None, fg=None, bg=None, **kw):
    return tk.Label(parent, text=text,
                    font=font or FONT_UI,
                    fg=fg or C["fg"],
                    bg=bg or C["panel"],
                    **kw)


def _entry(parent, textvariable, width=38):
    e = tk.Entry(parent, textvariable=textvariable, width=width,
                 bg=C["entry_bg"], fg=C["fg"],
                 insertbackground=C["fg"],
                 relief="flat", bd=0, font=FONT_UI,
                 highlightthickness=1,
                 highlightcolor=C["accent"],
                 highlightbackground=C["border"])
    return e


def _separator(parent, bg=None):
    return tk.Frame(parent, height=1, bg=bg or C["border"])


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        root.title("Document Cleaner")
        root.configure(bg=C["bg"])
        root.geometry("860x780")
        root.minsize(700, 600)

        self._q: queue.Queue = queue.Queue()
        self._stop_event: threading.Event = threading.Event()
        self._running = False
        self._last_docx: str = ""

        self.var_input  = tk.StringVar(value=str(Path("./input").resolve()))
        self.var_output = tk.StringVar(value=str(Path("./output").resolve()))
        self.var_model  = tk.StringVar(value="qwen3-8b")
        self.var_url    = tk.StringVar(value="http://localhost:1234/v1")
        self.var_minw   = tk.IntVar(value=8)
        self.var_dryrun = tk.BooleanVar(value=False)

        self.var_flagged   = tk.StringVar(value="0")
        self.var_rewritten = tk.StringVar(value="0")
        self.var_clean     = tk.StringVar(value="0")
        self.var_total     = tk.StringVar(value="0")

        self._build_ui()

    def _build_ui(self):
        self._build_header()
        self._build_settings()
        self._build_actions()
        self._build_stats_bar()
        self._build_progress()
        self._build_log()
        self._build_statusbar()

    def _build_header(self):
        hdr = tk.Frame(self.root, bg=C["accent"], height=54)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="  🧹  Document Cleaner",
                 font=FONT_TITLE, fg="#ffffff", bg=C["accent"],
                 anchor="w").pack(side="left", padx=18, pady=10)
        tk.Label(hdr, text="powered by LM Studio",
                 font=("Segoe UI", 9), fg="#d4b8ff", bg=C["accent"]
                 ).pack(side="right", padx=18)

    def _build_settings(self):
        outer = tk.Frame(self.root, bg=C["bg"], padx=14, pady=10)
        outer.pack(fill="x")

        card = tk.Frame(outer, bg=C["panel"], padx=16, pady=14,
                        highlightthickness=1, highlightbackground=C["border"])
        card.pack(fill="x")

        _label(card, "⚙  Settings", font=FONT_HEAD, fg=C["accent"],
               bg=C["panel"]).grid(row=0, column=0, columnspan=3,
                                    sticky="w", pady=(0, 10))

        rows = [
            ("Input folder",  self.var_input,  "folder"),
            ("Output folder", self.var_output, "folder"),
            ("Model ID",      self.var_model,  None),
            ("LM Studio URL", self.var_url,    None),
        ]
        for r, (lbl, var, kind) in enumerate(rows, start=1):
            _label(card, lbl + ":", bg=C["panel"],
                   fg=C["fg_dim"]).grid(row=r, column=0, sticky="w", pady=3)
            ent = _entry(card, var, width=44)
            ent.grid(row=r, column=1, sticky="ew", padx=8, pady=3)
            card.columnconfigure(1, weight=1)
            if kind == "folder":
                _var = var
                btn = _btn(card, "Browse…",
                           lambda v=_var: self._browse(v),
                           bg=C["border"], width=8)
                btn.grid(row=r, column=2, sticky="w")
            self._settings_entries_append(ent)

        opt_row = tk.Frame(card, bg=C["panel"])
        opt_row.grid(row=len(rows) + 1, column=0, columnspan=3,
                     sticky="w", pady=(8, 0))
        _label(opt_row, "Min words:", bg=C["panel"],
               fg=C["fg_dim"]).pack(side="left")
        sp = tk.Spinbox(opt_row, from_=1, to=50,
                        textvariable=self.var_minw, width=4,
                        bg=C["entry_bg"], fg=C["fg"],
                        buttonbackground=C["border"],
                        relief="flat", font=FONT_UI)
        sp.pack(side="left", padx=(6, 20))
        self._spin = sp

        cb = tk.Checkbutton(opt_row, text="Dry run  (preview only, no files written)",
                            variable=self.var_dryrun,
                            bg=C["panel"], fg=C["fg"], selectcolor=C["accent"],
                            activebackground=C["panel"],
                            font=FONT_UI)
        cb.pack(side="left")
        self._cb_dry = cb

    _settings_entries: list[tk.Entry] = []
    def _settings_entries_append(self, e): self._settings_entries.append(e)

    def _build_actions(self):
        bar = tk.Frame(self.root, bg=C["bg"], padx=14, pady=4)
        bar.pack(fill="x")
        self.btn_start = _btn(bar, "▶   START PROCESSING",
                              self._start, bg=C["btn_start"], width=22)
        self.btn_start.pack(side="left", padx=(0, 10))

        self.btn_stop = _btn(bar, "■   STOP",
                             self._stop, bg=C["btn_stop"], width=10)
        self.btn_stop.pack(side="left", padx=(0, 10))
        self.btn_stop.config(state="disabled")

        self.btn_open_out = _btn(bar, "📂  Open Output Folder",
                                 self._open_output, bg=C["btn_open"])
        self.btn_open_out.pack(side="left", padx=(0, 10))

        self.btn_clear = _btn(bar, "Clear Log",
                              self._clear_log, bg=C["border"])
        self.btn_clear.pack(side="right")

    def _build_stats_bar(self):
        bar = tk.Frame(self.root, bg=C["panel"], padx=14, pady=6,
                       highlightthickness=1, highlightbackground=C["border"])
        bar.pack(fill="x", padx=14, pady=(0, 4))

        def stat_block(parent, label, var, colour):
            f = tk.Frame(parent, bg=C["panel"])
            f.pack(side="left", padx=18)
            tk.Label(f, textvariable=var, font=("Segoe UI", 18, "bold"),
                     fg=colour, bg=C["panel"]).pack()
            tk.Label(f, text=label, font=("Segoe UI", 8),
                     fg=C["fg_dim"], bg=C["panel"]).pack()

        stat_block(bar, "flagged",   self.var_flagged,   C["log_flag"])
        stat_block(bar, "rewritten", self.var_rewritten, C["log_rewrite"])
        stat_block(bar, "clean",     self.var_clean,     C["log_clean"])
        stat_block(bar, "scanned",   self.var_total,     C["fg_dim"])

    def _build_progress(self):
        pf = tk.Frame(self.root, bg=C["bg"], padx=14, pady=2)
        pf.pack(fill="x")

        self.progress_label = tk.Label(pf, text="Ready",
                                       font=FONT_UI, fg=C["fg_dim"], bg=C["bg"],
                                       anchor="w")
        self.progress_label.pack(fill="x", pady=(0, 4))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Custom.Horizontal.TProgressbar",
                        troughcolor=C["panel"],
                        background=C["accent"],
                        bordercolor=C["border"],
                        lightcolor=C["accent"],
                        darkcolor=C["accent"],
                        thickness=12)
        self.progress = ttk.Progressbar(pf, style="Custom.Horizontal.TProgressbar",
                                        orient="horizontal", length=100,
                                        mode="determinate", maximum=100)
        self.progress.pack(fill="x")

    def _build_log(self):
        lf = tk.Frame(self.root, bg=C["bg"], padx=14, pady=6)
        lf.pack(fill="both", expand=True)

        header = tk.Frame(lf, bg=C["bg"])
        header.pack(fill="x", pady=(0, 4))
        _label(header, "📋  Log", font=FONT_HEAD,
               fg=C["accent"], bg=C["bg"]).pack(side="left")
        self.autoscroll_var = tk.BooleanVar(value=True)
        tk.Checkbutton(header, text="Auto-scroll",
                       variable=self.autoscroll_var,
                       bg=C["bg"], fg=C["fg_dim"],
                       selectcolor=C["accent"],
                       activebackground=C["bg"],
                       font=("Segoe UI", 9)).pack(side="right")

        frame = tk.Frame(lf, bg=C["border"], padx=1, pady=1)
        frame.pack(fill="both", expand=True)

        self.log = tk.Text(frame, bg=C["panel"], fg=C["fg"],
                           font=FONT_MONO, wrap="word",
                           relief="flat", bd=0,
                           insertbackground=C["fg"],
                           selectbackground=C["accent"],
                           state="disabled", padx=10, pady=8)
        scroll = tk.Scrollbar(frame, command=self.log.yview,
                              bg=C["border"], troughcolor=C["panel"],
                              activebackground=C["accent"])
        self.log.config(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        tag_map = {
            "header":  {"foreground": C["log_header"], "font": ("Consolas", 10, "bold")},
            "info":    {"foreground": C["log_info"]},
            "clean":   {"foreground": C["log_clean"]},
            "flag":    {"foreground": C["log_flag"]},
            "rewrite": {"foreground": C["log_rewrite"]},
            "error":   {"foreground": C["log_error"]},
            "dim":     {"foreground": C["log_dim"]},
            "label":   {"foreground": C["log_label"], "font": ("Consolas", 9, "bold")},
        }
        for tag, cfg in tag_map.items():
            self.log.tag_configure(tag, **cfg)

    def _build_statusbar(self):
        sb = tk.Frame(self.root, bg=C["border"], height=1)
        sb.pack(fill="x")
        self.status_bar = tk.Label(self.root, text="  Ready",
                                   font=("Segoe UI", 9), fg=C["fg_dim"],
                                   bg=C["panel"], anchor="w", pady=5)
        self.status_bar.pack(fill="x")

    def _browse(self, var: tk.StringVar):
        path = filedialog.askdirectory(initialdir=var.get() or ".")
        if path:
            var.set(path)

    def _open_output(self):
        import subprocess, os
        out = self.var_output.get()
        if Path(out).exists():
            subprocess.Popen(f'explorer "{out}"')
        else:
            self._set_status("Output folder does not exist yet.")

    def _clear_log(self):
        self.log.config(state="normal")
        self.log.delete("1.0", "end")
        self.log.config(state="disabled")
        for v in (self.var_flagged, self.var_rewritten, self.var_clean, self.var_total):
            v.set("0")

    def _set_status(self, msg: str):
        self.status_bar.config(text=f"  {msg}")

    def _set_running(self, running: bool):
        self._running = running
        state_norm = "normal" if not running else "disabled"
        self.btn_start.config(state=state_norm)
        self.btn_stop.config(state="normal" if running else "disabled")
        for e in self._settings_entries:
            e.config(state=state_norm)
        self._spin.config(state=state_norm)
        self._cb_dry.config(state=state_norm)

    def _start(self):
        input_dir  = Path(self.var_input.get())
        output_dir = Path(self.var_output.get())
        model      = self.var_model.get().strip()
        base_url   = self.var_url.get().strip()
        min_words  = self.var_minw.get()
        dry_run    = self.var_dryrun.get()

        if not model:
            self._set_status("Model ID is required.")
            return
        if not input_dir.exists():
            input_dir.mkdir(parents=True)
            self._append_log(f"Created input folder: {input_dir}\n"
                             "Place your .docx files there, then click START again.\n",
                             "flag")
            self._set_status("Input folder created — add .docx files and try again.")
            return

        self._clear_log()
        self._stop_event.clear()
        self._set_running(True)
        self._set_status("Processing…")
        self.progress["value"] = 0

        self._append_log(
            f"  Model     : {model}\n"
            f"  Server    : {base_url}\n"
            f"  Input     : {input_dir.resolve()}\n"
            f"  Output    : {output_dir.resolve()}\n"
            f"  Min words : {min_words}\n"
            f"  Dry run   : {'YES' if dry_run else 'no'}\n",
            "info"
        )

        threading.Thread(
            target=worker,
            args=(input_dir, output_dir, model, base_url,
                  min_words, dry_run, self._q, self._stop_event),
            daemon=True,
        ).start()
        self._poll()

    def _stop(self):
        self._stop_event.set()
        self.btn_stop.config(state="disabled",
                             text="■   Stopping…")
        self._set_status("Stop requested…")

    def _poll(self):
        try:
            while True:
                msg = self._q.get_nowait()
                mtype = msg.get("type")

                if mtype == "log":
                    self._append_log(msg["text"] + "\n", msg.get("tag", "info"))

                elif mtype == "progress":
                    pct = max(0.0, min(1.0, msg["value"])) * 100
                    self.progress["value"] = pct
                    lbl = msg.get("label", "")
                    self.progress_label.config(text=lbl)
                    if lbl:
                        self._set_status(lbl)

                elif mtype == "stats":
                    self.var_flagged.set(str(msg["flagged"]))
                    self.var_rewritten.set(str(msg["rewritten"]))
                    self.var_clean.set(str(msg["clean"]))
                    self.var_total.set(str(msg["total"]))

                elif mtype == "file_done":
                    self._last_docx = msg.get("docx", "")

                elif mtype == "done":
                    self._finish(msg.get("success", False),
                                 msg.get("output_dir", ""))
                    return

        except queue.Empty:
            pass

        if self._running:
            self.root.after(60, self._poll)

    def _finish(self, success: bool, output_dir: str):
        self.progress["value"] = 100
        self._set_running(False)
        self.btn_stop.config(text="■   STOP")
        if success:
            self._set_status(f"Done!  Results saved to: {output_dir}")
            self.progress_label.config(text="Completed ✔")
        else:
            self._set_status("Finished with errors — check the log above.")
            self.progress_label.config(text="Finished with errors")

    def _append_log(self, text: str, tag: str = "info"):
        self.log.config(state="normal")
        self.log.insert("end", text, tag)
        self.log.config(state="disabled")
        if self.autoscroll_var.get():
            self.log.see("end")


# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────

def main():
    root = tk.Tk()
    root.configure(bg=C["bg"])

    # DPI awareness on Windows
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App(root)

    # Centre window on screen
    root.update_idletasks()
    w, h = root.winfo_width(), root.winfo_height()
    sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

    root.mainloop()


if __name__ == "__main__":
    main()

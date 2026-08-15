#!/usr/bin/env python3
"""Export cloud-agent transcript to a PDF with formatted Markdown tables."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from xml.sax.saxutils import escape

from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    KeepTogether,
    ListFlowable,
    ListItem,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

TIMESTAMP_RE = re.compile(r"<timestamp>.*?</timestamp>\s*", re.S)
TAG_RE = re.compile(r"<[^>]+>")
MD_TABLE_LINE_RE = re.compile(r"^\|.+\|\s*$")
MD_SEP_RE = re.compile(r"^\|?[\s:-]+\|[\s|:-]+\|?\s*$")
BOLD_RE = re.compile(r"\*\*(.+?)\*\*")
ITALIC_RE = re.compile(r"(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)")
CODE_RE = re.compile(r"`([^`]+)`")


def clean_text(t: str) -> str:
    if not t:
        return ""
    if "<system_notification>" in t:
        return ""
    t = TIMESTAMP_RE.sub("", t)
    t = TAG_RE.sub("", t)
    t = t.replace("\r\n", "\n").replace("\r", "\n")
    t = re.sub(r"\n{3,}", "\n\n", t).strip()
    return t


def inline_md_to_rl(text: str) -> str:
    """Convert a subset of Markdown inline markup to ReportLab rich text."""
    # Protect code spans first so wildcards like *_* in filenames are not touched.
    placeholders: list[str] = []

    def _code_sub(m: re.Match) -> str:
        placeholders.append(m.group(1))
        return f"\x00CODE{len(placeholders) - 1}\x00"

    text = CODE_RE.sub(_code_sub, text)
    text = escape(text)
    text = BOLD_RE.sub(r"<b>\1</b>", text)
    # Only treat *italic* when both sides are word-like (avoid *_filename* wildcards).
    text = re.sub(
        r"(?<![\w*])\*(?!\*)([^*\n]+?)\*(?!\*)(?![\w*])",
        r"<i>\1</i>",
        text,
    )
    for i, code in enumerate(placeholders):
        text = text.replace(
            f"\x00CODE{i}\x00",
            f'<font face="Courier" size="8">{escape(code)}</font>',
        )
    return text


def parse_md_table(lines: list[str]) -> list[list[str]] | None:
    if len(lines) < 2:
        return None
    rows = []
    for i, line in enumerate(lines):
        if i == 1 and MD_SEP_RE.match(line.strip()):
            continue
        cells = [c.strip() for c in line.strip().strip("|").split("|")]
        rows.append(cells)
    if not rows:
        return None
    width = max(len(r) for r in rows)
    return [r + [""] * (width - len(r)) for r in rows]


def split_blocks(text: str) -> list[tuple[str, object]]:
    """Split message into ('text', str) and ('table', rows) blocks."""
    lines = text.split("\n")
    blocks: list[tuple[str, object]] = []
    buf: list[str] = []
    i = 0
    while i < len(lines):
        if MD_TABLE_LINE_RE.match(lines[i]):
            if buf:
                blocks.append(("text", "\n".join(buf).strip()))
                buf = []
            table_lines = []
            while i < len(lines) and MD_TABLE_LINE_RE.match(lines[i]):
                table_lines.append(lines[i])
                i += 1
            rows = parse_md_table(table_lines)
            if rows:
                blocks.append(("table", rows))
            else:
                buf.extend(table_lines)
            continue
        buf.append(lines[i])
        i += 1
    if buf:
        blocks.append(("text", "\n".join(buf).strip()))
    return [(k, v) for k, v in blocks if v]


def make_styles():
    styles = getSampleStyleSheet()
    return {
        "title": ParagraphStyle(
            "TitlePT",
            parent=styles["Title"],
            fontSize=15,
            leading=19,
            spaceAfter=10,
        ),
        "meta": ParagraphStyle(
            "Meta",
            parent=styles["Normal"],
            fontSize=8.5,
            leading=11,
            textColor=colors.HexColor("#444444"),
            spaceAfter=14,
        ),
        "role_user": ParagraphStyle(
            "RoleUser",
            parent=styles["Heading2"],
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#0B3D91"),
            spaceBefore=12,
            spaceAfter=4,
        ),
        "role_asst": ParagraphStyle(
            "RoleAsst",
            parent=styles["Heading2"],
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#0B6E4F"),
            spaceBefore=12,
            spaceAfter=4,
        ),
        "body": ParagraphStyle(
            "BodyPT",
            parent=styles["Normal"],
            fontSize=9,
            leading=12,
            alignment=TA_JUSTIFY,
            spaceAfter=6,
        ),
        "cell": ParagraphStyle(
            "Cell",
            parent=styles["Normal"],
            fontSize=7.5,
            leading=9.5,
            alignment=TA_LEFT,
        ),
        "cell_header": ParagraphStyle(
            "CellHeader",
            parent=styles["Normal"],
            fontSize=7.5,
            leading=9.5,
            alignment=TA_LEFT,
            textColor=colors.white,
            fontName="Helvetica-Bold",
        ),
    }


def build_table(rows: list[list[str]], styles, page_width: float) -> Table:
    n_cols = max(len(r) for r in rows)
    usable = page_width - 3.2 * cm
    col_w = usable / n_cols

    data = []
    for r_i, row in enumerate(rows):
        style = styles["cell_header"] if r_i == 0 else styles["cell"]
        data.append([Paragraph(inline_md_to_rl(c), style) for c in row])

    tbl = Table(data, colWidths=[col_w] * n_cols, repeatRows=1, hAlign="LEFT")
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F4E79")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F7F9FC")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.HexColor("#F7F9FC"), colors.white]),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#9AA7B5")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 4),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    return tbl


def text_to_flowables(text: str, styles) -> list:
    out = []
    # preserve paragraphs; handle bullet-ish lines simply
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        # keep single newlines as <br/>
        out.append(Paragraph(inline_md_to_rl(para).replace("\n", "<br/>"), styles["body"]))
    return out


def export_pdf(
    transcript_path: Path,
    out_path: Path,
    *,
    title: str,
    landscape_mode: bool = False,
    stop_before_user_contains: str | None = None,
) -> dict:
    data = json.loads(transcript_path.read_text(encoding="utf-8"))
    styles = make_styles()
    pagesize = landscape(A4) if landscape_mode else A4
    page_width = pagesize[0]

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=pagesize,
        leftMargin=1.5 * cm,
        rightMargin=1.5 * cm,
        topMargin=1.3 * cm,
        bottomMargin=1.3 * cm,
        title=title,
        author="Cursor Cloud Agent",
    )

    story = [
        Paragraph(escape(title), styles["title"]),
        Paragraph(
            "Exportação com tabelas Markdown convertidas para tabelas PDF.<br/>"
            "Mensagens de ferramenta omitidas; markup <b>negrito</b>/<i>itálico</i> preservado.",
            styles["meta"],
        ),
    ]

    n_user = n_asst = n_tables = 0
    for m in data["messages"]:
        role = m.get("role")
        if role not in ("user", "assistant"):
            continue
        text = clean_text(m.get("text") or "")
        if not text:
            continue
        if (
            stop_before_user_contains
            and role == "user"
            and stop_before_user_contains.lower() in text.lower()
        ):
            break

        if role == "user":
            n_user += 1
            story.append(Paragraph(f"Usuário #{n_user}", styles["role_user"]))
        else:
            n_asst += 1
            story.append(Paragraph(f"Assistente #{n_asst}", styles["role_asst"]))

        for kind, payload in split_blocks(text):
            if kind == "text":
                story.extend(text_to_flowables(str(payload), styles))
            else:
                n_tables += 1
                story.append(Spacer(1, 4))
                story.append(KeepTogether([build_table(payload, styles, page_width)]))  # type: ignore[arg-type]
                story.append(Spacer(1, 8))

    story.append(Spacer(1, 16))
    story.append(
        Paragraph(
            f"Fim — {n_user} mensagens de usuário, {n_asst} de assistente, {n_tables} tabelas formatadas.",
            styles["meta"],
        )
    )
    doc.build(story)
    return {
        "path": str(out_path),
        "size_kb": round(out_path.stat().st_size / 1024, 1),
        "n_user": n_user,
        "n_asst": n_asst,
        "n_tables": n_tables,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--transcript", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--title", required=True)
    ap.add_argument("--landscape", action="store_true")
    ap.add_argument("--stop-before-user-contains")
    args = ap.parse_args()
    info = export_pdf(
        Path(args.transcript),
        Path(args.out),
        title=args.title,
        landscape_mode=args.landscape,
        stop_before_user_contains=args.stop_before_user_contains,
    )
    print(json.dumps(info, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

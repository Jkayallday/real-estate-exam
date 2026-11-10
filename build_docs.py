import html
import re
import shutil
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
MD_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.md"
DOCX_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.docx"
HTML_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.html"
INDEX_PATH = BASE_DIR / "index.html"

ARTIFACT_PATHS = [
    Path(r"C:\Users\16162\.gemini\antigravity\brain\70c815b2-b8c1-4557-9bbf-0f881349c69b\Holloways_Real_Estate_Audio_Transcripts.md"),
    Path(r"C:\Users\16162\.gemini\antigravity\brain\91b0d95b-bf8a-4c68-94cd-c6f26bb8a447\Holloways_Real_Estate_Audio_Transcripts.md"),
]

SLIDE_FILES = {
    1: "slides/Lesson-1-Student-v2.ppt",
    2: "slides/Lesson-2-Student-v2.ppt",
    3: "slides/Lesson-3-Student-v2.ppt",
    4: "slides/Lesson-4-Student-v2.ppt",
    5: "slides/Lesson-5-Student-v2.ppt",
    6: "slides/Lesson-6-Student-v2.ppt",
    7: "slides/Lesson-7-Student-v2.ppt",
    8: "slides/Lesson-8-Student-v2.ppt",
    9: "slides/lesson-9-Student-v2.pptx",
}


def parse_markdown(md_text: str):
    lessons = []
    toc_rows = []
    current_lesson = None
    current_track = None

    lesson_re = re.compile(r"^##\s+(Lesson\s+(\d+):\s+.+)$")
    track_re = re.compile(r"^###\s+Track\s+(\d+)\s+—\s+(.+?)\s+\((\d+:\d+)\)$")
    toc_row_re = re.compile(
        r"^\|\s*\*\*Track\s+(\d+)\*\*\s*\|\s*(.+?)\s*\|\s*(.+?)\s*\|\s*(\d+:\d+)\s*\|\s*`([^`]+)`\s*\|$"
    )

    lines = md_text.splitlines()
    for line in lines:
        m_toc = toc_row_re.match(line.strip())
        if m_toc and current_lesson is None:
            toc_rows.append(
                {
                    "number": int(m_toc.group(1)),
                    "website_lesson": m_toc.group(2).strip(),
                    "content_summary": m_toc.group(3).strip(),
                    "duration": m_toc.group(4).strip(),
                    "audio_file": m_toc.group(5).strip(),
                }
            )
            continue

        m_lesson = lesson_re.match(line)
        if m_lesson:
            if current_track and current_lesson:
                current_lesson["tracks"].append(current_track)
                current_track = None
            if current_lesson:
                lessons.append(current_lesson)
            lesson_num = int(m_lesson.group(2))
            current_lesson = {
                "number": lesson_num,
                "heading": m_lesson.group(1),
                "slide": SLIDE_FILES.get(lesson_num, ""),
                "tracks": [],
            }
            continue

        m_track = track_re.match(line)
        if m_track:
            if current_track and current_lesson:
                current_lesson["tracks"].append(current_track)
            track_num = int(m_track.group(1))
            current_track = {
                "number": track_num,
                "id": f"Track{track_num:02d}",
                "title": m_track.group(2),
                "duration": m_track.group(3),
                "local_audio": f"audio/Track{track_num:02d}.mp3",
                "remote_audio": f"https://hollowaysinstitute.biz/wp-content/uploads/Track{track_num:02d}.mp3",
                "notes": [],
                "paragraphs": [],
            }
            continue

        if current_track is not None:
            if line.startswith("- *[Note:"):
                note_text = re.sub(r"^-\s*\*", "", line).rstrip("*").strip()
                current_track["notes"].append(note_text)
            elif line.startswith(">"):
                content = line[1:].strip()
                if content:
                    current_track["paragraphs"].append(content)

    if current_track and current_lesson:
        current_lesson["tracks"].append(current_track)
    if current_lesson:
        lessons.append(current_lesson)

    return lessons, toc_rows


def format_inline_md(text: str) -> str:
    escaped = html.escape(text)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"\*(.+?)\*", r"<em>\1</em>", escaped)
    return escaped


def render_blocks_html(paragraphs: list[str]) -> str:
    html_blocks = []
    list_type = None
    list_items = []

    def flush_list():
        nonlocal list_type, list_items
        if list_type and list_items:
            items_str = "\n".join(f"<li>{item}</li>" for item in list_items)
            html_blocks.append(f"<{list_type}>\n{items_str}\n</{list_type}>")
        list_type = None
        list_items = []

    ol_re = re.compile(r"^(\d+)\.\s+(.+)$")
    for p in paragraphs:
        if p.startswith("- "):
            if list_type != "ul":
                flush_list()
                list_type = "ul"
            list_items.append(format_inline_md(p[2:].strip()))
        else:
            m_ol = ol_re.match(p)
            if m_ol:
                if list_type != "ol":
                    flush_list()
                    list_type = "ol"
                list_items.append(format_inline_md(m_ol.group(2).strip()))
            else:
                flush_list()
                html_blocks.append(f"<p>{format_inline_md(p)}</p>")

    flush_list()
    return "\n".join(html_blocks)


def md_to_docx_runs(text: str, base_bold: bool = False, base_italic: bool = False) -> str:
    token_re = re.compile(r"(\*\*.+?\*\*|\*.+?\*)")
    tokens = token_re.split(text)
    runs = []
    for tok in tokens:
        if not tok:
            continue
        b = base_bold
        i = base_italic
        content = tok
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            b = True
            content = tok[2:-2]
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            i = True
            content = tok[1:-1]
        rPr_parts = []
        if b:
            rPr_parts.append("<w:b/>")
        if i:
            rPr_parts.append("<w:i/>")
        rPr = f"<w:rPr>{''.join(rPr_parts)}</w:rPr>" if rPr_parts else ""
        runs.append(
            f"<w:r>{rPr}<w:t xml:space=\"preserve\">{html.escape(content)}</w:t></w:r>"
        )
    return "".join(runs)


def build_docx(lessons, toc_rows, output_path: Path):
    def w_p(
        text: str,
        style: str = "",
        bold: bool = False,
        italic: bool = False,
        indent_left: int = 0,
        hanging: int = 0,
    ):
        pPr_parts = []
        if style:
            pPr_parts.append(f'<w:pStyle w:val="{style}"/>')
        if indent_left or hanging:
            hang_attr = f' w:hanging="{hanging}"' if hanging else ""
            pPr_parts.append(f'<w:ind w:left="{indent_left}"{hang_attr}/>')
        pPr = f"<w:pPr>{''.join(pPr_parts)}</w:pPr>" if pPr_parts else ""
        runs_xml = md_to_docx_runs(text, base_bold=bold, base_italic=italic)
        return f"<w:p>{pPr}{runs_xml}</w:p>"

    body_parts = [
        w_p(
            "Holloway's Real Estate Institute — Complete Audio Program Transcripts (Lessons 1–9, Tracks 01–57)",
            style="Title",
            bold=True,
        ),
        w_p(
            "Source: https://hollowaysinstitute.biz/audio/ | Speaker: Mike Holloway (Copyright 2019) | Total Tracks: 57 | Total Duration: ~1 hr 38 min (98m 12s)",
            style="Subtitle",
            italic=True,
        ),
        w_p("Table of Contents & Master Track Index", style="Heading1", bold=True),
    ]

    for row in toc_rows:
        body_parts.append(
            w_p(
                f"• **Track {row['number']:02d}** ({row['duration']}) — *{row['website_lesson']}*: {row['content_summary']} [{row['audio_file']}]",
                style="ListParagraph",
                indent_left=360,
                hanging=180,
            )
        )

    ol_re = re.compile(r"^(\d+)\.\s+(.+)$")

    for lesson in lessons:
        body_parts.append(w_p(lesson["heading"], style="Heading1", bold=True))
        if lesson["slide"]:
            body_parts.append(
                w_p(f"Companion Slide Deck: {lesson['slide']}", italic=True)
            )
        for tr in lesson["tracks"]:
            body_parts.append(
                w_p(
                    f"Track {tr['number']:02d} — {tr['title']} ({tr['duration']})",
                    style="Heading2",
                    bold=True,
                )
            )
            body_parts.append(
                w_p(
                    f"Audio File: {tr['local_audio']} | Source: {tr['remote_audio']}",
                    italic=True,
                )
            )
            for note in tr["notes"]:
                body_parts.append(w_p(note, style="IntenseQuote", italic=True))
            for para in tr["paragraphs"]:
                if para.startswith("- "):
                    body_parts.append(
                        w_p(
                            f"• {para[2:].strip()}",
                            style="ListParagraph",
                            indent_left=720,
                            hanging=360,
                        )
                    )
                else:
                    m_ol = ol_re.match(para)
                    if m_ol:
                        body_parts.append(
                            w_p(
                                f"{m_ol.group(1)}. {m_ol.group(2).strip()}",
                                style="ListParagraph",
                                indent_left=720,
                                hanging=360,
                            )
                        )
                    else:
                        body_parts.append(w_p(para))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body_parts)}<w:sectPr/></w:body></w:document>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="140" w:line="276" w:lineRule="auto"/></w:pPr>'
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Title">'
        '<w:name w:val="Title"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:before="120" w:after="80"/></w:pPr>'
        '<w:rPr><w:b/><w:color w:val="0F172A"/><w:sz w:val="36"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Subtitle">'
        '<w:name w:val="Subtitle"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="240"/></w:pPr>'
        '<w:rPr><w:i/><w:color w:val="475569"/><w:sz w:val="20"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading1">'
        '<w:name w:val="heading 1"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="320" w:after="120"/><w:outlineLvl w:val="0"/></w:pPr>'
        '<w:rPr><w:b/><w:color w:val="1E3A8A"/><w:sz w:val="28"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="Heading2">'
        '<w:name w:val="heading 2"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:keepNext/><w:spacing w:before="240" w:after="80"/><w:outlineLvl w:val="1"/></w:pPr>'
        '<w:rPr><w:b/><w:color w:val="1D4ED8"/><w:sz w:val="24"/></w:rPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="ListParagraph">'
        '<w:name w:val="List Paragraph"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:spacing w:after="80"/></w:pPr>'
        '</w:style>'
        '<w:style w:type="paragraph" w:styleId="IntenseQuote">'
        '<w:name w:val="Intense Quote"/>'
        '<w:basedOn w:val="Normal"/>'
        '<w:pPr><w:ind w:left="360"/><w:spacing w:after="120"/></w:pPr>'
        '<w:rPr><w:i/><w:color w:val="92400E"/></w:rPr>'
        '</w:style>'
        '</w:styles>'
    )

    doc_rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
        '</Relationships>'
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '</Types>'
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        '</Relationships>'
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)
        zf.writestr("word/styles.xml", styles_xml)
        zf.writestr("word/_rels/document.xml.rels", doc_rels_xml)


def build_html(lessons, toc_rows, output_path: Path):
    total_tracks = sum(len(l["tracks"]) for l in lessons)
    lesson_buttons = [
        '<button class="filter-btn active" data-lesson="all">All 9 Lessons (57 Tracks)</button>'
    ]
    for l in lessons:
        lesson_buttons.append(
            f'<button class="filter-btn" data-lesson="{l["number"]}">'
            f'Lesson {l["number"]} ({len(l["tracks"])})</button>'
        )

    toc_rows_html = []
    for row in toc_rows:
        tnum = f"{row['number']:02d}"
        toc_rows_html.append(
            f"""
            <tr>
              <td><a href="#track-{tnum}" class="toc-track-link"><strong>Track {tnum}</strong></a></td>
              <td>{html.escape(row['website_lesson'])}</td>
              <td>{format_inline_md(row['content_summary'])}</td>
              <td class="toc-duration">{html.escape(row['duration'])}</td>
              <td><a href="{html.escape(row['audio_file'])}" download><code>{html.escape(row['audio_file'])}</code></a></td>
            </tr>
            """
        )

    sections_html = []
    for l in lessons:
        tracks_cards = []
        for tr in l["tracks"]:
            paras_html = render_blocks_html(tr["paragraphs"])
            notes_html = "\n".join(
                f'<div class="track-note">{html.escape(n)}</div>' for n in tr["notes"]
            )
            tracks_cards.append(
                f"""
                <article class="track-card" id="track-{tr['number']:02d}" data-lesson="{l['number']}" data-track="{tr['number']:02d}">
                  <div class="track-header">
                    <div>
                      <span class="track-badge">Track {tr['number']:02d}</span>
                      <h3 class="track-title">{html.escape(tr['title'])}</h3>
                    </div>
                    <span class="track-duration">{html.escape(tr['duration'])}</span>
                  </div>
                  <div class="audio-bar">
                    <audio controls preload="none">
                      <source src="{html.escape(tr['local_audio'])}" type="audio/mpeg">
                      <source src="{html.escape(tr['remote_audio'])}" type="audio/mpeg">
                    </audio>
                    <div class="audio-links">
                      <a href="{html.escape(tr['local_audio'])}" download>Local MP3</a>
                      <span>·</span>
                      <a href="{html.escape(tr['remote_audio'])}" target="_blank" rel="noopener">Remote MP3</a>
                    </div>
                  </div>
                  {notes_html}
                  <div class="transcript-body">
                    {paras_html}
                  </div>
                </article>
                """
            )

        slide_link = (
            f'<a class="slide-btn" href="{html.escape(l["slide"])}" download>Download Lesson {l["number"]} Slides</a>'
            if l["slide"]
            else ""
        )
        sections_html.append(
            f"""
            <section class="lesson-section" id="lesson-{l['number']}" data-lesson="{l['number']}">
              <div class="lesson-header">
                <h2>{html.escape(l['heading'])}</h2>
                {slide_link}
              </div>
              {''.join(tracks_cards)}
            </section>
            """
        )

    html_doc = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Holloway's Real Estate Institute — Complete Audio Program Transcripts (Lessons 1–9, Tracks 01–57)</title>
  <style>
    :root {{
      --bg: #f8fafc;
      --card: #ffffff;
      --text: #0f172a;
      --muted: #475569;
      --accent: #1d4ed8;
      --accent-soft: #dbeafe;
      --border: #e2e8f0;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Inter, sans-serif;
      background: var(--bg);
      color: var(--text);
      line-height: 1.65;
    }}
    header {{
      background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
      color: #fff;
      padding: 2.5rem 1.5rem;
    }}
    .container {{
      max-width: 1080px;
      margin: 0 auto;
    }}
    header h1 {{
      margin: 0 0 0.5rem 0;
      font-size: 1.85rem;
    }}
    header p {{
      margin: 0.25rem 0;
      color: #cbd5e1;
    }}
    .top-actions {{
      margin-top: 1.25rem;
      display: flex;
      flex-wrap: wrap;
      gap: 0.75rem;
    }}
    .action-btn {{
      display: inline-block;
      background: #fff;
      color: #0f172a;
      padding: 0.5rem 1rem;
      border-radius: 6px;
      font-weight: 600;
      font-size: 0.9rem;
      text-decoration: none;
      border: none;
      cursor: pointer;
    }}
    .action-btn.secondary {{
      background: rgba(255,255,255,0.15);
      color: #fff;
      border: 1px solid rgba(255,255,255,0.3);
    }}
    .controls {{
      position: sticky;
      top: 0;
      z-index: 20;
      background: rgba(248, 250, 252, 0.95);
      backdrop-filter: blur(8px);
      border-bottom: 1px solid var(--border);
      padding: 1rem 1.5rem;
    }}
    .search-row {{
      display: flex;
      gap: 0.75rem;
      align-items: center;
      margin-bottom: 0.75rem;
    }}
    #search-input {{
      flex: 1;
      padding: 0.65rem 1rem;
      font-size: 1rem;
      border: 1px solid #cbd5e1;
      border-radius: 8px;
    }}
    #match-count {{
      font-size: 0.9rem;
      color: var(--muted);
      font-weight: 600;
      white-space: nowrap;
    }}
    .filter-bar {{
      display: flex;
      flex-wrap: wrap;
      gap: 0.4rem;
    }}
    .filter-btn {{
      background: #fff;
      border: 1px solid var(--border);
      color: var(--text);
      padding: 0.35rem 0.75rem;
      border-radius: 999px;
      font-size: 0.82rem;
      cursor: pointer;
      font-weight: 500;
    }}
    .filter-btn.active {{
      background: var(--accent);
      color: #fff;
      border-color: var(--accent);
    }}
    main {{
      padding: 1.5rem;
    }}
    .toc-panel {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 1rem 1.25rem;
      margin-bottom: 2rem;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }}
    .toc-panel summary {{
      cursor: pointer;
      font-weight: 700;
      color: #1e3a8a;
      font-size: 1.05rem;
    }}
    .toc-table-wrap {{
      overflow-x: auto;
      margin-top: 1rem;
    }}
    .toc-table {{
      width: 100%;
      border-collapse: collapse;
      font-size: 0.88rem;
    }}
    .toc-table th, .toc-table td {{
      text-align: left;
      padding: 0.5rem 0.65rem;
      border-bottom: 1px solid var(--border);
      vertical-align: top;
    }}
    .toc-table th {{
      background: #f1f5f9;
      font-weight: 700;
      color: #0f172a;
      white-space: nowrap;
    }}
    .toc-track-link {{
      color: var(--accent);
      text-decoration: none;
      white-space: nowrap;
    }}
    .toc-track-link:hover {{
      text-decoration: underline;
    }}
    .toc-duration {{
      font-family: monospace;
      white-space: nowrap;
    }}
    .lesson-section {{
      margin-bottom: 2.5rem;
    }}
    .lesson-header {{
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 0.75rem;
      border-bottom: 2px solid var(--accent);
      padding-bottom: 0.5rem;
      margin-bottom: 1.25rem;
    }}
    .lesson-header h2 {{
      margin: 0;
      font-size: 1.4rem;
      color: #1e3a8a;
    }}
    .slide-btn {{
      background: var(--accent-soft);
      color: var(--accent);
      text-decoration: none;
      padding: 0.4rem 0.85rem;
      border-radius: 6px;
      font-size: 0.85rem;
      font-weight: 600;
    }}
    .track-card {{
      background: var(--card);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 1.25rem 1.5rem;
      margin-bottom: 1rem;
      box-shadow: 0 1px 2px rgba(0,0,0,0.03);
    }}
    .track-header {{
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 1rem;
      margin-bottom: 0.75rem;
    }}
    .track-badge {{
      display: inline-block;
      background: var(--accent);
      color: #fff;
      font-weight: 700;
      font-size: 0.78rem;
      padding: 0.15rem 0.55rem;
      border-radius: 4px;
      margin-right: 0.5rem;
      text-transform: uppercase;
    }}
    .track-title {{
      display: inline;
      margin: 0;
      font-size: 1.1rem;
    }}
    .track-duration {{
      font-family: monospace;
      font-size: 0.9rem;
      color: var(--muted);
      background: #f1f5f9;
      padding: 0.15rem 0.5rem;
      border-radius: 4px;
    }}
    .audio-bar {{
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 0.75rem;
      background: #f8fafc;
      padding: 0.5rem 0.75rem;
      border-radius: 8px;
      margin-bottom: 0.85rem;
    }}
    audio {{
      height: 36px;
      flex: 1;
      min-width: 240px;
    }}
    .audio-links {{
      font-size: 0.82rem;
      color: var(--muted);
    }}
    .audio-links a {{
      color: var(--accent);
      text-decoration: none;
    }}
    .track-note {{
      font-size: 0.88rem;
      color: #92400e;
      background: #fef3c7;
      border-left: 3px solid #f59e0b;
      padding: 0.4rem 0.75rem;
      margin-bottom: 0.75rem;
      border-radius: 4px;
    }}
    .transcript-body p {{
      margin: 0.6rem 0;
    }}
    .transcript-body ul, .transcript-body ol {{
      margin: 0.5rem 0 0.75rem 1.5rem;
      padding: 0;
    }}
    .transcript-body li {{
      margin: 0.25rem 0;
    }}
    @media print {{
      .controls, .audio-bar, .top-actions, .slide-btn {{ display: none !important; }}
      header {{ background: none; color: #000; padding: 0 0 1rem 0; border-bottom: 2px solid #000; }}
      header p {{ color: #333; }}
      .track-card {{ break-inside: avoid; border: 1px solid #ccc; box-shadow: none; }}
    }}
  </style>
</head>
<body>
  <header>
    <div class="container">
      <h1>Holloway's Real Estate Institute — Complete Audio Program Transcripts</h1>
      <p>Verbatim Transcripts of All {total_tracks} Audio Tracks Across Lessons 1–9 (Speaker: Mike Holloway · Total Duration: ~1 hr 38 min)</p>
      <div class="top-actions">
        <a class="action-btn" href="Holloways_Real_Estate_Audio_Transcripts.docx" download>Download Word (.docx)</a>
        <a class="action-btn secondary" href="Holloways_Real_Estate_Audio_Transcripts.md" download>Download Markdown (.md)</a>
        <button class="action-btn secondary" onclick="window.print()">Print / Save as PDF</button>
      </div>
    </div>
  </header>

  <div class="controls">
    <div class="container">
      <div class="search-row">
        <input type="search" id="search-input" placeholder="Search transcripts across all 57 tracks (e.g., LARA, escrow, riparian, TRID, 1866, mills)...">
        <span id="match-count">Showing {total_tracks} of {total_tracks} tracks</span>
      </div>
      <div class="filter-bar">
        {''.join(lesson_buttons)}
      </div>
    </div>
  </div>

  <main class="container">
    <details class="toc-panel">
      <summary>Table of Contents &amp; Master Track Index (All {len(toc_rows)} Audio Tracks Across Lessons 1–9)</summary>
      <div class="toc-table-wrap">
        <table class="toc-table">
          <thead>
            <tr>
              <th>Track</th>
              <th>Website Lesson Listing</th>
              <th>Actual Audio Content</th>
              <th>Duration</th>
              <th>Audio File</th>
            </tr>
          </thead>
          <tbody>
            {''.join(toc_rows_html)}
          </tbody>
        </table>
      </div>
    </details>
    {''.join(sections_html)}
  </main>

  <script>
    const searchInput = document.getElementById('search-input');
    const matchCount = document.getElementById('match-count');
    const filterBtns = document.querySelectorAll('.filter-btn');
    const cards = document.querySelectorAll('.track-card');
    const sections = document.querySelectorAll('.lesson-section');
    let activeLesson = 'all';

    function applyFilters() {{
      const q = searchInput.value.trim().toLowerCase();
      let visibleCount = 0;

      cards.forEach(card => {{
        const lessonMatch = (activeLesson === 'all' || card.dataset.lesson === activeLesson);
        const textMatch = (!q || card.textContent.toLowerCase().includes(q));
        const show = lessonMatch && textMatch;
        card.style.display = show ? '' : 'none';
        if (show) visibleCount++;
      }});

      sections.forEach(sec => {{
        const hasVisible = Array.from(sec.querySelectorAll('.track-card')).some(c => c.style.display !== 'none');
        sec.style.display = hasVisible ? '' : 'none';
      }});

      matchCount.textContent = `Showing ${{visibleCount}} of {total_tracks} tracks`;
    }}

    searchInput.addEventListener('input', applyFilters);
    filterBtns.forEach(btn => {{
      btn.addEventListener('click', () => {{
        filterBtns.forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        activeLesson = btn.dataset.lesson;
        applyFilters();
      }});
    }});
  </script>
</body>
</html>
"""
    output_path.write_text(html_doc, encoding="utf-8")


def verify_outputs():
    # Verify HTML structure
    html_text = HTML_PATH.read_text(encoding="utf-8")
    assert html_text == INDEX_PATH.read_text(encoding="utf-8"), "index.html out of sync with HTML_PATH"
    assert "<p>- " not in html_text, "Found unparsed markdown bullet list in HTML!"
    assert not re.search(r"<p>\d+\.\s+", html_text), "Found unparsed markdown ordered list in HTML!"
    assert "<ul>" in html_text and "<ol>" in html_text, "Expected both <ul> and <ol> lists in HTML!"
    assert 'class="toc-table"' in html_text, "Missing Table of Contents table in HTML!"
    for i in range(1, 58):
        assert f'id="track-{i:02d}"' in html_text, f"Missing track-{i:02d} in HTML"
        assert f'href="#track-{i:02d}"' in html_text, f"Missing TOC link to #track-{i:02d} in HTML"

    # Verify DOCX OpenXML package and styles
    with zipfile.ZipFile(DOCX_PATH, "r") as zf:
        names = set(zf.namelist())
        expected_parts = {
            "[Content_Types].xml",
            "_rels/.rels",
            "word/document.xml",
            "word/styles.xml",
            "word/_rels/document.xml.rels",
        }
        assert expected_parts.issubset(names), f"Missing DOCX parts: {expected_parts - names}"
        for part in expected_parts:
            ET.fromstring(zf.read(part))
        doc_xml = zf.read("word/document.xml").decode("utf-8")
        styles_xml = zf.read("word/styles.xml").decode("utf-8")
        assert '<w:styleId="Heading1"' in styles_xml or 'w:styleId="Heading1"' in styles_xml
        assert 'w:styleId="Heading2"' in styles_xml
        assert "<w:b/>" in doc_xml and "<w:i/>" in doc_xml
        for i in range(1, 58):
            assert f"Track {i:02d}" in doc_xml, f"Missing Track {i:02d} in DOCX"


def main():
    md_text = MD_PATH.read_text(encoding="utf-8")
    lessons, toc_rows = parse_markdown(md_text)
    total_tracks = sum(len(l["tracks"]) for l in lessons)
    print(f"Parsed {len(lessons)} lessons, {total_tracks} tracks, and {len(toc_rows)} TOC rows.")
    assert len(lessons) == 9, f"Expected 9 lessons, got {len(lessons)}"
    assert total_tracks == 57, f"Expected 57 tracks, got {total_tracks}"
    assert len(toc_rows) == 57, f"Expected 57 TOC rows, got {len(toc_rows)}"
    for l in lessons:
        for tr in l["tracks"]:
            assert len(tr["paragraphs"]) > 0, f"Track {tr['number']} has no paragraphs!"

    build_docx(lessons, toc_rows, DOCX_PATH)
    build_html(lessons, toc_rows, HTML_PATH)
    build_html(lessons, toc_rows, INDEX_PATH)

    for artifact_path in ARTIFACT_PATHS:
        if artifact_path.parent.exists():
            shutil.copyfile(MD_PATH, artifact_path)
            print(f"Synced artifact copy: {artifact_path}")

    verify_outputs()
    print(f"Generated & verified {DOCX_PATH.name} ({DOCX_PATH.stat().st_size} bytes)")
    print(f"Generated & verified {HTML_PATH.name} ({HTML_PATH.stat().st_size} bytes)")
    print(f"Generated & verified {INDEX_PATH.name} ({INDEX_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

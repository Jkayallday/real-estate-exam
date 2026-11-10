import html
import json
import os
import re
import zipfile
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
MD_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.md"
DOCX_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.docx"
HTML_PATH = BASE_DIR / "Holloways_Real_Estate_Audio_Transcripts.html"
INDEX_PATH = BASE_DIR / "index.html"

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
    current_lesson = None
    current_track = None

    lesson_re = re.compile(r"^##\s+(Lesson\s+(\d+):\s+.+)$")
    track_re = re.compile(r"^###\s+Track\s+(\d+)\s+—\s+(.+?)\s+\((\d+:\d+)\)$")

    lines = md_text.splitlines()
    for line in lines:
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
                current_track["notes"].append(line.strip("- *"))
            elif line.startswith(">"):
                content = line[1:].strip()
                if content:
                    current_track["paragraphs"].append(content)

    if current_track and current_lesson:
        current_lesson["tracks"].append(current_track)
    if current_lesson:
        lessons.append(current_lesson)

    return lessons


def format_inline_md(text: str) -> str:
    escaped = html.escape(text)
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"\*(.+?)\*", r"<em>\1</em>", escaped)
    return escaped


def strip_md(text: str) -> str:
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    return text


def build_docx(lessons, output_path: Path):
    def w_p(text: str, style: str = "", bold: bool = False, italic: bool = False):
        pPr = f'<w:pPr><w:pStyle w:val="{style}"/></w:pPr>' if style else ""
        rPr_parts = []
        if bold:
            rPr_parts.append("<w:b/>")
        if italic:
            rPr_parts.append("<w:i/>")
        rPr = f"<w:rPr>{''.join(rPr_parts)}</w:rPr>" if rPr_parts else ""
        return (
            f"<w:p>{pPr}<w:r>{rPr}"
            f'<w:t xml:space="preserve">{html.escape(text)}</w:t>'
            f"</w:r></w:p>"
        )

    body_parts = [
        w_p(
            "Holloway's Real Estate Institute — Complete Audio Program Transcripts (Lessons 1–9, Tracks 01–57)",
            style="Title",
            bold=True,
        ),
        w_p(
            "Source: https://hollowaysinstitute.biz/audio/ | Speaker: Mike Holloway | Total Tracks: 57 (Lessons 1–9)",
            italic=True,
        ),
    ]

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
                body_parts.append(w_p(note, italic=True))
            for para in tr["paragraphs"]:
                body_parts.append(w_p(strip_md(para)))

    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{''.join(body_parts)}<w:sectPr/></w:body></w:document>"
    )

    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )

    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)


def build_html(lessons, output_path: Path):
    total_tracks = sum(len(l["tracks"]) for l in lessons)
    lesson_buttons = ['<button class="filter-btn active" data-lesson="all">All 9 Lessons (57 Tracks)</button>']
    for l in lessons:
        lesson_buttons.append(
            f'<button class="filter-btn" data-lesson="{l["number"]}">'
            f'Lesson {l["number"]} ({len(l["tracks"])})</button>'
        )

    sections_html = []
    for l in lessons:
        tracks_cards = []
        for tr in l["tracks"]:
            paras_html = "\n".join(
                f"<p>{format_inline_md(p)}</p>" for p in tr["paragraphs"]
            )
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
      <p>Verbatim Transcripts of All {total_tracks} Audio Tracks Across Lessons 1–9 (Speaker: Mike Holloway)</p>
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


def main():
    md_text = MD_PATH.read_text(encoding="utf-8")
    lessons = parse_markdown(md_text)
    total_tracks = sum(len(l["tracks"]) for l in lessons)
    print(f"Parsed {len(lessons)} lessons and {total_tracks} tracks.")
    assert len(lessons) == 9, f"Expected 9 lessons, got {len(lessons)}"
    assert total_tracks == 57, f"Expected 57 tracks, got {total_tracks}"
    for l in lessons:
        for tr in l["tracks"]:
            assert len(tr["paragraphs"]) > 0, f"Track {tr['number']} has no paragraphs!"

    build_docx(lessons, DOCX_PATH)
    build_html(lessons, HTML_PATH)
    build_html(lessons, INDEX_PATH)
    print(f"Generated {DOCX_PATH.name} ({DOCX_PATH.stat().st_size} bytes)")
    print(f"Generated {HTML_PATH.name} ({HTML_PATH.stat().st_size} bytes)")
    print(f"Generated {INDEX_PATH.name} ({INDEX_PATH.stat().st_size} bytes)")


if __name__ == "__main__":
    main()

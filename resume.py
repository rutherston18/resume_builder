"""
Resume Builder — Streamlit App
================================
Run:  streamlit run resume_builder.py

Requirements:
    pip install streamlit python-docx

Place iitm_logo.png next to this script (or any logo you want at the top-right).
All data is stored in resume_data.json next to this script.
"""

import streamlit as st
import json
import os
import re
import copy
import uuid
import base64
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, Cm, Emu, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# ════════════════════════════════════════════════════════════
#  DATA FILE
# ════════════════════════════════════════════════════════════

SCRIPT_DIR = Path(__file__).parent
DATA_FILE = SCRIPT_DIR / "resume_data.json"
LOGO_FILE = SCRIPT_DIR / "iitm_logo.png"

DEFAULT_DATA = {
    "header": {
        "name": "Arsh Mathur",
        "rollNumber": "EP24B024",
        "institute": "INDIAN INSTITUTE OF TECHNOLOGY MADRAS",
        "extra": "xx/EE/xx/xx",
        "linkedin": "",
        "github": "",
    },
    "education": [
        {"id": "ed1", "program": "Dual Degree in Electrical Engineering", "institute": "Indian Institute of Technology, Madras", "score": "9.25", "year": "2026", "enabled": True},
        {"id": "ed2", "program": "CFA Level 1", "institute": "CFA Institute, USA", "score": "Passed", "year": "2021", "enabled": True},
        {"id": "ed3", "program": "Class XII (CBSE)", "institute": "", "score": "", "year": "2017", "enabled": True},
    ],
    "educationBullets": [
        {"id": "eb1", "text": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna", "enabled": True},
        {"id": "eb2", "text": "aliquam erat volutpat. Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex", "enabled": True},
        {"id": "eb3", "text": "Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat", "enabled": True},
    ],
    "publications": [
        {"id": "pub1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
        {"id": "pub2", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
    ],
    "conferences": [
        {"id": "conf1", "name": "Global Summit", "description": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy", "enabled": True},
        {"id": "conf2", "name": "SW, Economist, USA", "description": "Duis autem vel eum iriure dolor in hendrerit in vulputate velit.", "enabled": True},
    ],
    "experience": [
        {
            "id": "exp1", "company": "Lorem ipsum", "role": "Index Structuring", "period": "May \u2013 Jul\u201921",
            "headline": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut aliquam",
            "bullets": [
                {"id": "exp1b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "exp1b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "exp1b3", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore"},
            ],
            "enabled": True,
        },
        {
            "id": "exp2", "company": "Research Project", "role": "(Dr. Lorem ipsum.)", "period": "",
            "headline": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit,",
            "bullets": [
                {"id": "exp2b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "exp2b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip molet."},
                {"id": "exp2b3", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dol eu."},
                {"id": "exp2b4", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis te."},
            ],
            "enabled": True,
        },
        {
            "id": "exp3", "company": "Lorem Ipsum", "role": "Data Science", "period": "May - July\u201920",
            "headline": "Lorem ipsum dolor sit amet, consectetuer",
            "bullets": [
                {"id": "exp3b1", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit."},
                {"id": "exp3b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum eu."},
                {"id": "exp3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis te."},
            ],
            "enabled": True,
        },
    ],
    "projects": [
        {
            "id": "proj1", "title": "Project Title", "tech": "Python, TensorFlow", "period": "",
            "headline": "Brief project description here",
            "bullets": [
                {"id": "proj1b1", "text": "Key contribution or result with **bold** emphasis on important metrics"},
                {"id": "proj1b2", "text": "Another achievement or technical detail about the project"},
            ],
            "enabled": True,
        },
    ],
    "positions": [
        {
            "id": "pos1", "title": "Lorem ipsum", "role": "Core", "period": "Apr\u201920  May\u201921",
            "bullets": [
                {"id": "pos1b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "pos1b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut."},
                {"id": "pos1b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor"},
            ],
            "enabled": True,
        },
        {
            "id": "pos2", "title": "Lorem Ipsum", "role": "Coordinator", "period": "Apr 18 \u2013 Apr 20",
            "bullets": [
                {"id": "pos2b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut."},
                {"id": "pos2b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl."},
                {"id": "pos2b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "pos3", "title": "Lorem Ipsum", "role": "Consultant", "period": "",
            "bullets": [
                {"id": "pos3b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "pos3b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "pos3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor"},
            ],
            "enabled": True,
        },
    ],
    "extracurriculars": [
        {
            "id": "ec1", "title": "Winner, Lorem Ipsum\u201919", "subtitle": "",
            "bullets": [
                {"id": "ec1b1", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "ec1b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
            ],
            "enabled": True,
        },
        {
            "id": "ec2", "title": "Lorem Ipsum", "subtitle": "(Institute Team)",
            "bullets": [
                {"id": "ec2b1", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu."},
                {"id": "ec2b2", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "ec3", "title": "Lorem Ipsum", "subtitle": "Mentor",
            "bullets": [
                {"id": "ec3b1", "text": "aliquam erat volutpat. **Uni wa savui** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl."},
                {"id": "ec3b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
                {"id": "ec3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "ec4", "title": "Lorem Ipsum", "subtitle": "",
            "bullets": [
                {"id": "ec4b1", "text": "aliquam erat volutpat. **Uni wa savui** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl"},
                {"id": "ec4b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
            ],
            "enabled": True,
        },
    ],
    "miscellaneous": [
        {"id": "misc1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
        {"id": "misc2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut", "enabled": True},
        {"id": "misc3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor", "enabled": True},
        {"id": "misc4", "text": "Ui wausau eros et ul praesent zzril riot uizzu dolor amet vizzu eros ul et al aua.", "enabled": True},
    ],
    "skills": {
        "advanced": "MS Excel, Python",
        "intermediate": "MS Powerpoint, MS Word, RStudio",
        "basic": "Bloomberg, Power BI",
    },
    "sectionToggles": {
        "education": True,
        "publications": True,
        "experience": True,
        "projects": True,
        "positions": True,
        "extracurriculars": True,
        "miscellaneous": True,
        "skills": True,
    },
}


def load_data():
    if DATA_FILE.exists():
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for key in DEFAULT_DATA:
            if key not in data:
                data[key] = copy.deepcopy(DEFAULT_DATA[key])
        if "projects" not in data.get("sectionToggles", {}):
            data.setdefault("sectionToggles", {})["projects"] = True
        if "linkedin" not in data.get("header", {}):
            data["header"]["linkedin"] = ""
        if "github" not in data.get("header", {}):
            data["header"]["github"] = ""
        return data
    return copy.deepcopy(DEFAULT_DATA)


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def uid():
    return uuid.uuid4().hex[:8]


# ════════════════════════════════════════════════════════════
#  DOCX GENERATION  (raw OoXML for pixel-perfect tables)
# ════════════════════════════════════════════════════════════

FONT_NAME = "Calibri"
FONT_SIZE_HP = "20"      # 10pt in half-points
HEADER_SIZE_HP = "25"    # 12.5pt
SECTION_FILL = "595959"
LABEL_FILL = "e7e6e6"

EDU_COL_WIDTHS = [2775, 1935, 105, 4005, 1575, 1521]  # total 11916
MAIN_COL_WIDTHS = [1635, 3075, 105, 4005, 1575, 1521]  # total 11916
PUB_COL_WIDTHS = [1620, 3090, 105, 4005, 1575, 1521]   # total 11916
TABLE_INDENT = "-1278"  # symmetric: 11916 = 9360 + 2*1278

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def _esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _rPr(bold=False, italic=False, size=FONT_SIZE_HP, color=None, font=FONT_NAME, underline=False):
    p = [f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:eastAsia="{font}"/>']
    if bold:
        p.append('<w:b/><w:bCs/>')
    if italic:
        p.append('<w:i/><w:iCs/>')
    if underline:
        p.append('<w:u w:val="single"/>')
    p.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    if color:
        p.append(f'<w:color w:val="{color}"/>')
    return "<w:rPr>" + "".join(p) + "</w:rPr>"


def _run(text, **kw):
    rpr = _rPr(**kw)
    t = _esc(text)
    sp = ' xml:space="preserve"' if t and (t[0] == ' ' or t[-1] == ' ') else ''
    return f'<w:r>{rpr}<w:t{sp}>{t}</w:t></w:r>'


def _bold_runs(text, base_bold=False, base_italic=False, size=FONT_SIZE_HP, color=None):
    parts = re.split(r"(\*\*.*?\*\*)", text)
    out = []
    for p in parts:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            out.append(_run(p[2:-2], bold=True, italic=base_italic, size=size, color=color))
        else:
            out.append(_run(p, bold=base_bold, italic=base_italic, size=size, color=color))
    return "".join(out)


def _para(runs_xml, align=None, justify=False, bullet=False, numId=None):
    pp = []
    if bullet and numId:
        pp.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numId}"/></w:numPr>')
    pp.append('<w:spacing w:line="240" w:lineRule="auto"/>')
    if justify:
        pp.append('<w:jc w:val="both"/>')
    elif align:
        pp.append(f'<w:jc w:val="{align}"/>')
    pp.append(_rPr())
    return f'<w:p><w:pPr>{"".join(pp)}</w:pPr>{runs_xml}</w:p>'


def _cell(width, content, shading=None, gridSpan=None, vMerge=None, valign="center"):
    tp = []
    if gridSpan and int(gridSpan) > 1:
        tp.append(f'<w:gridSpan w:val="{gridSpan}"/>')
    if vMerge is not None:
        tp.append(f'<w:vMerge w:val="{vMerge}"/>' if vMerge == "restart" else '<w:vMerge/>')
    tp.append(f'<w:tcW w:w="{width}" w:type="dxa"/>')
    if shading:
        tp.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>')
    tp.append(f'<w:vAlign w:val="{valign}"/>')
    return f'<w:tc><w:tcPr>{"".join(tp)}</w:tcPr>{content}</w:tc>'


def _row(cells):
    return f'<w:tr>{"".join(cells)}</w:tr>'


def _table(col_widths, rows):
    tw = sum(col_widths)
    gc = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    return f"""<w:tbl xmlns:w="{W}">
<w:tblPr>
  <w:tblW w:w="{tw}" w:type="dxa"/>
  <w:tblInd w:w="{TABLE_INDENT}" w:type="dxa"/>
  <w:tblBorders>
    <w:top w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:left w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:bottom w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:right w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:insideH w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:insideV w:val="single" w:sz="4" w:color="000000" w:space="0"/>
  </w:tblBorders>
  <w:tblLayout w:type="fixed"/>
  <w:jc w:val="left"/>
</w:tblPr>
<w:tblGrid>{gc}</w:tblGrid>
{"".join(rows)}
</w:tbl>"""


def _section_hdr(text, cw):
    p = _para(_run(text, bold=True, color="FFFFFF"), align="center")
    c = _cell(sum(cw), p, shading=SECTION_FILL, gridSpan=str(len(cw)))
    return _row([c])


def _label_row(label_lines, content_xml, cw):
    runs = []
    for i, (txt, b, it) in enumerate(label_lines):
        if i > 0:
            txt = "\n" + txt
        runs.append(_run(txt, bold=b, italic=it))
    lp = _para("".join(runs), align="center")
    lc = _cell(cw[0], lp, shading=LABEL_FILL, valign="center")
    cc = _cell(sum(cw[1:]), content_xml, gridSpan=str(len(cw) - 1), valign="top")
    return _row([lc, cc])


def _content_hl_bullets(headline, bullets, nid):
    ps = []
    if headline:
        ps.append(_para(_bold_runs(headline, base_bold=True, base_italic=True), justify=True))
    for b in bullets:
        t = b.get("text", "") if isinstance(b, dict) else b
        ps.append(_para(_bold_runs(t), bullet=True, numId=nid))
    return "".join(ps) or _para("")


def _content_bullets(bullets, nid):
    ps = []
    for b in bullets:
        t = b.get("text", "") if isinstance(b, dict) else b
        ps.append(_para(_bold_runs(t), bullet=True, numId=nid))
    return "".join(ps) or _para("")


def generate_docx(data):
    doc = Document()

    # Page setup
    sec = doc.sections[0]
    sec.page_width = Emu(7772400)
    sec.page_height = Emu(10058400)
    sec.left_margin = Emu(914400)
    sec.right_margin = Emu(914400)
    sec.top_margin = Emu(171450)
    sec.bottom_margin = Emu(0)

    # Default style
    ns = doc.styles["Normal"]
    ns.font.name = FONT_NAME
    ns.font.size = Pt(10)
    ns.paragraph_format.space_before = Pt(0)
    ns.paragraph_format.space_after = Pt(0)
    ns.paragraph_format.line_spacing = 1.0

    # Bullet numbering definition
    np = doc.part.numbering_part
    ne = np.numbering_definitions._numbering
    exab = ne.findall(qn("w:abstractNum"))
    naid = max((int(a.get(qn("w:abstractNumId"), 0)) for a in exab), default=-1) + 1
    exn = ne.findall(qn("w:num"))
    nnid = max((int(n.get(qn("w:numId"), 0)) for n in exn), default=0) + 1

    ne.append(parse_xml(f"""<w:abstractNum w:abstractNumId="{naid}" xmlns:w="{W}">
      <w:multiLevelType w:val="hybridMultilevel"/>
      <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val="\u2022"/>
        <w:lvlJc w:val="left"/>
        <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/><w:sz w:val="{FONT_SIZE_HP}"/></w:rPr>
      </w:lvl>
    </w:abstractNum>"""))
    ne.append(parse_xml(f"""<w:num w:numId="{nnid}" xmlns:w="{W}">
      <w:abstractNumId w:val="{naid}"/>
    </w:num>"""))
    BN = str(nnid)

    # ── Header paragraph ──
    h = data["header"]
    htxt = f'{h["name"]} | {h["rollNumber"]} | {h["institute"]} | {h["extra"]}'

    img_run = ""
    if LOGO_FILE.exists():
        rid, _img = doc.part.get_or_add_image(str(LOGO_FILE))
        img_run = f"""<w:r>{_rPr(bold=True, size=HEADER_SIZE_HP)}
          <w:drawing xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">
            <wp:anchor allowOverlap="1" behindDoc="0" distB="0" distT="0"
                       distL="114300" distR="114300" hidden="0"
                       layoutInCell="1" locked="0" relativeHeight="0" simplePos="0">
              <wp:simplePos x="0" y="0"/>
              <wp:positionH relativeFrom="page"><wp:posOffset>7239000</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="page"><wp:posOffset>57150</wp:posOffset></wp:positionV>
              <wp:extent cx="342900" cy="342900"/>
              <wp:effectExtent b="0" l="0" r="0" t="0"/>
              <wp:wrapSquare wrapText="bothSides" distB="0" distT="0" distL="114300" distR="114300"/>
              <wp:docPr id="1" name="logo.png"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr preferRelativeResize="0"/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed="{rid}"/><a:srcRect b="0" l="0" r="0" t="0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="342900" cy="342900"/></a:xfrm><a:prstGeom prst="rect"/><a:ln/></pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing>
        </w:r>"""

    link_runs = ""
    link_rpr = _rPr(bold=True, size=HEADER_SIZE_HP, color="0563C1", underline=True)

    link_runs = ""
    hl_reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    if h.get("linkedin"):
        rl = doc.part.relate_to(h["linkedin"], hl_reltype, is_external=True)
        link_runs += _run(" | ", bold=True, size=HEADER_SIZE_HP)
        link_runs += f'<w:hyperlink r:id="{rl}" xmlns:r="{R_NS}"><w:r>{link_rpr}<w:t>LinkedIn</w:t></w:r></w:hyperlink>'
    if h.get("github"):
        rg = doc.part.relate_to(h["github"], hl_reltype, is_external=True)
        link_runs += _run(" | ", bold=True, size=HEADER_SIZE_HP)
        link_runs += f'<w:hyperlink r:id="{rg}" xmlns:r="{R_NS}"><w:r>{link_rpr}<w:t>GitHub</w:t></w:r></w:hyperlink>'

    hdr_xml = f"""<w:p xmlns:w="{W}" xmlns:r="{R_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}">
      <w:pPr>
        <w:spacing w:before="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:left="-1080" w:right="-1080" w:firstLine="0"/>
        <w:rPr>{_rPr(bold=True, size=HEADER_SIZE_HP).replace("<w:rPr>","").replace("</w:rPr>","")}</w:rPr>
      </w:pPr>
      {img_run}
      {_run(htxt, bold=True, size=HEADER_SIZE_HP)}
      {link_runs}
    </w:p>"""

    # Save sectPr before clearing body (it contains page setup)
    sectPr = doc.element.body.find(qn("w:sectPr"))
    sectPr_copy = copy.deepcopy(sectPr) if sectPr is not None else None

    doc.element.body.clear()
    doc.element.body.append(parse_xml(hdr_xml))

    # Restore sectPr at end after all content is added (we'll do it at the very end)

    toggles = data.get("sectionToggles", {})

    # ═══ EDUCATION ═══
    if toggles.get("education", True):
        eis = [e for e in data.get("education", []) if e.get("enabled", True)]
        ebs = [b for b in data.get("educationBullets", []) if b.get("enabled", True)]
        rows = [_section_hdr("EDUCATION AND SCHOLASTIC ACHIEVEMENTS", EDU_COL_WIDTHS)]
        # Sub-header
        pw = sum(EDU_COL_WIDTHS[:3])
        hcells = [
            _cell(pw, _para(_run("Program", bold=True), align="center"), shading=LABEL_FILL, gridSpan="3"),
            _cell(EDU_COL_WIDTHS[3], _para(_run("Institute", bold=True), align="center"), shading=LABEL_FILL),
            _cell(EDU_COL_WIDTHS[4], _para(_run("% / CGPA", bold=True), align="center"), shading=LABEL_FILL),
            _cell(EDU_COL_WIDTHS[5], _para(_run("Year", bold=True), align="center"), shading=LABEL_FILL),
        ]
        rows.append(_row(hcells))
        for ed in eis:
            rows.append(_row([
                _cell(pw, _para(_run(ed["program"]), align="center"), gridSpan="3"),
                _cell(EDU_COL_WIDTHS[3], _para(_run(ed["institute"]), align="center")),
                _cell(EDU_COL_WIDTHS[4], _para(_run(ed["score"]), align="center")),
                _cell(EDU_COL_WIDTHS[5], _para(_run(ed["year"]), align="center")),
            ]))
        if ebs:
            rows.append(_row([_cell(sum(EDU_COL_WIDTHS), _content_bullets(ebs, BN), gridSpan="6", valign="top")]))
        doc.element.body.append(parse_xml(_table(EDU_COL_WIDTHS, rows)))

    # ═══ PUBLICATIONS ═══
    if toggles.get("publications", True):
        pubs = [p for p in data.get("publications", []) if p.get("enabled", True)]
        confs = [c for c in data.get("conferences", []) if c.get("enabled", True)]
        if pubs or confs:
            rows = [_section_hdr("PUBLICATIONS AND CONFERENCES", PUB_COL_WIDTHS)]
            if pubs:
                pxml = "".join(_para(_bold_runs(p["text"])) for p in pubs)
                rows.append(_row([
                    _cell(PUB_COL_WIDTHS[0], _para(_run("Journal Publication", bold=True), align="center"), shading=LABEL_FILL),
                    _cell(sum(PUB_COL_WIDTHS[1:]), pxml, gridSpan="5", valign="top"),
                ]))
            for ci, c in enumerate(confs):
                rows.append(_row([
                    _cell(PUB_COL_WIDTHS[0], _para(_run("Conferences", bold=True), align="center"),
                          shading=LABEL_FILL, vMerge="restart" if ci == 0 else "continue"),
                    _cell(PUB_COL_WIDTHS[1], _para(_run(c["name"], bold=True))),
                    _cell(sum(PUB_COL_WIDTHS[2:]), _para(_bold_runs(c["description"])), gridSpan="4", valign="top"),
                ]))
            doc.element.body.append(parse_xml(_table(PUB_COL_WIDTHS, rows)))

    # ═══ BIG TABLE ═══
    brows = []

    def _add_items(title, items, itype):
        fl = [x for x in items if x.get("enabled", True)]
        if not fl:
            return
        brows.append(_section_hdr(title, MAIN_COL_WIDTHS))
        for it in fl:
            lines = []
            if itype == "experience":
                lines.append((it.get("company", ""), True, False))
                if it.get("role"): lines.append((it["role"], False, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "project":
                lines.append((it.get("title", ""), True, False))
                if it.get("tech"): lines.append((it["tech"], False, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "position":
                lines.append((it.get("title", ""), True, False))
                if it.get("role"): lines.append((it["role"], True, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "extracurricular":
                lines.append((it.get("title", ""), True, False))
                if it.get("subtitle"): lines.append((it["subtitle"], False, False))
            cxml = _content_hl_bullets(it.get("headline", ""), it.get("bullets", []), BN)
            brows.append(_label_row(lines, cxml, MAIN_COL_WIDTHS))

    if toggles.get("experience", True):
        _add_items("PROFESSIONAL EXPERIENCE", data.get("experience", []), "experience")
    if toggles.get("projects", True):
        _add_items("PROJECTS", data.get("projects", []), "project")
    if toggles.get("positions", True):
        _add_items("POSITIONS OF RESPONSIBILITY", data.get("positions", []), "position")
    if toggles.get("extracurriculars", True):
        _add_items("EXTRA-CURRICULARS", data.get("extracurriculars", []), "extracurricular")

    if toggles.get("miscellaneous", True):
        mi = [m for m in data.get("miscellaneous", []) if m.get("enabled", True)]
        if mi:
            brows.append(_label_row([("Miscellaneous", True, False)], _content_bullets(mi, BN), MAIN_COL_WIDTHS))

    if toggles.get("skills", True):
        sk = data.get("skills", {})
        parts = []
        if sk.get("advanced"): parts.append(f'**Advanced** - {sk["advanced"]}')
        if sk.get("intermediate"): parts.append(f'**Intermediate** - {sk["intermediate"]}')
        if sk.get("basic"): parts.append(f'**Basic** -- {sk["basic"]}')
        if parts:
            brows.append(_label_row([("Software Skills", True, False)], _para(_bold_runs("; ".join(parts))), MAIN_COL_WIDTHS))

    if brows:
        doc.element.body.append(parse_xml(_table(MAIN_COL_WIDTHS, brows)))

    # Restore section properties (page size, margins)
    if sectPr_copy is not None:
        doc.element.body.append(sectPr_copy)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="Resume Builder", layout="wide", initial_sidebar_state="expanded")

    st.markdown("""
    <style>
    .block-container { max-width: 900px; padding-top: 1rem; }
    div[data-testid="stExpander"] { border: 1px solid #333; border-radius: 8px; margin-bottom: 0.5rem; }
    .stButton > button { font-size: 0.85rem; }
    h1 { font-size: 1.6rem !important; }
    h3 { font-size: 1.05rem !important; margin-bottom: 0.3rem !important; }
    </style>
    """, unsafe_allow_html=True)

    if "data" not in st.session_state:
        st.session_state.data = load_data()
    data = st.session_state.data

    with st.sidebar:
        st.title("Resume Builder")
        st.caption("IIT Madras Format")
        if not LOGO_FILE.exists():
            st.warning(f"Place iitm_logo.png next to this script for the logo.")
        st.divider()
        st.subheader("Section Toggles")
        for key, label in {
            "education": "Education & Achievements",
            "publications": "Publications & Conferences",
            "experience": "Professional Experience",
            "projects": "Projects",
            "positions": "Positions of Responsibility",
            "extracurriculars": "Extra-Curriculars",
            "miscellaneous": "Miscellaneous",
            "skills": "Software Skills",
        }.items():
            data["sectionToggles"][key] = st.checkbox(label, value=data["sectionToggles"].get(key, True), key=f"toggle_{key}")
        st.divider()
        if st.button("Generate DOCX", type="primary", use_container_width=True):
            save_data(data)
            try:
                st.session_state.docx_bytes = generate_docx(data)
                st.success("Resume generated!")
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc())
        if "docx_bytes" in st.session_state:
            st.download_button("Download Resume.docx", data=st.session_state.docx_bytes,
                file_name="Resume.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        st.divider()
        c1, c2 = st.columns(2)
        if c1.button("Export JSON", use_container_width=True):
            save_data(data); st.session_state.show_export = True
        if c2.button("Reset All", use_container_width=True):
            st.session_state.data = copy.deepcopy(DEFAULT_DATA); save_data(st.session_state.data); st.rerun()
        if st.session_state.get("show_export"):
            st.download_button("Download JSON", data=json.dumps(data, indent=2, ensure_ascii=False), file_name="resume_data.json", mime="application/json", use_container_width=True)
        uploaded = st.file_uploader("Import JSON", type=["json"], key="import_json")
        if uploaded:
            try:
                imp = json.loads(uploaded.read().decode("utf-8")); st.session_state.data = imp; save_data(imp); st.success("Imported!"); st.rerun()
            except Exception as e:
                st.error(f"Invalid JSON: {e}")

    # ── HEADER ──
    with st.expander("HEADER", expanded=True):
        c1, c2 = st.columns(2)
        data["header"]["name"] = c1.text_input("Full Name", data["header"]["name"], key="h_name")
        data["header"]["rollNumber"] = c2.text_input("Roll Number", data["header"]["rollNumber"], key="h_roll")
        c3, c4 = st.columns(2)
        data["header"]["institute"] = c3.text_input("Institute", data["header"]["institute"], key="h_inst")
        data["header"]["extra"] = c4.text_input("Extra Info", data["header"]["extra"], key="h_extra")
        c5, c6 = st.columns(2)
        data["header"]["linkedin"] = c5.text_input("LinkedIn URL", data["header"].get("linkedin", ""), key="h_li")
        data["header"]["github"] = c6.text_input("GitHub URL", data["header"].get("github", ""), key="h_gh")

    # ── EDUCATION ──
    if data["sectionToggles"].get("education", True):
        with st.expander("EDUCATION & SCHOLASTIC ACHIEVEMENTS"):
            for i, ed in enumerate(data["education"]):
                cols = st.columns([0.5, 3, 2.5, 1, 1, 0.5])
                ed["enabled"] = cols[0].checkbox("On", ed.get("enabled", True), key=f"ed_en_{i}", label_visibility="collapsed")
                ed["program"] = cols[1].text_input("Program", ed["program"], key=f"ed_prog_{i}", label_visibility="collapsed")
                ed["institute"] = cols[2].text_input("Institute", ed["institute"], key=f"ed_inst_{i}", label_visibility="collapsed")
                ed["score"] = cols[3].text_input("Score", ed["score"], key=f"ed_score_{i}", label_visibility="collapsed")
                ed["year"] = cols[4].text_input("Year", ed["year"], key=f"ed_year_{i}", label_visibility="collapsed")
                if cols[5].button("X", key=f"ed_del_{i}"): data["education"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Education", key="add_edu"):
                data["education"].append({"id": uid(), "program": "", "institute": "", "score": "", "year": "", "enabled": True}); save_data(data); st.rerun()
            st.markdown("---"); st.caption("Achievement Bullets")
            for i, b in enumerate(data["educationBullets"]):
                cols = st.columns([0.5, 8, 0.5])
                b["enabled"] = cols[0].checkbox("On", b.get("enabled", True), key=f"eb_en_{i}", label_visibility="collapsed")
                b["text"] = cols[1].text_input("Bullet", b["text"], key=f"eb_txt_{i}", label_visibility="collapsed")
                if cols[2].button("X", key=f"eb_del_{i}"): data["educationBullets"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Achievement Bullet", key="add_eb"):
                data["educationBullets"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()

    # ── PUBLICATIONS ──
    if data["sectionToggles"].get("publications", True):
        with st.expander("PUBLICATIONS & CONFERENCES"):
            st.caption("Journal Publications")
            for i, p in enumerate(data["publications"]):
                cols = st.columns([0.5, 8, 0.5])
                p["enabled"] = cols[0].checkbox("On", p.get("enabled", True), key=f"pub_en_{i}", label_visibility="collapsed")
                p["text"] = cols[1].text_area("Text", p["text"], key=f"pub_txt_{i}", height=60, label_visibility="collapsed")
                if cols[2].button("X", key=f"pub_del_{i}"): data["publications"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Publication", key="add_pub"):
                data["publications"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()
            st.markdown("---"); st.caption("Conferences")
            for i, c in enumerate(data["conferences"]):
                cols = st.columns([0.5, 3, 5, 0.5])
                c["enabled"] = cols[0].checkbox("On", c.get("enabled", True), key=f"conf_en_{i}", label_visibility="collapsed")
                c["name"] = cols[1].text_input("Name", c["name"], key=f"conf_name_{i}", label_visibility="collapsed")
                c["description"] = cols[2].text_input("Description", c["description"], key=f"conf_desc_{i}", label_visibility="collapsed")
                if cols[3].button("X", key=f"conf_del_{i}"): data["conferences"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Conference", key="add_conf"):
                data["conferences"].append({"id": uid(), "name": "", "description": "", "enabled": True}); save_data(data); st.rerun()

    # ── Generic section editor ──
    def edit_section(sk, sl, fields, has_hl=False):
        items = data.get(sk, [])
        if not data["sectionToggles"].get(sk, True): return
        with st.expander(sl.upper()):
            for i, item in enumerate(items):
                lbl = " - ".join(filter(None, [item.get(fields[0][0], ""), item.get(fields[1][0], "") if len(fields) > 1 else ""]))
                st.markdown(f"**{lbl or 'New Item'}**")
                tc = st.columns([0.5] + [1]*len(fields) + [0.5])
                item["enabled"] = tc[0].checkbox("On", item.get("enabled", True), key=f"{sk}_en_{i}", label_visibility="collapsed")
                for fi, (fk, fl) in enumerate(fields):
                    item[fk] = tc[fi+1].text_input(fl, item.get(fk, ""), key=f"{sk}_{fk}_{i}")
                if tc[-1].button("X", key=f"{sk}_del_{i}"): items.pop(i); save_data(data); st.rerun()
                if has_hl:
                    item["headline"] = st.text_input("Headline (bold italic)", item.get("headline", ""), key=f"{sk}_hl_{i}")
                st.caption("Bullets (use **bold** for emphasis)")
                for bi, b in enumerate(item.get("bullets", [])):
                    bc = st.columns([9, 0.5])
                    b["text"] = bc[0].text_input(f"Bullet {bi+1}", b["text"], key=f"{sk}_b_{i}_{bi}", label_visibility="collapsed")
                    if bc[1].button("X", key=f"{sk}_bd_{i}_{bi}"): item["bullets"].pop(bi); save_data(data); st.rerun()
                if st.button("+ Add Bullet", key=f"{sk}_ab_{i}"):
                    item.setdefault("bullets", []).append({"id": uid(), "text": ""}); save_data(data); st.rerun()
                st.divider()
            ni = {"id": uid(), "enabled": True, "bullets": [{"id": uid(), "text": ""}]}
            for fk, _ in fields: ni[fk] = ""
            if has_hl: ni["headline"] = ""
            if st.button(f"+ Add {sl}", key=f"add_{sk}"): items.append(ni); save_data(data); st.rerun()

    edit_section("experience", "Professional Experience", [("company", "Company"), ("role", "Role"), ("period", "Period")], has_hl=True)
    edit_section("projects", "Project", [("title", "Title"), ("tech", "Tech / Context"), ("period", "Period")], has_hl=True)
    edit_section("positions", "Position of Responsibility", [("title", "Title"), ("role", "Role"), ("period", "Period")])
    edit_section("extracurriculars", "Extra-Curricular", [("title", "Title"), ("subtitle", "Subtitle")])

    if data["sectionToggles"].get("miscellaneous", True):
        with st.expander("MISCELLANEOUS"):
            for i, m in enumerate(data["miscellaneous"]):
                cols = st.columns([0.5, 8, 0.5])
                m["enabled"] = cols[0].checkbox("On", m.get("enabled", True), key=f"misc_en_{i}", label_visibility="collapsed")
                m["text"] = cols[1].text_input("Item", m["text"], key=f"misc_txt_{i}", label_visibility="collapsed")
                if cols[2].button("X", key=f"misc_del_{i}"): data["miscellaneous"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Misc Item", key="add_misc"):
                data["miscellaneous"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()

    if data["sectionToggles"].get("skills", True):
        with st.expander("SOFTWARE SKILLS"):
            data["skills"]["advanced"] = st.text_input("Advanced", data["skills"].get("advanced", ""), key="sk_adv")
            data["skills"]["intermediate"] = st.text_input("Intermediate", data["skills"].get("intermediate", ""), key="sk_int")
            data["skills"]["basic"] = st.text_input("Basic", data["skills"].get("basic", ""), key="sk_bas")

    save_data(data)

if __name__ == "__main__":
    main()
    
"""
Resume Builder — Streamlit App
================================
Run:  streamlit run resume_builder.py

Requirements:
    pip install streamlit python-docx

Place iitm_logo.png next to this script (or any logo you want at the top-right).
All data is stored in resume_data.json next to this script.
"""

import streamlit as st
import json
import os
import re
import copy
import uuid
import base64
from io import BytesIO
from pathlib import Path

from docx import Document
from docx.shared import Pt, Inches, Cm, Emu, Twips, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

# ════════════════════════════════════════════════════════════
#  DATA FILE
# ════════════════════════════════════════════════════════════

SCRIPT_DIR = Path(__file__).parent
DATA_FILE = SCRIPT_DIR / "resume_data.json"
LOGO_FILE = SCRIPT_DIR / "iitm_logo.png"

DEFAULT_DATA = {
    "header": {
        "name": "Arsh Mathur",
        "rollNumber": "EP24B024",
        "institute": "INDIAN INSTITUTE OF TECHNOLOGY MADRAS",
        "extra": "xx/EE/xx/xx",
        "linkedin": "",
        "github": "",
    },
    "education": [
        {"id": "ed1", "program": "Dual Degree in Electrical Engineering", "institute": "Indian Institute of Technology, Madras", "score": "9.25", "year": "2026", "enabled": True},
        {"id": "ed2", "program": "CFA Level 1", "institute": "CFA Institute, USA", "score": "Passed", "year": "2021", "enabled": True},
        {"id": "ed3", "program": "Class XII (CBSE)", "institute": "", "score": "", "year": "2017", "enabled": True},
    ],
    "educationBullets": [
        {"id": "eb1", "text": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna", "enabled": True},
        {"id": "eb2", "text": "aliquam erat volutpat. Ut wisi enim ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip ex", "enabled": True},
        {"id": "eb3", "text": "Duis autem vel eum iriure dolor in hendrerit in vulputate velit esse molestie consequat, vel illum dolore eu feugiat", "enabled": True},
    ],
    "publications": [
        {"id": "pub1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
        {"id": "pub2", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
    ],
    "conferences": [
        {"id": "conf1", "name": "Global Summit", "description": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy", "enabled": True},
        {"id": "conf2", "name": "SW, Economist, USA", "description": "Duis autem vel eum iriure dolor in hendrerit in vulputate velit.", "enabled": True},
    ],
    "experience": [
        {
            "id": "exp1", "company": "Lorem ipsum", "role": "Index Structuring", "period": "May \u2013 Jul\u201921",
            "headline": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut aliquam",
            "bullets": [
                {"id": "exp1b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "exp1b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "exp1b3", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore"},
            ],
            "enabled": True,
        },
        {
            "id": "exp2", "company": "Research Project", "role": "(Dr. Lorem ipsum.)", "period": "",
            "headline": "Lorem ipsum dolor sit amet, consectetuer adipiscing elit,",
            "bullets": [
                {"id": "exp2b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "exp2b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut aliquip molet."},
                {"id": "exp2b3", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dol eu."},
                {"id": "exp2b4", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis te."},
            ],
            "enabled": True,
        },
        {
            "id": "exp3", "company": "Lorem Ipsum", "role": "Data Science", "period": "May - July\u201920",
            "headline": "Lorem ipsum dolor sit amet, consectetuer",
            "bullets": [
                {"id": "exp3b1", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit."},
                {"id": "exp3b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum eu."},
                {"id": "exp3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis te."},
            ],
            "enabled": True,
        },
    ],
    "projects": [
        {
            "id": "proj1", "title": "Project Title", "tech": "Python, TensorFlow", "period": "",
            "headline": "Brief project description here",
            "bullets": [
                {"id": "proj1b1", "text": "Key contribution or result with **bold** emphasis on important metrics"},
                {"id": "proj1b2", "text": "Another achievement or technical detail about the project"},
            ],
            "enabled": True,
        },
    ],
    "positions": [
        {
            "id": "pos1", "title": "Lorem ipsum", "role": "Core", "period": "Apr\u201920  May\u201921",
            "bullets": [
                {"id": "pos1b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "pos1b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut."},
                {"id": "pos1b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor"},
            ],
            "enabled": True,
        },
        {
            "id": "pos2", "title": "Lorem Ipsum", "role": "Coordinator", "period": "Apr 18 \u2013 Apr 20",
            "bullets": [
                {"id": "pos2b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut."},
                {"id": "pos2b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl."},
                {"id": "pos2b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "pos3", "title": "Lorem Ipsum", "role": "Consultant", "period": "",
            "bullets": [
                {"id": "pos3b1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut"},
                {"id": "pos3b2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "pos3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor"},
            ],
            "enabled": True,
        },
    ],
    "extracurriculars": [
        {
            "id": "ec1", "title": "Winner, Lorem Ipsum\u201919", "subtitle": "",
            "bullets": [
                {"id": "ec1b1", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut"},
                {"id": "ec1b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
            ],
            "enabled": True,
        },
        {
            "id": "ec2", "title": "Lorem Ipsum", "subtitle": "(Institute Team)",
            "bullets": [
                {"id": "ec2b1", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu."},
                {"id": "ec2b2", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "ec3", "title": "Lorem Ipsum", "subtitle": "Mentor",
            "bullets": [
                {"id": "ec3b1", "text": "aliquam erat volutpat. **Uni wa savui** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl."},
                {"id": "ec3b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
                {"id": "ec3b3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor."},
            ],
            "enabled": True,
        },
        {
            "id": "ec4", "title": "Lorem Ipsum", "subtitle": "",
            "bullets": [
                {"id": "ec4b1", "text": "aliquam erat volutpat. **Uni wa savui** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl"},
                {"id": "ec4b2", "text": "Duis autem vel eum iriure dolor in hendrerit in **vulputate** velit esse molestie consequat, vel illum dolore eu"},
            ],
            "enabled": True,
        },
    ],
    "miscellaneous": [
        {"id": "misc1", "text": "Lorem ipsum **dolor** sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut", "enabled": True},
        {"id": "misc2", "text": "aliquam erat volutpat. **Ut wisi enim** ad minim veniam, quis nostrud exerci tation ullamcorper suscipit lobortis nisl ut", "enabled": True},
        {"id": "misc3", "text": "vero eros et accumsan et **iusto odio dignissim** qui blandit praesent luptatum zzril delenit augue duis dolor", "enabled": True},
        {"id": "misc4", "text": "Ui wausau eros et ul praesent zzril riot uizzu dolor amet vizzu eros ul et al aua.", "enabled": True},
    ],
    "skills": {
        "advanced": "MS Excel, Python",
        "intermediate": "MS Powerpoint, MS Word, RStudio",
        "basic": "Bloomberg, Power BI",
    },
    "sectionToggles": {
        "education": True,
        "publications": True,
        "experience": True,
        "projects": True,
        "positions": True,
        "extracurriculars": True,
        "miscellaneous": True,
        "skills": True,
    },
}


def load_data():
    if DATA_FILE.exists():
        with open(DATA_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        for key in DEFAULT_DATA:
            if key not in data:
                data[key] = copy.deepcopy(DEFAULT_DATA[key])
        if "projects" not in data.get("sectionToggles", {}):
            data.setdefault("sectionToggles", {})["projects"] = True
        if "linkedin" not in data.get("header", {}):
            data["header"]["linkedin"] = ""
        if "github" not in data.get("header", {}):
            data["header"]["github"] = ""
        return data
    return copy.deepcopy(DEFAULT_DATA)


def save_data(data):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def uid():
    return uuid.uuid4().hex[:8]


# ════════════════════════════════════════════════════════════
#  DOCX GENERATION  (raw OoXML for pixel-perfect tables)
# ════════════════════════════════════════════════════════════

FONT_NAME = "Calibri"
FONT_SIZE_HP = "20"      # 10pt in half-points
HEADER_SIZE_HP = "25"    # 12.5pt
SECTION_FILL = "595959"
LABEL_FILL = "e7e6e6"

EDU_COL_WIDTHS = [2775, 1935, 105, 4005, 1575, 1500]
MAIN_COL_WIDTHS = [1635, 3075, 105, 4005, 1575, 1515]
PUB_COL_WIDTHS = [1620, 3090, 105, 4005, 1575, 1515]
TABLE_INDENT = "-1278"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def _esc(text):
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _rPr(bold=False, italic=False, size=FONT_SIZE_HP, color=None, font=FONT_NAME, underline=False):
    p = [f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" w:eastAsia="{font}"/>']
    if bold:
        p.append('<w:b/><w:bCs/>')
    if italic:
        p.append('<w:i/><w:iCs/>')
    if underline:
        p.append('<w:u w:val="single"/>')
    p.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    if color:
        p.append(f'<w:color w:val="{color}"/>')
    return "<w:rPr>" + "".join(p) + "</w:rPr>"


def _run(text, **kw):
    rpr = _rPr(**kw)
    t = _esc(text)
    sp = ' xml:space="preserve"' if t and (t[0] == ' ' or t[-1] == ' ') else ''
    return f'<w:r>{rpr}<w:t{sp}>{t}</w:t></w:r>'


def _bold_runs(text, base_bold=False, base_italic=False, size=FONT_SIZE_HP, color=None):
    parts = re.split(r"(\*\*.*?\*\*)", text)
    out = []
    for p in parts:
        if not p:
            continue
        if p.startswith("**") and p.endswith("**"):
            out.append(_run(p[2:-2], bold=True, italic=base_italic, size=size, color=color))
        else:
            out.append(_run(p, bold=base_bold, italic=base_italic, size=size, color=color))
    return "".join(out)


def _para(runs_xml, align=None, justify=False, bullet=False, numId=None):
    pp = []
    if bullet and numId:
        pp.append(f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{numId}"/></w:numPr>')
    pp.append('<w:spacing w:line="240" w:lineRule="auto"/>')
    if justify:
        pp.append('<w:jc w:val="both"/>')
    elif align:
        pp.append(f'<w:jc w:val="{align}"/>')
    pp.append(_rPr())
    return f'<w:p><w:pPr>{"".join(pp)}</w:pPr>{runs_xml}</w:p>'


def _cell(width, content, shading=None, gridSpan=None, vMerge=None, valign="center"):
    tp = []
    if gridSpan and int(gridSpan) > 1:
        tp.append(f'<w:gridSpan w:val="{gridSpan}"/>')
    if vMerge is not None:
        tp.append(f'<w:vMerge w:val="{vMerge}"/>' if vMerge == "restart" else '<w:vMerge/>')
    tp.append(f'<w:tcW w:w="{width}" w:type="dxa"/>')
    if shading:
        tp.append(f'<w:shd w:val="clear" w:color="auto" w:fill="{shading}"/>')
    tp.append(f'<w:vAlign w:val="{valign}"/>')
    return f'<w:tc><w:tcPr>{"".join(tp)}</w:tcPr>{content}</w:tc>'


def _row(cells):
    return f'<w:tr>{"".join(cells)}</w:tr>'


def _table(col_widths, rows):
    tw = sum(col_widths)
    gc = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths)
    return f"""<w:tbl xmlns:w="{W}">
<w:tblPr>
  <w:tblW w:w="{tw}" w:type="dxa"/>
  <w:tblInd w:w="{TABLE_INDENT}" w:type="dxa"/>
  <w:tblBorders>
    <w:top w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:left w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:bottom w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:right w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:insideH w:val="single" w:sz="4" w:color="000000" w:space="0"/>
    <w:insideV w:val="single" w:sz="4" w:color="000000" w:space="0"/>
  </w:tblBorders>
  <w:tblLayout w:type="fixed"/>
  <w:jc w:val="left"/>
</w:tblPr>
<w:tblGrid>{gc}</w:tblGrid>
{"".join(rows)}
</w:tbl>"""


def _section_hdr(text, cw):
    p = _para(_run(text, bold=True, color="FFFFFF"), align="center")
    c = _cell(sum(cw), p, shading=SECTION_FILL, gridSpan=str(len(cw)))
    return _row([c])


def _label_row(label_lines, content_xml, cw):
    runs = []
    for i, (txt, b, it) in enumerate(label_lines):
        if i > 0:
            txt = "\n" + txt
        runs.append(_run(txt, bold=b, italic=it))
    lp = _para("".join(runs), align="center")
    lc = _cell(cw[0], lp, shading=LABEL_FILL, valign="center")
    cc = _cell(sum(cw[1:]), content_xml, gridSpan=str(len(cw) - 1), valign="top")
    return _row([lc, cc])


def _content_hl_bullets(headline, bullets, nid):
    ps = []
    if headline:
        ps.append(_para(_bold_runs(headline, base_bold=True, base_italic=True), justify=True))
    for b in bullets:
        t = b.get("text", "") if isinstance(b, dict) else b
        ps.append(_para(_bold_runs(t), bullet=True, numId=nid))
    return "".join(ps) or _para("")


def _content_bullets(bullets, nid):
    ps = []
    for b in bullets:
        t = b.get("text", "") if isinstance(b, dict) else b
        ps.append(_para(_bold_runs(t), bullet=True, numId=nid))
    return "".join(ps) or _para("")


def generate_docx(data):
    doc = Document()

    # Page setup
    sec = doc.sections[0]
    sec.page_width = Emu(7772400)
    sec.page_height = Emu(10058400)
    sec.left_margin = Emu(914400)
    sec.right_margin = Emu(914400)
    sec.top_margin = Emu(171450)
    sec.bottom_margin = Emu(0)

    # Default style
    ns = doc.styles["Normal"]
    ns.font.name = FONT_NAME
    ns.font.size = Pt(10)
    ns.paragraph_format.space_before = Pt(0)
    ns.paragraph_format.space_after = Pt(0)
    ns.paragraph_format.line_spacing = 1.0

    # Bullet numbering definition
    np = doc.part.numbering_part
    ne = np.numbering_definitions._numbering
    exab = ne.findall(qn("w:abstractNum"))
    naid = max((int(a.get(qn("w:abstractNumId"), 0)) for a in exab), default=-1) + 1
    exn = ne.findall(qn("w:num"))
    nnid = max((int(n.get(qn("w:numId"), 0)) for n in exn), default=0) + 1

    ne.append(parse_xml(f"""<w:abstractNum w:abstractNumId="{naid}" xmlns:w="{W}">
      <w:multiLevelType w:val="hybridMultilevel"/>
      <w:lvl w:ilvl="0">
        <w:start w:val="1"/>
        <w:numFmt w:val="bullet"/>
        <w:lvlText w:val="\u2022"/>
        <w:lvlJc w:val="left"/>
        <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
        <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/><w:sz w:val="{FONT_SIZE_HP}"/></w:rPr>
      </w:lvl>
    </w:abstractNum>"""))
    ne.append(parse_xml(f"""<w:num w:numId="{nnid}" xmlns:w="{W}">
      <w:abstractNumId w:val="{naid}"/>
    </w:num>"""))
    BN = str(nnid)

    # ── Header paragraph ──
    h = data["header"]
    htxt = f'{h["name"]} | {h["rollNumber"]} | {h["institute"]} | {h["extra"]}'

    img_run = ""
    if LOGO_FILE.exists():
        rid, _img = doc.part.get_or_add_image(str(LOGO_FILE))
        img_run = f"""<w:r>{_rPr(bold=True, size=HEADER_SIZE_HP)}
          <w:drawing xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}" xmlns:r="{R_NS}">
            <wp:anchor allowOverlap="1" behindDoc="0" distB="0" distT="0"
                       distL="114300" distR="114300" hidden="0"
                       layoutInCell="1" locked="0" relativeHeight="0" simplePos="0">
              <wp:simplePos x="0" y="0"/>
              <wp:positionH relativeFrom="page"><wp:posOffset>7239000</wp:posOffset></wp:positionH>
              <wp:positionV relativeFrom="page"><wp:posOffset>57150</wp:posOffset></wp:positionV>
              <wp:extent cx="342900" cy="342900"/>
              <wp:effectExtent b="0" l="0" r="0" t="0"/>
              <wp:wrapSquare wrapText="bothSides" distB="0" distT="0" distL="114300" distR="114300"/>
              <wp:docPr id="1" name="logo.png"/>
              <a:graphic>
                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <pic:pic>
                    <pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr preferRelativeResize="0"/></pic:nvPicPr>
                    <pic:blipFill><a:blip r:embed="{rid}"/><a:srcRect b="0" l="0" r="0" t="0"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
                    <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="342900" cy="342900"/></a:xfrm><a:prstGeom prst="rect"/><a:ln/></pic:spPr>
                  </pic:pic>
                </a:graphicData>
              </a:graphic>
            </wp:anchor>
          </w:drawing>
        </w:r>"""

    link_runs = ""
    link_rpr = _rPr(bold=True, size=HEADER_SIZE_HP, color="0563C1", underline=True)

    link_runs = ""
    hl_reltype = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    if h.get("linkedin"):
        rl = doc.part.relate_to(h["linkedin"], hl_reltype, is_external=True)
        link_runs += _run(" | ", bold=True, size=HEADER_SIZE_HP)
        link_runs += f'<w:hyperlink r:id="{rl}" xmlns:r="{R_NS}"><w:r>{link_rpr}<w:t>LinkedIn</w:t></w:r></w:hyperlink>'
    if h.get("github"):
        rg = doc.part.relate_to(h["github"], hl_reltype, is_external=True)
        link_runs += _run(" | ", bold=True, size=HEADER_SIZE_HP)
        link_runs += f'<w:hyperlink r:id="{rg}" xmlns:r="{R_NS}"><w:r>{link_rpr}<w:t>GitHub</w:t></w:r></w:hyperlink>'

    hdr_xml = f"""<w:p xmlns:w="{W}" xmlns:r="{R_NS}" xmlns:wp="{WP_NS}" xmlns:a="{A_NS}" xmlns:pic="{PIC_NS}">
      <w:pPr>
        <w:spacing w:before="120" w:line="240" w:lineRule="auto"/>
        <w:ind w:left="-1080" w:right="-1080" w:firstLine="0"/>
        <w:rPr>{_rPr(bold=True, size=HEADER_SIZE_HP).replace("<w:rPr>","").replace("</w:rPr>","")}</w:rPr>
      </w:pPr>
      {img_run}
      {_run(htxt, bold=True, size=HEADER_SIZE_HP)}
      {link_runs}
    </w:p>"""

    # Save sectPr before clearing body (it contains page setup)
    sectPr = doc.element.body.find(qn("w:sectPr"))
    sectPr_copy = copy.deepcopy(sectPr) if sectPr is not None else None

    doc.element.body.clear()
    doc.element.body.append(parse_xml(hdr_xml))

    # Restore sectPr at end after all content is added (we'll do it at the very end)

    toggles = data.get("sectionToggles", {})

    # ═══ EDUCATION ═══
    if toggles.get("education", True):
        eis = [e for e in data.get("education", []) if e.get("enabled", True)]
        ebs = [b for b in data.get("educationBullets", []) if b.get("enabled", True)]
        rows = [_section_hdr("EDUCATION AND SCHOLASTIC ACHIEVEMENTS", EDU_COL_WIDTHS)]
        # Sub-header
        pw = sum(EDU_COL_WIDTHS[:3])
        hcells = [
            _cell(pw, _para(_run("Program", bold=True), align="center"), shading=LABEL_FILL, gridSpan="3"),
            _cell(EDU_COL_WIDTHS[3], _para(_run("Institute", bold=True), align="center"), shading=LABEL_FILL),
            _cell(EDU_COL_WIDTHS[4], _para(_run("% / CGPA", bold=True), align="center"), shading=LABEL_FILL),
            _cell(EDU_COL_WIDTHS[5], _para(_run("Year", bold=True), align="center"), shading=LABEL_FILL),
        ]
        rows.append(_row(hcells))
        for ed in eis:
            rows.append(_row([
                _cell(pw, _para(_run(ed["program"]), align="center"), gridSpan="3"),
                _cell(EDU_COL_WIDTHS[3], _para(_run(ed["institute"]), align="center")),
                _cell(EDU_COL_WIDTHS[4], _para(_run(ed["score"]), align="center")),
                _cell(EDU_COL_WIDTHS[5], _para(_run(ed["year"]), align="center")),
            ]))
        if ebs:
            rows.append(_row([_cell(sum(EDU_COL_WIDTHS), _content_bullets(ebs, BN), gridSpan="6", valign="top")]))
        doc.element.body.append(parse_xml(_table(EDU_COL_WIDTHS, rows)))

    # ═══ PUBLICATIONS ═══
    if toggles.get("publications", True):
        pubs = [p for p in data.get("publications", []) if p.get("enabled", True)]
        confs = [c for c in data.get("conferences", []) if c.get("enabled", True)]
        if pubs or confs:
            rows = [_section_hdr("PUBLICATIONS AND CONFERENCES", PUB_COL_WIDTHS)]
            if pubs:
                pxml = "".join(_para(_bold_runs(p["text"])) for p in pubs)
                rows.append(_row([
                    _cell(PUB_COL_WIDTHS[0], _para(_run("Journal Publication", bold=True), align="center"), shading=LABEL_FILL),
                    _cell(sum(PUB_COL_WIDTHS[1:]), pxml, gridSpan="5", valign="top"),
                ]))
            for ci, c in enumerate(confs):
                rows.append(_row([
                    _cell(PUB_COL_WIDTHS[0], _para(_run("Conferences", bold=True), align="center"),
                          shading=LABEL_FILL, vMerge="restart" if ci == 0 else "continue"),
                    _cell(PUB_COL_WIDTHS[1], _para(_run(c["name"], bold=True))),
                    _cell(sum(PUB_COL_WIDTHS[2:]), _para(_bold_runs(c["description"])), gridSpan="4", valign="top"),
                ]))
            doc.element.body.append(parse_xml(_table(PUB_COL_WIDTHS, rows)))

    # ═══ BIG TABLE ═══
    brows = []

    def _add_items(title, items, itype):
        fl = [x for x in items if x.get("enabled", True)]
        if not fl:
            return
        brows.append(_section_hdr(title, MAIN_COL_WIDTHS))
        for it in fl:
            lines = []
            if itype == "experience":
                lines.append((it.get("company", ""), True, False))
                if it.get("role"): lines.append((it["role"], False, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "project":
                lines.append((it.get("title", ""), True, False))
                if it.get("tech"): lines.append((it["tech"], False, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "position":
                lines.append((it.get("title", ""), True, False))
                if it.get("role"): lines.append((it["role"], True, False))
                if it.get("period"): lines.append((f'({it["period"]})', False, False))
            elif itype == "extracurricular":
                lines.append((it.get("title", ""), True, False))
                if it.get("subtitle"): lines.append((it["subtitle"], False, False))
            cxml = _content_hl_bullets(it.get("headline", ""), it.get("bullets", []), BN)
            brows.append(_label_row(lines, cxml, MAIN_COL_WIDTHS))

    if toggles.get("experience", True):
        _add_items("PROFESSIONAL EXPERIENCE", data.get("experience", []), "experience")
    if toggles.get("projects", True):
        _add_items("PROJECTS", data.get("projects", []), "project")
    if toggles.get("positions", True):
        _add_items("POSITIONS OF RESPONSIBILITY", data.get("positions", []), "position")
    if toggles.get("extracurriculars", True):
        _add_items("EXTRA-CURRICULARS", data.get("extracurriculars", []), "extracurricular")

    if toggles.get("miscellaneous", True):
        mi = [m for m in data.get("miscellaneous", []) if m.get("enabled", True)]
        if mi:
            brows.append(_label_row([("Miscellaneous", True, False)], _content_bullets(mi, BN), MAIN_COL_WIDTHS))

    if toggles.get("skills", True):
        sk = data.get("skills", {})
        parts = []
        if sk.get("advanced"): parts.append(f'**Advanced** - {sk["advanced"]}')
        if sk.get("intermediate"): parts.append(f'**Intermediate** - {sk["intermediate"]}')
        if sk.get("basic"): parts.append(f'**Basic** -- {sk["basic"]}')
        if parts:
            brows.append(_label_row([("Software Skills", True, False)], _para(_bold_runs("; ".join(parts))), MAIN_COL_WIDTHS))

    if brows:
        doc.element.body.append(parse_xml(_table(MAIN_COL_WIDTHS, brows)))

    # Restore section properties (page size, margins)
    if sectPr_copy is not None:
        doc.element.body.append(sectPr_copy)

    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ════════════════════════════════════════════════════════════
#  STREAMLIT UI
# ════════════════════════════════════════════════════════════

def main():
    st.set_page_config(page_title="Resume Builder", layout="wide", initial_sidebar_state="expanded")

    st.markdown("""
    <style>
    .block-container { max-width: 900px; padding-top: 1rem; }
    div[data-testid="stExpander"] { border: 1px solid #333; border-radius: 8px; margin-bottom: 0.5rem; }
    .stButton > button { font-size: 0.85rem; }
    h1 { font-size: 1.6rem !important; }
    h3 { font-size: 1.05rem !important; margin-bottom: 0.3rem !important; }
    </style>
    """, unsafe_allow_html=True)

    if "data" not in st.session_state:
        st.session_state.data = load_data()
    data = st.session_state.data

    with st.sidebar:
        st.title("Resume Builder")
        st.caption("IIT Madras Format")
        if not LOGO_FILE.exists():
            st.warning(f"Place iitm_logo.png next to this script for the logo.")
        st.divider()
        st.subheader("Section Toggles")
        for key, label in {
            "education": "Education & Achievements",
            "publications": "Publications & Conferences",
            "experience": "Professional Experience",
            "projects": "Projects",
            "positions": "Positions of Responsibility",
            "extracurriculars": "Extra-Curriculars",
            "miscellaneous": "Miscellaneous",
            "skills": "Software Skills",
        }.items():
            data["sectionToggles"][key] = st.checkbox(label, value=data["sectionToggles"].get(key, True), key=f"toggle_{key}")
        st.divider()
        if st.button("Generate DOCX", type="primary", use_container_width=True):
            save_data(data)
            try:
                st.session_state.docx_bytes = generate_docx(data)
                st.success("Resume generated!")
            except Exception as e:
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc())
        if "docx_bytes" in st.session_state:
            st.download_button("Download Resume.docx", data=st.session_state.docx_bytes,
                file_name="Resume.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
        st.divider()
        c1, c2 = st.columns(2)
        if c1.button("Export JSON", use_container_width=True):
            save_data(data); st.session_state.show_export = True
        if c2.button("Reset All", use_container_width=True):
            st.session_state.data = copy.deepcopy(DEFAULT_DATA); save_data(st.session_state.data); st.rerun()
        if st.session_state.get("show_export"):
            st.download_button("Download JSON", data=json.dumps(data, indent=2, ensure_ascii=False), file_name="resume_data.json", mime="application/json", use_container_width=True)
        uploaded = st.file_uploader("Import JSON", type=["json"], key="import_json")
        if uploaded:
            try:
                imp = json.loads(uploaded.read().decode("utf-8")); st.session_state.data = imp; save_data(imp); st.success("Imported!"); st.rerun()
            except Exception as e:
                st.error(f"Invalid JSON: {e}")

    # ── HEADER ──
    with st.expander("HEADER", expanded=True):
        c1, c2 = st.columns(2)
        data["header"]["name"] = c1.text_input("Full Name", data["header"]["name"], key="h_name")
        data["header"]["rollNumber"] = c2.text_input("Roll Number", data["header"]["rollNumber"], key="h_roll")
        c3, c4 = st.columns(2)
        data["header"]["institute"] = c3.text_input("Institute", data["header"]["institute"], key="h_inst")
        data["header"]["extra"] = c4.text_input("Extra Info", data["header"]["extra"], key="h_extra")
        c5, c6 = st.columns(2)
        data["header"]["linkedin"] = c5.text_input("LinkedIn URL", data["header"].get("linkedin", ""), key="h_li")
        data["header"]["github"] = c6.text_input("GitHub URL", data["header"].get("github", ""), key="h_gh")

    # ── EDUCATION ──
    if data["sectionToggles"].get("education", True):
        with st.expander("EDUCATION & SCHOLASTIC ACHIEVEMENTS"):
            for i, ed in enumerate(data["education"]):
                cols = st.columns([0.5, 3, 2.5, 1, 1, 0.5])
                ed["enabled"] = cols[0].checkbox("On", ed.get("enabled", True), key=f"ed_en_{i}", label_visibility="collapsed")
                ed["program"] = cols[1].text_input("Program", ed["program"], key=f"ed_prog_{i}", label_visibility="collapsed")
                ed["institute"] = cols[2].text_input("Institute", ed["institute"], key=f"ed_inst_{i}", label_visibility="collapsed")
                ed["score"] = cols[3].text_input("Score", ed["score"], key=f"ed_score_{i}", label_visibility="collapsed")
                ed["year"] = cols[4].text_input("Year", ed["year"], key=f"ed_year_{i}", label_visibility="collapsed")
                if cols[5].button("X", key=f"ed_del_{i}"): data["education"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Education", key="add_edu"):
                data["education"].append({"id": uid(), "program": "", "institute": "", "score": "", "year": "", "enabled": True}); save_data(data); st.rerun()
            st.markdown("---"); st.caption("Achievement Bullets")
            for i, b in enumerate(data["educationBullets"]):
                cols = st.columns([0.5, 8, 0.5])
                b["enabled"] = cols[0].checkbox("On", b.get("enabled", True), key=f"eb_en_{i}", label_visibility="collapsed")
                b["text"] = cols[1].text_input("Bullet", b["text"], key=f"eb_txt_{i}", label_visibility="collapsed")
                if cols[2].button("X", key=f"eb_del_{i}"): data["educationBullets"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Achievement Bullet", key="add_eb"):
                data["educationBullets"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()

    # ── PUBLICATIONS ──
    if data["sectionToggles"].get("publications", True):
        with st.expander("PUBLICATIONS & CONFERENCES"):
            st.caption("Journal Publications")
            for i, p in enumerate(data["publications"]):
                cols = st.columns([0.5, 8, 0.5])
                p["enabled"] = cols[0].checkbox("On", p.get("enabled", True), key=f"pub_en_{i}", label_visibility="collapsed")
                p["text"] = cols[1].text_area("Text", p["text"], key=f"pub_txt_{i}", height=60, label_visibility="collapsed")
                if cols[2].button("X", key=f"pub_del_{i}"): data["publications"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Publication", key="add_pub"):
                data["publications"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()
            st.markdown("---"); st.caption("Conferences")
            for i, c in enumerate(data["conferences"]):
                cols = st.columns([0.5, 3, 5, 0.5])
                c["enabled"] = cols[0].checkbox("On", c.get("enabled", True), key=f"conf_en_{i}", label_visibility="collapsed")
                c["name"] = cols[1].text_input("Name", c["name"], key=f"conf_name_{i}", label_visibility="collapsed")
                c["description"] = cols[2].text_input("Description", c["description"], key=f"conf_desc_{i}", label_visibility="collapsed")
                if cols[3].button("X", key=f"conf_del_{i}"): data["conferences"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Conference", key="add_conf"):
                data["conferences"].append({"id": uid(), "name": "", "description": "", "enabled": True}); save_data(data); st.rerun()

    # ── Generic section editor ──
    def edit_section(sk, sl, fields, has_hl=False):
        items = data.get(sk, [])
        if not data["sectionToggles"].get(sk, True): return
        with st.expander(sl.upper()):
            for i, item in enumerate(items):
                lbl = " - ".join(filter(None, [item.get(fields[0][0], ""), item.get(fields[1][0], "") if len(fields) > 1 else ""]))
                st.markdown(f"**{lbl or 'New Item'}**")
                tc = st.columns([0.5] + [1]*len(fields) + [0.5])
                item["enabled"] = tc[0].checkbox("On", item.get("enabled", True), key=f"{sk}_en_{i}", label_visibility="collapsed")
                for fi, (fk, fl) in enumerate(fields):
                    item[fk] = tc[fi+1].text_input(fl, item.get(fk, ""), key=f"{sk}_{fk}_{i}")
                if tc[-1].button("X", key=f"{sk}_del_{i}"): items.pop(i); save_data(data); st.rerun()
                if has_hl:
                    item["headline"] = st.text_input("Headline (bold italic)", item.get("headline", ""), key=f"{sk}_hl_{i}")
                st.caption("Bullets (use **bold** for emphasis)")
                for bi, b in enumerate(item.get("bullets", [])):
                    bc = st.columns([9, 0.5])
                    b["text"] = bc[0].text_input(f"Bullet {bi+1}", b["text"], key=f"{sk}_b_{i}_{bi}", label_visibility="collapsed")
                    if bc[1].button("X", key=f"{sk}_bd_{i}_{bi}"): item["bullets"].pop(bi); save_data(data); st.rerun()
                if st.button("+ Add Bullet", key=f"{sk}_ab_{i}"):
                    item.setdefault("bullets", []).append({"id": uid(), "text": ""}); save_data(data); st.rerun()
                st.divider()
            ni = {"id": uid(), "enabled": True, "bullets": [{"id": uid(), "text": ""}]}
            for fk, _ in fields: ni[fk] = ""
            if has_hl: ni["headline"] = ""
            if st.button(f"+ Add {sl}", key=f"add_{sk}"): items.append(ni); save_data(data); st.rerun()

    edit_section("experience", "Professional Experience", [("company", "Company"), ("role", "Role"), ("period", "Period")], has_hl=True)
    edit_section("projects", "Project", [("title", "Title"), ("tech", "Tech / Context"), ("period", "Period")], has_hl=True)
    edit_section("positions", "Position of Responsibility", [("title", "Title"), ("role", "Role"), ("period", "Period")])
    edit_section("extracurriculars", "Extra-Curricular", [("title", "Title"), ("subtitle", "Subtitle")])

    if data["sectionToggles"].get("miscellaneous", True):
        with st.expander("MISCELLANEOUS"):
            for i, m in enumerate(data["miscellaneous"]):
                cols = st.columns([0.5, 8, 0.5])
                m["enabled"] = cols[0].checkbox("On", m.get("enabled", True), key=f"misc_en_{i}", label_visibility="collapsed")
                m["text"] = cols[1].text_input("Item", m["text"], key=f"misc_txt_{i}", label_visibility="collapsed")
                if cols[2].button("X", key=f"misc_del_{i}"): data["miscellaneous"].pop(i); save_data(data); st.rerun()
            if st.button("+ Add Misc Item", key="add_misc"):
                data["miscellaneous"].append({"id": uid(), "text": "", "enabled": True}); save_data(data); st.rerun()

    if data["sectionToggles"].get("skills", True):
        with st.expander("SOFTWARE SKILLS"):
            data["skills"]["advanced"] = st.text_input("Advanced", data["skills"].get("advanced", ""), key="sk_adv")
            data["skills"]["intermediate"] = st.text_input("Intermediate", data["skills"].get("intermediate", ""), key="sk_int")
            data["skills"]["basic"] = st.text_input("Basic", data["skills"].get("basic", ""), key="sk_bas")

    save_data(data)

if __name__ == "__main__":
    main()
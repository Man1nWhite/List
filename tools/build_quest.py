# -*- coding: utf-8 -*-
"""Опросный лист v2.2 — сборка документа.

Правки v2.2: сокращено вступление, убраны подсказки «можно выбрать несколько»,
удалён вопрос 2.1 (переформулирован 2.2), уточнены 3.1–3.7, возвращена исходная
таблица сигналов, убраны 5.1–5.2 и «дополнительные резервы», справка по резервам
и подписной блок, пункты 7.1 — построчно.
"""
import os, re, copy
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement, parse_xml
from docx.table import _Cell

SRC = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "Опросный_лист.docx")
if not os.path.exists(SRC):
    SRC = [os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), x)
           for x in os.listdir(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
           if x.endswith('.docx')][0]

NAVY = RGBColor(0x1F, 0x38, 0x64)
TEXT = RGBColor(0x1A, 0x1A, 0x1A)
SOFT = RGBColor(0x3A, 0x3A, 0x3A)
GREY = RGBColor(0x6B, 0x7A, 0x8C)
HDR_FILL = "1F3864"
BAND_FILL = "F4F7FB"
BOX_FILL = "EDF2F9"
SUB_FILL = "E4EBF5"
SUB_FILL2 = "F0F4FA"
LINE = "C9D5E4"

PAGE_W = 17.5
FONT = "Times New Roman"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"

_sdt_id = [4700]


def _nid():
    _sdt_id[0] += 1
    return _sdt_id[0]


def el(tag, **attrs):
    e = OxmlElement(tag)
    for k, v in attrs.items():
        e.set(qn(k), v)
    return e


def set_run_font(run, size=None, bold=None, color=None, italic=None, spacing=None):
    run.font.name = FONT
    rPr = run._element.get_or_add_rPr()
    rf = rPr.find(qn('w:rFonts'))
    if rf is None:
        rf = el('w:rFonts'); rPr.insert(0, rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), FONT)
    if size is not None:
        run.font.size = Pt(size)
        rPr.append(el('w:szCs', **{'w:val': str(int(size * 2))}))
    if bold is not None:
        run.font.bold = bold
    if italic is not None:
        run.font.italic = italic
    if color is not None:
        run.font.color.rgb = color
    if spacing is not None:
        rPr.append(el('w:spacing', **{'w:val': str(spacing)}))
    return run


def para(container, text="", size=10.5, bold=False, color=TEXT, italic=False,
         align=None, before=0, after=0, left=0.0, first=0.0, line=1.0,
         keep_next=False, tab_to=None):
    p = None
    if isinstance(container, _Cell) and len(container.paragraphs) == 1:
        cand = container.paragraphs[0]
        if (not cand.text and not cand._p.findall(qn('w:r'))
                and not cand._p.findall(qn('w:sdt')) and not getattr(cand, "_used", False)):
            p = cand
    if p is None:
        p = container.add_paragraph()
    p._used = True
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.line_spacing = line
    if left:
        pf.left_indent = Cm(left)
    if first:
        pf.first_line_indent = Cm(first)
    if align is not None:
        p.alignment = align
    if keep_next:
        pf.keep_with_next = True
    if tab_to:
        pf.tab_stops.add_tab_stop(Cm(tab_to), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.LINES)
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold, color=color, italic=italic)
    return p


def shade(cell, hexcolor):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:shd')):
        tcPr.remove(old)
    tcPr.append(el('w:shd', **{'w:val': 'clear', 'w:color': 'auto', 'w:fill': hexcolor}))


def cell_width(cell, cm):
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = el('w:tcW'); tcPr.insert(0, tcW)
    tcW.set(qn('w:w'), str(int(cm * 567)))
    tcW.set(qn('w:type'), 'dxa')


def table_setup(table, widths, indent=0.0, borders=True, cellmar=(0.22, 0.14, 0.18, 0.14)):
    tbl = table._tbl
    tblPr = tbl.tblPr
    for tag in ('w:tblW', 'w:tblLayout', 'w:tblBorders', 'w:tblCellMar', 'w:tblInd', 'w:tblLook'):
        for old in tblPr.findall(qn(tag)):
            tblPr.remove(old)
    total = sum(widths)
    tblPr.append(el('w:tblW', **{'w:w': str(int(total * 567)), 'w:type': 'dxa'}))
    if indent:
        tblPr.append(el('w:tblInd', **{'w:w': str(int(indent * 567)), 'w:type': 'dxa'}))
    tblPr.append(el('w:tblLayout', **{'w:type': 'fixed'}))
    if borders:
        b = el('w:tblBorders')
        for side in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b.append(el('w:' + side, **{'w:val': 'single', 'w:sz': '4', 'w:space': '0', 'w:color': LINE}))
        tblPr.append(b)
    l, t, r, btm = cellmar
    cm_ = el('w:tblCellMar')
    for side, v in (('top', t), ('left', l), ('bottom', btm), ('right', r)):
        cm_.append(el('w:' + side, **{'w:w': str(int(v * 567)), 'w:type': 'dxa'}))
    tblPr.append(cm_)
    tblPr.append(el('w:tblLook', **{'w:val': '04A0', 'w:firstRow': '1', 'w:lastRow': '0',
                                    'w:firstColumn': '0', 'w:lastColumn': '0',
                                    'w:noHBand': '0', 'w:noVBand': '1'}))
    grid = tbl.find(qn('w:tblGrid'))
    if grid is not None:
        tbl.remove(grid)
    grid = el('w:tblGrid')
    for w in widths:
        grid.append(el('w:gridCol', **{'w:w': str(int(w * 567))}))
    tbl.insert(list(tbl).index(tblPr) + 1, grid)
    for row in table.rows:
        for i, c in enumerate(row.cells):
            if i < len(widths):
                cell_width(c, widths[i])


def row_height(row, cm, exact=False):
    trPr = row._tr.get_or_add_trPr()
    trPr.append(el('w:trHeight', **{'w:val': str(int(cm * 567)),
                                    'w:hRule': 'exact' if exact else 'atLeast'}))


def no_split(row):
    row._tr.get_or_add_trPr().append(el('w:cantSplit'))


def header_row(row):
    row._tr.get_or_add_trPr().append(el('w:tblHeader'))


def vcenter(cell):
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def clear_cell(cell):
    for p in list(cell.paragraphs)[1:]:
        p._element.getparent().remove(p._element)
    p0 = cell.paragraphs[0]
    for r in list(p0.runs):
        r._element.getparent().remove(r._element)
    for s in p0._p.findall(qn('w:sdt')):
        p0._p.remove(s)
    p0.paragraph_format.space_before = Pt(0)
    p0.paragraph_format.space_after = Pt(0)
    p0.paragraph_format.line_spacing = 1.0
    return p0


def add_field(p, instr):
    for kind in ('begin', 'instr', 'separate', 'text', 'end'):
        r = p.add_run()
        set_run_font(r, size=8, color=GREY)
        if kind in ('begin', 'separate', 'end'):
            r._r.append(el('w:fldChar', **{'w:fldCharType': kind}))
        elif kind == 'instr':
            it = el('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
            r._r.append(it)
        else:
            r.text = "1"


def p_border_bottom(p, color=LINE, sz=6, space=2):
    pPr = p._p.get_or_add_pPr()
    pbdr = el('w:pBdr')
    pbdr.append(el('w:bottom', **{'w:val': 'single', 'w:sz': str(sz), 'w:space': str(space), 'w:color': color}))
    pPr.append(pbdr)


def spacer(doc, pt=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(pt)
    p.paragraph_format.line_spacing = Pt(pt)
    set_run_font(p.add_run(""), size=1)
    return p


# ---------------- элементы формы ----------------
def _rpr_xml(run_el):
    rPr = run_el.find(qn('w:rPr')) if run_el is not None else None
    if rPr is None:
        return ""
    from lxml import etree
    return etree.tostring(rPr, encoding='unicode')


def sdt_checkbox(run_el, char="☐", title="Выбор варианта"):
    rpr = _rpr_xml(run_el)
    rpr = re.sub(r'<w:rFonts[^>]*/>', '', rpr) if rpr else ""
    rpr = (rpr or f'<w:rPr xmlns:w="{W_NS}"></w:rPr>').replace(
        '</w:rPr>',
        '<w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic" w:cs="MS Gothic" '
        'w:eastAsia="MS Gothic"/></w:rPr>')
    xml = (
        f'<w:sdt xmlns:w="{W_NS}" xmlns:w14="{W14_NS}">'
        '<w:sdtPr>'
        f'{rpr}'
        f'<w:alias w:val="{title}"/>'
        '<w:tag w:val="checkbox"/>'
        f'<w:id w:val="{_nid()}"/>'
        '<w14:checkbox>'
        '<w14:checked w14:val="0"/>'
        '<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        '<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        '</w14:checkbox>'
        '</w:sdtPr>'
        '<w:sdtContent>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{char}</w:t></w:r>'
        '</w:sdtContent>'
        '</w:sdt>'
    )
    return parse_xml(xml)


def sdt_text(run_el, placeholder, title="Поле для заполнения"):
    rpr = _rpr_xml(run_el)
    rpr = rpr.replace('</w:rPr>', '<w:u w:val="single"/></w:rPr>') if rpr else (
        f'<w:rPr xmlns:w="{W_NS}"><w:u w:val="single"/></w:rPr>')
    ph = placeholder.replace('_', '\u00A0\u00A0')
    ph = (ph.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))
    xml = (
        f'<w:sdt xmlns:w="{W_NS}">'
        '<w:sdtPr>'
        f'{rpr}'
        f'<w:alias w:val="{title}"/>'
        '<w:tag w:val="text"/>'
        f'<w:id w:val="{_nid()}"/>'
        '<w:showingPlcHdr/>'
        '<w:text w:multiLine="0"/>'
        '</w:sdtPr>'
        '<w:sdtContent>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{ph}</w:t></w:r>'
        '</w:sdtContent>'
        '</w:sdt>'
    )
    return parse_xml(xml)



def split_multi_text_runs(root):
    """python-docx превращает \n в <w:br/>, из-за чего в одном <w:r> оказывается
    несколько <w:t>. Делим такие run'ы на отдельные, чтобы при замене на элементы
    формы не терять части текста."""
    for r in list(root.iter(qn('w:r'))):
        kids = [k for k in r if k.tag in (qn('w:t'), qn('w:br'), qn('w:tab'))]
        if len(kids) <= 1:
            continue
        parent = r.getparent()
        if parent is None:
            continue
        idx = list(parent).index(r)
        rPr = r.find(qn('w:rPr'))
        new_runs = []
        for k in kids:
            nr = OxmlElement('w:r')
            if rPr is not None:
                nr.append(copy.deepcopy(rPr))
            nr.append(copy.deepcopy(k))
            new_runs.append(nr)
        parent.remove(r)
        for i, nr in enumerate(new_runs):
            parent.insert(idx + i, nr)


SPLIT_RE = re.compile(r'(☐|_{4,})')


def fieldify_paragraphs(root):
    n_chk = n_txt = 0
    for p in list(root.iter(qn('w:p'))):
        for child in list(p):
            if child.tag != qn('w:r'):
                continue
            t = child.find(qn('w:t'))
            if t is None or not t.text:
                continue
            parts = SPLIT_RE.split(t.text)
            if len(parts) < 2:
                continue
            try:
                idx = list(p).index(child)
            except ValueError:
                continue
            new_nodes = []
            for part in parts:
                if not part:
                    continue
                if part == '☐':
                    new_nodes.append(sdt_checkbox(child))
                    n_chk += 1
                elif set(part) == {'_'} and len(part) >= 4:
                    new_nodes.append(sdt_text(child, part))
                    n_txt += 1
                else:
                    r = copy.deepcopy(child)
                    for tt in r.findall(qn('w:t')):
                        r.remove(tt)
                    nt = el('w:t')
                    nt.set(qn('xml:space'), 'preserve')
                    nt.text = part
                    r.append(nt)
                    new_nodes.append(r)
            p.remove(child)
            for i, node in enumerate(new_nodes):
                p.insert(idx + i, node)
    return n_chk, n_txt


# ---------------- нормализация OOXML ----------------
ORDERS = {
    'pPr': ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
            'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
            'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
            'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
            'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
            'sectPr', 'pPrChange'],
    'rPr': ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline',
            'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing',
            'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText',
            'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'],
    'tblPr': ['tblStyle', 'tblpPr', 'tblOverlap', 'bidiVisual', 'tblStyleRowBandSize', 'tblStyleColBandSize',
              'tblW', 'jc', 'tblCellSpacing', 'tblInd', 'tblBorders', 'shd', 'tblLayout', 'tblCellMar',
              'tblLook', 'tblCaption', 'tblDescription'],
    'trPr': ['cnfStyle', 'divId', 'gridBefore', 'gridAfter', 'wBefore', 'wAfter', 'cantSplit', 'trHeight',
             'tblHeader', 'tblCellSpacing', 'jc', 'hidden', 'ins', 'del', 'trPrChange'],
    'tcPr': ['cnfStyle', 'tcW', 'gridSpan', 'hMerge', 'vMerge', 'tcBorders', 'shd', 'noWrap', 'tcMar',
             'textDirection', 'tcFitText', 'vAlign', 'hideMark', 'headers', 'cellIns', 'cellDel', 'cellMerge',
             'tcPrChange'],
}
W = '{' + W_NS + '}'


def normalize(root):
    for tag, order in ORDERS.items():
        idx = {W + t: i for i, t in enumerate(order)}
        for props in root.iter(W + tag):
            kids = list(props)
            seen = {}
            for k in kids:
                if k.tag in idx:
                    seen.setdefault(k.tag, []).append(k)
            for k, lst in seen.items():
                for dup in lst[:-1]:
                    props.remove(dup)
            kids = [k for k in props]
            kids.sort(key=lambda e: idx.get(e.tag, 999))
            for k in kids:
                props.remove(k)
            for k in kids:
                props.append(k)


# ================= документ =================
doc = Document(SRC)
body = doc.element.body
for child in list(body):
    if not child.tag.endswith('}sectPr'):
        body.remove(child)

normal = doc.styles['Normal']
normal.font.name = FONT
normal.font.size = Pt(10.5)
normal.font.color.rgb = TEXT
rpr = normal.element.get_or_add_rPr()
rf = rpr.find(qn('w:rFonts'))
if rf is None:
    rf = el('w:rFonts'); rpr.insert(0, rf)
for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
    rf.set(qn(a), FONT)
normal.paragraph_format.space_before = Pt(0)
normal.paragraph_format.space_after = Pt(0)
normal.paragraph_format.line_spacing = 1.0

h1 = doc.styles['Heading 1']
h1.font.name = FONT
h1.font.size = Pt(13)
h1.font.bold = True
h1.font.color.rgb = NAVY
h1.paragraph_format.space_before = Pt(18)
h1.paragraph_format.space_after = Pt(8)
h1.paragraph_format.keep_with_next = True
h1.paragraph_format.line_spacing = 1.0
h1rpr = h1.element.get_or_add_rPr()
h1rf = h1rpr.find(qn('w:rFonts'))
if h1rf is None:
    h1rf = el('w:rFonts'); h1rpr.insert(0, h1rf)
for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
    h1rf.set(qn(a), FONT)

sec = doc.sections[0]
sec.page_width = Cm(21.0)
sec.page_height = Cm(29.7)
sec.left_margin = Cm(2.0)
sec.right_margin = Cm(1.5)
sec.top_margin = Cm(1.8)
sec.bottom_margin = Cm(1.8)
sec.header_distance = Cm(1.0)
sec.footer_distance = Cm(0.9)


def heading(doc, text):
    p = doc.add_heading(level=1)
    set_run_font(p.add_run(text), size=13, bold=True, color=NAVY)
    p_border_bottom(p, color="9DB3CE", sz=6, space=3)
    return p


UND_W = 0.19


def field_run(p, width_cm, size=10.5):
    n = max(8, min(80, int(width_cm / UND_W)))
    set_run_font(p.add_run("_" * n), size=size, color=TEXT)


def q_table(doc, rows, widths=(1.25, PAGE_W - 1.25), indent=0.0):
    t = doc.add_table(rows=len(rows) + 1, cols=2)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_setup(t, list(widths), indent=indent, cellmar=(0.22, 0.14, 0.18, 0.14))
    inner = widths[1] - 0.40

    hdr = t.rows[0]
    header_row(hdr); no_split(hdr); row_height(hdr, 0.62)
    for i, txt in enumerate(("№", "Вопрос")):
        c = hdr.cells[i]
        shade(c, HDR_FILL); vcenter(c)
        p = clear_cell(c)
        if i == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(txt), size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    for ri, (num, blocks) in enumerate(rows, start=1):
        row = t.rows[ri]
        no_split(row)
        tnum, tbody = row.cells[0], row.cells[1]
        vcenter(tnum)
        pn = clear_cell(tnum)
        pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if num:
            set_run_font(pn.add_run(num), size=9.5, bold=True, color=NAVY)
        if ri % 2 == 0:
            shade(tnum, BAND_FILL); shade(tbody, BAND_FILL)

        first = True
        for blk in blocks:
            kind = blk[0]
            text = blk[1] if len(blk) > 1 else ""
            if first:
                p = clear_cell(tbody)
                first = False
            else:
                p = para(tbody, size=10, after=1.5)
            if kind == 'q':
                p.paragraph_format.space_after = Pt(2.5)
                set_run_font(p.add_run(text), size=10.5, color=TEXT)
            elif kind == 'sub':
                p.paragraph_format.space_after = Pt(1.5)
                set_run_font(p.add_run(text), size=10, color=GREY, italic=True)
            elif kind == 'opt':
                p.paragraph_format.left_indent = Cm(0.42)
                p.paragraph_format.space_after = Pt(1.5)
                set_run_font(p.add_run(text), size=10, color=SOFT)
            elif kind == 'note':
                p.paragraph_format.space_after = Pt(2)
                set_run_font(p.add_run(text), size=9.5, color=GREY, italic=True)
            elif kind == 'fill':
                p.paragraph_format.space_after = Pt(3)
                used = 0.0
                if text:
                    set_run_font(p.add_run(text + " "), size=10.5)
                    used = len(text) * UND_W + 0.3
                field_run(p, inner - used)
    return t


def grid_table(doc, header, rows, widths, indent=1.25, header_size=8.8, body_size=9.5,
               row_cm=0.62, band=True, center_from=1):
    t = doc.add_table(rows=len(rows) + 1, cols=len(widths))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_setup(t, list(widths), indent=indent, cellmar=(0.14, 0.10, 0.10, 0.10))
    h = t.rows[0]
    header_row(h); no_split(h); row_height(h, 0.6)
    for i, txt in enumerate(header):
        c = h.cells[i]
        shade(c, HDR_FILL); vcenter(c)
        p = clear_cell(c)
        p.paragraph_format.line_spacing = 1.0
        if i >= center_from:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(txt), size=header_size, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    for ri, values in enumerate(rows, start=1):
        row = t.rows[ri]
        no_split(row); row_height(row, row_cm)
        for i, txt in enumerate(values):
            c = row.cells[i]
            vcenter(c)
            if band and ri % 2 == 0:
                shade(c, BAND_FILL)
            p = clear_cell(c)
            p.paragraph_format.line_spacing = 1.0
            if i >= center_from:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if txt:
                set_run_font(p.add_run(txt), size=body_size, color=TEXT)
    return t


def matrix_header_cell(cell, text, fill, color, size=7.5, bold=False, center=True):
    shade(cell, fill)
    vcenter(cell)
    p = clear_cell(cell)
    p.paragraph_format.line_spacing = 0.95
    if center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if text:
        set_run_font(p.add_run(text), size=size, bold=bold, color=color)


def signals_table(doc):
    """Таблица сигналов: 9 шкафов × 9 типов, пятистрочная шапка."""
    widths = [2.4, 1.3, 2.2, 1.6, 1.6, 1.6, 1.6, 1.6, 1.6, 1.9]
    n_cab = 9
    t = doc.add_table(rows=5 + n_cab, cols=10)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_setup(t, widths, indent=0.0, cellmar=(0.08, 0.05, 0.08, 0.05))

    corner = t.cell(0, 0)
    corner.merge(t.cell(4, 0))

    r0 = ['AI', 'AI', 'AI', 'AO', 'DI', 'DI', 'DO', 'DO', 'FI']
    r1 = ['TC', 'RTD', '0–10 В /\n4–20 мА', '4–20 мА /\n0–10 В', 'сухой конт.', 'NAMUR',
          'сухой конт.', 'сухой конт.', 'частотно-импульсный']
    r2 = ['☐ Exi', 'Exi', '☐ Exi', '☐ Exi', '☐ Exi', '', '', '☐ Exi', 'Exi']
    r3 = ['', '', '☐ HART', '☐ HART', '', '', '', '', '']
    r4 = ['2-х пров.сх.', '☐ 3-х пров.сх.\n☐ 4-х пров.сх.', '2-х пров.сх.', '2-х пров.сх.',
          '24 V DC', '24 V DC', '220 V AC', '24 V DC', '2-х пров.сх.']
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)

    for ri, vals in enumerate((r0, r1, r2, r3, r4)):
        row = t.rows[ri]
        no_split(row); header_row(row); row_height(row, 0.34)
        for ci, txt in enumerate(vals, start=1):
            if ri == 0:
                matrix_header_cell(row.cells[ci], txt, HDR_FILL, WHITE, size=8.5, bold=True)
            else:
                fill = SUB_FILL if ri in (1, 2) else SUB_FILL2
                matrix_header_cell(row.cells[ci], txt, fill, TEXT, size=7.5)
    matrix_header_cell(corner, "Шкаф", HDR_FILL, WHITE, size=8.5, bold=True)

    for i in range(n_cab):
        row = t.rows[5 + i]
        no_split(row); row_height(row, 0.62)
        cell = row.cells[0]
        shade(cell, SUB_FILL2); vcenter(cell)
        p = clear_cell(cell)
        set_run_font(p.add_run("Шкаф %d" % (i + 1)), size=8.5, color=TEXT)
        for ci in range(1, 10):
            vcenter(row.cells[ci])
    return t


def interfaces_table(doc):
    """Таблица цифровых интерфейсов: 9 шкафов × 5 колонок."""
    widths = [1.7, 2.6, 2.7, 3.3, 2.7, 4.5]
    header = ['', 'Ethernet', 'Profibus', 'Foundation Fieldbus', 'Modbus RTU (RS-485)',
              'Другое: ______________']
    t = doc.add_table(rows=10, cols=6)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_setup(t, widths, indent=0.0, cellmar=(0.10, 0.06, 0.10, 0.06))
    h = t.rows[0]
    header_row(h); no_split(h); row_height(h, 0.6)
    WHITE = RGBColor(0xFF, 0xFF, 0xFF)
    for i, txt in enumerate(header):
        c = h.cells[i]
        if i == 0:
            shade(c, HDR_FILL); vcenter(c)
            clear_cell(c)
            set_run_font(c.paragraphs[0].add_run("Шкаф"), size=8.5, bold=True, color=WHITE)
        else:
            matrix_header_cell(c, txt, HDR_FILL, WHITE, size=8, bold=True)
    for i in range(9):
        row = t.rows[1 + i]
        no_split(row); row_height(row, 0.62)
        cell = row.cells[0]
        shade(cell, SUB_FILL2); vcenter(cell)
        p = clear_cell(cell)
        set_run_font(p.add_run("Шкаф %d" % (i + 1)), size=8.5, color=TEXT)
        for ci in range(1, 6):
            vcenter(row.cells[ci])
    return t


# ---------- шапка ----------
p = para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, after=2)
set_run_font(p.add_run("ОПРОСНЫЙ ЛИСТ"), size=19, bold=True, color=NAVY, spacing=30)
p = para(doc, align=WD_ALIGN_PARAGRAPH.CENTER, after=10)
set_run_font(p.add_run("на модернизацию систем автоматизации и диспетчеризации"),
             size=12.5, color=RGBColor(0x44, 0x54, 0x6A))

t = doc.add_table(rows=1, cols=1)
t.alignment = WD_TABLE_ALIGNMENT.LEFT
table_setup(t, [PAGE_W], borders=False, cellmar=(0.35, 0.25, 0.35, 0.25))
tblPr = t._tbl.tblPr
b = el('w:tblBorders')
b.append(el('w:top', **{'w:val': 'single', 'w:sz': '8', 'w:space': '0', 'w:color': '1F3864'}))
b.append(el('w:bottom', **{'w:val': 'single', 'w:sz': '2', 'w:space': '0', 'w:color': 'DCE4F0'}))
b.append(el('w:left', **{'w:val': 'single', 'w:sz': '2', 'w:space': '0', 'w:color': 'DCE4F0'}))
b.append(el('w:right', **{'w:val': 'single', 'w:sz': '2', 'w:space': '0', 'w:color': 'DCE4F0'}))
b.append(el('w:insideH', **{'w:val': 'nil'}))
b.append(el('w:insideV', **{'w:val': 'nil'}))
tblPr.append(b)
c = t.rows[0].cells[0]
shade(c, BOX_FILL)
clear_cell(c)
p = para(c, size=10.5, after=3, line=1.05)
set_run_font(p.add_run("Уважаемые коллеги!"), size=10.5, bold=True, color=NAVY)
p = para(c, size=9.5, color=SOFT, after=0, line=1.06)
set_run_font(p.add_run("Заполнение занимает около 10 минут. Анкета предназначена для предварительной оценки "
                       "объёма работ; точный состав оборудования определяется после уточнения исходных данных."),
             size=9.5, color=SOFT)

# ---------- 1 ----------
heading(doc, "1. Общие сведения")
q_table(doc, [
    ("1.1", [("fill", "Предприятие:")]),
    ("1.2", [("fill", "Контактное лицо (ФИО, должность):")]),
    ("1.3", [("fill", "Телефон / e-mail:")]),
    ("1.4", [("q", "Объект автоматизации — участок / оборудование:"),
             ("opt", "☐  Основное производство        ☐  Вспомогательные системы"),
             ("opt", "☐  Склад и отгрузка        ☐  Другое: ______________________"),
             ("fill", "Наименование участка / оборудования:")]),
    ("1.5", [("q", "Тип задачи:"),
             ("opt", "☐  Новая система"),
             ("opt", "☐  Замена устаревшей системы"),
             ("opt", "☐  Расширение существующей системы"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("1.6", [("q", "Планируемые сроки:"),
             ("opt", "☐  До 6 месяцев"),
             ("opt", "☐  6–12 месяцев"),
             ("opt", "☐  Более года"),
             ("opt", "☐  Не определены")]),
    ("1.7", [("q", "Какие исходные данные есть в наличии?"),
             ("opt", "☐  Перечень сигналов"),
             ("opt", "☐  Электрические схемы"),
             ("opt", "☐  Программа существующего контроллера"),
             ("opt", "☐  Техническое задание / стандарты предприятия"),
             ("opt", "☐  Фотографии шкафов и оборудования"),
             ("opt", "☐  Затрудняюсь ответить")]),
])

# ---------- 2 ----------
heading(doc, "2. Существующая система и логика управления")
q_table(doc, [
    ("2.1", [("q", "Что используется сейчас (марки контроллеров, панелей оператора, "
                    "названия программ, если известны):"),
             ("fill", ""), ("fill", ""), ("fill", "")]),
    ("2.2", [("q", "Существующую логику работы необходимо:"),
             ("opt", "☐  Сохранить без изменений"),
             ("opt", "☐  Перенести с доработкой"),
             ("opt", "☐  Разработать заново"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("2.3", [("q", "Сложность технологического процесса:"),
             ("opt", "☐  Простая — пуск/стоп, блокировки, сигнализация"),
             ("opt", "☐  Средняя — регулирование, аналоговые сигналы, счётчики"),
             ("opt", "☐  Высокая — каскадное регулирование, позиционирование, расчётные модели"),
             ("opt", "☐  Затрудняюсь ответить")]),
])

# ---------- 3 ----------
heading(doc, "3. Управление и обмен данными")
q_table(doc, [
    ("3.1", [("q", "Где оператор должен управлять оборудованием и видеть его состояние?"),
             ("opt", "☐  На местном пульте с кнопками и индикацией"),
             ("opt", "☐  На компьютере оператора / диспетчера (АРМ)"),
             ("opt", "☐  Решение необходимо предложить поставщику")]),
    ("3.2", [("q", "Количество панелей оператора и рабочих мест:"),
             ("fill", "Панелей оператора, шт.:"),
             ("fill", "Рабочих мест (АРМ), шт.:"),
             ("fill", "Размеры экранов (если известны):")]),
    ("3.3", [("q", "Нужен ли обмен данными с другими системами?"),
             ("opt", "☐  Да        ☐  Нет        ☐  Затрудняюсь ответить")]),
    ("3.4", [("q", "С какими системами и какие данные нужно передавать:"),
             ("fill", ""), ("fill", "")]),
    ("3.5", [("q", "Требуемый протокол связи с верхним уровнем:"),
             ("opt", "☐  Modbus TCP        ☐  OPC UA"),
             ("opt", "☐  Modbus RTU (RS-485)        ☐  Другое: ______________________"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("3.6", [("q", "Какая линия связи уже проложена в месте установки шкафа?"),
             ("opt", "☐  Витая пара"),
             ("opt", "☐  Оптика"),
             ("opt", "☐  Линии связи нет, требуется прокладка"),
             ("opt", "☐  Другое: ______________________"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("3.7", [("q", "Где должен храниться архив данных и журнал событий?"),
             ("opt", "☐  Отдельный сервер архива"),
             ("opt", "☐  Достаточно хранения на АРМ / панели оператора"),
             ("opt", "☐  Архивирование не требуется"),
             ("opt", "☐  Затрудняюсь ответить")]),
])

# ---------- 4 ----------
heading(doc, "4. Сигналы ввода/вывода")
q_table(doc, [
    ("4.1", [("q", "Количество шкафов управления:"),
             ("fill", "Количество, шт.:")]),
    ("4.2", [("q", "Количество сигналов"),
             ("sub", "Укажите фактическое количество без учёта резерва — требования к резерву задаются в разделе 5. "
                     "Допускаются ориентировочные значения.")]),
])
spacer(doc, 4)
signals_table(doc)
spacer(doc, 4)
q_table(doc, [
    ("4.3", [("q", "Цифровые интерфейсы по шкафам"),
             ("sub", "Укажите количество подключаемых устройств по каждому интерфейсу.")]),
])
spacer(doc, 4)
interfaces_table(doc)

# ---------- 5 ----------
heading(doc, "5. Резервы для расширения")


def reserve_table(doc):
    widths = [1.25, 8.13, 8.13]
    rows = [
        ("5.1", "Резервное место в шкафу, % полезной монтажной площади",
         "☐ 10 %   ☐ 20 %   ☐ 30 %   ☐ другое: _____ %\n☐ предложить и согласовать"),
        ("5.2", "Свободные каналы ввода/вывода, % от установленного количества каналов каждого типа",
         "☐ 10 %   ☐ 20 %   ☐ 30 %   ☐ другое: _____ %\n☐ предложить и согласовать"),
        ("5.3", "Ограничения по габаритам шкафа:",
         "☐  Нет ограничений\n☐  Есть, место ограничено: Ш × В × Г не более ____________ мм\n"
         "☐  Затрудняюсь ответить"),
    ]
    t = doc.add_table(rows=len(rows) + 1, cols=3)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    table_setup(t, widths, indent=0.0, cellmar=(0.16, 0.12, 0.16, 0.12))
    hdr = t.rows[0]
    header_row(hdr); no_split(hdr); row_height(hdr, 0.62)
    for i, txt in enumerate(("№", "Вопрос", "Требование")):
        c = hdr.cells[i]
        shade(c, HDR_FILL); vcenter(c)
        p = clear_cell(c)
        if i == 0:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(p.add_run(txt), size=9.5, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))
    for ri, (num, question, answer) in enumerate(rows, start=1):
        row = t.rows[ri]
        no_split(row)
        band = ri % 2 == 0
        for ci, cell in enumerate(row.cells):
            vcenter(cell)
            if band:
                shade(cell, BAND_FILL)
        pn = clear_cell(row.cells[0])
        pn.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_run_font(pn.add_run(num), size=9.5, bold=True, color=NAVY)
        pq = clear_cell(row.cells[1])
        set_run_font(pq.add_run(question), size=10, color=TEXT)
        pa = clear_cell(row.cells[2])
        for k, line in enumerate(answer.split("\n")):
            pp = pa if k == 0 else para(row.cells[2], size=9.5, after=1.5)
            set_run_font(pp.add_run(line), size=9.5, color=SOFT)
    # 5.3 — ответ занимает обе правые колонки
    t.cell(3, 1).merge(t.cell(3, 2))
    return t


reserve_table(doc)

# ---------- 6 ----------
heading(doc, "6. Питание и условия эксплуатации")
q_table(doc, [
    ("6.1", [("q", "Вводное напряжение питания шкафа:"),
             ("opt", "☐  380 В AC (3 фазы)        ☐  220 В AC (1 фаза)"),
             ("opt", "☐  24 В DC        ☐  Другое: __________________"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("6.2", [("q", "Бесперебойное питание (ИБП):"),
             ("opt", "☐  Не требуется"),
             ("opt", "☐  ИБП внутри шкафа управления"),
             ("opt", "☐  Отдельный шкаф ИБП"),
             ("opt", "☐  Используется существующий ИБП предприятия"),
             ("opt", "Время работы от ИБП:  ☐ 10 мин   ☐ 30 мин   ☐ другое: ________   ☐ Затрудняюсь ответить")]),
    ("6.3", [("q", "Место установки шкафа:"),
             ("opt", "☐  Отапливаемое помещение (электропомещение, операторская)"),
             ("opt", "☐  Неотапливаемое помещение"),
             ("opt", "☐  Улица / под навесом"),
             ("opt", "☐  Другое: ______________________")]),
    ("6.4", [("q", "Особые условия:"),
             ("opt", "☐  Пыль        ☐  Влага / мойка оборудования        ☐  Высокая температура"),
             ("opt", "☐  Агрессивная среда        ☐  Взрывоопасная зона"),
             ("opt", "☐  Особых условий нет        ☐  Затрудняюсь ответить")]),
    ("6.5", [("q", "Температура окружающей среды в месте установки:"),
             ("opt", "от ______________ до ______________ °C"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("6.6", [("q", "Требуемая степень защиты:"),
             ("opt", "☐  IP20        ☐  IP31        ☐  IP54        ☐  IP65        ☐  Другое: __________"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("6.7", [("q", "Требуется ли внутреннее освещение шкафа?"),
             ("opt", "☐  Да        ☐  Нет        ☐  Затрудняюсь ответить")]),
])

# ---------- 7 ----------
heading(doc, "7. Объём работ")
q_table(doc, [
    ("7.1", [("q", "Какие работы нужны от поставщика?"),
             ("opt", "☐  Обследование объекта"),
             ("opt", "☐  Проектирование"),
             ("opt", "☐  Изготовление и поставка шкафов"),
             ("opt", "☐  Разработка программного обеспечения"),
             ("opt", "☐  Заводские приёмочные испытания (FAT)"),
             ("opt", "☐  Монтаж оборудования и кабелей"),
             ("opt", "☐  Шефмонтаж"),
             ("opt", "☐  Пусконаладочные работы"),
             ("opt", "☐  Обучение персонала"),
             ("opt", "☐  Метрологическая поверка"),
             ("opt", "☐  Комплексное выполнение работ «под ключ»"),
             ("opt", "☐  Другое: ______________________")]),
    ("7.2", [("q", "Прокладка кабеля:"),
             ("opt", "☐  Силами заказчика"),
             ("opt", "☐  Силами поставщика, лотки и трубы уже проложены"),
             ("opt", "☐  Силами поставщика, включая проектирование и монтаж трасс"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("7.3", [("q", "Возможность остановки оборудования для проведения работ:"),
             ("opt", "☐  Остановка допускается"),
             ("opt", "☐  Только в согласованное технологическое окно"),
             ("opt", "☐  Остановка не допускается"),
             ("opt", "☐  Затрудняюсь ответить")]),
    ("7.4", [("q", "Дополнительные ограничения:"),
             ("sub", "Режим работы, пропускной режим, требования к подрядчикам и персоналу."),
             ("fill", ""), ("fill", "")]),
])

# ---------- 8 ----------
heading(doc, "8. Дополнительно")
q_table(doc, [
    ("8.1", [("q", "Основные проблемы текущей системы, пожелания, особые требования:"),
             ("fill", ""), ("fill", ""), ("fill", "")]),
])

# ---------- колонтитулы ----------
def build_footer(part):
    root = part._element
    for child in list(root):
        if child.tag != qn('w:sectPr'):
            root.remove(child)
    try:
        fp = part.paragraphs[0]
    except IndexError:
        fp = part.add_paragraph()
    pf = fp.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.tab_stops.add_tab_stop(Cm(PAGE_W), WD_TAB_ALIGNMENT.RIGHT)
    pPr = fp._p.get_or_add_pPr()
    pbdr = el('w:pBdr')
    pbdr.append(el('w:top', **{'w:val': 'single', 'w:sz': '4', 'w:space': '4', 'w:color': 'C9D5E4'}))
    pPr.append(pbdr)
    set_run_font(fp.add_run("Опросный лист на модернизацию систем автоматизации и диспетчеризации · версия 2.3"),
                 size=8, color=GREY)
    fp.add_run("\t")
    set_run_font(fp.add_run("стр. "), size=8, color=GREY)
    add_field(fp, " PAGE ")
    set_run_font(fp.add_run(" из "), size=8, color=GREY)
    add_field(fp, " NUMPAGES ")


build_footer(sec.footer)
build_footer(sec.even_page_footer)
sec.different_first_page_header_footer = False

# ---------- элементы формы ----------
split_multi_text_runs(body)
n_chk, n_txt = fieldify_paragraphs(body)

normalize(doc.element.body)
for st in doc.styles:
    normalize(st.element)
for part in (sec.footer, sec.even_page_footer):
    normalize(part._element)

cp = doc.core_properties
cp.title = "Опросный лист на модернизацию систем автоматизации и диспетчеризации"
cp.subject = "Опросный лист, версия 2.3 (заполняемая форма Word)"
cp.comments = ("Редакция 2.3: сокращённое вступление, уточнённые вопросы по управлению и обмену данными, "
               "исходная таблица сигналов, резервы без дополнительных опций.")
cp.revision = 4

out = SRC
doc.save(out)
print("saved", os.path.basename(out), "| флажков:", n_chk, "| полей:", n_txt)

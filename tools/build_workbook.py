#!/usr/bin/env python3
"""Przebudowuje UI skoroszytu PW_integrationg.xlsm na poziomie XML (zachowuje makra, przyciski i pole wyszukiwania).

Uzycie:  python3 tools/build_workbook.py <wejsciowy.xlsm> <wyjsciowy.xlsm>
Wejscie = ORYGINALNY plik (arkusze Sheet_1, Sheet_2, Sheet_3, PW_Diagnostics, Log).
"""
import re, sys, zipfile, math
from xml.sax.saxutils import escape

SRC, DST = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(SRC)
files = {n: zin.read(n) for n in zin.namelist()}
T = lambda n: files[n].decode('utf-8')

NAVY, BLUE, LBLUE, YEL, RED, GREY, GREEN = '1F3864', '2F5597', 'DDEBF7', 'FFEB9C', 'FFC7CE', 'F2F2F2', 'E2EFDA'

# ------------------------------------------------------------------ styles
st = T('xl/styles.xml')
def cnt(tag): return int(re.search(r'<%s count="(\d+)"' % tag, st).group(1))
base_font, base_fill, base_border, base_xf = cnt('fonts'), cnt('fills'), cnt('borders'), cnt('cellXfs')
new_fonts, new_fills, new_borders, new_xfs = [], [], [], []

def font(b=False, i=False, u=False, sz=11, color='FF262626', name='Aptos Narrow', strike=False):
    x = '<font>' + ('<b/>' if b else '') + ('<i/>' if i else '') + ('<u/>' if u else '') + '<sz val="%s"/><color rgb="%s"/><name val="%s"/><family val="2"/>' % (sz, color, name)
    x += '<scheme val="minor"/></font>' if name == 'Aptos Narrow' else '</font>'
    new_fonts.append(x); return base_font + len(new_fonts) - 1
def fill(rgb):
    new_fills.append('<fill><patternFill patternType="solid"><fgColor rgb="FF%s"/><bgColor indexed="64"/></patternFill></fill>' % rgb)
    return base_fill + len(new_fills) - 1
def border(kind):
    c = '<color rgb="FFBFBFBF"/>'
    if kind == 'box': x = '<border><left style="thin">%s</left><right style="thin">%s</right><top style="thin">%s</top><bottom style="thin">%s</bottom><diagonal/></border>' % (c, c, c, c)
    else: x = '<border><left/><right/><top/><bottom style="medium"><color rgb="FF%s"/></bottom><diagonal/></border>' % NAVY
    new_borders.append(x); return base_border + len(new_borders) - 1
def xf(f=0, fl=0, b=0, num=0, h=None, v=None, wrap=False, indent=0):
    al = ''
    if h or v or wrap or indent:
        al = '<alignment%s%s%s%s/>' % (' horizontal="%s"' % h if h else '', ' vertical="%s"' % v if v else '', ' wrapText="1"' if wrap else '', ' indent="%d"' % indent if indent else '')
    x = '<xf numFmtId="%d" fontId="%d" fillId="%d" borderId="%d" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">%s</xf>' % (num, f, fl, b, al)
    new_xfs.append(x); return base_xf + len(new_xfs) - 1

fl_navy, fl_blue, fl_lblue, fl_yel, fl_red, fl_grey, fl_green, fl_white = map(fill, (NAVY, BLUE, LBLUE, YEL, RED, GREY, GREEN, 'FFFFFF'))
b_box, b_line = border('box'), border('line')
f_title = font(b=True, sz=20, color='FFFFFFFF'); f_sub = font(sz=11, color='FFD9E1F2')
f_hdr = font(b=True, color='FFFFFFFF'); f_h2 = font(b=True, sz=13, color='FF' + NAVY)
f_txt = font(); f_bold = font(b=True); f_note = font(i=True, sz=10, color='FF7F7F7F')
f_mono = font(sz=10, color='FF' + NAVY, name='Consolas'); f_link = font(u=True, color='FF0563C1')
f_ok = font(b=True, color='FF375623'); f_bad = font(b=True, color='FFC00000')

S = dict(
  title=xf(f_title, fl_navy, 0, h='left', v='center', indent=1),
  sub=xf(f_sub, fl_navy, 0, h='left', v='center', indent=1),
  band=xf(f_txt, fl_navy, 0),
  hdr=xf(f_hdr, fl_navy, b_box, h='left', v='center', wrap=True, indent=1),
  hdrc=xf(f_hdr, fl_navy, b_box, h='center', v='center', wrap=True),
  h2=xf(f_h2, 0, b_line, h='left', v='center'),
  text=xf(f_txt, 0, 0, h='left', v='top', wrap=True),
  note=xf(f_note, 0, 0, h='left', v='center', wrap=True),
  cell=xf(f_txt, 0, b_box, h='left', v='top', wrap=True, indent=1),
  cellb=xf(f_bold, fl_grey, b_box, h='left', v='top', wrap=True, indent=1),
  cellc=xf(f_txt, 0, b_box, h='center', v='top', wrap=True),
  mono=xf(f_mono, fl_grey, b_box, h='left', v='top', wrap=True, indent=1),
  link=xf(f_link, fl_grey, b_box, h='left', v='top', wrap=True, indent=1),
  ok=xf(f_ok, fl_green, b_box, h='left', v='top', wrap=True, indent=1),
  bad=xf(f_bad, fl_red, b_box, h='left', v='top', wrap=True, indent=1),
  in_c=xf(f_txt, 0, b_box, h='center', v='center'),
  in_l=xf(f_txt, 0, b_box, h='left', v='center', wrap=True, indent=1),
  in_mono=xf(f_mono, 0, b_box, h='left', v='center', wrap=False, indent=1),
  lab=xf(f_bold, fl_grey, b_box, h='left', v='center', indent=1),
  sw_blue=xf(f_bold, fl_lblue, b_box, h='center', v='center'),
  sw_yel=xf(f_txt, fl_yel, b_box, h='center', v='center'),
  sw_red=xf(f_txt, fl_red, b_box, h='center', v='center'),
  sw_green=xf(f_txt, fl_green, b_box, h='center', v='center'),
  sw_white=xf(f_txt, 0, b_box, h='center', v='center'),
  legend=xf(f_bold, 0, 0, h='right', v='center'),
  search=xf(f_txt, 0, b_box, h='left', v='center', indent=1),
  status=xf(font(b=True, color='FF0066CC'), 0, 0, h='left', v='center'),
  date=xf(f_txt, 0, 0, num=165, h='left', v='top'),
  logtxt=xf(f_txt, 0, 0, h='left', v='top'),
)
numfmt = '<numFmts count="1"><numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/></numFmts>'
st = re.sub(r'(<styleSheet[^>]*>)', r'\1' + numfmt, st, count=1)
def append(tag, items):
    global st
    n = cnt(tag)
    st = re.sub(r'<%s count="\d+"' % tag, '<%s count="%d"' % (tag, n + len(items)), st, count=1)
    st = st.replace('</%s>' % tag, ''.join(items) + '</%s>' % tag, 1)
append('fonts', new_fonts); append('fills', new_fills); append('borders', new_borders); append('cellXfs', new_xfs)
dxfs = ('<dxfs count="3">'
        '<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>'
        '<dxf><font><strike/><color rgb="FF808080"/></font></dxf>'
        '<dxf><fill><patternFill><bgColor rgb="FFFFEB9C"/></patternFill></fill></dxf></dxfs>')
st = st.replace('<dxfs count="0"/>', dxfs)
files['xl/styles.xml'] = st.encode('utf-8')

# ------------------------------------------------------------------ helpers
def col_letter(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s
def lines(text, width):
    return sum(max(1, math.ceil(len(p) / max(1, width * 1.05))) for p in str(text).split('\n'))

class Sheet:
    def __init__(self): self.rows = {}; self.ht = {}; self.merges = []; self.links = []
    def put(self, col, row, val, style, ht=None):
        ref = '%s%d' % (col_letter(col), row)
        s = S[style] if isinstance(style, str) else style
        if val is None or val == '': x = '<c r="%s" s="%d"/>' % (ref, s)
        elif isinstance(val, (int, float)): x = '<c r="%s" s="%d"><v>%s</v></c>' % (ref, s, val)
        else: x = '<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (ref, s, escape(str(val)))
        self.rows.setdefault(row, {})[col] = x
        if ht: self.ht[row] = max(self.ht.get(row, 0), ht)
    def fillrow(self, row, c1, c2, style):
        for c in range(c1, c2 + 1):
            self.rows.setdefault(row, {}).setdefault(c, '<c r="%s%d" s="%d"/>' % (col_letter(c), row, S[style]))
    def merge(self, c1, r1, c2, r2): self.merges.append('%s%d:%s%d' % (col_letter(c1), r1, col_letter(c2), r2))
    def link(self, col, row, sheet): self.links.append('<hyperlink ref="%s%d" location="\'%s\'!A1" display="%s"/>' % (col_letter(col), row, sheet, sheet))
    def data(self):
        out = []
        for r in sorted(set(self.rows) | set(self.ht)):
            cells = ''.join(self.rows.get(r, {})[c] for c in sorted(self.rows.get(r, {})))
            h = ' ht="%s" customHeight="1"' % self.ht[r] if r in self.ht else ''
            out.append('<row r="%d"%s>%s</row>' % (r, h, cells))
        return '<sheetData>' + ''.join(out) + '</sheetData>'
    def tail(self):
        x = ''
        if self.merges: x += '<mergeCells count="%d">%s</mergeCells>' % (len(self.merges), ''.join('<mergeCell ref="%s"/>' % m for m in self.merges))
        return x

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
def cols_xml(widths, styles=None):
    return '<cols>' + ''.join('<col min="%d" max="%d" width="%s" customWidth="1"%s/>' % (i + 1, i + 1, w, (' style="%d"' % styles[i]) if styles and styles.get(i) else '') for i, w in enumerate(widths)) + '</cols>'
def banner(sh, ncols, title, subtitle, h1=34, h2=22):
    for c in range(1, ncols + 1): sh.put(c, 1, None, 'band'); sh.put(c, 2, None, 'band')
    sh.put(1, 1, title, 'title', ht=h1); sh.put(1, 2, subtitle, 'sub', ht=h2)

def sheet_xml(sh, widths, tab, view='', cf='', dv='', extra_head='', gridlines=False, codename=None, tail_extra='', col_styles=None, landscape=True):
    cn = ' codeName="%s"' % codename if codename else ''
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet %s><sheetPr%s><tabColor rgb="FF%s"/><pageSetUpPr fitToPage="1"/></sheetPr>' % (NS, cn, tab) +
            '<sheetViews><sheetView %sworkbookViewId="0">%s</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>' % ('' if gridlines else 'showGridLines="0" ', view) +
            cols_xml(widths, col_styles) + sh.data() + sh.tail() + cf + dv +
            ('<hyperlinks>%s</hyperlinks>' % ''.join(sh.links) if sh.links else '') +
            '<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="%s" fitToHeight="0"/>' % ('landscape' if landscape else 'portrait') + tail_extra + '</worksheet>')

# ------------------------------------------------------------------ Folders (old Sheet_2)
sst = re.findall(r'<si><t[^>]*>(.*?)</t></si>', T('xl/sharedStrings.xml'), re.S)
unesc = lambda s: s.replace('&amp;', '&')
L = {i: unesc(v) for i, v in enumerate(sst)}
links = [('pw:\\\\pa7PJWW2K01.egis.racine.local:ENQE01301_SZC_Sizewell_EPR\\Documents\\SZC_Sizewell_EPR\\01-ICOSH\\04-INTERNAL DOCUMENTS\\04.1-DOCUMENT\\', 'SZC Sizewell EPR – 01-ICOSH / 04.1-DOCUMENT')]
sample_folder = L[12]

f = Sheet()
banner(f, 3, 'Folders to scan', 'ProjectWise folders the macro downloads files from (including all subfolders).')
for c in range(1, 4): f.put(c, 3, None, 'text')
f.put(1, 3, 'How to fill in: paste a ProjectWise folder link into column B (one folder = one row). Column A: YES = scan, NO = skip without deleting the row. '
            'Column C is a free note. AUTOMATIC: after you paste a link the macro discovers the attributes (sheet Attributes) and downloads the data.', 'text', ht=36); f.merge(1, 3, 3, 3)
f.put(1, 4, 'Link pattern:', 'lab', ht=22)
f.put(2, 4, 'pw:\\\\<server>:<Datasource>\\Documents\\<Project>\\<Folder>\\<Subfolder>\\', 'mono'); f.put(3, 4, 'Details and examples: sheet Guide', 'note')
f.put(1, 5, 'Active', 'hdrc', ht=24); f.put(2, 5, 'ProjectWise folder link', 'hdr'); f.put(3, 5, 'Description (optional)', 'hdr')
for i in range(100):
    r = 6 + i
    lk, ds = links[i] if i < len(links) else (None, None)
    f.put(1, r, 'YES' if lk else None, 'in_c', ht=20); f.put(2, r, lk, 'in_mono'); f.put(3, r, ds, 'in_l')
f_cf = ('<conditionalFormatting sqref="B6:B105"><cfRule type="expression" dxfId="0" priority="1"><formula>AND(B6&lt;&gt;"",LOWER(LEFT(TRIM(B6),3))&lt;&gt;"pw:")</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="A6:C105"><cfRule type="expression" dxfId="1" priority="2"><formula>UPPER(TRIM($A6))="NO"</formula></cfRule></conditionalFormatting>')
f_dv = ('<dataValidations count="2">'
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Active" error="Choose YES or NO." promptTitle="Active" prompt="YES = folder is scanned, NO = skipped." sqref="A6:A105"><formula1>"YES,NO"</formula1></dataValidation>'
        '<dataValidation type="custom" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="This does not look like a PW link" error="The folder link should start with pw:\\\\ (copy it from ProjectWise Explorer). Save anyway?" promptTitle="PW folder link" prompt="Format: pw:\\\\serwer:Datasource\\Documents\\Projekt\\Folder\\" sqref="B6:B105"><formula1>LOWER(LEFT(TRIM(B6),3))="pw:"</formula1></dataValidation>'
        '</dataValidations>')
files['xl/worksheets/sheet2.xml'] = sheet_xml(f, [15, 118, 50], '2F5597', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B6" sqref="B6"/>', f_cf, f_dv, codename='Sheet2').encode('utf-8')
files['xl/worksheets/_rels/sheet2.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

# ------------------------------------------------------------------ Attributes (old Sheet_3)
a = Sheet()
banner(a, 5, 'Attributes to download', 'Filled in automatically after you paste a link in sheet Folders. Attributes marked YES in column B become columns in PW_Data.')
for c in range(1, 5): a.put(c, 3, None, 'text')
a.put(1, 3, 'The list fills itself (new attributes are appended, nothing is deleted). Change YES/NO to enable or disable a column; you can add your own attributes at the end.', 'note', ht=34); a.merge(1, 3, 3, 3)
a.put(1, 4, None, 'text', ht=6)
for i, h in enumerate(['Attribute name', 'Download', 'Type', 'Example value', 'Notes (for you)']): a.put(i + 1, 5, h, 'hdr' if i != 1 else 'hdrc', ht=24)
for i in range(300):
    r = 6 + i
    a.put(1, r, None, 'in_l', ht=20); a.put(2, r, None, 'in_c'); a.put(3, r, None, 'in_c'); a.put(4, r, None, 'in_l'); a.put(5, r, None, 'in_l')
a_cf = ('<conditionalFormatting sqref="A6:E305"><cfRule type="expression" dxfId="1" priority="1"><formula>UPPER(TRIM($B6))="NO"</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="A6:A305"><cfRule type="duplicateValues" dxfId="2" priority="2"/></conditionalFormatting>')
a_dv = ('<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Download" error="Choose YES or NO." sqref="B6:B305"><formula1>"YES,NO"</formula1></dataValidation></dataValidations>')
files['xl/worksheets/sheet3.xml'] = sheet_xml(a, [40, 11, 12, 55, 45], '2F5597', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/>', a_cf, a_dv, codename='Sheet3', tail_extra='<drawing r:id="rId1"/>').encode('utf-8')

# ------------------------------------------------------------------ Diagnostics (old PW_Diagnostics)
d = Sheet()
d.put(1, 1, 'Attribute name (copy to column A in sheet Attributes)', 'hdr', ht=24); d.put(2, 1, 'Example value (from 1 file)', 'hdr')
d.put(1, 2, 'No data - run DISCOVER ATTRIBUTES on sheet Attributes.', 'note', ht=20)
files['xl/worksheets/sheet4.xml'] = sheet_xml(d, [58, 110], '7F7F7F', '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', gridlines=True, codename='Sheet4').encode('utf-8')

# ------------------------------------------------------------------ Log
lg = Sheet()
for i, h in enumerate(['Date and Time', 'Action', 'File Name', 'Changed Attribute', 'OLD Value', 'NEW Value', 'User']):
    lg.put(i + 1, 1, h, 'hdr', ht=24)
files['xl/worksheets/sheet5.xml'] = sheet_xml(lg, [21, 12, 60, 28, 42, 42, 30], 'A6A6A6', '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', gridlines=True, codename='Sheet6', col_styles={0: S['date'], 1: S['logtxt'], 2: S['logtxt'], 3: S['logtxt'], 4: S['logtxt'], 5: S['logtxt'], 6: S['logtxt']}).encode('utf-8')

# ------------------------------------------------------------------ PW_Data (old Sheet_1)
s1 = T('xl/worksheets/sheet1.xml')
root = re.match(r'(<\?xml[^>]*\?>\s*<worksheet[^>]*>)', s1).group(1)
tail = s1[s1.index('<pageMargins'):]
b = Sheet()
for r in (1, 2): b.fillrow(r, 1, 11, 'band')
b.put(1, 1, 'ProjectWise file database', 'title', ht=28); b.put(1, 2, 'File list from the selected PW folders with attributes. Changes since the previous update are colour-coded.', 'sub', ht=18)
b.ht[3] = 17; b.ht[4] = 17
b.put(1, 5, 'Status column  →', 'legend', ht=18); b.put(2, 5, 'NEW = added', 'sw_green'); b.put(3, 5, 'CHANGED = value changed (see Change Info)', 'sw_yel'); b.put(4, 5, 'DELETED = removed from PW', 'sw_red')
b.put(1, 6, None, 'search', ht=26)
b.put(1, 7, '↑ This sheet is the database (one row per file). Use the filter arrows in row 11; search and browse on sheet Dashboard.', 'note', ht=15)
b.put(1, 8, 'Last Database Update: - (no update has been run yet)', 'status', ht=22)
b.ht[9] = 4; b.ht[10] = 4
for i, h in enumerate(['File Name', 'PW Source Path', 'PW Link', 'Updater', 'Updated']): b.put(i + 1, 11, h, 'hdr', ht=24)
for c in range(6, 11): b.put(c, 11, None, 'hdr')
oldcols = re.search(r'<cols>(.*?)</cols>', s1).group(1)
keep = ''.join(re.findall(r'<col min="(?:[6-9]|\d{2,})"[^>]*/>', oldcols))
newcols = '<cols><col min="1" max="1" width="62" customWidth="1"/><col min="2" max="2" width="55" customWidth="1"/><col min="3" max="3" width="46" customWidth="1"/><col min="4" max="4" width="30" customWidth="1"/><col min="5" max="5" width="13" customWidth="1"/>' + keep + '</cols>'
sheetpr = '<sheetPr codeName="Sheet1"><tabColor rgb="FF1F7A3D"/></sheetPr>'
views = '<sheetViews><sheetView zoomScale="90" zoomScaleNormal="90" workbookViewId="0"><pane ySplit="11" topLeftCell="A12" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A12" sqref="A12"/></sheetView></sheetViews>'
# nowe pozycje przyciskow i pola wyszukiwania
tail = tail.replace('<xdr:col>0</xdr:col><xdr:colOff>790575</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>47625</xdr:rowOff>', '<xdr:col>0</xdr:col><xdr:colOff>47625</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>19050</xdr:rowOff>')
tail = tail.replace('<xdr:col>0</xdr:col><xdr:colOff>7477125</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>628650</xdr:rowOff>', '<xdr:col>0</xdr:col><xdr:colOff>3900000</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>311150</xdr:rowOff>')
files['xl/worksheets/sheet1.xml'] = (root + sheetpr + '<dimension ref="A1:K11"/>' + views + '<sheetFormatPr defaultRowHeight="15"/>' + newcols + b.data() +
                                     '<mergeCells count="2"><mergeCell ref="A2:E2"/><mergeCell ref="A7:E7"/></mergeCells>' + tail).encode('utf-8')
# drawing1: przyciski + pole tekstowe
d1 = T('xl/drawings/drawing1.xml')
tb = d1[d1.index('<mc:AlternateContent'):d1.index('</mc:AlternateContent>') + len('</mc:AlternateContent>')]
tb = tb.replace('<xdr:col>0</xdr:col><xdr:colOff>790575</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>47625</xdr:rowOff>', '<xdr:col>0</xdr:col><xdr:colOff>47625</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>19050</xdr:rowOff>')
tb = tb.replace('<xdr:col>0</xdr:col><xdr:colOff>7477125</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>628650</xdr:rowOff>', '<xdr:col>0</xdr:col><xdr:colOff>3900000</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>311150</xdr:rowOff>')
def button(i, name, macro, text, c1, c2, r1, r2, color, edit='oneCell'):
    return ('<xdr:twoCellAnchor editAs="%s"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:to>'
            '<xdr:sp macro="[0]!%s" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="pl-PL" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
            % (edit, c1[0], c1[1], r1[0], r1[1], c2[0], c2[1], r2[0], r2[1], macro, i, name, color, escape(text)))
hdr_d = d1[:d1.index('<xdr:twoCellAnchor')]
new_d1 = (hdr_d + tb + '</xdr:wsDr>')   # PW_Data: tylko pole wyszukiwania, bez przyciskow
files['xl/drawings/drawing1.xml'] = new_d1.encode('utf-8')
# drawing2 (Attributes): jeden przycisk w C3
d2 = T('xl/drawings/drawing2.xml')
new_d2 = d2[:d2.index('<xdr:twoCellAnchor')] + button(2, 'btnProbe', 'ZbadajAttributesPW', 'DISCOVER ATTRIBUTES ▶', (3, 95250), (3, 2300000), (2, 30000), (2, 400000), '2F5597') + button(3, 'btnBench', 'TestSzybkosciPW', 'PW SPEED TEST', (4, 95250), (4, 2300000), (2, 30000), (2, 400000), '7F7F7F') + button(4, 'btnCopyOut', 'TestPobieraniaPW', 'TEST POBIERANIA PLIKU', (5, 95250), (5, 2300000), (2, 30000), (2, 400000), '7F7F7F') + '</xdr:wsDr>'
files['xl/drawings/drawing2.xml'] = new_d2.encode('utf-8')

# ------------------------------------------------------------------ Views (definicje) + Widok (wynik)
w = Sheet()
banner(w, 3, 'Views - table templates', 'A view = which PW_Data columns appear and in what order. You pick it on sheet Dashboard (VIEW drop-down).')
for c in range(1, 4): w.put(c, 3, None, 'text')
w.put(1, 3, 'In column B list the columns in display order, separated by commas: header names (e.g. File Name, Version) or PW_Data column letters (e.g. A, B, F, L, TK - spaces work too: A B F L TK). Asterisk * = all columns.', 'note', ht=48); w.merge(1, 3, 3, 3)
w.put(1, 4, None, 'text', ht=6)
w.put(1, 5, 'View name', 'hdr', ht=24); w.put(2, 5, 'Columns (order = order in the view)', 'hdr'); w.put(3, 5, 'Description', 'hdr')
vs = [('All', '*', 'All PW_Data columns in database order.'),
      ('Revisions', 'File Name, Version, VersionSequence, WorkflowState, Updated', 'File name and revision/version with workflow state.'),
      ('Links', 'File Name, PW Link, PW Source Path', 'File name and clickable links.'),
      ('Letters (example)', 'A, B, F', 'The same columns given as letters: A = File Name, B = PW Source Path, F = first attribute.')]
for i in range(20):
    r = 6 + i; e = vs[i] if i < len(vs) else (None, None, None)
    w.put(1, r, e[0], 'in_l', ht=22); w.put(2, r, e[1], 'in_l'); w.put(3, r, e[2], 'in_l')
files['xl/worksheets/sheet7.xml'] = sheet_xml(w, [30, 70, 80], '7030A0', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/>').encode('utf-8')

v = Sheet()
banner(v, 12, 'PW Tracker - Dashboard', 'Main working sheet: PW_Data shown according to the selected view. It is a copy - the macro downloads the data, you define views on sheet Views.')
v.put(1, 3, 'View:', 'lab', ht=24); v.put(2, 3, 'All', 'in_l'); v.put(7, 3, 'Download to:', 'lab'); v.put(8, 3, None, 'in_l')
v.put(1, 4, 'Search:', 'lab', ht=24); v.put(2, 4, None, 'search'); v.put(3, 4, '← type a phrase (min. 2 characters) - the list filters instantly on any column', 'note')
v.put(1, 5, None, 'text', ht=38)
v.put(1, 6, 'Colour legend  →', 'legend', ht=20); v.put(2, 6, 'Blue row = folder', 'sw_blue'); v.put(3, 6, 'Yellow cell = change', 'sw_yel'); v.put(4, 6, 'Red row = deleted from PW', 'sw_red')
v.put(1, 7, 'Last Database Update: - (no update has been run yet)', 'status', ht=22)
v.put(1, 8, 'No data - paste a link in sheet Folders or click UPDATE DATA.', 'hdr', ht=24)
for cc in range(2, 7): v.put(cc, 8, None, 'hdr')
v.merge(1, 8, 5, 8)
v_dv = '<dataValidations count="1"><dataValidation type="list" allowBlank="0" showErrorMessage="1" errorTitle="View" error="Pick a view from the list (defined on sheet Views)." sqref="B3"><formula1>Views!$A$6:$A$25</formula1></dataValidation></dataValidations>'
files['xl/worksheets/sheet8.xml'] = sheet_xml(v, [30, 36, 28, 40, 18, 18, 18, 18, 18, 18, 18, 18], '7030A0', '<pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B3" sqref="B3"/>', dv=v_dv, gridlines=True, tail_extra='<drawing r:id="rId1"/>').replace('<sheetView ', '<sheetView tabSelected="1" ', 1).encode('utf-8')
files['xl/worksheets/_rels/sheet8.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing3.xml"/></Relationships>'
files['xl/drawings/drawing3.xml'] = (d2[:d2.index('<xdr:twoCellAnchor')] +
    button(2, 'btnUpdateV', 'DownloadFromProjectWise_PWPS', 'UPDATE DATA', (0, 47625), (0, 1950000), (4, 38100), (4, 450000), '1F7A3D', 'absolute') +
    button(3, 'btnNavV', 'ShowNavigator', 'FOLDER NAVIGATOR', (1, 47625), (1, 2250000), (4, 38100), (4, 450000), '2F5597', 'absolute') +
    button(4, 'btnGetFiles', 'PobierzPlikiPW', 'DOWNLOAD FROM PW', (2, 47625), (2, 1950000), (4, 38100), (4, 450000), '7F4F00', 'absolute') + '</xdr:wsDr>').encode('utf-8')

# ------------------------------------------------------------------ Guide (nowy arkusz)
I = Sheet()
W = [2, 26, 66, 66]
banner(I, 4, 'PW → Excel  ·  User guide', 'This workbook downloads the file list and attributes from ProjectWise and shows what changed since the last update.', h1=40, h2=24)
row = [4]
def nxt(n=1): row[0] += n
def h2(t): nxt(); I.put(2, row[0], t, 'h2', ht=26); I.put(3, row[0], None, 'h2'); I.put(4, row[0], None, 'h2'); nxt()
def para(t, style='text', cols=(2, 4)):
    I.put(cols[0], row[0], t, style, ht=max(18, lines(t, sum(W[cols[0] - 1:cols[1]])) * 15 + 4)); I.merge(cols[0], row[0], cols[1], row[0]); nxt()
def tr(c1, c2, c3, st=('cellb', 'cell', 'cell'), link=None):
    h = max(lines(c1, W[1]), lines(c2, W[2]), lines(c3, W[3])) * 15 + 6
    I.put(2, row[0], c1, st[0], ht=max(21, h)); I.put(3, row[0], c2, st[1]); I.put(4, row[0], c3, st[2])
    if link: I.link(2, row[0], link)
    nxt()
def th(c1, c2, c3): I.put(2, row[0], c1, 'hdr', ht=22); I.put(3, row[0], c2, 'hdr'); I.put(4, row[0], c3, 'hdr'); nxt()

h2('1.  Quick start (the process is automatic)')
th('Step', 'What to do', 'Result')
tr('1. Folders', 'Paste a ProjectWise folder link into column B (format: see section 3). Sign in in the PowerShell window (Bentley IMS) if asked.', 'The macro discovers the attributes of the folder (sheet Diagnostics) and appends them to sheet Attributes.', link='Folders')
tr('2. Attributes', 'Optional: in column B switch YES/NO for attributes that should (not) go into the database.', 'Only attributes marked YES become columns.', link='Attributes')
tr('3. Update', 'After attribute discovery the data downloads by itself. Manually: the UPDATE DATA button on sheet Dashboard. Wait until the PowerShell window closes.', 'File list (each revision in its own row); changes highlighted.', link='Dashboard')
tr('4. Dashboard', 'Work on sheet Dashboard: pick a view with the VIEW drop-down (e.g. Revisions, Links), use SEARCH; the folder tree sits in column A (click a folder to filter). Define your own views on sheet Views.', 'A table with the chosen columns in the chosen order.', link='Dashboard')
tr('5. Log', 'Full change history (ADDED / CHANGED / DELETED). Each run is a separate row group - expand with the plus on the left.', 'Audit trail.', link='Log')

h2('2.  Sheets - what they are for and what to enter')
th('Sheet', 'Purpose', 'What to enter / notes')
tr('Folders', 'List of PW folders to scan.', 'A: YES/NO · B: folder link (pw:\\\\…) - pasting starts the automation · C: description. Data from row 6. An invalid link turns red.', link='Folders')
tr('Attributes', 'List of PW attributes (fills itself).', 'A: name · B: Download YES/NO · C: type (ATRYBUT = project attribute, SYSTEM = PW property) · D: example · E: your notes. Nothing is deleted on the next discovery.', link='Attributes')
tr('Views', 'Table templates: which columns and in what order.', 'A: view name · B: comma-separated columns (header names or letters, e.g. A, B, F, L, TK) · C: description. See section 5.', link='Views')
tr('PW_Data', 'DATABASE - full file list with attributes, grouped by folder (the source for views).', 'Treat it as a database: do not edit by hand, the macro overwrites it on every update. Work on sheet Dashboard instead.', link='PW_Data')
tr('Dashboard', 'MAIN WORKING SHEET: PW_Data shown by the selected view, with buttons, search, KPI cards and legend.', 'VIEW / SEARCH / THEME / DOWNLOAD TO controls and KPI cards at the top; the folder tree is in column A. It is a copy - refreshed when you change the view and after every data update.', link='Dashboard')
tr('Diagnostics', 'Attributes of one file from the folder discovered last.', 'Read-only; its content is cleared on every discovery.', link='Diagnostics')
tr('Log', 'Change history from all updates.', 'Read-only. Do not delete the header in row 1.', link='Log')

h2('3.  What a folder link looks like (sheet Folders, column B)')
para('Easiest: in ProjectWise Explorer copy the folder path from the address bar - it starts with pw:\\\\. The link has four parts:')
th('Part', 'Meaning', 'Example')
tr('pw:\\\\', 'Prefix - required. Without it the row is ignored.', 'pw:\\\\', st=('cellb', 'cell', 'mono'))
tr('<server>:<Datasource>', 'Server address and datasource name separated by a colon. The macro uses it to log in.', 'pwserver.firma.local:PW_Projekty', st=('cellb', 'cell', 'mono'))
tr('\\Documents\\', 'Fixed root segment. Everything after it is the folder path.', '\\Documents\\', st=('cellb', 'cell', 'mono'))
tr('<folder path>', 'Folders separated by \\. The trailing \\ is optional.', 'PROJ01\\04_DELIVERABLES\\03-HR\\', st=('cellb', 'cell', 'mono'))
I.put(2, row[0], 'Full example', 'cellb', ht=34); I.put(3, row[0], 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\04_DELIVERABLES\\03-HR\\', 'ok'); I.merge(3, row[0], 4, row[0]); I.put(4, row[0], None, 'ok'); nxt()
th('Examples', 'Link', 'Verdict')
tr('Valid', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\Deliverables\\', 'OK - full format.', st=('ok', 'mono', 'cell'))
tr('Valid (short)', 'pw://PW_Projekty/Documents/PROJ01/Deliverables', 'OK - forward slashes work too; the server can be omitted if the datasource is known in ProjectWise.', st=('ok', 'mono', 'cell'))
tr('Valid (with suffix)', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\Deliverables (Working folder)', 'OK - the " (Working folder)" suffix is removed automatically.', st=('ok', 'mono', 'cell'))
tr('Invalid', '\\\\fileserver\\share\\PROJ01\\Deliverables', 'This is a plain network path, not a PW link - the row is skipped.', st=('bad', 'mono', 'cell'))
tr('Invalid', 'https://pwserver.firma.local/PROJ01', 'A web address is not supported. Use a pw: path\\\\.', st=('bad', 'mono', 'cell'))
tr('Invalid', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\', 'No folder path after \\Documents\\ - nothing to scan.', st=('bad', 'mono', 'cell'))

h2('4.  Columns in sheet PW_Data')
th('Column', 'Meaning', 'What it looks like')
tr('File Name', 'File name in ProjectWise.', 'PROJ01-STR-0001_Plan.dwg', st=('cellb', 'cell', 'mono'))
tr('PW Source Path', 'The folder of the file, counted from \\Documents\\. The navigator tree and the folder filter are built from this column.', 'PROJ01\\04_DELIVERABLES\\03-HR', st=('cellb', 'cell', 'mono'))
tr('PW Link', 'Clickable document link - opens it in ProjectWise Explorer (must be installed). The link is also the key the macro uses to recognise the file - do not edit it.', 'pw://pwserver.firma.local:PW_Projekty/Documents/D:{3F2A1B4C-0000-0000-0000-000000000000}', st=('cellb', 'cell', 'mono'))
tr('Updater', 'Who last changed the file in PW.', 'Jan Kowalski', st=('cellb', 'cell', 'mono'))
tr('Updated', 'Date of the last change, format YYYY-MM-DD.', '2026-05-27', st=('cellb', 'cell', 'mono'))
tr('F, G, H …', 'Attributes from sheet Attributes - in the same order.', 'e.g. Version = 3', st=('cellb', 'cell', 'mono'))

h2('5.  Views - how to define a table')
para('Sheet Views: one row = one view. In column B list the PW_Data columns in the order they should appear. You can use header names or column letters (letters depend on the attribute order, names are more stable).')
th('View name', 'Columns (column B on sheet Views)', 'Effect')
tr('Revisions', 'File Name, Version, VersionSequence, WorkflowState, Updated', 'File name, revision, version number, state and date - in that order.', st=('cellb', 'mono', 'cell'))
tr('Links', 'File Name, PW Link, PW Source Path', 'Name, clickable link and folder.', st=('cellb', 'mono', 'cell'))
tr('Letters', 'A, B, F, L, TK   (or: A B F L TK)', 'PW_Data columns given as letters, in the given order.', st=('cellb', 'mono', 'cell'))
tr('All', '*', 'All columns.', st=('cellb', 'mono', 'cell'))
para('You pick the view on sheet Dashboard (VIEW drop-down). If a name/letter does not exist, the Dashboard shows it as a warning under the title.', 'note')

h2('6.  Search, folders and colours')
th('Item', 'Meaning', 'Notes')
tr('Folder tree', 'The navigator on the left lists the folders found in PW_Data (with file counts). Click a folder = the table shows the files of that folder (if it has none: everything below it).', 'Click the FOLDERS header (A8) or the page bar to show everything again.')
I.put(2, row[0], 'Yellow row', 'sw_yel', ht=36); I.put(3, row[0], 'Status CHANGED: a value changed since the last update. Column Change Info lists old -> new values.', 'cell'); I.put(4, row[0], 'Overwrites the value in PW_Data. A CHANGED entry in Log.', 'cell'); nxt()
I.put(2, row[0], 'Red row', 'sw_red', ht=36); I.put(3, row[0], 'Status DELETED: the file disappeared from ProjectWise (in a scanned folder).', 'cell'); I.put(4, row[0], 'Marks the row, does not delete it. If the file comes back the mark disappears. A DELETED entry in Log.', 'cell'); nxt()
I.put(2, row[0], 'Green row', 'sw_green', ht=24); I.put(3, row[0], 'Status NEW: the file was added in the last update.', 'cell'); I.put(4, row[0], 'An ADDED entry in Log.', 'cell'); nxt()
I.put(2, row[0], 'White row', 'sw_white', ht=24); I.put(3, row[0], 'No change.', 'cell'); I.put(4, row[0], '-', 'cell'); nxt()

h2('6b.  Search syntax (SEARCH box on the Dashboard)')
th('You type', 'Meaning', 'Example')
tr('HRX', 'contains the text, in any column', 'HRX', st=('cellb', 'cell', 'mono'))
tr('* and ?', 'wildcards: * = any text, ? = any single character', 'HRX-0??-*', st=('cellb', 'cell', 'mono'))
tr('column:value', 'search only in one column (filename, path, link, updater, updated, status or any PW_Data column name)', 'filename:*HRX*   path:Detailing   state:approved', st=('cellb', 'cell', 'mono'))
tr('"phrase"', 'quotes keep spaces together', 'filename:"my file*"', st=('cellb', 'cell', 'mono'))
tr('-term', 'exclude', '-pdf   -path:Templates', st=('cellb', 'cell', 'mono'))
tr('several terms', 'all of them must match (AND)', 'filename:*HRX* -pdf status:CHANGED', st=('cellb', 'cell', 'mono'))
para('The Dashboard shows one page of results at a time (500 rows); use PREV / NEXT. The whole database stays in sheet PW_Data - the Dashboard does not copy it.', 'note')

h2('7.  Requirements')
for t in ['Windows + Excel with macros enabled (.xlsm saved on disk, not opened from an e-mail attachment).',
          'PowerShell module PWPS_DAB (Bentley) installed and access to the ProjectWise server; sign-in via Bentley IMS.',
          'The macro creates temporary files next to the workbook: PW_Folders.txt, PW_GetFiles.ps1, PW_Output.csv, PW_Structure_Log.txt. They are overwritten on the next update.']:
    para('•  ' + t)

h2('8.  Troubleshooting')
th('Symptom', 'Cause', 'What to do')
tr('In Log: WARNING „Scope not scanned”', 'Could not log in or find the folder.', 'Check the link on sheet Folders and your PW login. Files of this folder were NOT marked as deleted.')
tr('PowerShell window: FOLDER NOT FOUND', 'Wrong path or no permission for the folder.', 'Copy the link again from ProjectWise Explorer.')
tr('An attribute column is empty', 'The attribute name does not exist in this folder or is spelled differently.', 'Check the name on sheet Diagnostics / Attributes; attributes depend on the folder and document class.')
tr('„Please save the Excel workbook…”', 'The workbook is not saved on disk.', 'Save the workbook in a normal folder and run again.')
tr('The download takes very long', 'The PW server (attributes and revisions) takes the time, not Excel.', 'Sheet Attributes → "PW SPEED TEST": compares download methods and writes timings to sheet Test_PW. Disabling revisions (GET_ALL_VERSIONS in modConfig) also saves time.')
tr('Macros do not run', 'The file is blocked after download.', 'File properties → Unblock; in Excel enable content (macros).')
I.merge(2, 3, 4, 3)
I.put(2, 3, 'Workbook version: V1.7  ·  Paste a PW folder link in sheet Folders - the macro does the rest. Settings: Attributes (YES/NO) and Views.', 'note', ht=22)
I.put(3, 3, None, 'note'); I.put(4, 3, None, 'note')
files['xl/worksheets/sheet6.xml'] = sheet_xml(I, W, '1F3864', '<selection activeCell="B6" sqref="B6"/>', landscape=False).encode('utf-8')

# ------------------------------------------------------------------ workbook-level
wb = T('xl/workbook.xml')
wb = re.sub(r'<sheets>.*?</sheets>',
            '<sheets><sheet name="Dashboard" sheetId="11" r:id="rId13"/><sheet name="Guide" sheetId="9" r:id="rId11"/><sheet name="Folders" sheetId="2" r:id="rId2"/><sheet name="Attributes" sheetId="3" r:id="rId3"/><sheet name="Views" sheetId="10" r:id="rId12"/>'
            '<sheet name="PW_Data" sheetId="8" r:id="rId1"/><sheet name="Diagnostics" sheetId="4" r:id="rId4"/><sheet name="Log" sheetId="7" r:id="rId5"/></sheets>', wb, flags=re.S)
wb = wb.replace('activeTab="3"', 'activeTab="0"')
files['xl/workbook.xml'] = wb.encode('utf-8')
rels = T('xl/_rels/workbook.xml.rels')
rels = re.sub(r'<Relationship Id="rId9"[^>]*calcChain[^>]*/>', '', rels)
rels = rels.replace('</Relationships>', '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/><Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet8.xml"/></Relationships>')
files['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')
ct = T('[Content_Types].xml')
ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
ct = ct.replace('<Override PartName="/xl/theme/theme1.xml"', '<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet8.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml"')
ct = ct.replace('<Override PartName="/xl/theme/theme1.xml"', '<Override PartName="/xl/drawings/drawing3.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/theme/theme1.xml"', 1)
files['[Content_Types].xml'] = ct.encode('utf-8')
del files['xl/calcChain.xml']
app = T('docProps/app.xml')
app = app.replace('<vt:i4>5</vt:i4>', '<vt:i4>8</vt:i4>')
app = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>', '<TitlesOfParts><vt:vector size="8" baseType="lpstr"><vt:lpstr>Dashboard</vt:lpstr><vt:lpstr>Guide</vt:lpstr><vt:lpstr>Folders</vt:lpstr><vt:lpstr>Attributes</vt:lpstr><vt:lpstr>Views</vt:lpstr><vt:lpstr>PW_Data</vt:lpstr><vt:lpstr>Diagnostics</vt:lpstr><vt:lpstr>Log</vt:lpstr></vt:vector></TitlesOfParts>', app, flags=re.S)
files['docProps/app.xml'] = app.encode('utf-8')

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
    for n, v in files.items(): z.writestr(n, v)
print('OK ->', DST)

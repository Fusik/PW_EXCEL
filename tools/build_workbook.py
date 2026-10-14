#!/usr/bin/env python3
"""Przebudowuje UI skoroszytu PW_integrationg.xlsm na poziomie XML (zachowuje makra, przyciski i pole wyszukiwania).

Uzycie:  python3 tools/build_workbook.py <wejsciowy.xlsm> <wyjsciowy.xlsm>
Wejscie = ORYGINALNY plik (arkusze Sheet_1, Sheet_2, Sheet_3, PW_Diagnostyka, Log).
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

# ------------------------------------------------------------------ Foldery (old Sheet_2)
sst = re.findall(r'<si><t[^>]*>(.*?)</t></si>', T('xl/sharedStrings.xml'), re.S)
unesc = lambda s: s.replace('&amp;', '&')
L = {i: unesc(v) for i, v in enumerate(sst)}
links = [(L[1], 'Deliverables dla JDO Site – 03-HR'), (L[0], 'Deliverables dla Edvance – 03-HR'), (L[8], 'Procedury i narzędzia (oficjalne)'),
         (L[10], 'Deliverables dla JDO Site – 02-HK'), (L[11], 'Deliverables dla JDO Site – 04-HL')]
sample_folder = L[12]

f = Sheet()
banner(f, 3, 'Foldery do skanowania', 'Lista folderów ProjectWise, z których makro pobiera pliki (razem ze wszystkimi podfolderami).')
for c in range(1, 4): f.put(c, 3, None, 'text')
f.put(1, 3, 'Jak wypełniać: wklej link do folderu z ProjectWise do kolumny B (jeden folder = jeden wiersz). Kolumna A: TAK = skanuj, NIE = pomiń bez kasowania wiersza. '
            'Kolumna C to dowolna notatka. AUTOMAT: po wklejeniu linku makro samo wykrywa atrybuty (arkusz Atrybuty) i pobiera dane.', 'text', ht=36); f.merge(1, 3, 3, 3)
f.put(1, 4, 'Wzorzec linku:', 'lab', ht=22)
f.put(2, 4, 'pw:\\\\<serwer>:<Datasource>\\Documents\\<Projekt>\\<Folder>\\<Podfolder>\\', 'mono'); f.put(3, 4, 'Szczegóły i przykłady: arkusz Instrukcja', 'note')
f.put(1, 5, 'Aktywny', 'hdrc', ht=24); f.put(2, 5, 'Link do folderu ProjectWise', 'hdr'); f.put(3, 5, 'Opis (opcjonalnie)', 'hdr')
for i in range(100):
    r = 6 + i
    lk, ds = links[i] if i < len(links) else (None, None)
    f.put(1, r, 'TAK' if lk else None, 'in_c', ht=20); f.put(2, r, lk, 'in_mono'); f.put(3, r, ds, 'in_l')
f_cf = ('<conditionalFormatting sqref="B6:B105"><cfRule type="expression" dxfId="0" priority="1"><formula>AND(B6&lt;&gt;"",LOWER(LEFT(TRIM(B6),3))&lt;&gt;"pw:")</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="A6:C105"><cfRule type="expression" dxfId="1" priority="2"><formula>UPPER(TRIM($A6))="NIE"</formula></cfRule></conditionalFormatting>')
f_dv = ('<dataValidations count="2">'
        '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="Aktywny" error="Wybierz TAK lub NIE." promptTitle="Aktywny" prompt="TAK = folder jest skanowany, NIE = pomijany." sqref="A6:A105"><formula1>"TAK,NIE"</formula1></dataValidation>'
        '<dataValidation type="custom" errorStyle="warning" allowBlank="1" showInputMessage="1" showErrorMessage="1" errorTitle="To nie wygląda na link PW" error="Link do folderu powinien zaczynać się od pw:\\\\ (skopiuj go z ProjectWise Explorer). Zapisać mimo to?" promptTitle="Link do folderu PW" prompt="Format: pw:\\\\serwer:Datasource\\Documents\\Projekt\\Folder\\" sqref="B6:B105"><formula1>LOWER(LEFT(TRIM(B6),3))="pw:"</formula1></dataValidation>'
        '</dataValidations>')
files['xl/worksheets/sheet2.xml'] = sheet_xml(f, [15, 118, 50], '2F5597', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B6" sqref="B6"/>', f_cf, f_dv, codename='Sheet2').encode('utf-8')
files['xl/worksheets/_rels/sheet2.xml.rels'] = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'

# ------------------------------------------------------------------ Atrybuty (old Sheet_3)
a = Sheet()
banner(a, 5, 'Atrybuty do pobrania', 'Wypełniane automatycznie po wklejeniu linku w arkuszu Foldery. Atrybuty z TAK w kolumnie B stają się kolumnami w Baza_PW.')
for c in range(1, 5): a.put(c, 3, None, 'text')
a.put(1, 3, 'Lista uzupełnia się sama (nowe atrybuty są dopisywane, nic nie jest kasowane). Zmień TAK/NIE, aby włączyć lub wyłączyć kolumnę; własne atrybuty możesz dopisać na końcu.', 'note', ht=34); a.merge(1, 3, 3, 3)
a.put(1, 4, None, 'text', ht=6)
for i, h in enumerate(['Nazwa atrybutu', 'Pobierać', 'Typ', 'Przykładowa wartość', 'Opis (dla Ciebie)']): a.put(i + 1, 5, h, 'hdr' if i != 1 else 'hdrc', ht=24)
for i in range(300):
    r = 6 + i
    a.put(1, r, None, 'in_l', ht=20); a.put(2, r, None, 'in_c'); a.put(3, r, None, 'in_c'); a.put(4, r, None, 'in_l'); a.put(5, r, None, 'in_l')
a_cf = ('<conditionalFormatting sqref="A6:E305"><cfRule type="expression" dxfId="1" priority="1"><formula>UPPER(TRIM($B6))="NIE"</formula></cfRule></conditionalFormatting>'
        '<conditionalFormatting sqref="A6:A305"><cfRule type="duplicateValues" dxfId="2" priority="2"/></conditionalFormatting>')
a_dv = ('<dataValidations count="1"><dataValidation type="list" allowBlank="1" showErrorMessage="1" errorTitle="Pobierać" error="Wybierz TAK lub NIE." sqref="B6:B305"><formula1>"TAK,NIE"</formula1></dataValidation></dataValidations>')
files['xl/worksheets/sheet3.xml'] = sheet_xml(a, [40, 11, 12, 55, 45], '2F5597', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/>', a_cf, a_dv, codename='Sheet3', tail_extra='<drawing r:id="rId1"/>').encode('utf-8')

# ------------------------------------------------------------------ Diagnostyka (old PW_Diagnostyka)
d = Sheet()
d.put(1, 1, 'Nazwa atrybutu (skopiuj do kolumny A w arkuszu Atrybuty)', 'hdr', ht=24); d.put(2, 1, 'Przykładowa wartość (z 1 pliku)', 'hdr')
d.put(1, 2, 'Brak danych – uruchom „Zbadaj atrybuty” na arkuszu Atrybuty.', 'note', ht=20)
files['xl/worksheets/sheet4.xml'] = sheet_xml(d, [58, 110], '7F7F7F', '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', gridlines=True, codename='Sheet4').encode('utf-8')

# ------------------------------------------------------------------ Log
lg = Sheet()
for i, h in enumerate(['Date and Time', 'Action', 'File Name', 'Changed Attribute', 'OLD Value', 'NEW Value', 'User']):
    lg.put(i + 1, 1, h, 'hdr', ht=24)
files['xl/worksheets/sheet5.xml'] = sheet_xml(lg, [21, 12, 60, 28, 42, 42, 30], 'A6A6A6', '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/>', gridlines=True, codename='Sheet6', col_styles={0: S['date'], 1: S['logtxt'], 2: S['logtxt'], 3: S['logtxt'], 4: S['logtxt'], 5: S['logtxt'], 6: S['logtxt']}).encode('utf-8')

# ------------------------------------------------------------------ Baza_PW (old Sheet_1)
s1 = T('xl/worksheets/sheet1.xml')
root = re.match(r'(<\?xml[^>]*\?>\s*<worksheet[^>]*>)', s1).group(1)
tail = s1[s1.index('<pageMargins'):]
b = Sheet()
for r in (1, 2): b.fillrow(r, 1, 11, 'band')
b.put(1, 1, 'Baza plików ProjectWise', 'title', ht=28); b.put(1, 2, 'Lista plików z wybranych folderów PW wraz z atrybutami. Zmiany względem poprzedniej aktualizacji są oznaczone kolorami.', 'sub', ht=18)
b.ht[3] = 17; b.ht[4] = 17
b.put(1, 5, 'Legenda kolorów  →', 'legend', ht=18); b.put(2, 5, 'Niebieski wiersz = folder', 'sw_blue'); b.put(3, 5, 'Żółta komórka = zmieniona wartość', 'sw_yel'); b.put(4, 5, 'Czerwony wiersz = usunięty z PW', 'sw_red')
b.put(1, 6, None, 'search', ht=26)
b.put(1, 7, '↑ Pole wyszukiwania: wpisz frazę, aby przefiltrować listę (po dowolnej kolumnie). Wyczyść pole, aby pokazać wszystko.', 'note', ht=15)
b.put(1, 8, 'Last Database Update: – (aktualizacja jeszcze nie była uruchamiana)', 'status', ht=22)
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
def button(i, name, macro, text, c1, c2, r1, r2, color):
    return ('<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:to>'
            '<xdr:sp macro="[0]!%s" textlink=""><xdr:nvSpPr><xdr:cNvPr id="%d" name="%s"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/></a:avLst></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="pl-PL" sz="1100" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>%s</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
            % (c1[0], c1[1], r1[0], r1[1], c2[0], c2[1], r2[0], r2[1], macro, i, name, color, escape(text)))
hdr_d = d1[:d1.index('<xdr:twoCellAnchor')]
new_d1 = (hdr_d + button(2, 'btnUpdate', 'DownloadFromProjectWise_PWPS', 'AKTUALIZUJ DANE', (0, 47625), (0, 1900000), (2, 19050), (3, 311150), '1F7A3D') +
          button(3, 'btnNavigator', 'ShowNavigator', 'NAWIGATOR FOLDERÓW', (0, 2000000), (0, 3900000), (2, 19050), (3, 311150), '2F5597') + tb + '</xdr:wsDr>')
files['xl/drawings/drawing1.xml'] = new_d1.encode('utf-8')
# drawing2 (Atrybuty): jeden przycisk w C3
d2 = T('xl/drawings/drawing2.xml')
new_d2 = d2[:d2.index('<xdr:twoCellAnchor')] + button(2, 'btnProbe', 'ZbadajAtrybutyPW', 'WYKRYJ ATRYBUTY ▶', (3, 95250), (3, 2300000), (2, 30000), (2, 400000), '2F5597') + button(3, 'btnBench', 'TestSzybkosciPW', 'TEST SZYBKOŚCI PW', (4, 95250), (4, 2300000), (2, 30000), (2, 400000), '7F7F7F') + '</xdr:wsDr>'
files['xl/drawings/drawing2.xml'] = new_d2.encode('utf-8')

# ------------------------------------------------------------------ Widoki (definicje) + Widok (wynik)
w = Sheet()
banner(w, 3, 'Widoki – szablony tabeli', 'Widok = które kolumny z Baza_PW i w jakiej kolejności. Wybierasz go w arkuszu Widok (komórka B3).')
for c in range(1, 4): w.put(c, 3, None, 'text')
w.put(1, 3, 'W kolumnie B wypisz kolumny w kolejności wyświetlania, oddzielone przecinkiem: nazwami nagłówków (np. File Name, Version) albo literami kolumn Baza_PW (np. A, B, F, L, TK – można też bez przecinków: A B F L TK). Gwiazdka * = wszystkie kolumny.', 'note', ht=48); w.merge(1, 3, 3, 3)
w.put(1, 4, None, 'text', ht=6)
w.put(1, 5, 'Nazwa widoku', 'hdr', ht=24); w.put(2, 5, 'Kolumny (kolejność = kolejność w widoku)', 'hdr'); w.put(3, 5, 'Opis', 'hdr')
vs = [('Wszystko', '*', 'Wszystkie kolumny Baza_PW w kolejności z bazy.'),
      ('Rewizje', 'File Name, Version, VersionSequence, WorkflowState, Updated', 'Nazwa pliku i rewizja/wersja ze stanem workflow.'),
      ('Linki', 'File Name, PW Link, PW Source Path', 'Nazwa pliku i klikalne linki.'),
      ('Litery (przykład)', 'A, B, F', 'Te same kolumny wskazane literami: A = File Name, B = PW Source Path, F = pierwszy atrybut.')]
for i in range(20):
    r = 6 + i; e = vs[i] if i < len(vs) else (None, None, None)
    w.put(1, r, e[0], 'in_l', ht=22); w.put(2, r, e[1], 'in_l'); w.put(3, r, e[2], 'in_l')
files['xl/worksheets/sheet7.xml'] = sheet_xml(w, [30, 70, 80], '7030A0', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A6" sqref="A6"/>').encode('utf-8')

v = Sheet()
banner(v, 12, 'Widok', 'Baza_PW wyświetlona wg wybranego widoku. To kopia – zmiany rób w Baza_PW, a widoki definiuj w arkuszu Widoki.')
v.put(1, 3, 'Widok:', 'lab', ht=26); v.put(2, 3, 'Wszystko', 'in_l')
v.put(1, 4, None, 'note', ht=18); v.put(2, 4, 'Wybierz widok z listy (strzałka przy B3).', 'note')
v.put(1, 5, None, 'text', ht=4)
v_dv = '<dataValidations count="1"><dataValidation type="list" allowBlank="0" showErrorMessage="1" errorTitle="Widok" error="Wybierz widok z listy (zdefiniowane w arkuszu Widoki)." sqref="B3"><formula1>Widoki!$A$6:$A$25</formula1></dataValidation></dataValidations>'
files['xl/worksheets/sheet8.xml'] = sheet_xml(v, [30, 36, 22, 40, 18, 18, 18, 18, 18, 18, 18, 18], '7030A0', '<pane ySplit="5" topLeftCell="A6" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="B3" sqref="B3"/>', dv=v_dv, gridlines=True).encode('utf-8')

# ------------------------------------------------------------------ Instrukcja (nowy arkusz)
I = Sheet()
W = [2, 26, 66, 66]
banner(I, 4, 'PW → Excel  ·  Instrukcja obsługi', 'Skoroszyt pobiera listę plików i ich atrybutów z ProjectWise i pokazuje, co zmieniło się od ostatniej aktualizacji.', h1=40, h2=24)
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

h2('1.  Szybki start (proces jest automatyczny)')
th('Krok', 'Co zrobić', 'Efekt')
tr('1. Foldery', 'Wklej link do folderu ProjectWise do kolumny B (format: patrz punkt 3). Zaloguj się w oknie PowerShell (Bentley IMS), jeśli o to poprosi.', 'Makro samo wykrywa atrybuty folderu (arkusz Diagnostyka) i dopisuje je do arkusza Atrybuty.', link='Foldery')
tr('2. Atrybuty', 'Opcjonalnie: w kolumnie B przełącz TAK/NIE przy atrybutach, które mają (nie) trafić do bazy.', 'Tylko atrybuty z TAK stają się kolumnami.', link='Atrybuty')
tr('3. Baza_PW', 'Po wykryciu atrybutów dane pobierają się same. Ręcznie: przycisk „AKTUALIZUJ DANE”. Poczekaj, aż okno PowerShell się zamknie.', 'Lista plików (każda rewizja w osobnym wierszu); zmiany podświetlone.', link='Baza_PW')
tr('4. Widok', 'Wybierz widok w komórce B3 (np. Rewizje, Linki). Własne widoki definiujesz w arkuszu Widoki.', 'Tabela z wybranymi kolumnami w wybranej kolejności.', link='Widok')
tr('5. Log', 'Pełna historia zmian (ADDED / CHANGED / DELETED). Każdy przebieg to osobna grupa wierszy – rozwiń plusem po lewej.', 'Ślad audytowy.', link='Log')

h2('2.  Arkusze – do czego służą i co w nich wpisać')
th('Arkusz', 'Do czego służy', 'Co wpisać / uwagi')
tr('Foldery', 'Lista folderów PW do skanowania.', 'A: TAK/NIE · B: link do folderu (pw:\\\\…) – wklejenie uruchamia automat · C: opis. Dane od wiersza 6. Błędny link podświetla się na czerwono.', link='Foldery')
tr('Atrybuty', 'Lista atrybutów PW (uzupełnia się sama).', 'A: nazwa · B: Pobierać TAK/NIE · C: typ (ATRYBUT = atrybut projektu, SYSTEM = właściwość PW) · D: przykład · E: Twoje notatki. Nic nie jest kasowane przy kolejnym wykryciu.', link='Atrybuty')
tr('Widoki', 'Szablony tabeli: które kolumny i w jakiej kolejności.', 'A: nazwa widoku · B: kolumny po przecinku (nazwy nagłówków lub litery, np. A, B, F, L, TK) · C: opis. Patrz punkt 5.', link='Widoki')
tr('Baza_PW', 'WYNIK – pełna lista plików z atrybutami, pogrupowana w foldery.', 'Nie edytuj ręcznie kolumn A–E – makro nadpisuje je przy każdej aktualizacji. Własne kolumny dopisuj na prawo od atrybutów.', link='Baza_PW')
tr('Widok', 'Baza_PW wyświetlona wg wybranego widoku.', 'Tylko B3 (lista wyboru). To kopia – odświeża się po zmianie widoku i po każdej aktualizacji danych.', link='Widok')
tr('Diagnostyka', 'Atrybuty jednego pliku z ostatnio wykrywanego folderu.', 'Tylko do odczytu; zawartość jest kasowana przy każdym wykryciu.', link='Diagnostyka')
tr('Log', 'Historia zmian ze wszystkich aktualizacji.', 'Tylko do odczytu. Nie usuwaj nagłówka w wierszu 1.', link='Log')

h2('3.  Jak ma wyglądać link do folderu (arkusz Foldery, kolumna B)')
para('Najprościej: w ProjectWise Explorer skopiuj ścieżkę folderu z paska adresu – zaczyna się od pw:\\\\. Link składa się z czterech części:')
th('Część', 'Znaczenie', 'Przykład')
tr('pw:\\\\', 'Prefiks – wymagany. Bez niego wiersz jest ignorowany.', 'pw:\\\\', st=('cellb', 'cell', 'mono'))
tr('<serwer>:<Datasource>', 'Adres serwera i nazwa datasource oddzielone dwukropkiem. Makro używa tego do zalogowania.', 'pwserver.firma.local:PW_Projekty', st=('cellb', 'cell', 'mono'))
tr('\\Documents\\', 'Stały segment korzenia. Wszystko po nim to ścieżka folderu.', '\\Documents\\', st=('cellb', 'cell', 'mono'))
tr('<ścieżka folderu>', 'Kolejne foldery oddzielone \\. Końcowy \\ jest opcjonalny.', 'PROJ01\\04_DELIVERABLES\\03-HR\\', st=('cellb', 'cell', 'mono'))
I.put(2, row[0], 'Pełny przykład', 'cellb', ht=34); I.put(3, row[0], 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\04_DELIVERABLES\\03-HR\\', 'ok'); I.merge(3, row[0], 4, row[0]); I.put(4, row[0], None, 'ok'); nxt()
th('Przykłady', 'Link', 'Ocena')
tr('Poprawny', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\Deliverables\\', 'OK – pełny format.', st=('ok', 'mono', 'cell'))
tr('Poprawny (skrót)', 'pw://PW_Projekty/Documents/PROJ01/Deliverables', 'OK – ukośniki / też działają, serwer można pominąć, jeśli datasource jest znany w ProjectWise.', st=('ok', 'mono', 'cell'))
tr('Poprawny (z dopiskiem)', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\PROJ01\\Deliverables (Working folder)', 'OK – dopisek „ (Working folder)” jest usuwany automatycznie.', st=('ok', 'mono', 'cell'))
tr('Błędny', '\\\\fileserver\\share\\PROJ01\\Deliverables', 'To zwykła ścieżka sieciowa, nie link PW – wiersz zostanie pominięty.', st=('bad', 'mono', 'cell'))
tr('Błędny', 'https://pwserver.firma.local/PROJ01', 'Adres WWW nie jest obsługiwany. Użyj ścieżki pw:\\\\.', st=('bad', 'mono', 'cell'))
tr('Błędny', 'pw:\\\\pwserver.firma.local:PW_Projekty\\Documents\\', 'Brak ścieżki folderu po \\Documents\\ – nie wiadomo, co skanować.', st=('bad', 'mono', 'cell'))

h2('4.  Kolumny w arkuszu Baza_PW')
th('Kolumna', 'Znaczenie', 'Jak wygląda')
tr('File Name', 'Nazwa pliku w ProjectWise.', 'PROJ01-STR-0001_Plan.dwg', st=('cellb', 'cell', 'mono'))
tr('PW Source Path', 'Folder pliku liczony od \\Documents\\. Wiersz folderu (niebieski) ma tę kolumnę pustą.', 'PROJ01\\04_DELIVERABLES\\03-HR', st=('cellb', 'cell', 'mono'))
tr('PW Link', 'Klikalny link do dokumentu – otwiera go w ProjectWise Explorer (musi być zainstalowany). W wierszu folderu jest tu link „Open folder in PW”. Link jest też kluczem, po którym makro rozpoznaje plik – nie edytuj go.', 'pw://pwserver.firma.local:PW_Projekty/Documents/D:{3F2A1B4C-0000-0000-0000-000000000000}', st=('cellb', 'cell', 'mono'))
tr('Updater', 'Kto ostatnio zmienił plik w PW.', 'Jan Kowalski', st=('cellb', 'cell', 'mono'))
tr('Updated', 'Data ostatniej zmiany, format RRRR-MM-DD.', '2026-05-27', st=('cellb', 'cell', 'mono'))
tr('F, G, H …', 'Atrybuty z arkusza Atrybuty – w tej samej kolejności.', 'np. Version = 3', st=('cellb', 'cell', 'mono'))

h2('5.  Widoki – jak zdefiniować tabelę')
para('Arkusz Widoki: jeden wiersz = jeden widok. W kolumnie B wypisz kolumny Baza_PW w takiej kolejności, w jakiej mają się pokazać. Możesz użyć nazw nagłówków albo liter kolumn (litery zależą od kolejności atrybutów, nazwy są trwalsze).')
th('Nazwa widoku', 'Kolumny (kolumna B w arkuszu Widoki)', 'Efekt')
tr('Rewizje', 'File Name, Version, VersionSequence, WorkflowState, Updated', 'Nazwa pliku, rewizja, numer wersji, stan i data – w tej kolejności.', st=('cellb', 'mono', 'cell'))
tr('Linki', 'File Name, PW Link, PW Source Path', 'Nazwa, klikalny link i folder.', st=('cellb', 'mono', 'cell'))
tr('Litery', 'A, B, F, L, TK   (albo: A B F L TK)', 'Kolumny Baza_PW wskazane literami, w podanej kolejności.', st=('cellb', 'mono', 'cell'))
tr('Wszystko', '*', 'Wszystkie kolumny.', st=('cellb', 'mono', 'cell'))
para('Widok wybierasz w arkuszu Widok (B3). Jeśli któraś nazwa/litera nie istnieje, arkusz Widok pokaże ją w komórce D4.', 'note')

h2('6.  Kolory i oznaczenia')
th('Oznaczenie', 'Znaczenie', 'Co robi makro')
I.put(2, row[0], 'Niebieski wiersz', 'sw_blue', ht=24); I.put(3, row[0], 'Folder – nagłówek grupy plików.', 'cell'); I.put(4, row[0], 'Dodaje przy pierwszym pliku z nowego folderu.', 'cell'); nxt()
I.put(2, row[0], 'Żółta komórka', 'sw_yel', ht=36); I.put(3, row[0], 'Wartość zmieniła się od ostatniej aktualizacji. Najedź na komórkę – komentarz pokazuje starą wartość, datę i osobę.', 'cell'); I.put(4, row[0], 'Nadpisuje wartość nową i dopisuje komentarz. Wpis CHANGED w Log.', 'cell'); nxt()
I.put(2, row[0], 'Czerwony, przekreślony wiersz', 'sw_red', ht=36); I.put(3, row[0], 'Plik zniknął z ProjectWise (w skanowanym folderze).', 'cell'); I.put(4, row[0], 'Zaznacza wiersz, nie kasuje go. Jeśli plik wróci – kolor znika. Wpis DELETED w Log.', 'cell'); nxt()
I.put(2, row[0], 'Biały wiersz', 'sw_white', ht=24); I.put(3, row[0], 'Bez zmian.', 'cell'); I.put(4, row[0], '–', 'cell'); nxt()

h2('7.  Wymagania')
for t in ['Windows + Excel z włączonymi makrami (plik .xlsm zapisany na dysku, nie otwarty z załącznika e-mail).',
          'Zainstalowany moduł PowerShell PWPS_DAB (Bentley) oraz dostęp do serwera ProjectWise; logowanie przez Bentley IMS.',
          'Makro tworzy obok skoroszytu pliki tymczasowe: PW_Folders.txt, PW_GetFiles.ps1, PW_Output.csv, PW_Structure_Log.txt. Po aktualizacji zapyta, czy je usunąć.']:
    para('•  ' + t)

h2('8.  Rozwiązywanie problemów')
th('Objaw', 'Przyczyna', 'Co zrobić')
tr('W Log: WARNING „Scope not scanned”', 'Nie udało się zalogować lub znaleźć folderu.', 'Sprawdź link w arkuszu Foldery i logowanie do PW. Pliki tego folderu NIE zostały oznaczone jako usunięte.')
tr('Okno PowerShell: FOLDER NOT FOUND', 'Zła ścieżka lub brak uprawnień do folderu.', 'Skopiuj link ponownie z ProjectWise Explorer.')
tr('Kolumna atrybutu jest pusta', 'Nazwa atrybutu nie istnieje w tym folderze albo ma inną pisownię.', 'Sprawdź nazwę w arkuszu Diagnostyka / Atrybuty; atrybuty zależą od folderu i klasy dokumentu.')
tr('„Please save the Excel workbook…”', 'Plik nie jest zapisany na dysku.', 'Zapisz skoroszyt w zwykłym folderze i uruchom ponownie.')
tr('Pobieranie trwa bardzo długo', 'Czas zjada serwer PW (pobieranie atrybutów i rewizji), nie Excel.', 'Arkusz Atrybuty → „TEST SZYBKOŚCI PW”: porównuje sposoby pobierania i zapisuje czasy w arkuszu Test_PW. Mniej czasu daje też wyłączenie rewizji (GET_ALL_VERSIONS w modConfig).')
tr('Makra nie działają', 'Plik zablokowany po pobraniu.', 'Właściwości pliku → Odblokuj; w Excelu włącz zawartość (makra).')
I.merge(2, 3, 4, 3)
I.put(2, 3, 'Wersja skoroszytu: V1.6  ·  Wystarczy wkleić link w arkuszu Foldery – resztę robi makro. Własne ustawienia: Atrybuty (TAK/NIE) i Widoki.', 'note', ht=22)
I.put(3, 3, None, 'note'); I.put(4, 3, None, 'note')
files['xl/worksheets/sheet6.xml'] = sheet_xml(I, W, '1F3864', '<selection activeCell="B6" sqref="B6"/>', landscape=False).replace('<sheetView showGridLines="0" ', '<sheetView showGridLines="0" tabSelected="1" ').encode('utf-8')

# ------------------------------------------------------------------ workbook-level
wb = T('xl/workbook.xml')
wb = re.sub(r'<sheets>.*?</sheets>',
            '<sheets><sheet name="Instrukcja" sheetId="9" r:id="rId11"/><sheet name="Foldery" sheetId="2" r:id="rId2"/><sheet name="Atrybuty" sheetId="3" r:id="rId3"/><sheet name="Widoki" sheetId="10" r:id="rId12"/>'
            '<sheet name="Baza_PW" sheetId="8" r:id="rId1"/><sheet name="Widok" sheetId="11" r:id="rId13"/><sheet name="Diagnostyka" sheetId="4" r:id="rId4"/><sheet name="Log" sheetId="7" r:id="rId5"/></sheets>', wb, flags=re.S)
wb = wb.replace('activeTab="3"', 'activeTab="0"')
files['xl/workbook.xml'] = wb.encode('utf-8')
rels = T('xl/_rels/workbook.xml.rels')
rels = re.sub(r'<Relationship Id="rId9"[^>]*calcChain[^>]*/>', '', rels)
rels = rels.replace('</Relationships>', '<Relationship Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet6.xml"/><Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet7.xml"/><Relationship Id="rId13" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet8.xml"/></Relationships>')
files['xl/_rels/workbook.xml.rels'] = rels.encode('utf-8')
ct = T('[Content_Types].xml')
ct = re.sub(r'<Override PartName="/xl/calcChain.xml"[^>]*/>', '', ct)
ct = ct.replace('<Override PartName="/xl/theme/theme1.xml"', '<Override PartName="/xl/worksheets/sheet6.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet7.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet8.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/theme/theme1.xml"')
files['[Content_Types].xml'] = ct.encode('utf-8')
del files['xl/calcChain.xml']
app = T('docProps/app.xml')
app = app.replace('<vt:i4>5</vt:i4>', '<vt:i4>8</vt:i4>')
app = re.sub(r'<TitlesOfParts>.*?</TitlesOfParts>', '<TitlesOfParts><vt:vector size="8" baseType="lpstr"><vt:lpstr>Instrukcja</vt:lpstr><vt:lpstr>Foldery</vt:lpstr><vt:lpstr>Atrybuty</vt:lpstr><vt:lpstr>Widoki</vt:lpstr><vt:lpstr>Baza_PW</vt:lpstr><vt:lpstr>Widok</vt:lpstr><vt:lpstr>Diagnostyka</vt:lpstr><vt:lpstr>Log</vt:lpstr></vt:vector></TitlesOfParts>', app, flags=re.S)
files['docProps/app.xml'] = app.encode('utf-8')

with zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', files.pop('[Content_Types].xml'))
    for n, v in files.items(): z.writestr(n, v)
print('OK ->', DST)

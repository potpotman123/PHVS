import json, re, pathlib
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT

ROOT = pathlib.Path(__file__).resolve().parents[1]
data = json.loads((ROOT/'research/memo_content.json').read_text())
doc = Document()
sec = doc.sections[0]
sec.top_margin = sec.bottom_margin = Inches(.7)
sec.left_margin = sec.right_margin = Inches(.72)
sec.page_width = Inches(8.5); sec.page_height = Inches(11)
styles = doc.styles
for name in ['Normal','Title','Subtitle','Heading 1','Heading 2','Heading 3']:
    st=styles[name]; st.font.name='Calibri'; st.font.color.rgb=RGBColor(0,0,0)
for st in styles:
    for border in list(st.element.iter(qn('w:pBdr'))):
        border.getparent().remove(border)
styles['Normal'].font.size=Pt(10.5)
styles['Normal'].paragraph_format.space_after=Pt(6)
styles['Normal'].paragraph_format.line_spacing=1.08
styles['Title'].font.size=Pt(24)
styles['Subtitle'].font.size=Pt(11)
styles['Heading 1'].font.size=Pt(16)
styles['Heading 1'].paragraph_format.space_before=Pt(15)
styles['Heading 1'].paragraph_format.space_after=Pt(7)
styles['Heading 2'].font.size=Pt(12)
styles['Heading 2'].paragraph_format.space_before=Pt(10)
styles['Heading 2'].paragraph_format.space_after=Pt(5)
for name in ['Heading 1','Heading 2','Heading 3']:
    styles[name].paragraph_format.keep_with_next=True
sources={s['id']:s for s in data['sources']}
def hyperlink(p,label,url):
    rid=p.part.relate_to(url,'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',is_external=True)
    h=OxmlElement('w:hyperlink');h.set(qn('r:id'),rid)
    r=OxmlElement('w:r'); rp=OxmlElement('w:rPr'); color=OxmlElement('w:color');color.set(qn('w:val'),'1A5276');rp.append(color);r.append(rp)
    t=OxmlElement('w:t');t.text=label;r.append(t);h.append(r);p._p.append(h)
def text(p,s):
    for token in re.split(r'(\[[A-Z]\d+\]|\*\*.*?\*\*)',s):
        if re.fullmatch(r'\[[A-Z]\d+\]',token) and token[1:-1] in sources:
            hyperlink(p,token,sources[token[1:-1]]['url'])
        elif token.startswith('**') and token.endswith('**'):
            p.add_run(token[2:-2]).bold=True
        else:p.add_run(token)
def paragraph(s,style=None):
    p=doc.add_paragraph(style=style);text(p,s);return p
def table(block):
    headers=block['headers'];rows=block['rows']
    t=doc.add_table(rows=1,cols=len(headers));t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    widths=block.get('widths',[7.06/len(headers)]*len(headers))
    for i,w in enumerate(widths):t.columns[i].width=Inches(w)
    for i,h in enumerate(headers):text(t.rows[0].cells[i].paragraphs[0],h)
    for vals in rows:
        cells=t.add_row().cells
        for i,s in enumerate(vals):text(cells[i].paragraphs[0],s)
    pr=t._tbl.tblPr
    borders=OxmlElement('w:tblBorders')
    for edge in ['top','left','bottom','right','insideH','insideV']:
        b=OxmlElement('w:'+edge);b.set(qn('w:val'),'single');b.set(qn('w:sz'),'4');b.set(qn('w:color'),'D9D9D9');borders.append(b)
    pr.append(borders)
    for ri,row in enumerate(t.rows):
        for ci,c in enumerate(row.cells):
            c.width=Inches(widths[ci]);c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            cp=c._tc.get_or_add_tcPr();m=OxmlElement('w:tcMar')
            for e in ['top','left','bottom','right']:
                q=OxmlElement('w:'+e);q.set(qn('w:w'),'80');q.set(qn('w:type'),'dxa');m.append(q)
            cp.append(m)
            sh=OxmlElement('w:shd');sh.set(qn('w:fill'),'E4EBEF' if ri==0 else ('F5F7F8' if ri%2==0 else 'FFFFFF'));cp.append(sh)
            for p in c.paragraphs:
                p.paragraph_format.space_after=Pt(2);p.paragraph_format.space_before=Pt(2);p.paragraph_format.line_spacing=1.04
                for r in p.runs:r.font.size=Pt(9.5);r.bold=(ri==0)
        trpr=row._tr.get_or_add_trPr()
        if ri==0:
            repeat=OxmlElement('w:tblHeader');trpr.append(repeat)
        cant=OxmlElement('w:cantSplit');trpr.append(cant)
    doc.add_paragraph().paragraph_format.space_after=Pt(1)

title=paragraph(data['title'],'Title')
for border in list(title._p.iter(qn('w:pBdr'))):
    border.getparent().remove(border)
paragraph(data['subtitle'],'Subtitle')
closing_at=next(i for i,b in enumerate(data['blocks']) if b.get('text')=='What is ready for the eventual model')
for block in data['blocks'][:closing_at]:
    if block['type']=='h1':paragraph(block['text'],'Heading 1')
    elif block['type']=='h2':paragraph(block['text'],'Heading 2')
    elif block['type']=='p':paragraph(block['text'])
    elif block['type']=='table':table(block)
    elif block['type']=='equation':
        p=paragraph(block['text']);p.paragraph_format.left_indent=Inches(.15)
        for r in p.runs:r.font.name='Consolas';r.font.size=Pt(9.5)
    elif block['type']=='bullet':paragraph(block['text'],'List Bullet')

paragraph('Evidence ledger','Heading 1')
paragraph('All web sources were accessed on 6 October 2026. O = observed disclosure, D = derived arithmetic, P = proxy, A = assumed scenario, U = unknown. Sponsor materials are primary disclosures, not independent validation. Repeated releases, posters and filings of the same dataset count as one underlying trial. Click a source identifier in the text or a source title below to open the original.')
for s in data['sources']:
    p=doc.add_paragraph();p.paragraph_format.space_before=Pt(5)
    p.paragraph_format.keep_with_next=True
    p.add_run(s['id']+'  ').bold=True;hyperlink(p,s['title'],s['url'])
    paragraph(s['note'])
paragraph('Research boundary and reproducibility','Heading 2')
paragraph(data['method'])
for block in data['blocks'][closing_at:]:
    paragraph(block['text'],'Heading 1' if block['type']=='h1' else None)

footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
footer.add_run('PHVS research  |  6 October 2026  |  ').font.size=Pt(8)
field=OxmlElement('w:fldSimple');field.set(qn('w:instr'),'PAGE');footer._p.append(field)
doc.core_properties.title=data['title'];doc.core_properties.subject='Clinical and commercial research assessment and quantitative valuation framework'
doc.core_properties.author='Investment research'
out=ROOT/'output/Pharvaris_research_memo_2026-10-06.docx';doc.save(out)
print(out)
print('Words in content:',len(' '.join(str(x) for x in data['blocks']).split()))

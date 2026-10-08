from pathlib import Path
from zipfile import ZipFile

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from lxml import etree


ROOT = Path(__file__).resolve().parents[1]
DEST = ROOT / 'output' / 'centered_equation_demo.docx'
DEST.parent.mkdir(exist_ok=True)


def elem(tag, *children):
    node = OxmlElement(tag)
    for child in children:
        node.append(child)
    return node


def text(value):
    node = elem('m:t')
    node.text = value
    return elem('m:r', node)


def squared(value):
    return elem('m:sSup', elem('m:e', text(value)), elem('m:sup', text('2')))


def derivative(variable, second=False):
    numerator = elem('m:num')
    denominator = elem('m:den')
    if second:
        numerator.extend([squared('∂'), text('u')])
        denominator.extend([text('∂'), squared(variable)])
    else:
        numerator.append(text('∂u'))
        denominator.append(text('∂' + variable))
    return elem('m:f', numerator, denominator)


doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.top_margin = sec.bottom_margin = Cm(2.5)
sec.left_margin = sec.right_margin = Cm(2.5)
normal = doc.styles['Normal']
normal.font.name = 'Times New Roman'
normal.font.size = Pt(12)
fonts = normal.element.get_or_add_rPr().get_or_add_rFonts()
for slot, font in [('eastAsia', '標楷體'), ('ascii', 'Times New Roman'), ('hAnsi', 'Times New Roman')]:
    fonts.set(qn('w:' + slot), font)

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('Word 原生方程式示範')
run.bold = True
run.font.size = Pt(18)
title.paragraph_format.space_after = Pt(24)
doc.add_paragraph('以下為置中的二維熱傳導方程式，可在 Word 中直接點選並編輯。')

paragraph = doc.add_paragraph()
paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
paragraph.paragraph_format.space_before = Pt(24)
paragraph.paragraph_format.space_after = Pt(24)
paragraph.paragraph_format.line_spacing = 1.5
jc = elem('m:jc')
jc.set(qn('m:val'), 'center')
math = elem('m:oMath', derivative('t'), text('=α('), derivative('x', True), text('+'), derivative('y', True), text(')'))
paragraph._p.append(elem('m:oMathPara', elem('m:oMathParaPr', jc), math))

doc.save(DEST)
with ZipFile(DEST) as package:
    root = etree.fromstring(package.read('word/document.xml'))
    ns = {'m': 'http://schemas.openxmlformats.org/officeDocument/2006/math', 'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    assert len(root.xpath('//m:oMath', namespaces=ns)) == 1
    assert len(root.xpath('//m:f', namespaces=ns)) == 3
    assert root.xpath('//m:oMathPara/m:oMathParaPr/m:jc/@m:val', namespaces=ns) == ['center']
    assert root.xpath('//w:p[m:oMathPara]/w:pPr/w:jc/@w:val', namespaces=ns) == ['center']
    assert not root.xpath('//w:drawing', namespaces=ns)
print(f'Created and verified native centered equation: {DEST}')

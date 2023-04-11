from docx import Document
from docx.oxml import OxmlElement, ns
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL

class Directive:

    

    def _create_element(self, name):
        return OxmlElement(name)

    def _create_attribute(self, element, name, value):
        element.set(ns.qn(name), value)


    def _add_page_number(self, run):
        fldChar1 = create_element('w:fldChar')
        create_attribute(fldChar1, 'w:fldCharType', 'begin')

        instrText = create_element('w:instrText')
        create_attribute(instrText, 'xml:space', 'preserve')
        instrText.text = "PAGE"

        fldChar2 = create_element('w:fldChar')
        create_attribute(fldChar2, 'w:fldCharType', 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)
    
def _set_page_properties(self, document):
    header = self._document.sections[0]
    header.page_width = Cm(21)
    header.page_height = Cm(29.7)
    header.left_margin = Cm(3)
    header.right_margin = Cm(1.5)
    header.top_margin = Cm(2)
    header.bottom_margin = Cm(2)

def _set_directive_styles(self, document):
    main_style = document.styles.add_style(
        'Directive Text', WD_STYLE_TYPE.PARAGRAPH)
    main_style.next_paragraph_style = main_style
    main_style.font.size = Pt(14)
    main_style.font.name = 'Times New Roman'
    p_f = main_style.paragraph_format
    p_f.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_f.first_line_indent = Cm(1.25)
    p_f.line_spacing = 1.15
    p_f.space_before = Pt(0)
    p_f.space_after = Pt(0)

    title_style = document.styles.add_style(
        'Directive Title', WD_STYLE_TYPE.PARAGRAPH)
    title_style.base_style = main_style
    title_style.next_paragraph_style = main_style
    title_style.font.bold = True
    p_f = title_style.paragraph_format
    p_f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_f.line_spacing = 1.15
    p_f.first_line_indent = Cm(0)
    p_f.space_before = Pt(238)
    p_f.space_after = Pt(42)

    position_style = document.styles.add_style(
        'Directive Position', WD_STYLE_TYPE.PARAGRAPH)
    position_style.base_style = main_style
    p_f = position_style.paragraph_format
    p_f.first_line_indent = Cm(0)
    p_f.alignment = WD_ALIGN_PARAGRAPH.LEFT


    name_style = document.styles.add_style(
        'Directive Name', WD_STYLE_TYPE.PARAGRAPH)
    name_style.base_style = main_style
    p_f = name_style.paragraph_format
    p_f.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    return [title_style, main_style, position_style, name_style]

def _create_position_table(self, document, data, styles):
    table = document.add_table(rows=1, cols=2)
    cells = table._cells
    for d, s, c in zip(data, styles, cells):
        c.text = d
        c.vertical_alignment = WD_ALIGN_VERTICAL.BOTTOM
        c.paragraphs[0].style = s

def __init__(self):
    self._document = Document()
    self._set_page_properties(self._document)
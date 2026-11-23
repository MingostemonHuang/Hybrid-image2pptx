"""Convert LaTeX to native DrawingML math with a source-image fallback."""
from copy import deepcopy
import re

from lxml import etree

MATH = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
DRAWING = 'http://schemas.openxmlformats.org/drawingml/2006/main'
A14 = 'http://schemas.microsoft.com/office/drawing/2010/main'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
NS = {'m': MATH, 'a': DRAWING, 'a14': A14, 'mc': MC}


def compile_equation(latex, size_pt, color='#111111', bold=False):
    """Return Office Math, preserving its structured mathematical operators."""
    import latex2mathml.converter
    import mathml2omml

    if not latex.strip():
        raise ValueError('Equation LaTeX is empty')
    mathml = latex2mathml.converter.convert(latex)
    mathml_root = etree.fromstring(mathml.encode())
    tokens = ''.join(mathml_root.itertext())
    if re.search(r'\\[A-Za-z]+', tokens):
        raise ValueError(f'Unconverted LaTeX command in equation: {latex}')
    omml = mathml2omml.convert(mathml)
    wrapper = etree.fromstring(
        (f'<root xmlns:m="{MATH}">{omml}</root>').encode())
    root = wrapper[0]
    if root.tag not in [f'{{{MATH}}}oMath', f'{{{MATH}}}oMathPara']:
        raise ValueError('LaTeX conversion did not produce Office Math')
    style = etree.Element(f'{{{DRAWING}}}rPr',
                          sz=str(round(size_pt * 100)), b='1' if bold else '0')
    fill = etree.SubElement(style, f'{{{DRAWING}}}solidFill')
    etree.SubElement(fill, f'{{{DRAWING}}}srgbClr', val=color.lstrip('#'))
    for name in ['latin', 'ea', 'cs']:
        etree.SubElement(style, f'{{{DRAWING}}}{name}', typeface='Cambria Math')
    for run in root.iter(f'{{{MATH}}}r'):
        index = 1 if len(run) and run[0].tag == f'{{{MATH}}}rPr' else 0
        run.insert(index, deepcopy(style))
    for properties in root.iter():
        if properties.tag.startswith(f'{{{MATH}}}') and properties.tag.endswith('Pr'):
            if etree.QName(properties).localname in ['rPr', 'ctrlPr', 'oMathParaPr']:
                continue
            control = etree.SubElement(properties, f'{{{MATH}}}ctrlPr')
            control.append(deepcopy(style))
    return root


def wrap_equation(shape, fallback_shape, item, pt_scale):
    """Use native math in PowerPoint and an exclusive fallback in other viewers."""
    math = compile_equation(item['latex'], item.get('size_px', 25) * pt_scale,
                            item.get('color', '#111111'), item.get('bold', False))
    paragraph = shape.text_frame.paragraphs[0]._p
    for child in list(paragraph):
        if child.tag != f'{{{DRAWING}}}pPr':
            paragraph.remove(child)
    carrier = etree.SubElement(paragraph, f'{{{A14}}}m', nsmap={'a14': A14, 'm': MATH})
    carrier.append(math)
    shape._element.nvSpPr.cNvPr.set('descr', 'LaTeX: ' + item['latex'])
    fallback_shape.name = item.get('id', 'equation') + '-preview'
    parent = shape._element.getparent()
    index = parent.index(shape._element)
    alternate = etree.Element(f'{{{MC}}}AlternateContent', nsmap={'mc': MC})
    choice = etree.SubElement(alternate, f'{{{MC}}}Choice',
                              nsmap={'a14': A14}, Requires='a14')
    choice.append(shape._element)
    fallback = etree.SubElement(alternate, f'{{{MC}}}Fallback')
    fallback.append(fallback_shape._element)
    parent.insert(index, alternate)
    return shape

"""Build editable slide structures while preserving source-image artwork."""
from __future__ import annotations

import argparse
import io
import json
from pathlib import Path

from PIL import Image, ImageDraw
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.xmlchemy import OxmlElement
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt, Emu
from equations import wrap_equation


def xml(tag, **attrs):
    node = OxmlElement(tag)
    for key, value in attrs.items():
        node.set(key, str(value))
    return node


def rgb(value):
    return RGBColor.from_string(value.lstrip('#'))


def clear_theme_effects(shape):
    for node in shape._element.iter(qn('a:effectRef')):
        node.set('idx','0')


def paint(shape, item):
    clear_theme_effects(shape)
    fill = item.get('fill')
    if isinstance(fill, str):
        shape.fill.solid()
        shape.fill.fore_color.rgb = rgb(fill)
    elif isinstance(fill, dict) and fill.get('gradient'):
        sp = shape._element.spPr
        for name in ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']:
            for old in list(sp.findall(qn('a:'+name))):
                sp.remove(old)
        grad = xml('a:gradFill', rotWithShape='1')
        stops = xml('a:gsLst')
        values = fill['gradient']
        for i, color in enumerate(values):
            stop = xml('a:gs', pos=round(i*100000/(len(values)-1)))
            stop.append(xml('a:srgbClr', val=color.lstrip('#')))
            stops.append(stop)
        grad.append(stops)
        grad.append(xml('a:lin', ang=round(fill.get('angle', 90)*60000), scaled='1'))
        sp.insert_element_before(grad, 'a:ln', 'a:effectLst', 'a:effectDag')
    else:
        shape.fill.background()
    line = item.get('stroke')
    if line:
        shape.line.color.rgb = rgb(line)
        shape.line.width = Pt(item.get('stroke_pt', 0.6))
    else:
        shape.line.fill.background()
    sp=shape._element.spPr
    for old in list(sp.findall(qn('a:effectLst'))): sp.remove(old)
    sp.append(xml('a:effectLst'))
    if item.get('shadow'):
        effects = sp.find(qn('a:effectLst'))
        shadow = xml('a:outerShdw', blurRad=50800, dist=19050, dir=5400000,
                     algn='ctr', rotWithShape='0')
        color = xml('a:srgbClr', val=item.get('shadow_color', '888888').lstrip('#'))
        color.append(xml('a:alpha', val=item.get('shadow_alpha', 13000)))
        shadow.append(color)
        effects.append(shadow)


def add_text(slide, item, xscale, yscale, pt_scale):
    x,y,w,h = item['box']
    shape = slide.shapes.add_textbox(Emu(x*xscale), Emu(y*yscale), Emu(w*xscale), Emu(h*yscale))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = item.get('wrap', False)
    frame.margin_left = frame.margin_right = frame.margin_top = frame.margin_bottom = 0
    frame.vertical_anchor = {'top':MSO_ANCHOR.TOP,'middle':MSO_ANCHOR.MIDDLE,'bottom':MSO_ANCHOR.BOTTOM}.get(item.get('valign'), MSO_ANCHOR.TOP)
    lines = item.get('text','').split('\n')
    for i,line in enumerate(lines):
        p = frame.paragraphs[0] if i==0 else frame.add_paragraph()
        p.alignment = {'left':PP_ALIGN.LEFT,'center':PP_ALIGN.CENTER,'right':PP_ALIGN.RIGHT}.get(item.get('align'),PP_ALIGN.LEFT)
        p.space_before = p.space_after = Pt(0)
        p.line_spacing = item.get('line_spacing',1.05)
        runs = item.get('runs') if i==0 and len(lines)==1 else None
        for data in runs or [{'text':line}]:
            run = p.add_run(); run.text = data['text']
            run.font.size = Pt(data.get('size_px',item.get('size_px',22))*pt_scale)
            run.font.bold = data.get('bold', item.get('bold',False))
            run.font.italic = data.get('italic',item.get('italic',False))
            run.font.color.rgb = rgb(data.get('color',item.get('color','#222222')))
            font = data.get('font',item.get('font','Microsoft YaHei'))
            run.font.name = font
            rpr = run._r.get_or_add_rPr()
            for tag in ['a:ea','a:cs']:
                node = rpr.find(qn(tag))
                if node is None:
                    node = xml(tag); rpr.append(node)
                node.set('typeface',font)
    return shape


def add_picture(slide, item, source, xscale, yscale):
    x,y,w,h = item['box']
    sx,sy,sw,sh = item.get('source_box',item['box'])
    image = source.crop((round(sx),round(sy),round(sx+sw),round(sy+sh))).convert('RGBA')
    # Exclusion masks separate text/plates from the retained source artwork.
    if item.get('exclude_boxes'):
        mask = Image.new('L',image.size,255)
        draw = ImageDraw.Draw(mask)
        for bx,by,bw,bh in item['exclude_boxes']:
            draw.rectangle((round(bx-sx),round(by-sy),round(bx+bw-sx),round(by+bh-sy)), fill=0)
        image.putalpha(mask)
    buf=io.BytesIO(); image.save(buf,format='PNG'); buf.seek(0)
    shape=slide.shapes.add_picture(buf,Emu(x*xscale),Emu(y*yscale),width=Emu(w*xscale),height=Emu(h*yscale))
    shape._element.spPr.append(xml('a:effectLst'))
    shape.line.fill.background()
    style=xml('p:style')
    for tag in ['lnRef','fillRef','effectRef']:
        ref=xml('a:'+tag,idx='0')
        ref.append(xml('a:srgbClr',val='000000')); style.append(ref)
    font_ref=xml('a:fontRef',idx='minor')
    font_ref.append(xml('a:schemeClr',val='lt1')); style.append(font_ref)
    shape._element.append(style)
    clear_theme_effects(shape)
    return shape


def add_connector(slide, item, xscale, yscale):
    points = item['points']
    a,b = points[0],points[-1]
    kind = {'straight':MSO_CONNECTOR.STRAIGHT,'elbow':MSO_CONNECTOR.ELBOW,'curve':MSO_CONNECTOR.CURVE}[item.get('kind','straight')]
    if len(points)==4:
        left=min(p[0] for p in points); top=min(p[1] for p in points)
        width=max(max(p[0] for p in points)-left,1)
        height=max(max(p[1] for p in points)-top,1)
        shape=slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,Emu(left*xscale),Emu(top*yscale),Emu(width*xscale),Emu(height*yscale))
        shape.fill.background()
    else:
        shape = slide.shapes.add_connector(kind,Emu(a[0]*xscale),Emu(a[1]*yscale),Emu(b[0]*xscale),Emu(b[1]*yscale))
    shape.line.color.rgb = rgb(item.get('color','#C20B15'))
    shape.line.width = Pt(item.get('width_pt',1.3))
    shape._element.spPr.append(xml('a:effectLst'))
    if len(points)==4:
        left=min(p[0] for p in points); top=min(p[1] for p in points)
        width=max(p[0] for p in points)-left; height=max(p[1] for p in points)-top
        width=max(width,1); height=max(height,1)
        sp=shape._element.spPr
        for old in list(sp.findall(qn('a:prstGeom'))): sp.remove(old)
        geom=xml('a:custGeom')
        for name in ['avLst','gdLst','ahLst','cxnLst']: geom.append(xml('a:'+name))
        geom.append(xml('a:rect',l='0',t='0',r='r',b='b'))
        paths=xml('a:pathLst'); path=xml('a:path',w=round(width*1000),h=round(height*1000),fill='none',stroke='1')
        move=xml('a:moveTo'); move.append(xml('a:pt',x=round((a[0]-left)*1000),y=round((a[1]-top)*1000))); path.append(move)
        curve=xml('a:cubicBezTo')
        for point in points[1:]: curve.append(xml('a:pt',x=round((point[0]-left)*1000),y=round((point[1]-top)*1000)))
        path.append(curve); paths.append(path); geom.append(paths)
        sp.insert_element_before(geom,'a:ln','a:effectLst')
        xfrm=sp.xfrm
        xfrm.off.x=round(left*xscale); xfrm.off.y=round(top*yscale)
        xfrm.ext.cx=round(width*xscale); xfrm.ext.cy=round(height*yscale)
        for key in ['flipH','flipV']: xfrm.attrib.pop(key,None)
    # Preserve measured paths; semantic endpoint IDs remain in the scene file.
    clear_theme_effects(shape)
    ln=shape.line._get_or_add_ln()
    if item.get('arrow',True): ln.append(xml('a:tailEnd',type='triangle',w='sm',len='sm'))
    return shape


def build(scene_path, output):
    scene_path=Path(scene_path); scene=json.loads(scene_path.read_text())
    source_path=Path(scene['source'])
    if not source_path.is_absolute(): source_path=scene_path.parent/source_path
    source=Image.open(source_path).convert('RGB')
    width,height=scene['canvas']
    if source.size!=(width,height): raise ValueError('Scene canvas must match source pixels')
    prs=Presentation(); prs.slide_width=Inches(scene.get('width_in',13.333333)); prs.slide_height=Inches(scene.get('height_in',7.5))
    slide=prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid(); slide.background.fill.fore_color.rgb=rgb(scene.get('background','#FFFFFF'))
    xs=prs.slide_width/width; ys=prs.slide_height/height; ps=prs.slide_height/12700/height
    presets={'rect':MSO_SHAPE.RECTANGLE,'round_rect':MSO_SHAPE.ROUNDED_RECTANGLE,'ellipse':MSO_SHAPE.OVAL,'hexagon':MSO_SHAPE.HEXAGON,'arrow_up':MSO_SHAPE.UP_ARROW,'arrow_down':MSO_SHAPE.DOWN_ARROW,'parallelogram':MSO_SHAPE.PARALLELOGRAM}
    for i,item in enumerate(scene['elements']):
        kind=item['type']
        if kind=='text': shape=add_text(slide,item,xs,ys,ps)
        elif kind=='equation':
            shape=add_text(slide,{**item,'text':'','runs':None},xs,ys,ps)
            fallback=add_picture(slide,{
                'box':item['box'], 'source_box':item.get('source_box',item['box'])
            },source,xs,ys)
            shape=wrap_equation(shape,fallback,item,ps)
        elif kind=='picture': shape=add_picture(slide,item,source,xs,ys)
        elif kind=='connector': shape=add_connector(slide,item,xs,ys)
        elif kind=='shape':
            x,y,w,h=item['box']
            shape=slide.shapes.add_shape(presets[item.get('preset','round_rect')],Emu(x*xs),Emu(y*ys),Emu(w*xs),Emu(h*ys))
            if item.get('preset','round_rect') in ['round_rect','hexagon']: shape.adjustments[0]=item.get('radius',0.08)
            paint(shape,item)
        else: raise ValueError('Unknown element type: '+kind)
        key=item.get('id',f'object-{i:03d}'); shape.name=key
    prs.core_properties.title=scene.get('title','')
    prs.core_properties.author=''; prs.core_properties.last_modified_by=''; prs.core_properties.comments=''
    if scene.get('notes'): slide.notes_slide.notes_text_frame.text=scene['notes']
    output=Path(output); output.parent.mkdir(parents=True,exist_ok=True); prs.save(output)
    return output


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('scene'); parser.add_argument('output')
    args=parser.parse_args(); print(build(args.scene,args.output))

"""Report native objects, Office Math and connector references in a PPTX."""
import argparse
import json
from pathlib import Path
import re
from zipfile import ZipFile

from lxml import etree
from equations import NS

NS = {**NS, 'p':'http://schemas.openxmlformats.org/presentationml/2006/main'}


def inspect(path):
    pages=[]
    with ZipFile(path) as archive:
        names=sorted((name for name in archive.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',name)),
                     key=lambda name:int(re.search(r'slide(\d+)\.xml',name).group(1)))
        for name in names:
            root=etree.fromstring(archive.read(name))
            ids=root.xpath('.//p:cNvPr/@id',namespaces=NS)
            if len(ids)!=len(set(ids)):
                raise ValueError(f'Duplicate shape IDs in {name}')
            counts={'text':0,'equation':0,'shape':0,'connector':0,'native_curve':0,
                    'picture':0,'equation_preview':0,'bound_ends':0}
            tree=root.find('.//p:spTree',NS)
            selected=[]
            for node in tree:
                if node.tag==f"{{{NS['mc']}}}AlternateContent":
                    math=node.xpath('./mc:Choice//m:oMath',namespaces=NS)
                    if not math:
                        raise ValueError('AlternateContent has no native equation')
                    counts['equation']+=len(math)
                    previews=node.xpath('./mc:Fallback/p:pic',namespaces=NS)
                    if not previews:
                        raise ValueError('Native equation has no source preview')
                    counts['equation_preview']+=len(previews)
                else:
                    selected.append(node)
            for node in selected:
                tag=etree.QName(node).localname
                if tag=='pic': counts['picture']+=1
                elif tag=='cxnSp':
                    counts['connector']+=1
                    for ref in node.xpath('.//a:stCxn|.//a:endCxn',namespaces=NS):
                        if ref.get('id') not in ids:
                            raise ValueError('Connector references a missing node')
                        counts['bound_ends']+=1
                elif node.xpath('.//a:cubicBezTo',namespaces=NS): counts['native_curve']+=1
                elif node.xpath('.//a:t/text()',namespaces=NS): counts['text']+=1
                elif tag=='sp': counts['shape']+=1
            pages.append(counts)
    return {'file':str(Path(path).resolve()),'slides':pages}


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__); parser.add_argument('pptx')
    print(json.dumps(inspect(parser.parse_args().pptx),ensure_ascii=False,indent=2))

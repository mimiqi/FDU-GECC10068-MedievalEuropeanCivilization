from pathlib import Path
from zipfile import ZipFile
import json, re
from lxml import etree
from docx import Document

root=Path(r'D:\text\02_中世纪欧洲文明\讨论课')
work=root/'.制作过程'/'整合张智语'
data=json.loads((work/'content.revised.json').read_text(encoding='utf-8'))
doc=Document(work/'讲稿修订候选.docx')
paragraphs=[p.text for p in doc.paragraphs]
assert len(data['slides'])==12
for s in data['slides'][:10]+data['slides'][11:]:
    for paragraph in s['script']:
        assert paragraphs.count(paragraph)==1, ('Speech not synchronized',s['n'])
for qa in data['qa']:
    assert qa['q'] in paragraphs and qa['a'] in paragraphs
with ZipFile(work/'candidate.pptx') as z:
    for s in data['slides']:
        node=etree.fromstring(z.read(f"ppt/notesSlides/notesSlide{s['n']}.xml"))
        text=''.join(node.xpath('//*[local-name()="t"]/text()'))
        for paragraph in s['script']:
            assert paragraph in text, ('Notes not synchronized',s['n'])
    assert len([p for p in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',p)])==12
with ZipFile(work/'讲稿修订候选.docx') as z:
    comments=etree.fromstring(z.read('word/comments.xml'))
    assert len(comments)==5
    assert z.testzip() is None
print('Verified matching speech and notes, 7 Q&A answers, 5 revision comments, and 12 slides.')

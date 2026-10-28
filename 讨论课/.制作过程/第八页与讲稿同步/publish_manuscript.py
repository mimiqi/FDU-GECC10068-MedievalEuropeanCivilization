from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import json, posixpath
from lxml import etree
from docx import Document
ROOT=Path(r'D:\text\02_中世纪欧洲文明\讨论课')
WORK=ROOT/'.制作过程'/'第八页与讲稿同步'
source=ROOT/'交付文件'/'查理曼_12分钟汇报讲稿.docx'
candidate=WORK/'讲稿修订候选.docx'
final=ROOT/'交付文件'/'查理曼_12分钟汇报讲稿_同步修订版.docx'
ppt=ROOT/'交付文件'/'查理曼_政治治理视角_12分钟汇报_同步修订版.pptx'
validation=json.loads((WORK/'manuscript-validation.json').read_text(encoding='utf-8'))
assert sha256(source.read_bytes()).hexdigest()==validation['sourceSha256']
assert sha256(candidate.read_bytes()).hexdigest()==validation['candidateSha256']
data=json.loads((WORK/'content.json').read_text(encoding='utf-8'))
NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main','a':'http://schemas.openxmlformats.org/drawingml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def relation_map(z,path):
    return {r.get('Id'):r.get('Target') for r in etree.fromstring(z.read(path))}
def resolve(base,target):
    return target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join(base,target))
with ZipFile(ppt) as z:
    assert z.testzip() is None
    rels=relation_map(z,'ppt/_rels/presentation.xml.rels')
    ids=etree.fromstring(z.read('ppt/presentation.xml')).findall('p:sldIdLst/p:sldId',NS)
    for s,slideid in zip(data['slides'],ids):
        slidepath=resolve('ppt',rels[slideid.get('{'+NS['r']+'}id')])
        folder=posixpath.dirname(slidepath)
        sr=relation_map(z,folder+'/_rels/'+posixpath.basename(slidepath)+'.rels')
        note=next(v for v in sr.values() if 'notesSlide' in v)
        xml=etree.fromstring(z.read(resolve(folder,note)))
        text=''.join(xml.xpath('//p:sp[p:nvSpPr/p:nvPr/p:ph[@type="body"]]//a:t/text()',namespaces=NS))
        assert f"第{s['n']}页  {s['time']}" in text,(s['n'],'note heading')
        for paragraph in s['script']:assert paragraph in text,(s['n'],'script mismatch')
with ZipFile(candidate) as z:assert z.testzip() is None
with final.open('xb') as f:f.write(candidate.read_bytes())
assert final.read_bytes()==candidate.read_bytes()
record={'docx':str(final),'pptx':str(ppt),'docxSha256':sha256(final.read_bytes()).hexdigest(),'slides':11,'docPages':9,'mainSpeechHan':validation['han'],'mainSpeechMatchesFinalPpt':True,'renderedFileMatchesFinal':True}
(WORK/'配套交付记录.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'filesPublished':2,'finalNotesMatched':10,'docPages':9,'sha256':record['docxSha256']}))

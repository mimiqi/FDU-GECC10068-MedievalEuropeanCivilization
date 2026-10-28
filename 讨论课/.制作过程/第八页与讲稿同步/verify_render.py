from pathlib import Path
from PIL import Image, ImageChops
import pypdfium2 as pdfium
import json
WORK=Path(r'D:\text\02_中世纪欧洲文明\讨论课\.制作过程\第八页与讲稿同步')
changed=[]
for n in range(1,12):
    a=Image.open(WORK/'PPT预览'/'修改前'/f'slide-{n:02}.png').convert('RGB')
    b=Image.open(WORK/'PPT预览'/'修改后'/f'slide-{n:02}.png').convert('RGB')
    if a.size!=b.size or ImageChops.difference(a,b).getbbox():changed.append(n)
assert changed==[8],changed
pdf=pdfium.PdfDocument(WORK/'讲稿预览'/'讲稿.pdf')
report=[]
changed_doc=[]
for i in range(len(pdf)):
    page=pdf[i]
    out=WORK/'讲稿预览'/f'page-{i+1:02}.png'
    bitmap=page.render(scale=1.35).to_pil()
    if out.exists():
        old=Image.open(out).convert('RGB')
        if old.size!=bitmap.size or ImageChops.difference(old,bitmap.convert('RGB')).getbbox():changed_doc.append(i+1)
    bitmap.save(out)
    text=page.get_textpage().get_text_range()
    report.append({'page':i+1,'text':text})
(WORK/'rendered-text.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'changedVisibleSlides':changed,'docPages':len(pdf),'changedDocPagesSinceLastRender':changed_doc},ensure_ascii=False))

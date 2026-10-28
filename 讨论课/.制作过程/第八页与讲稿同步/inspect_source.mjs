import fs from 'node:fs/promises';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/第八页与讲稿同步';
const source=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_深化修订版.pptx';
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const {records}=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.position','elements.style','notes.text']});
await fs.writeFile(work+'/source-inspection.json',JSON.stringify(records,null,2));
await fs.mkdir(work+'/PPT预览/修改前',{recursive:true});
for(let i=0;i<p.slides.items.length;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
 await fs.writeFile(work+`/PPT预览/修改前/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
console.log(JSON.stringify({slides:p.slides.items.length,slide8:records[7],cover:records[0].elements.filter(e=>e.text).map(e=>e.text)}));

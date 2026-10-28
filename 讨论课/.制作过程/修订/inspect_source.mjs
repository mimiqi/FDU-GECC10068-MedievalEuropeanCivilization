import fs from 'node:fs/promises';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';
const root='D:/text/02_中世纪欧洲文明/讨论课';
const p=await PresentationFile.importPptx(await FileBlob.load(root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx'));
const r=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.position','elements.style'],fileName:root+'/.制作过程/修订/source-inspect.ndjson'});
await fs.writeFile(root+'/.制作过程/修订/anchors.json',JSON.stringify(r.records,null,2));
console.log(JSON.stringify({slides:p.slides.items.length,recordCount:r.recordCount,records:r.records?.slice(0,1),shapes:p.slides.items[0].shapes.items.map(x=>({id:x.id,name:x.name,text:String(x.text)}))}));

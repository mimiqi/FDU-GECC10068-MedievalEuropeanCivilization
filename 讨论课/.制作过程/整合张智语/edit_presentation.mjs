import fs from 'node:fs/promises';
import {PresentationFile, FileBlob} from '@oai/artifact-tool';

const root='D:/text/02_中世纪欧洲文明/讨论课', work=root+'/.制作过程/整合张智语';
const old=JSON.parse(await fs.readFile(root+'/.制作过程/修订/content.revised.json','utf8'));
const data=JSON.parse(await fs.readFile(work+'/content.revised.json','utf8'));
const p=await PresentationFile.importPptx(await FileBlob.load(root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx'));
const {records}=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.style','elements.position']});
await fs.writeFile(work+'/source-inspection.json',JSON.stringify(records,null,2));
function replace(n,before,after){
  if(before===after)return;
  const hits=records[n-1].elements.filter(e=>e.kind==='textbox'&&e.text===before);
  if(hits.length!==1)throw new Error(`Slide ${n}: expected exactly 1 target for ${before}, got ${hits.length}`);
  const rec=hits[0],shape=p.resolve(rec.id);
  shape.text=after;
  shape.text.style={...rec.style,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.12};
}
replace(7,old.slides[6].title,data.slides[6].title);
replace(7,old.slides[6].ref,data.slides[6].ref);
replace(7,'查理曼支持的工作','教育与书写的变化');
replace(7,'对日常政务的帮助','与帝国治理的联系');
replace(7,'拉丁文教育\n校订礼仪文本\n抄写和保存书籍','要求教会办学，学习拉丁文\n校订《圣经》与礼仪文本\n清晰易读的小写字体逐渐传播');
replace(7,'培养能读写的人\n便于传达敕令和书写文书\n加强不同地区之间的往来','培养能处理文书的教士\n减少抄写错误，便于阅读\n方便各地使用和传递文本');
replace(7,'这种做法有晚期罗马的先例，查理曼进一步支持教会教育。','教会教育和书写规范，有助于各地交流和处理政务。');
replace(10,old.slides[9].ref.replace(/^结论依据：/,''),data.slides[9].ref.replace(/^结论依据：/,''));
replace(10,'支持教育和书籍抄写\n加强不同地区的往来','支持教育和书籍抄写\n文化影响在分裂后延续');
replace(11,'113、146—149、206—210、229、234、239、271—272页（PDF＝书内页＋9）','113、146—149、206—210、229、231—237、271—272页（PDF＝书内页＋9）');
replace(11,'译者前言4页（PDF 10）；仅用于经济答辩的争议背景','290页（PDF 322）：文化延续；译者前言4页（PDF 10）：经济解释的争议');
for(const idx of [6,9,10]){
  const d=data.slides[idx];
  p.slides.items[idx].speakerNotes.text=`第${idx+1}页  ${d.time}\n\n${d.script.join('\n\n')}\n\n本页参考：${d.ref}\n页码说明：中文书默认书内页码，括号内为PDF顺序页码。Nelson采用PDF页码。`+
  (idx===6?'\n史料提示：789年的训令说明办学要求，不等于各地均已落实。巴尔贝罗第237页指出，没有直接证据证明小写字体出自查理曼本人的特别命令。':'')+
  (idx===9?'\n答辩补充：'+data.qa.at(-1).q+'\n'+data.qa.at(-1).a+'\n依据：'+data.qa.at(-1).ref:'')+
  (idx===10?'\n完整书目：\n'+data.bibliography.map(b=>b.details+b.title+'；'+b.pages).join('\n'):'');
}
await fs.mkdir(work+'/PPT预览',{recursive:true});
await (await PresentationFile.exportPptx(p)).save(work+'/candidate.pptx');
for(let i=0;i<12;i++){
  const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
  await fs.writeFile(work+`/PPT预览/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
console.log('Updated slides 7, 10, 11 and matching notes. Rendered all 12 slides.');

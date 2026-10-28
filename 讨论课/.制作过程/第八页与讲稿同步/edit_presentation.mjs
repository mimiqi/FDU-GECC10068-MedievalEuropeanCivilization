import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/第八页与讲稿同步';
const source=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_深化修订版.pptx';
const expected='2d2e6a8874785ed24b78d0e25f3731381230604d0d2e197ced4f908a761b75da';
if(createHash('sha256').update(await fs.readFile(source)).digest('hex')!==expected)throw new Error('Source changed.');
const p=await PresentationFile.importPptx(await FileBlob.load(source));
const {records}=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.position','elements.style','notes.text']});
const r8=records[7];
function edit(id,text,position={},style={}){
 const rec=r8.elements.find(e=>e.id===id);if(!rec)throw new Error('Missing '+id);
 const shape=p.resolve(id);shape.text=text;shape.position={...rec.position,...position};
 shape.text.style={...rec.style,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.15,...style};
}
function add(text,x,y,w,h,size,bold=false,color='#203443'){
 const sh=p.slides.items[7].shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 sh.text=text;sh.text.style={typeface:'Microsoft YaHei',fontSize:size,bold,color,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.15};
}
const title='查理曼统治的“欧洲性”体现在哪里';
edit('sh/m5gvmhwn','中世纪欧洲文明   /   地区之间的联系');
edit('sh/g72x4zyd',title);
edit('sh/tkby9kzm','政治重心',{left:64,top:188,width:206,height:78},{fontSize:31});
edit('sh/sjix0zy1','古罗马围绕地中海；查理曼帝国以莱茵河流域为重心，\n亚琛成为重要的政治中心。',{left:300,top:184,width:916,height:94},{fontSize:28});
edit('sh/id0fu50z','地方治理',{left:64,top:306,width:206,height:78},{fontSize:31});
edit('sh/o7ydoret','伯爵、地方领主和教会参与治理，\n使西欧、中欧部分地区与意大利北部服从同一位君主。',{left:300,top:302,width:916,height:94},{fontSize:28});
add('文化往来',64,424,206,78,31,true,'#9A443A');
add('不列颠的阿尔昆、比萨的彼得在宫廷任教；\n拉丁文书籍在宫廷与修道院之间流通。',300,420,916,94,28);
const rule=r8.elements.find(e=>e.id==='sh/je9g3ahk');
p.resolve(rule.id).position={...rule.position,top:535};
edit('sh/98rehwve','“欧洲性”体现为这些地区的政治与文化联系。',{top:551,height:43},{fontSize:29,bold:true,color:'#9A443A'});
add('判断君主性质，还要结合前述治理方式；这不等于现代欧洲已经形成。',64,594,1152,35,23,false,'#59646B');

const scripts8=[
 '前面讲了治理和教育，这些做法怎样支持“欧洲君主”的判断？先看统治重心。古罗马围绕地中海，查理曼帝国则以莱茵河流域为重心，亚琛成为重要的政治中心。但位置在欧洲，还不足以说明君主的性质。关键是他依靠伯爵、地方领主和教会，把西欧、中欧部分地区与意大利北部联系在同一统治之下。',
 '文化往来还超出了帝国疆界。艾因哈德记载，来自不列颠的阿尔昆和来自比萨的彼得，都曾在宫廷任教；拉丁文书籍也在宫廷和修道院之间流通。因此，这里的“欧洲性”，指这些地区在法兰克王权和拉丁基督教文化中的联系。结合前面介绍的治理办法，我们才更倾向于称他为新型欧洲君主。这不意味着他创造了现代欧洲，也不意味着他统治了整个欧洲。'
];
const ref8='巴尔贝罗《查理大帝：欧洲之父》113页（PDF 122，地理重心）、234页（PDF 243，书籍流通）；艾因哈德《查理大帝传》28页（PDF 37，彼得与阿尔昆）；布洛赫《封建社会》上卷266—269页（PDF 288—291，地方联系）；Nelson, King and Emperor, PDF 408页（亚琛作为新的权力中心）、515页（反对以“欧洲之父”为预定结论）。';
const times=['0:00—0:20','0:20—1:30','1:30—2:10','2:10—3:20','3:20—5:10','5:10—7:00','7:00—8:10','8:10—9:30','9:30—10:40','10:40—11:30'];
const slides=records.slice(0,10).map((rec,i)=>{
 const parts=rec.notes.text.split('\n\n本页参考：');
 const oldNumber=Number(parts[0].match(/^第(\d+)页/)[1]);
 const scripts=parts[0].split('\n\n').slice(1);
 const ref=parts[1].split('\n页码说明：')[0];
 return {n:i+1,oldNumber,time:times[i],title:i===0?'查理曼：罗马皇帝，还是新型欧洲君主？':rec.elements.find(e=>e.position?.top===78&&e.text)?.text,script:scripts,ref};
});
for(const n of [2,3])p.slides.items[n-1].speakerNotes.text=records[n-1].notes.text.replace(/^第\d+页[^\n]*/,`第${n}页  ${times[n-1]}`);
slides[7]={...slides[7],title,script:scripts8,ref:ref8};
p.slides.items[7].speakerNotes.text=`第8页  ${times[7]}\n\n${scripts8.join('\n\n')}\n\n本页参考：${ref8}\n页码说明：中文书使用书内页码，括号内为PDF顺序页码；Nelson使用PDF顺序页码。\n论证说明：地理重心只是比较的起点；君主性质的判断须结合第5—7页的治理方式。学者跨地区活动说明文化联系，不表示不列颠属于查理曼帝国。艾因哈德中译本称阿尔昆为“阿尔昆或阿尔宾”，此处统一采用阿尔昆。\n材料衔接：延续张元赫稿的地理比较，并把张智语稿中的文化教育论述落实到人员与文本的往来。`;
await fs.writeFile(work+'/content.json',JSON.stringify({source,sourceSha256:expected,slides},null,2));
await fs.mkdir(work+'/PPT预览/修改后',{recursive:true});
await (await PresentationFile.exportPptx(p)).save(work+'/candidate.pptx');
for(let i=0;i<p.slides.items.length;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
 await fs.writeFile(work+`/PPT预览/修改后/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
console.log(JSON.stringify({slides:p.slides.items.length,changedVisiblePages:[8],changedNotes:[2,3,8],han:slides.map(s=>(s.script.join('').match(/[\u4e00-\u9fff]/g)||[]).length)}));

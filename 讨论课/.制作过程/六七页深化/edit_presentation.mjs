import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';

const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/六七页深化';
const source=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_第2至4页修订.pptx';
const expected='0e8c4c2d76839143a909ae2ff235288c23582a126d47b5095f345dd74cd49287';
if(createHash('sha256').update(await fs.readFile(source)).digest('hex')!==expected)throw new Error('Source changed during revision.');
const p=await PresentationFile.importPptx(await FileBlob.load(source));
if(p.slides.items.length!==11)throw new Error('Expected 11 slides.');
const {records}=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.position','elements.style','notes.text']});
const C={ink:'#203443',red:'#9A443A',muted:'#59646B'};
function edit(n,before,after,position={},style={}){
 const hits=records[n-1].elements.filter(e=>e.text===before);
 if(hits.length!==1)throw new Error(`Slide ${n}: ambiguous target ${before}`);
 const rec=hits[0],shape=p.resolve(rec.id);
 shape.text=after;
 shape.position={...rec.position,...position};
 shape.text.style={...rec.style,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.12,...style};
}
function add(n,text,x,y,w,h,size=27,bold=false,color=C.ink){
 const shape=p.slides.items[n-1].shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 shape.text=text;
 shape.text.style={typeface:'Microsoft YaHei',fontSize:size,bold,color,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.12};
}

// Page 6: distinguish the relationships, then show why and how lords helped the king.
edit(6,'臣民向君主宣誓','臣民向君主宣誓',{top:173});
edit(6,'臣民对君主负有效忠义务','表达对君主的效忠义务\n并不使所有臣民都成为附庸',{top:234,height:76});
edit(6,'领主与附庸的关系','领主与附庸的关系',{top:173});
edit(6,'附庸承诺服务，领主给予保护','附庸承诺忠诚与服务\n领主承担保护责任',{top:234,height:76});
edit(6,'≠','≠',{top:225,height:58},{fontSize:40});
const rule=records[5].elements.find(e=>e.kind==='shape'&&e.position.top===337);
if(!rule)throw new Error('Page 6 separator missing.');
p.resolve(rule.id).position={...rule.position,top:318};
add(6,'为何借助领主',64,341,180,40,25,true,C.muted);
add(6,'官员少、交通不便，君主难以直接联系每个臣民。',260,339,950,45,27);
edit(6,'810年的规定（据布洛赫转述）','810年的规定',{left:64,top:394,width:180,height:40},{fontSize:25,bold:true,color:C.muted});
edit(6,'领主负责约束属从，使他们服从皇帝命令','领主要约束属从，使他们服从皇帝的命令。',{left:260,top:391,width:950,height:45},{fontSize:28,bold:false});
add(6,'具体怎样落实',64,445,180,40,25,true,C.muted);
add(6,'督促附庸出庭、服军役。领主参战时率领附庸，\n领主不在场时，由伯爵直接指挥。',260,443,950,83,26);
edit(6,'君主能借领主召集人员，也必须争取他们的合作。','王权借助个人忠诚组织军队，也更依赖领主的合作。',{top:558},{fontSize:28});

// Page 7: pair each concrete practice with its relevance to government and communication.
edit(7,'教育与书写的变化','具体做法',{left:300,top:166,width:460,height:36},{fontSize:23,color:C.muted});
edit(7,'与帝国治理的联系','与治理的联系',{left:810,top:166,width:406,height:36},{fontSize:23,color:C.muted});
const rows=[
 ['教会学校','789年训令要求教会设校\n教授语法、诗篇和历法计算','培养熟悉拉丁文的教士\n为教会与文书事务提供人员'],
 ['校订与抄写','阿尔昆在图尔校订《圣经》\n抄本通过宫廷与修道院传播','校订有误的文本\n方便各地取得、使用书籍'],
 ['书写规范','加洛林小写字体逐渐传播\n字形清晰，更容易辨认','便于跨地区阅读、抄写\n方便教会之间交流书面信息']
];
rows.forEach((row,i)=>{
 const y=217+i*113;
 add(7,row[0],64,y,205,78,29,true,C.red);
 if(i===0){
  edit(7,'要求教会办学，学习拉丁文\n校订《圣经》与礼仪文本\n清晰易读的小写字体逐渐传播',row[1],{left:300,top:y,width:460,height:94},{fontSize:27});
  edit(7,'培养能处理文书的教士\n减少抄写错误，便于阅读\n方便各地使用和传递文本',row[2],{left:810,top:y,width:406,height:94},{fontSize:25.5});
 }else{
  add(7,row[1],300,y,460,94,27);
  add(7,row[2],810,y,406,94,25.5);
 }
});
edit(7,'教会教育和书写规范，有助于各地交流和处理政务。','教会学校培养读写人才，书籍和书写规范方便各地交流。',{top:568,height:57},{fontSize:28});

const scripts={
 6:[
  '上一页说的是怎样监督地方官。这里的问题是，君主靠谁把人组织起来。先分清两种关系：臣民对君主的效忠，是一种政治义务；附庸关系则发生在特定的领主和属从之间，附庸承诺忠诚和服务，领主承担保护责任。两者可以重叠，但不能把所有宣誓的臣民都叫作附庸。',
  '为什么要借助领主？布洛赫指出，当时官员少，交通又不方便，君主很难直接联系到每个人。810年的规定要求领主管束自己的属从，使他们服从皇帝命令。这把已有的个人关系与王权的要求联系起来。具体来说，领主要督促附庸出庭、履行军役；如果领主亲自参战，附庸由他率领，领主不在场时，则由伯爵直接指挥。',
  '所以，这种关系能够帮助查理曼组织军队、落实义务，说明他在尝试利用地方力量加强统治。不过，命令能否落实也取决于领主是否合作。我们不能因此认定国家权力已被私人关系完全取代，也不能把当时的伯爵都看成后来的世袭封建领主。'
 ],
 7:[
  '教育和书写改革可以用三件具体的事来说明。第一，789年训令要求教会办学，教授语法、诗篇和历法计算等知识。目的首先是让教士更准确地理解信仰，同时也培养了能读写拉丁文的人。第二，阿尔昆在图尔校订《圣经》，抄本通过宫廷和修道院传播，使不同地区更容易取得、使用较可靠的文本。',
  '第三，加洛林小写字体逐渐传播，字形清晰，方便阅读和抄写。它没有马上取代所有地方字体，也不能算查理曼个人的发明。这些改革与治理的联系，在于学校培养读写人员，书籍和书写规范方便交流，教会也能提供处理文书的人才。不过，办学训令并不意味着各地都已落实。'
 ]
};
const sources={
 6:'布洛赫《封建社会》上卷266—269页（PDF 288—291）：267页解释官员与交通条件的限制，并转述810年的规定；268页说明出庭、军役和战场指挥关系。臣民效忠与附庸关系的区分，另参巴尔贝罗《查理大帝：欧洲之父》146—149页（PDF 155—158）。',
 7:'巴尔贝罗《查理大帝：欧洲之父》231—237页（PDF 240—246）：232—233页为办学要求及落实情况；231、234页为阿尔昆的《圣经》校订及抄本传播；236—237页为加洛林小写字体。'
};
for(const n of [6,7]){
 const time=records[n-1].notes.text.split('\n')[0];
 p.slides.items[n-1].speakerNotes.text=time+'\n\n'+scripts[n].join('\n\n')+'\n\n本页参考：'+sources[n]+'\n页码说明：中文书使用书内页码，括号内为PDF顺序页码。'+
 (n===6?'\n答辩补充：军役和出庭的具体安排据布洛赫第268页概括，不将这些细节都当作810年法令的原文。一般臣民的效忠誓言不自动构成附庸关系，伯爵的公共职务也不能直接等同于私人依附。':'\n答辩补充：这些材料首先描述宗教教育与书籍生产；“有助于政务沟通”是结合读写人才和跨地区文本交流作出的分析，不能据此断言所有官府文书都已采用同一种字体。');
}
const content={source,sourceSha256:expected,scripts,sources};
await fs.writeFile(work+'/content.json',JSON.stringify(content,null,2));
await fs.mkdir(work+'/PPT预览/修改后',{recursive:true});
await (await PresentationFile.exportPptx(p)).save(work+'/candidate.pptx');
for(let i=0;i<p.slides.items.length;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
 await fs.writeFile(work+`/PPT预览/修改后/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
const count=text=>(text.match(/[\u4e00-\u9fff]/g)||[]).length;
console.log(JSON.stringify({slides:p.slides.items.length,speechHan:{6:count(scripts[6].join('')),7:count(scripts[7].join(''))}}));

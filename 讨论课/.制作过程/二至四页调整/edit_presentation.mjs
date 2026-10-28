import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';

const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/二至四页调整';
const input=root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx';
const expectedSourceSha='9152837795e5c214003f3d56bb822eb720f4c4f971f9e8282f907507b8f779ef';
if(createHash('sha256').update(await fs.readFile(input)).digest('hex')!==expectedSourceSha)throw new Error('Source changed during this edit.');
const p=await PresentationFile.importPptx(await FileBlob.load(input));
if(p.slides.items.length!==11)throw new Error('Unexpected source slide count.');
const {records}=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.position','elements.style','notes.text']});
const C={ink:'#203443',red:'#9A443A',muted:'#59646B'};
function edit(n,before,after,position={},style={}){
 const hits=records[n-1].elements.filter(e=>e.text===before && !(n===2 && ['01','02','03'].includes(before) && e.position.left!==64));
 if(hits.length!==1)throw new Error(`Slide ${n}: ambiguous anchor ${before}`);
 const rec=hits[0],shape=p.resolve(rec.id);
 shape.text=after;
 shape.position={...rec.position,...position};
 shape.text.style={...rec.style,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.12,...style};
}
function addText(n,text,x,y,w,h,size,color){
 const shape=p.slides.items[n-1].shapes.add({geometry:'textbox',position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
 shape.text=text;
 shape.text.style={typeface:'Microsoft YaHei',fontSize:size,color,bold:true,autoFit:'none',wrap:'square',insets:0,lineSpacing:1.12};
}

// Page 2: identify the inherited arrangements and their concrete problems.
addText(2,'原有的安排',170,161,420,29,22,C.muted);
addText(2,'面临的困难',635,161,565,29,22,C.muted);
const page2=[
 ['01','司法','怎样管理地方？','伯爵主持地方司法','君主怎样让伯爵和领主听命','伯爵可能徇私，君主难以直接监督'],
 ['02','出庭','怎样约束伯爵？','自由民参加公共法庭','谁来审判，谁来监督伯爵','出庭过于频繁，加重普通人的负担'],
 ['03','动员','怎样召集军队？','领主与属从的关系','领主怎样为君主召集人员','官员有限，传令和征兵需借助领主']
];
page2.forEach((row,i)=>{
 const y=205+113*i;
 edit(2,row[0],row[1],{top:y+2,width:84},{fontSize:29});
 edit(2,row[2],row[3],{top:y},{fontSize:32});
 edit(2,row[4],row[5],{top:y+4});
});
for(const [oldTop,newTop] of [[260,289],[363,402]]){
 const rule=records[1].elements.find(e=>e.kind==='shape'&&e.position.top===oldTop);
 if(!rule)throw new Error('Expected existing row separator.');
 p.resolve(rule.id).position={...rule.position,top:newTop};
}
edit(2,'他调整了原有的统治办法，回应当时的实际问题。','他加强对伯爵的监督，减轻出庭负担，并借领主召集人员。',{top:550},{fontSize:28});

// Page 3: retain editable elements while replacing oversized date columns with evidence rows.
edit(3,'罗马传统在查理曼时代继续发挥作用','查理曼继承了哪些罗马传统');
const page3=[
 ['800','罗马帝国的名号','在罗马接受加冕','800年在罗马加冕，801年开始使用帝国头衔','以皇帝身份\n统治不同地区','文书称他为“治理罗马帝国的皇帝”'],
 ['亚琛','皇权与教会事务','修建礼拜堂','794年召集法兰克福会议，参与裁断教义争议','从罗马和拉文纳\n运来柱子及大理石','君主召集教会会议，在晚期罗马帝国已有先例'],
 ['812','具体案件中的罗马法','东方承认皇帝地位','加冕后，按罗马法审判袭击教皇的人','东西方两个皇帝\n达成政治妥协','罗马法用于具体案件，各地法律仍然并存']
];
page3.forEach((row,i)=>{
 const y=184+120*i;
 edit(3,row[0],row[1],{left:64,top:y,width:300,height:74},{fontSize:30,color:C.red,bold:true});
 edit(3,row[2],row[3],{left:390,top:y-2,width:826,height:45},{fontSize:29,color:C.ink,bold:false});
 edit(3,row[4],row[5],{left:390,top:y+42,width:826,height:65},{fontSize:26,color:C.muted,bold:false});
});
edit(3,'他借用罗马的称号、仪式和建筑，确立自己的皇帝身份。','罗马传统既涉及皇帝身份，也体现在教会事务和司法实践中。',{top:556},{fontSize:28});

// Page 4: use region-first parallel headings and lead directly into the judicial evidence.
edit(4,'774：征服伦巴第王国','伦巴第：征服与法律延续');
edit(4,'伦巴第人仍保留自己的法律和一些政治传统。','774年征服伦巴第王国后，当地仍保留自己的法律和一些政治传统。');
edit(4,'君主怎样让各地都执行自己的命令？','地方伯爵掌管审判。\n君主怎样防止他们徇私滥权？');

const scripts={
 2:['判断查理曼的性质，既要看皇帝称号，也要看他怎样治理地方。他接手的法兰克社会，已有伯爵主持地方司法、自由民参加公共法庭，以及领主与属从之间的关系。困难也很具体：伯爵可能徇私，普通人频繁出庭负担很重，君主又缺少人手直接组织各地人员。接下来我们看，他怎样加强监督、调整出庭义务，并借领主执行命令。'],
 3:['先看罗马传统。800年加冕以后，查理曼从801年起使用包含“治理罗马帝国的皇帝”的文书称号。这表明，他希望以罗马帝国的继承者自居。其次，是君主介入教会事务。晚期罗马皇帝曾召集教会会议；查理曼也在794年召集法兰克福会议，参与教义争议的处理。','第三，是罗马法的实际运用。Nelson记载，加冕后，查理曼按罗马法审理袭击教皇的人。因此，罗马传统在他的统治中有具体内容。不过，某些案件援用罗马法，不等于整个帝国实行统一的罗马法。各地法律继续并存，正是他实际统治时必须面对的条件。'],
 4:['查理曼最显眼的功绩，是通过战争扩大法兰克统治。774年，他征服伦巴第王国，把意大利北部纳入自己的统治。但征服之后，伦巴第人仍保留了自己的法律和一些政治传统。服从同一位君主，并不代表各地马上采用同一套制度。','萨克森战争则显示了征服的另一面。艾因哈德记载，查理曼迁走了大批萨克森人，并要求归顺者接受基督教。这些战争扩大了帝国，也伴随着强迫和暴力。领土扩大以后，君主必须更多地依靠地方官员。伯爵掌管审判，如果他们徇私受贿，君主要求的公正就很难实现。因此，我们先看查理曼怎样监督地方司法，再看他怎样借领主组织人员。']
};
const sources={
 2:'巴尔贝罗《查理大帝：欧洲之父》206—210页（PDF 215—219）；布洛赫《封建社会》上卷266—269页（PDF 288—291）。',
 3:'Nelson, King and Emperor, PDF 251页（君士坦丁召集325年尼西亚会议）、405页（按罗马法审理袭击教皇者）、407—408页（801年文书称号）；巴尔贝罗《查理大帝：欧洲之父》238—239页（PDF 247—248，794年法兰克福会议）、210页（PDF 219，多种法律并存）。',
 4:'艾因哈德《查理大帝传》10、12页（PDF 19、21）；Nelson, King and Emperor, PDF 175页；巴尔贝罗《查理大帝：欧洲之父》206—210页（PDF 215—219）；地图：巴尔贝罗书前地图（PDF 5）。'
};
for(const n of [2,3,4]){
 const time=records[n-1].notes.text.split('\n')[0];
 p.slides.items[n-1].speakerNotes.text=time+'\n\n'+scripts[n].join('\n\n')+'\n\n本页参考：'+sources[n]+'\n页码说明：中文书使用书内页码，括号内为PDF顺序页码；Nelson统一使用PDF顺序页码。'+
 (n===3?'\n答辩补充：亚琛礼拜堂借用罗马与拉文纳的建筑材料，见艾因哈德《查理大帝传》28页（PDF 37）；812年东方承认其皇帝地位，见Nelson, PDF 480—481页。宗教会议的比较说明统治职责上的延续，并非断言法兰克福会议与罗马帝国会议具有完全相同的地位。':'');
}
const oldIntro='另一个办法，是借助领主和附庸之间的关系。';
if(!records[5].notes.text.includes(oldIntro))throw new Error('Page 6 note anchor is missing.');
p.slides.items[5].speakerNotes.text=records[5].notes.text.replace(oldIntro,'监督地方官员，只解决了一部分问题。君主还要组织人员服军役、执行命令，这就涉及领主与附庸之间的关系。');
edit(11,'113、146—149、206—210、229、231—237、271—272页','113、146—149、206—210、229、231—239、271—272页');
edit(11,'PDF 175、204、405—408、480—481、497—498、515页','PDF 175、204、251、405—408、480—481、497—498、515页');
p.slides.items[10].speakerNotes.text=records[10].notes.text.replace('每页页脚保留对应书名及定位页码。','各页演讲者备注保留对应书名及定位页码。').replace('231—237、271—272','231—239、271—272').replace('PDF 175、204、405—408','PDF 175、204、251、405—408');
await fs.writeFile(work+'/content.json',JSON.stringify({sourceSha256:expectedSourceSha,scripts,sources},null,2));
await fs.mkdir(work+'/PPT预览/修改后',{recursive:true});
await (await PresentationFile.exportPptx(p)).save(work+'/candidate.pptx');
for(let i=0;i<11;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
 await fs.writeFile(work+`/PPT预览/修改后/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
}
console.log('Edited 11-slide user version: pages 2, 3, 4, bibliography, and related speaker notes.');

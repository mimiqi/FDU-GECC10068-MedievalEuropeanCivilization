import fs from 'node:fs/promises';
import {PresentationFile,FileBlob} from '@oai/artifact-tool';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/修订';
const old=JSON.parse(await fs.readFile(root+'/.制作过程/content.json','utf8'));
const data=JSON.parse(await fs.readFile(work+'/content.revised.json','utf8'));
const p=await PresentationFile.importPptx(await FileBlob.load(root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx'));
const inspected=await p.inspect({select:'$.slides[*]',include:['id','index','elements.id','elements.kind','elements.text','elements.style','elements.position']});
const records=inspected.records;
const swaps={
 1:[['从政治治理看罗马传统的延续与重组',data.subtitle]],
 2:[['公共职权与地方精英如何衔接','君主怎样让伯爵和领主听命'],['正义要求如何进入地方审判','谁来审判，谁来监督伯爵'],['个人依附怎样服务君主目标','领主怎样为君主召集人员'],['“新型”：重组统治资源；不要求每一项制度都从零开始。','新意体现在怎样调整旧有办法，适应当时的统治条件。']],
 3:[['801','亚琛'],['罗马加冕','在罗马接受加冕'],['帝国称号与\n基督教普遍秩序','以皇帝身份\n统治不同地区'],['叠加的头衔','修建礼拜堂'],['皇帝，同时仍是\n法兰克与伦巴第国王','从罗马和拉文纳\n运来柱子及大理石'],['帝国权威得到承认，\n称号与地位仍有区别','东西方两个皇帝\n达成政治妥协'],['皇帝身份具有实质意义；但称号还不能独自解释治理。','他借用罗马的称号、仪式和建筑，确立自己的皇帝身份。']],
 4:[['政治纳入后，地方的法律与传统仍然存在。','伦巴第人仍保留自己的法律和一些政治传统。'],['军事整合的成就，伴随人口迁徙与征服代价。','战争扩大了帝国，也伴随人口迁徙和强迫改宗。'],['核心难题：怎样统治差异很大的地区？','君主怎样让各地都执行自己的命令？']],
 5:[['针对伯爵徇私、受贿与弱者求诉无门','检查伯爵怎样办案，干预徇私和受贿'],['提高熟悉法律和处理案件的可能性','让经常参与审判的人积累法律知识和经验'],['由特定人员承担更多审判事务','普通人减少到场次数，由指定人员承担更多事务'],['加强公共司法的尝试；执行仍受地方权力与选任条件限制。','伯爵仍有很大影响，改革的落实取决于地方上的配合。']],
 6:[['广泛的臣民效忠','臣民向君主宣誓'],['表达君主对臣民的公共权威','臣民对君主负有效忠义务'],['特定的附庸关系','领主与附庸的关系'],['领主与属从之间的依附和服务','附庸承诺服务，领主给予保护'],['810年规定的动员逻辑','810年的规定（据布洛赫转述）'],['君主命令  →  领主约束属从  →  军役与司法动员','领主负责约束属从，使他们服从皇帝命令'],['王权借助个人关系传令，也因此更加依赖地方有力者。','君主能借领主召集人员，也必须争取他们的合作。']],
 7:[['具体工作','查理曼支持的工作'],['治理意义','对日常政务的帮助'],['教士教育与学习\n礼仪文本校订\n手稿抄写与保存','拉丁文教育\n校订礼仪文本\n抄写和保存书籍'],['识字人员\n可跨地区传递的拉丁文规范\n教会组织之间的联系','培养能读写的人\n便于传达敕令和书写文书\n加强不同地区之间的往来'],['宗教目的与政治作用相互交织；这种关联有晚期罗马的前史。','教会参与政治有罗马传统，教育改革让这种合作更便于开展。']],
 8:[['新型','古罗马'],['公共职权 ＋ 个人依附 ＋ 教会网络\n在法兰克政治条件下重新组合','围绕地中海展开\n领土跨越欧洲、北非和亚洲沿岸'],['欧洲性','查理曼'],['立足法兰克，连接西欧、中欧部分地区与意大利\n形成跨地区的政治与拉丁基督教联系','以莱茵河流域为重心\n连接西欧、中欧部分地区与意大利北部'],['历史性的区域联系，不等同于现代欧洲认同或今日欧洲边界。','“欧洲性”指这些地区的政治与文化联系，不能套用今天的欧洲边界。']],
 9:[['更能解释实际统治的称谓：新型欧洲君主。','罗马皇帝的理想，在法兰克社会中有了不同的实现方式。']],
 10:[['扩展与整合','军事征服'],['军事征服\n多地区政治纳入','扩大领土\n把多个地区纳入统治'],['司法与动员','地方治理'],['监督地方审判\n组织人员与服务','监督伯爵审判\n借领主召集人员'],['教会与文化','文化教育'],['教育与文本整理\n维系跨地区联系','支持教育和书籍抄写\n加强不同地区的往来'],['帝国认同上：罗马皇帝\n治理结构上：新型欧洲君主','他继承了罗马的皇帝理想\n也依靠地方贵族和教会治理帝国']],
 11:[['10—12、28—30页（PDF 19—21、37—39）','10—12、28—31页（PDF 19—21、37—40）'],['146—149、206—210、229、234、239、271—272页（PDF＝书内页＋9）；地图PDF 5','113、146—149、206—210、229、234、239、271—272页（PDF＝书内页＋9）'],['PDF 175、204、405—408、480—481、515页','PDF 175、204、405—408、480—481、497—498、515页']],
 12:[['庄园经济能直接证明“新型”吗？','庄园经济是否完全封闭'],['农业与庄园占重要地位  ≠  经济完全封闭','庄园生产与市场交换可以并存'],['可确认的事实','书中的证据'],['可使用的判断','对我们论证的帮助'],['解释资源组织的条件\n不能单独证明“完整封建制度”','说明帝国如何取得资源\n不足以单独判断君主的性质'],['皮朗的解释具有影响力，也存在争议。经济只作补充，不替代政治治理证据。','皮朗的解释仍有争议，经济材料在这里作为补充。']]
};
swaps[2].push(['谁执行命令？','怎样管理地方？'],['司法怎样落实？','怎样约束伯爵？'],['力量怎样动员？','怎样召集军队？']);
swaps[2].find(x=>x[0].startsWith('“新型”：'))[1]='他调整了原有的统治办法，回应当时的实际问题。';
swaps[5].push(['稳定人员','固定人员'],['调整负担','减轻负担']);
swaps[7].find(x=>x[0].startsWith('宗教目的'))[1]='这种做法有晚期罗马的先例，查理曼进一步支持教会教育。';
const labels=['小组讨论汇报','讨论思路','罗马传统','军事征服','地方司法','附庸关系','文化与政务','地理重心','比较与判断','结论','附录 · 书目','附录 · 答辩备用'];
function replace(n,before,after){
 if(before===after)return;
 const hits=records[n-1].elements.filter(e=>e.kind==='textbox'&&e.text===before);
 if(hits.length!==1)throw new Error(`Slide ${n}: expected exactly 1 target for ${before}, got ${hits.length}`);
 const rec=hits[0],shape=p.resolve(rec.id);
 shape.text=after;
 shape.text.style={...rec.style,autoFit:'none',wrap:'square',insets:0,verticalAlignment:'top',lineSpacing:1.12};
}
for(let i=0;i<12;i++){
 if(i>0&&i<10)replace(i+1,old.slides[i].title,data.slides[i].title);
 replace(i+1,old.slides[i].ref,data.slides[i].ref.replace(/^结论依据：/,''));
 replace(i+1,records[i].elements.find(e=>e.text?.startsWith('中世纪欧洲文明   /   ')).text,'中世纪欧洲文明   /   '+labels[i]);
 for(const [before,after] of swaps[i+1]||[])replace(i+1,before,after);
 const d=data.slides[i];
 p.slides.items[i].speakerNotes.text=`第${i+1}页  ${d.time}\n\n${d.script.join('\n\n')}\n\n本页参考：${d.ref}\n页码说明：中文书默认书内页码，括号内为PDF顺序页码。Nelson采用PDF页码。`+
 (i===3?'\n地图：巴尔贝罗《查理大帝：欧洲之父》书前地图，PDF第5页，保留原图地理内容。':'')+
 (i===2?'\n补充依据：吸收张元赫稿中的建筑比较，具体核对艾因哈德第28页。':'')+
 (i===7?'\n补充依据：吸收张元赫稿中的地理重心比较，具体核对巴尔贝罗第113页。':'')+
 (i===10?'\n完整书目：\n'+data.bibliography.map(b=>b.details+b.title+'；'+b.pages).join('\n'):'')+
 (i===11?'\n此页仅用于答辩，不计入主讲。':'');
}
const tableRec=records[8].elements.find(e=>e.kind==='table');
if(!tableRec)throw new Error('Editable evidence table missing');
const table=p.resolve(tableRec.id);
const values=[['比较内容','继承的罗马传统','查理曼的具体做法'],['皇帝身份','皇帝称号与基督教帝国理想','接受加冕，同时保留法兰克与伦巴第王权'],['地方司法','君主有维护正义的责任','巡察使监督伯爵，固定人员辅助审判'],['政务与征兵','教会参与帝国事务','借领主召集人员，依靠教士处理文书']];
for(let r=0;r<4;r++)for(let c=0;c<3;c++){
 const cell=table.getCell(r,c);cell.value=values[r][c];
 cell.text.style={typeface:'Microsoft YaHei',fontSize:r===0?26:25,bold:r===0||c===0,color:r===0?'#FFFFFF':'#203443',autoFit:'none',verticalAlignment:'middle',wrap:'square'};
}
await fs.mkdir(work+'/PPT预览',{recursive:true});
await (await PresentationFile.exportPptx(p)).save(work+'/candidate.pptx');
for(let i=0;i<12;i++){
 const blob=await p.export({slide:p.slides.items[i],format:'png',scale:1.25});
 await fs.writeFile(work+`/PPT预览/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await blob.arrayBuffer()));
 const layout=await p.slides.items[i].export({format:'layout'});
 await fs.writeFile(work+`/PPT预览/slide-${String(i+1).padStart(2,'0')}.json`,await layout.text());
}
console.log('Edited and rendered 12 slides; reused original shapes, map, table, and layout.');

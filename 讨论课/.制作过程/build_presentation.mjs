import fs from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { Presentation, PresentationFile } from '@oai/artifact-tool';

const root = 'D:/text/02_中世纪欧洲文明/讨论课';
const build = root + '/.制作过程';
const data = JSON.parse(await fs.readFile(build + '/content.json', 'utf8'));
const C = {paper:'#F5F1E8',ink:'#203443',red:'#9A443A',muted:'#59646B',rule:'#C9C4B9',white:'#FFFFFF',pale:'#E8E3D8'};
const FONT = 'Microsoft YaHei';
const p = Presentation.create({slideSize:{width:1280,height:720}});
let serial=0;
function text(s,str,x,y,w,h,size=28,opt={}) {
  const sh=s.shapes.add({geometry:'textbox',name:opt.name||`text-${++serial}`,position:{left:x,top:y,width:w,height:h},fill:'none',line:{fill:'none',width:0}});
  sh.text=str;
  sh.text.style={typeface:FONT,fontSize:size,bold:opt.bold??false,color:opt.color||C.ink,alignment:opt.align||'left',verticalAlignment:'top',autoFit:'none',wrap:'square',insets:0,lineSpacing:1.12};
  return sh;
}
function rule(s,y,x=64,w=1152) {
  s.shapes.add({geometry:'line',name:`rule-${++serial}`,position:{left:x,top:y,width:w,height:0},fill:'none',line:{fill:C.rule,width:1}});
}
function base(i,kicker,title) {
  const s=p.slides.add(); s.background.fill=C.paper;
  text(s,`中世纪欧洲文明   /   ${kicker}`,64,27,1030,27,17,{color:C.muted});
  text(s,String(i).padStart(2,'0'),1152,26,64,29,18,{align:'right',color:C.red});
  if(title) text(s,title,64,78,1152,72,44,{bold:true,name:`heading-${i}`});
  rule(s,632);
  text(s,data.slides[i-1].ref,64,649,1152,49,16.3,{color:C.muted});
  const info=data.slides[i-1];
  s.speakerNotes.text=`第${i}页｜${info.time}\n${info.script.join('\n\n')}\n\n本页参考：${info.ref}\n\n页码说明：中文书默认书内页码；括号中PDF为文件顺序页码，从1计数。Nelson统一使用PDF页码。\n${i===4?'地图来自课程所供巴尔贝罗《查理大帝：欧洲之父》PDF第5页；仅裁去白边，无地理内容修改。\n':''}${i===8?'本页“创新”与“欧洲性”的关联是本组基于上述材料的分析，不作为史料直接引语。\n':''}${i===12?'此页为答辩备用，不计入12分钟主讲。\n':''}`;
  return s;
}
function heading(s,str,x,y,w=1100,size=30){return text(s,str,x,y,w,51,size,{bold:true});}
function body(s,str,x,y,w,h=120,size=27){return text(s,str,x,y,w,h,size);}
function takeaway(s,str,y=552,size=30){return text(s,str,64,y,1152,66,size,{bold:true,color:C.red});}

// 01 — Cover: a clear thesis and the required member placeholders.
{
 const s=base(1,'小组讨论汇报',null);
 text(s,'政治治理视角',64,104,1080,38,24,{color:C.red,bold:true});
 text(s,data.slides[0].title,64,171,1152,170,64,{bold:true,name:'cover-title'});
 text(s,data.subtitle,64,369,1110,60,31);
 text(s,'小班：[待填写]    小组：[待填写]',64,471,1140,42,24);
 text(s,'成员及分工：[姓名—分工]  /  [姓名—分工]  /  [姓名—分工]\n                         [姓名—分工]  /  [姓名—分工]  /  [姓名—分工]',64,526,1152,72,22,{color:C.muted});
}
// 02 — Criteria, deliberately confined to governance.
{
 const s=base(2,'论证标准',data.slides[1].title);
 const ys=[186,289,392];
 [['01','谁执行命令？','公共职权与地方精英如何衔接'],['02','司法怎样落实？','正义要求如何进入地方审判'],['03','力量怎样动员？','个人依附怎样服务君主目标']].forEach((a,i)=>{
   text(s,a[0],64,ys[i],75,52,35,{color:C.red});
   heading(s,a[1],170,ys[i],420,34);
   body(s,a[2],635,ys[i]+4,565,62,27);
   if(i<2) rule(s,ys[i]+74,170,1046);
 });
 takeaway(s,'“新型”：重组统治资源；不要求每一项制度都从零开始。',534,29);
}
// 03 — Do not make Roman legitimacy a straw man.
{
 const s=base(3,'承认延续',data.slides[2].title);
 const xs=[64,462,860];
 const columns=[['800','罗马加冕','帝国称号与\n基督教普遍秩序'],['801','叠加的头衔','皇帝，同时仍是\n法兰克与伦巴第国王'],['812','东方承认皇帝地位','帝国权威得到承认，\n称号与地位仍有区别']];
 columns.forEach((a,i)=>{
   text(s,a[0],xs[i],185,350,93,66,{color:C.red,bold:true});
   heading(s,a[1],xs[i],305,350,28);
   body(s,a[2],xs[i],368,350,122,27);
 });
 takeaway(s,'皇帝身份具有实质意义；但称号还不能独自解释治理。',552,30);
}
// 04 — Actual map from the supplied book, with all original legends retained.
{
 const s=base(4,'功绩与问题',data.slides[3].title);
 const bytes=new Uint8Array(await fs.readFile(build+'/assets/查理曼疆域_原书地图.png'));
 s.images.add({blob:bytes,contentType:'image/png',alt:'巴尔贝罗书前加洛林帝国地图，显示771年法兰克王国、771—814年征服地区及统治范围',fit:'contain',position:{left:64,top:161,width:458,height:451}});
 heading(s,'774：征服伦巴第王国',572,184,640,31);
 body(s,'政治纳入后，地方的法律与传统仍然存在。',572,246,634,90,28);
 heading(s,'萨克森：战争与强制改宗',572,354,640,31);
 body(s,'军事整合的成就，伴随人口迁徙与征服代价。',572,416,634,90,28);
 text(s,'核心难题：怎样统治差异很大的地区？',572,541,640,76,29,{bold:true,color:C.red});
}
// 05 — Three concrete institutional measures, with limits visible.
{
 const s=base(5,'核心证据一 · 司法',data.slides[4].title);
 const rows=[
  ['监督地方','巡察使介入司法','针对伯爵徇私、受贿与弱者求诉无门'],
  ['稳定人员','固定审判辅助人员（scabini）','提高熟悉法律和处理案件的可能性'],
  ['调整负担','809：减轻普通自由民出庭义务','由特定人员承担更多审判事务']
 ];
 rows.forEach((a,i)=>{
  const y=182+i*112;
  text(s,a[0],64,y+2,175,51,28,{color:C.red,bold:true});
  heading(s,a[1],263,y,940,30);
  body(s,a[2],263,y+49,940,50,26);
  if(i<2) rule(s,y+96,263,953);
 });
 takeaway(s,'加强公共司法的尝试；执行仍受地方权力与选任条件限制。',546,28);
}
// 06 — Distinguish general fidelity from particular vassal relationships.
{
 const s=base(6,'核心证据二 · 附庸',data.slides[5].title);
 heading(s,'广泛的臣民效忠',64,179,530,31);
 body(s,'表达君主对臣民的公共权威',64,242,530,76,27);
 heading(s,'特定的附庸关系',670,179,530,31);
 body(s,'领主与属从之间的依附和服务',670,242,530,76,27);
 text(s,'≠',586,210,80,78,46,{align:'center',color:C.red});
 rule(s,337);
 text(s,'810年规定的动员逻辑',64,369,1140,40,23,{color:C.muted});
 text(s,'君主命令  →  领主约束属从  →  军役与司法动员',64,426,1152,70,36,{bold:true});
 takeaway(s,'王权借助个人关系传令，也因此更加依赖地方有力者。',549,29);
}
// 07 — Church and cultural work are evaluated through the political lens.
{
 const s=base(7,'功绩与支撑条件',data.slides[6].title);
 heading(s,'具体工作',64,185,500,30);
 body(s,'教士教育与学习\n礼仪文本校订\n手稿抄写与保存',64,260,490,226,32);
 heading(s,'治理意义',670,185,500,30);
 body(s,'识字人员\n可跨地区传递的拉丁文规范\n教会组织之间的联系',670,260,550,226,32);
 takeaway(s,'宗教目的与政治作用相互交织；这种关联有晚期罗马的前史。',548,28);
}
// 08 — Explicitly connect innovation to a bounded meaning of Europe.
{
 const s=base(8,'由史实走向判断',data.slides[7].title);
 text(s,'新型',64,194,240,76,48,{color:C.red,bold:true});
 body(s,'公共职权 ＋ 个人依附 ＋ 教会网络\n在法兰克政治条件下重新组合',324,189,879,118,32);
 rule(s,344);
 text(s,'欧洲性',64,384,240,76,48,{color:C.red,bold:true});
 body(s,'立足法兰克，连接西欧、中欧部分地区与意大利\n形成跨地区的政治与拉丁基督教联系',324,381,879,137,29);
 text(s,'历史性的区域联系，不等同于现代欧洲认同或今日欧洲边界。',64,560,1152,53,25,{color:C.muted});
}
// 09 — Native editable table owns the central evidence comparison.
{
 const s=base(9,'证据汇总',data.slides[8].title);
 const vals=[
 ['比较维度','罗马延续','查理曼时期的组织方式'],
 ['政治权威','帝国称号、基督教普遍秩序','皇帝头衔与法兰克、伦巴第王权叠加'],
 ['司法治理','君主主持正义的理想','巡察使监督、固定辅助人员、地方法律并存'],
 ['动员与沟通','公共义务、基督教政治传统','附庸联系与教会网络共同支撑王权']
 ];
 const tb=s.tables.add({rows:4,columns:3,left:64,top:182,width:1152,height:310,columnWidths:[180,385,587],values:vals});
 tb.styleOptions={headerRow:true,bandedRows:false};
 tb.borders.assign({fill:C.rule,width:1,style:'solid'});
 tb.cells.block({row:0,column:0,rowCount:4,columnCount:3}).assign({margins:{left:15,right:15,top:12,bottom:12},anchor:'center'});
 for(let r=0;r<4;r++) for(let c=0;c<3;c++){
  const cell=tb.getCell(r,c);cell.fill=r===0?C.ink:(c===0?C.pale:C.paper);
  cell.text.style={typeface:FONT,fontSize:r===0?26:25,bold:r===0||c===0,color:r===0?C.white:C.ink,autoFit:'none',verticalAlignment:'middle',wrap:'square'};
 }
 tb.rows[0].height=58;
 tb.rows[1].height=84;tb.rows[2].height=84;tb.rows[3].height=84;
 takeaway(s,'更能解释实际统治的称谓：新型欧洲君主。',545,33);
}
// 10 — A pointed conclusion including achievements and their limits.
{
 const s=base(10,'结论',data.slides[9].title);
 const xs=[64,462,860];
 [['扩展与整合','军事征服\n多地区政治纳入'],['司法与动员','监督地方审判\n组织人员与服务'],['教会与文化','教育与文本整理\n维系跨地区联系']].forEach((a,i)=>{
  heading(s,a[0],xs[i],190,350,31);
  body(s,a[1],xs[i],256,350,121,29);
 });
 rule(s,418);
 text(s,'帝国认同上：罗马皇帝\n治理结构上：新型欧洲君主',64,462,1152,120,38,{bold:true,color:C.red});
}
// 11 — Bibliography remains in the deck but is outside the speaking sequence.
{
 const s=base(11,'附录 · 书目',data.slides[10].title);
 const entries=[
 ['艾因哈德《查理大帝传》','10—12、28—30页（PDF 19—21、37—39）'],
 ['巴尔贝罗《查理大帝：欧洲之父》','146—149、206—210、229、234、239、271—272页（PDF＝书内页＋9）；地图PDF 5'],
 ['布洛赫《封建社会》上卷','266—269页（合并PDF 288—291）'],
 ['Janet L. Nelson, King and Emperor: A New Life of Charlemagne','PDF 175、204、405—408、480—481、515页'],
 ['皮朗《穆罕默德和查理曼》','译者前言4页（PDF 10）；仅用于经济答辩的争议背景']
 ];
 entries.forEach((a,i)=>{
   const y=171+i*76;
   text(s,a[0],64,y,1152,36,i===3?23:25,{bold:true});
   text(s,a[1],64,y+35,1152,34,19.7,{color:C.muted});
 });
 text(s,'页码以课程所供版本为准；中文书默认书内页码，Nelson采用PDF顺序页码（从1计）。',64,575,1152,42,20,{color:C.red});
 s.speakerNotes.text += '\n完整书目与定位：\n'+data.bibliography.map(x=>x.details+x.title+'；'+x.pages).join('\n');
}
// 12 — Backup economy: distinguish agriculture from complete autarky.
{
 const s=base(12,'附录 · 答辩备用','庄园经济能直接证明“新型”吗？');
 text(s,'农业与庄园占重要地位  ≠  经济完全封闭',64,190,1152,76,39,{bold:true,color:C.red});
 heading(s,'可确认的事实',64,315,500,30);
 body(s,'庄园可能生产盈余\n交换与市场并未消失',64,376,510,132,30);
 heading(s,'可使用的判断',670,315,500,30);
 body(s,'解释资源组织的条件\n不能单独证明“完整封建制度”',670,376,548,132,29);
 text(s,'皮朗的解释具有影响力，也存在争议。经济只作补充，不替代政治治理证据。',64,555,1152,56,25,{color:C.muted});
}

const candidate=build+'/candidate.pptx';
await (await PresentationFile.exportPptx(p)).save(candidate);
await fs.writeFile(build+'/source.snapshot.json',JSON.stringify(p.toProto(),null,2));
for(let i=0;i<p.slides.items.length;i++){
  const slide=p.slides.items[i];
  const out=await p.export({slide,format:'png',scale:1.25});
  await fs.writeFile(build+`/预览/slide-${String(i+1).padStart(2,'0')}.png`,new Uint8Array(await out.arrayBuffer()));
  const lay=await slide.export({format:'layout'});
  await fs.writeFile(build+`/预览/slide-${String(i+1).padStart(2,'0')}.layout.json`,await lay.text());
}
console.log(JSON.stringify({candidate,slides:p.slides.items.length,rendered:12}));

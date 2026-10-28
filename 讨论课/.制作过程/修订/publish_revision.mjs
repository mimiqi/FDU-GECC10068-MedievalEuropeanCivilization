import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
process.env.RUNTIME_NODE_MODULES='C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/修订';
const skill='C:/Users/mimiqizhang/.codex/plugins/cache/openai-primary-runtime/presentations/26.1017.22136/skills/presentations';
const ppt=root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx';
const doc=root+'/交付文件/查理曼_12分钟汇报讲稿.docx';
const backup=work+'/修改前';
const oldHash='1a78c26d8023d774ead5dfe54c37619e5df0770d7289cd17cda0796bcd4a982d';
const sha=async file=>createHash('sha256').update(await fs.readFile(file)).digest('hex');
if(await sha(ppt)!==oldHash)throw new Error('Source PPT changed during revision; do not overwrite.');
await fs.mkdir(backup,{recursive:true});
const backupPpt=backup+'/查理曼_政治治理视角_12分钟汇报.pptx';
const backupDoc=backup+'/查理曼_12分钟汇报讲稿.docx';
for(const name of [backupPpt,backupDoc]){if(await fs.access(name).then(()=>true,()=>false))throw new Error('Backup already exists: '+name);}
const {finalizePresentation}=await import(pathToFileURL(skill+'/container_tools/artifact_tool_utils.mjs').href);
await fs.rename(ppt,backupPpt);
let receipt;
try{
 receipt=await finalizePresentation({
  explicitTotalSlideCount:12,workspaceDir:root,candidatePath:work+'/candidate.pptx',finalPath:ppt,
  pythonExecutable:'C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
  integrityValidatorPath:skill+'/container_tools/inspect_presentation_package_integrity.py',
  layoutValidatorPath:skill+'/container_tools/inspect_presentation_layout_geometry.py',
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit','--require-native-table-slide','9'],
  requiredNativeTableOwnerSlides:[9],requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:'reference',families:['Microsoft YaHei'],referencePath:backupPpt,referenceSha256:oldHash},
  verifyArtifactToolImport:true,receiptPath:work+'/validation.json'
 });
}catch(error){
 if(!await fs.access(ppt).then(()=>true,()=>false))await fs.copyFile(backupPpt,ppt);
 throw error;
}
await fs.rename(doc,backupDoc);
try{await fs.copyFile(work+'/讲稿修订候选.docx',doc);}
catch(error){await fs.copyFile(backupDoc,doc);throw error;}
const record={ppt,doc,backup,slides:12,manuscriptPages:9,mainSpeechHanCharacters:2414,
 sourceData:work+'/content.revised.json',pptSha256:await sha(ppt),docxSha256:await sha(doc),
 renderedPptMatchesFinal:await sha(ppt)===await sha(work+'/candidate.pptx'),
 renderedDocxMatchesFinal:await sha(doc)===await sha(work+'/讲稿修订候选.docx'),
 packagePassed:receipt.packageIntegrity.status==='pass',layoutPassed:receipt.presentationLayout.findingCount===0};
await fs.writeFile(work+'/交付记录.json',JSON.stringify(record,null,2));
console.log(JSON.stringify(record));

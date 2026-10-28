import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
process.env.RUNTIME_NODE_MODULES='C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/二至四页调整';
const skill='C:/Users/mimiqizhang/.codex/plugins/cache/openai-primary-runtime/presentations/26.1017.22136/skills/presentations';
const source=root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx';
const final=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_第2至4页修订.pptx';
const backup=work+'/修改前/查理曼_政治治理视角_12分钟汇报.pptx';
const expected='9152837795e5c214003f3d56bb822eb720f4c4f971f9e8282f907507b8f779ef';
const sha=async name=>createHash('sha256').update(await fs.readFile(name)).digest('hex');
if(await sha(source)!==expected)throw new Error('The user edited the source again; do not overwrite.');
if(await fs.access(backup).then(()=>true,()=>false))throw new Error('The backup already exists.');
await fs.mkdir(work+'/修改前',{recursive:true});
const {finalizePresentation}=await import(pathToFileURL(skill+'/container_tools/artifact_tool_utils.mjs').href);
await fs.copyFile(source,backup);
let result;
try{
 result=await finalizePresentation({
  explicitTotalSlideCount:11,workspaceDir:root,candidatePath:work+'/candidate.pptx',finalPath:final,
  pythonExecutable:'C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
  integrityValidatorPath:skill+'/container_tools/inspect_presentation_package_integrity.py',
  layoutValidatorPath:skill+'/container_tools/inspect_presentation_layout_geometry.py',
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit','--require-native-table-slide','9'],
  requiredNativeTableOwnerSlides:[9],requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:'reference',families:['Microsoft YaHei'],referencePath:backup,referenceSha256:expected},
  verifyArtifactToolImport:true,receiptPath:work+'/validation.json'
 });
}catch(error){
 throw error;
}
const record={final,backup,slides:11,changedVisiblePages:[2,3,4,11],changedNotes:[2,3,4,6,11],
 sha256:await sha(final),renderedFileMatchesFinal:await sha(final)===await sha(work+'/candidate.pptx'),
 packagePassed:result.packageIntegrity.status==='pass',layoutPassed:result.presentationLayout.findingCount===0};
await fs.writeFile(work+'/交付记录.json',JSON.stringify(record,null,2));
console.log(JSON.stringify(record));

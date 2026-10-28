import fs from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {pathToFileURL} from 'node:url';
process.env.RUNTIME_NODE_MODULES='C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const root='D:/text/02_中世纪欧洲文明/讨论课',work=root+'/.制作过程/第八页与讲稿同步';
const source=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_深化修订版.pptx';
const final=root+'/交付文件/查理曼_政治治理视角_12分钟汇报_同步修订版.pptx';
const expected='2d2e6a8874785ed24b78d0e25f3731381230604d0d2e197ced4f908a761b75da';
const sha=async file=>createHash('sha256').update(await fs.readFile(file)).digest('hex');
if(await sha(source)!==expected)throw new Error('Source changed during revision.');
const skill='C:/Users/mimiqizhang/.codex/plugins/cache/openai-primary-runtime/presentations/26.1017.22136/skills/presentations';
const {finalizePresentation}=await import(pathToFileURL(skill+'/container_tools/artifact_tool_utils.mjs').href);
const result=await finalizePresentation({
 explicitTotalSlideCount:11,workspaceDir:root,candidatePath:work+'/candidate.pptx',finalPath:final,
 pythonExecutable:'C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
 integrityValidatorPath:skill+'/container_tools/inspect_presentation_package_integrity.py',
 layoutValidatorPath:skill+'/container_tools/inspect_presentation_layout_geometry.py',
 layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit','--require-native-table-slide','9'],
 requiredNativeTableOwnerSlides:[9],requiredNativeChartOwnerSlides:[],
 fontPolicy:{basis:'reference',families:['Microsoft YaHei'],referencePath:source,referenceSha256:expected},
 verifyArtifactToolImport:true,receiptPath:work+'/validation.json'
});
const record={final,source,slides:11,changedVisiblePages:[8],changedNotes:[2,3,8],
 sha256:await sha(final),renderedFileMatchesFinal:await sha(final)===await sha(work+'/candidate.pptx'),
 packagePassed:result.packageIntegrity.status==='pass',layoutPassed:result.presentationLayout.findingCount===0};
await fs.writeFile(work+'/交付记录.json',JSON.stringify(record,null,2));
console.log(JSON.stringify(record));

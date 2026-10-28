import {pathToFileURL} from 'node:url';
process.env.RUNTIME_NODE_MODULES='C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules';
const skill='C:/Users/mimiqizhang/.codex/plugins/cache/openai-primary-runtime/presentations/26.1017.22136/skills/presentations';
const root='D:/text/02_中世纪欧洲文明/讨论课';
const {finalizePresentation}=await import(pathToFileURL(skill+'/container_tools/artifact_tool_utils.mjs').href);
const result=await finalizePresentation({
  explicitTotalSlideCount:12,
  workspaceDir:root,
  candidatePath:root+'/.制作过程/candidate.pptx',
  finalPath:root+'/交付文件/查理曼_政治治理视角_12分钟汇报.pptx',
  pythonExecutable:'C:/Users/mimiqizhang/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/python.exe',
  integrityValidatorPath:skill+'/container_tools/inspect_presentation_package_integrity.py',
  layoutValidatorPath:skill+'/container_tools/inspect_presentation_layout_geometry.py',
  layoutArgs:['--expected-slide-size-emu','12192000,6858000','--validate-bullet-geometry','--validate-heading-fit','--require-native-table-slide','9'],
  requiredNativeTableOwnerSlides:[9],requiredNativeChartOwnerSlides:[],
  fontPolicy:{basis:'design',families:['Microsoft YaHei']},
  verifyArtifactToolImport:true,
  receiptPath:root+'/.制作过程/validation.json'
});
console.log(JSON.stringify(result));

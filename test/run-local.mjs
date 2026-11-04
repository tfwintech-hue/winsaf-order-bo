// 사용법: node test/run-local.mjs 파일1.xlsx [파일2.xlsx ...]
import { readFileSync, writeFileSync } from 'node:fs';
import { convertWorkbook, buildOutput, resolveModel } from '../src/convert.mjs';
const groups = [];
for (const f of process.argv.slice(2)) {
  const rows = await convertWorkbook(readFileSync(f));
  console.log(`\n# ${f} → ${rows.length}줄`);
  for (const r of rows) console.log(JSON.stringify(r));
  groups.push({ round: '1차', rows });
}
// 모델 판별 단위 시험
const cases = [
  ['스테이미니 헤어드라이어 WTH-1636 접이식 드라이기 화이트', '', ''],
  ['WC-40B 블랙 좌측 열림 (우힌지)', '', ''],
  ['WC-40B 우측열림', '', ''],
  ['JG-36B', '', 'S'],
  ['JG-36B', '', ''],
  ['[윈세프] 원터치 무선전기포트(WTK-2113CG)1L/커피포트/블랙', '', ''],
  ['WH-N1', '', '없음'],
];
console.log('\n# 모델 판별');
for (const [t, ty, co] of cases) { const r = resolveModel(t, ty, co); console.log(t, '→', r.m ? `${r.m.유형}|${r.m.옵션}|${r.m.모델명}` : '', r.notes.join(';')); }
writeFileSync('test/out-sample.xlsx', await buildOutput(groups));

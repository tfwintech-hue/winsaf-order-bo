// 거래처 발주서 엑셀 → 윈세프 발주서 "일반 주문건" 줄로 변환
// 기준: 윈세프 발주서 작성규칙(v2·AI인수인계) 4장·5-4·6장
import ExcelJS from 'exceljs';
import { readFileSync } from 'node:fs';

const here = new URL('.', import.meta.url);
const MODELS = JSON.parse(readFileSync(new URL('models.json', here), 'utf8')).모델;
const VENDORS = JSON.parse(readFileSync(new URL('vendors.json', here), 'utf8'));

export const OUT_HEADER = ['발주처', '유형', '옵션', '모델명', '수량', '고객명', '연락처 1', '연락처 2',
  '주소', '비고', '배송메세지', '택배사', '송장번호', '물류허브', '사방넷', 'P', '확인 필요 (붙여넣기 전에 지우세요)'];

// 머리글 이름 → 발주서 열 (띄어쓰기 없이 비교)
const HEADER_MAP = [
  ['발주처', ['발주처']],
  ['유형', ['유형']],
  ['옵션', ['색상', '옵션']],
  ['모델명', ['모델명', '상품명', '품목']],
  ['수량', ['수량']],
  ['고객명', ['고객명', '수취인', '수취인명', '수령인', '받는분']],
  ['연락처1', ['전화번호1', '전화번호', '연락처1', '연락처', '휴대폰', '수취인전화번호1']],
  ['연락처2', ['전화번호2', '연락처2', '수취인전화번호2']],
  ['주소', ['배송처', '주소', '배송지', '수취인주소']],
  ['비고', ['비고']],
  ['배송메세지', ['배송메세지', '배송메시지', '배송시요청사항', '요청사항']],
];

const squash = s => String(s ?? '').replace(/\s+/g, '');
const norm = s => String(s ?? '').toUpperCase().replace(/[\s\-_()\[\]]/g, '');

function cellText(cell) {
  const v = cell?.value;
  if (v == null) return '';
  if (typeof v === 'object') {
    if (v.richText) return v.richText.map(t => t.text).join('');
    if (v.text) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v instanceof Date) return v.toISOString().slice(0, 10);
  }
  return String(v).replace(/_x000D_/g, '').trim();
}

function findHeader(ws) {
  for (let r = 1; r <= Math.min(ws.rowCount, 15); r++) {
    const row = ws.getRow(r);
    const map = {};
    row.eachCell({ includeEmpty: false }, (cell, col) => {
      const h = squash(cellText(cell));
      for (const [key, names] of HEADER_MAP) {
        if (map[key]) continue;
        if (names.includes(h)) { map[key] = col; break; }
      }
    });
    if (map.모델명 && (map.고객명 || map.주소)) return { row: r, map };
  }
  return null;
}

// 전화번호: 원본 그대로 텍스트로. 엑셀이 숫자로 바꿔 맨 앞 0이 사라진 경우만 0을 붙이고 표시
function phone(cell, notes, label) {
  const raw = cellText(cell);
  if (!raw) return '';
  if (typeof cell.value === 'number' && /^[1-9]\d{8,10}$/.test(raw)) {
    notes.push(`${label} 맨 앞 0 복원(${raw})`);
    return '0' + raw;
  }
  return raw;
}

const COLOR_SUFFIX = [
  ['아이보리', ['IV', 'W']], ['화이트', ['W', 'IV']], ['블랙', ['BK', 'B']], ['실버', ['S']],
  ['골드', ['G']], ['핑크', ['PK', 'P']], ['소다', ['SD']], ['블루', ['BL']], ['그린', ['GN', 'GR']],
];
const HINGE_R =['우힌지', '좌측열림', '좌측방향', '좌측손잡이', '우경첩'];
const HINGE_L = ['좌힌지', '우측열림', '우측방향', '우측손잡이', '좌경첩'];

// 모델명 칸 글자(상품명이 통째로 와도 됨) → 매뉴얼 행
export function resolveModel(text, vendorType = '', vendorColor = '') {
  const t = norm(text);
  if (!t) return { notes: ['모델명 없음'] };
  let best = [];
  let bestLen = 0;
  for (const m of MODELS) {
    const key = norm(m.모델명);
    if (key.length < 3 || !t.includes(key)) continue;
    if (key.length > bestLen) { best = [m]; bestLen = key.length; }
    else if (key.length === bestLen) best.push(m);
  }
  if (!best.length) {
    // 모델 글자에 색상 글자가 빠진 경우: "WTH-1636 … 화이트" → WTH-1636W (규칙서 6-2)
    const words = squash(`${text} ${vendorColor}`);
    for (const [word, sufs] of COLOR_SUFFIX) {
      if (!words.includes(word)) continue;
      for (const m of MODELS) {
        const key = norm(m.모델명);
        const s = sufs.find(x => key.endsWith(x));
        if (s && key.length - s.length >= 4 && t.includes(key.slice(0, -s.length)) && !t.includes(key)) best.push(m);
      }
      if (best.length) break;
    }
    if (best.length === 1) return { m: best[0], notes: ['상품명 색상으로 모델 판단 — 확인 권장'] };
    if (!best.length) return { notes: ['매뉴얼에 없는 모델 — 직접 확인'] };
  }
  if (best.length === 1) return { m: best[0], notes: [] };

  // 같은 모델의 여러 행(R/L, 색상)이 있으면 힌지·색상 글자로 고름
  const all = squash(`${text} ${vendorType} ${vendorColor}`);
  const isR = HINGE_R.some(w => all.includes(w)) || /(^|\/)R($|\/)/.test(squash(vendorType));
  const isL = HINGE_L.some(w => all.includes(w)) || /(^|\/)L($|\/)/.test(squash(vendorType));
  let cands = best;
  if (isR !== isL) {
    const want = isR ? 'R' : 'L';
    const f = cands.filter(m => (want === 'R' ? (m.유형 === 'R' || !/L/.test(m.유형)) : /L/.test(m.유형)));
    if (f.length) cands = f;
  }
  if (cands.length > 1 && vendorColor) {
    const f = cands.filter(m => m.옵션 && squash(vendorColor).includes(squash(m.옵션)));
    if (f.length) cands = f;
  }
  if (cands.length === 1) return { m: cands[0], notes: [] };
  return { notes: [`후보 ${cands.length}개(${cands.map(m => [m.유형, m.옵션].filter(Boolean).join('/') || '기본').join(', ')}) — 힌지·색상 확인`] };
}

function vendorName(raw, fallback) {
  const v = (raw || fallback || '').trim();
  return VENDORS.발주처별칭[v] || v;
}

// buffer: 엑셀 파일 내용, fallbackVendor: 메일 보낸 거래처의 발주처 이름
export async function convertWorkbook(buffer, fallbackVendor = '') {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const rows = [];
  for (const ws of wb.worksheets) {
    if (/모델리스트|매뉴얼/.test(ws.name)) continue;
    const h = findHeader(ws);
    if (!h) continue;
    const c = (row, key) => (h.map[key] ? row.getCell(h.map[key]) : { value: null });
    for (let r = h.row + 1; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const model = cellText(c(row, '모델명'));
      const name = cellText(c(row, '고객명'));
      if (!model && !name) continue;
      if (squash(model) === '모델명' || squash(name) === '고객명') continue; // 머리글이 한 번 더 있는 경우
      const notes = [];
      const vType = cellText(c(row, '유형'));
      let vColor = cellText(c(row, '옵션'));
      if (vColor === '없음') vColor = '';
      const res = resolveModel(model, vType, vColor);
      notes.push(...res.notes);
      const tel1 = phone(c(row, '연락처1'), notes, '연락처1');
      let tel2 = phone(c(row, '연락처2'), notes, '연락처2');
      if (tel2 === tel1 || tel2 === '0') tel2 = '';
      const qty = Number(cellText(c(row, '수량'))) || cellText(c(row, '수량'));
      if (!qty) notes.push('수량 없음');
      if (!name) notes.push('고객명 없음');
      const addr = cellText(c(row, '주소'));
      if (!addr) notes.push('주소 없음');
      rows.push({
        발주처: vendorName(cellText(c(row, '발주처')), fallbackVendor),
        유형: res.m ? res.m.유형 : vType,
        옵션: res.m ? res.m.옵션 : vColor,
        모델명: res.m ? res.m.모델명 : model,
        수량: qty,
        고객명: name,
        연락처1: tel1,
        연락처2: tel2,
        주소: addr,
        비고: cellText(c(row, '비고')),
        배송메세지: cellText(c(row, '배송메세지')),
        확인: notes.join(' / '),
      });
    }
  }
  return rows;
}

// 변환 결과 → 붙여넣기용 엑셀 (연락처는 텍스트, 2차는 P열 보라)
export async function buildOutput(groups) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('붙여넣기용');
  ws.addRow(OUT_HEADER);
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).eachCell(cell => { cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } }; });
  for (const g of groups) {
    for (const r of g.rows) {
      const row = ws.addRow([r.발주처, r.유형, r.옵션, r.모델명, r.수량, r.고객명, r.연락처1, r.연락처2,
        r.주소, r.비고, r.배송메세지, '', '', '', '', '', r.확인]);
      row.getCell(7).numFmt = '@';
      row.getCell(8).numFmt = '@';
      row.getCell(10).font = { color: { argb: 'FFFF0000' } };
      if (g.round === '2차') row.getCell(16).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFB4A7D6' } };
      if (r.확인) row.getCell(17).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFF2CC' } };
      row.alignment = { vertical: 'middle', wrapText: true };
    }
  }
  [14, 8, 11, 16, 6, 9, 15, 15, 46, 12, 18, 8, 14, 10, 10, 6, 34].forEach((w, i) => { ws.getColumn(i + 1).width = w; });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

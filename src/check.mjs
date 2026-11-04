// 네이버 메일 '발주서' 메일함 확인 → 새 발주서 변환 → 변환 파일을 메일로 발송 → 현황판(docs/status.json) 갱신
// GitHub Actions가 평일 업무시간에 10분마다 실행합니다. (.github/workflows/check-orders.yml)
import { ImapFlow } from 'imapflow';
import { simpleParser } from 'mailparser';
import nodemailer from 'nodemailer';
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { convertWorkbook, buildOutput } from './convert.mjs';

const { NAVER_ID, NAVER_APP_PASSWORD } = process.env;
const MAIL_TO = process.env.MAIL_TO || `${NAVER_ID}@naver.com`;
if (!NAVER_ID || !NAVER_APP_PASSWORD) {
  console.error('NAVER_ID, NAVER_APP_PASSWORD 비밀값이 없습니다. README의 3단계를 확인하세요.');
  process.exit(1);
}

const VENDORS = JSON.parse(readFileSync(new URL('vendors.json', import.meta.url), 'utf8'));
const STATUS = new URL('../docs/status.json', import.meta.url);

// 한국 시간
const kst = d => new Date(d.getTime() + 9 * 3600e3);
const ymd = d => kst(d).toISOString().slice(0, 10);
const hm = d => kst(d).toISOString().slice(11, 16);
const now = new Date();
const today = ymd(now);

// 1차 10:30 / 2차 13:30 (규칙서 1장)
const roundOf = d => { const t = hm(d); return t < '10:30' ? '1차' : t < '13:30' ? '2차' : '마감 후'; };

function vendorOf(from) {
  const s = `${from?.name || ''} ${from?.address || ''}`.toLowerCase();
  return VENDORS.거래처.find(v => v.보낸사람.some(k => s.includes(k.toLowerCase())));
}

let status = existsSync(STATUS) ? JSON.parse(readFileSync(STATUS, 'utf8')) : {};
if (status.date !== today) status = { date: today, mails: [], processed: [] };

const imap = new ImapFlow({
  host: 'imap.naver.com', port: 993, secure: true,
  auth: { user: NAVER_ID, pass: NAVER_APP_PASSWORD }, logger: false,
});
const smtp = nodemailer.createTransport({
  host: 'smtp.naver.com', port: 465, secure: true,
  auth: { user: NAVER_ID, pass: NAVER_APP_PASSWORD },
});

await imap.connect();
const boxes = await imap.list();
const box = boxes.find(b => b.name === VENDORS.메일함 || b.path === VENDORS.메일함);
if (!box) {
  console.error(`'${VENDORS.메일함}' 메일함을 찾지 못했습니다. 있는 메일함: ${boxes.map(b => b.path).join(', ')}`);
  await imap.logout();
  process.exit(1);
}

const lock = await imap.getMailboxLock(box.path);
let newCount = 0;
try {
  const uidValidity = String(imap.mailbox.uidValidity);
  const uids = await imap.search({ since: new Date(`${today}T00:00:00+09:00`) }, { uid: true });
  for (const uid of uids) {
    const key = `${uidValidity}:${uid}`;
    if (status.processed.includes(key)) continue;

    // source 가져오기는 '읽음' 표시를 바꾸지 않습니다 (BODY.PEEK)
    const msg = await imap.fetchOne(uid, { source: true, internalDate: true, envelope: true }, { uid: true });
    const when = msg.internalDate || now;
    if (ymd(when) !== today) { status.processed.push(key); continue; }

    const mail = await simpleParser(msg.source);
    const from = mail.from?.value?.[0];
    const vendor = vendorOf(from);
    const round = roundOf(when);
    const entry = {
      time: hm(when), round,
      vendor: vendor ? vendor.이름 : '미등록 거래처',
      subject: (mail.subject || '').slice(0, 60),
      files: [], rows: 0, flags: 0, error: '',
    };

    const groups = [];
    for (const att of mail.attachments || []) {
      const fname = att.filename || '첨부';
      entry.files.push(fname);
      if (!/\.xlsx$/i.test(fname)) { entry.error ||= `변환 불가 형식(${fname})`; continue; }
      try {
        const rows = await convertWorkbook(att.content, vendor?.발주처 || '');
        groups.push({ round, rows });
        entry.rows += rows.length;
        entry.flags += rows.filter(r => r.확인).length;
      } catch (e) {
        entry.error ||= `엑셀 읽기 실패(${fname})`;
      }
    }
    if (!entry.files.length) entry.error = '첨부파일 없음';
    if (groups.length && entry.rows === 0) entry.error ||= '주문 줄을 찾지 못함(양식 확인)';

    if (entry.rows > 0) {
      const out = await buildOutput(groups);
      await smtp.sendMail({
        from: `${NAVER_ID}@naver.com`, to: MAIL_TO,
        subject: `[발주 자동변환] ${entry.vendor} ${entry.rows}줄 (${round})${entry.flags ? ` · 확인필요 ${entry.flags}` : ''}`,
        text: [
          `원본 메일: ${mail.subject || ''} (${entry.time} 수신)`,
          `첨부: ${entry.files.join(', ')}`,
          `변환: ${entry.rows}줄, 확인 필요 ${entry.flags}줄`,
          round === '2차' ? '2차 추가분입니다. P열이 보라색으로 칠해져 있습니다.' : '',
          '마지막 열(확인 필요)은 발주서에 붙여넣기 전에 지우세요.',
        ].filter(Boolean).join('\n'),
        attachments: [{ filename: `발주변환_${today}_${entry.vendor}_${entry.time.replace(':', '')}.xlsx`, content: out }],
      });
    }
    status.mails.push(entry);
    status.processed.push(key);
    newCount++;
  }
} finally {
  lock.release();
  await imap.logout();
}

status.updatedAt = `${today} ${hm(now)}`;
status.vendors = VENDORS.거래처.map(v => ({ 이름: v.이름, 매일: v.매일 }));
delete status.sample;
writeFileSync(STATUS, JSON.stringify(status, null, 1));
console.log(`확인 완료: 새 메일 ${newCount}건`);

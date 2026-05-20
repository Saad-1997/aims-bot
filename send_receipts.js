/**
 * send_receipts.js
 * ──────────────────────────────────────────────────────
 * Sends confirmation receipts to all 31 existing students
 * from your Google Forms Excel export.
 *
 * HOW TO USE:
 * 1. Export Google Forms responses as .xlsx
 * 2. Rename it: existing_students.xlsx
 * 3. Place it in same folder as this file
 * 4. Make sure bot.js ran at least once (auth_info folder exists)
 * 5. Run: node send_receipts.js
 * ──────────────────────────────────────────────────────
 */

import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
} from '@whiskeysockets/baileys';
import { Boom } from '@hapi/boom';
import pino from 'pino';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const EXCEL_FILE = path.join(__dirname, 'existing_students.xlsx');

// ── Receipt builder ──
function buildReceipt(d) {
  return `✅ *AIMS Academy D.I.Khan*
━━━━━━━━━━━━━━━━━━━━━━━
🎓 *Application Confirmed!*
━━━━━━━━━━━━━━━━━━━━━━━
🔢 *Roll Number:*  ${d.roll}
👤 *Full Name:*    ${d.name}
👨 *Father:*       ${d.father}
💼 *Occupation:*   ${d.occupation}
🏫 *College:*      ${d.college}
📚 *Program:*      ${d.program}
📝 *Matric:*       ${d.matric}
📗 *FSc Part I:*   ${d.fsc1}
📘 *FSc Part II:*  ${d.fsc2 || 'N/A'}
🩺 *MDCAT:*        ${d.mdcat || 'N/A'}
📱 *WhatsApp:*     ${d.whatsapp}
━━━━━━━━━━━━━━━━━━━━━━━
🏆 *Merit Score:*  ${d.merit}
_(FSc 40% + MDCAT 50% + Matric 10%)_
━━━━━━━━━━━━━━━━━━━━━━━
Thank you for applying! Our team will contact you soon. 🎉
*AIMS Academy D.I.Khan*`;
}

// ── Format phone number ──
function formatPhone(num) {
  let clean = String(num).replace(/[\s\-\+\(\)]/g, '');
  if (clean.startsWith('0')) clean = '92' + clean.slice(1);
  if (!clean.startsWith('92')) clean = '92' + clean;
  return clean + '@s.whatsapp.net';
}

// ── Read Excel ──
async function readStudents() {
  if (!fs.existsSync(EXCEL_FILE)) {
    console.error(`\n❌ File not found: existing_students.xlsx`);
    console.error('   Please place your Google Forms export in this folder.\n');
    process.exit(1);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_FILE);
  const sheet = workbook.worksheets[0];

  const students = [];
  let headers = [];

  sheet.eachRow((row, rowNum) => {
    const values = row.values.slice(1);
    if (rowNum === 1) {
      headers = values.map(v => String(v || '').trim());
      return;
    }

    const obj = {};
    headers.forEach((h, i) => { obj[h] = String(values[i] || '').trim(); });

    if (!obj['Full Name'] || !obj['WhatsApp Number']) return;

    students.push({
      roll:       obj['Roll Number']                    || '',
      name:       obj['Full Name']                      || '',
      father:     obj['Father Name']                    || '',
      occupation: obj['Father Occupation']              || '',
      college:    obj['College Name']                   || '',
      program:    obj['Program and Gender']             || '',
      matric:     obj['Marks in Matric']                || '',
      fsc1:       obj['Marks in Fsc i']                 || '',
      fsc2:       obj['Marks in Fsc ii (If Applicable)']|| 'N/A',
      mdcat:      obj['Marks in MDCAT (If Applicable)'] || 'N/A',
      whatsapp:   obj['WhatsApp Number']                || '',
      merit:      obj['Merit']                          || '',
    });
  });

  return students;
}

// ── Sleep helper ──
const sleep = ms => new Promise(r => setTimeout(r, ms));

// ── Main ──
async function main() {
  const students = await readStudents();
  console.log(`\n📋 Found ${students.length} students in Excel.\n`);

  if (students.length === 0) {
    console.log('No students found. Check your column names match exactly.');
    process.exit(0);
  }

  const { state, saveCreds } = await useMultiFileAuthState('auth_info');
  const { version }          = await fetchLatestBaileysVersion();

  const sock = makeWASocket({
    version,
    auth:   state,
    logger: pino({ level: 'silent' }),
    browser: ['AIMS Receipts', 'Chrome', '1.0.0'],
  });

  sock.ev.on('creds.update', saveCreds);

  sock.ev.on('connection.update', async ({ connection, lastDisconnect, qr }) => {
    if (qr) {
      console.log('⚠️  QR appeared — run bot.js first to connect WhatsApp, then run this.');
    }

    if (connection === 'open') {
      console.log('✅ Connected! Sending receipts...\n');

      let sent = 0, failed = 0;

      for (const student of students) {
        if (!student.whatsapp) {
          console.log(`⚠️  Skipped (no number): ${student.name}`);
          failed++;
          continue;
        }

        const jid     = formatPhone(student.whatsapp);
        const receipt = buildReceipt(student);

        try {
          await sock.sendMessage(jid, { text: receipt });
          console.log(`✅ Sent → ${student.name} (${student.roll}) | ${student.whatsapp}`);
          sent++;
          await sleep(4000); // 4 second gap to avoid spam detection
        } catch (err) {
          console.log(`❌ Failed → ${student.name} | ${student.whatsapp} | ${err.message}`);
          failed++;
          await sleep(2000);
        }
      }

      console.log(`\n━━━━━━━━━━━━━━━━━━━━━━━`);
      console.log(`✅ Successfully sent: ${sent}`);
      console.log(`❌ Failed:           ${failed}`);
      console.log(`━━━━━━━━━━━━━━━━━━━━━━━`);
      console.log('Done! You can close this window.\n');

      await sock.logout();
      process.exit(0);
    }

    if (connection === 'close') {
      const shouldReconnect =
        (lastDisconnect?.error instanceof Boom)
          ? lastDisconnect.error.output?.statusCode !== DisconnectReason.loggedOut
          : true;
      if (shouldReconnect) main();
    }
  });
}

main().catch(console.error);

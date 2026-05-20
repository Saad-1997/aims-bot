import makeWASocket, {
  useMultiFileAuthState,
  DisconnectReason,
  fetchLatestBaileysVersion,
} from '@whiskeysockets/baileys';
import { Boom } from '@hapi/boom';
import pino from 'pino';
import qrcode from 'qrcode-terminal';
import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ─────────────────────────────────────────
//  ROLL NUMBER COUNTERS
// ─────────────────────────────────────────
const DEFAULT_COUNTERS = {
  'Pre Medical - Boys':      26118,
  'Pre Medical - Girls':     27115,
  'Pre Engineering - Boys':  28101,
  'Pre Engineering - Girls': 29101,
  'ICS - Boys':              30103,
  'ICS - Girls':             31101,
};

const COUNTER_FILE = path.join(__dirname, 'roll_counters.json');
const rollCounters = fs.existsSync(COUNTER_FILE)
  ? JSON.parse(fs.readFileSync(COUNTER_FILE, 'utf8'))
  : { ...DEFAULT_COUNTERS };

function saveCounters() {
  fs.writeFileSync(COUNTER_FILE, JSON.stringify(rollCounters, null, 2));
}

function assignRoll(program) {
  const roll = rollCounters[program];
  rollCounters[program]++;
  saveCounters();
  return roll;
}

// ─────────────────────────────────────────
//  MERIT CALCULATION
// ─────────────────────────────────────────
function calcMerit(fscObt, fscTotal, matricObt, matricTotal, mdcatObt, mdcatTotal) {
  const fscPct    = (fscObt / fscTotal) * 100;
  const matricPct = (matricObt / matricTotal) * 100;

  let merit;
  if (mdcatObt !== null && mdcatTotal !== null) {
    const mdcatPct = (mdcatObt / mdcatTotal) * 100;
    merit = (fscPct * 0.40) + (mdcatPct * 0.50) + (matricPct * 0.10);
  } else {
    merit = (fscPct * 0.6667) + (matricPct * 0.3333);
  }
  return merit.toFixed(2);
}

// ─────────────────────────────────────────
//  EXCEL SAVE — exact Google Forms columns
// ─────────────────────────────────────────
const EXCEL_FILE = path.join(__dirname, 'AIMS_Admissions.xlsx');

async function saveToExcel(data) {
  const workbook = new ExcelJS.Workbook();
  let sheet;

  if (fs.existsSync(EXCEL_FILE)) {
    await workbook.xlsx.readFile(EXCEL_FILE);
    sheet = workbook.getWorksheet('Admissions');
  }

  if (!sheet) {
    sheet = workbook.addWorksheet('Admissions');
    sheet.columns = [
      { header: 'Timestamp',                       key: 'timestamp',  width: 22 },
      { header: 'Email Address',                   key: 'email',      width: 26 },
      { header: 'Full Name',                       key: 'fullName',   width: 22 },
      { header: 'Father Name',                     key: 'fatherName', width: 22 },
      { header: 'Father Occupation',               key: 'fatherOcc',  width: 22 },
      { header: 'College Name',                    key: 'college',    width: 24 },
      { header: 'Program and Gender',              key: 'program',    width: 24 },
      { header: 'Marks in Matric',                 key: 'matric',     width: 16 },
      { header: 'Marks in Fsc i',                  key: 'fsc1',       width: 14 },
      { header: 'Marks in Fsc ii (If Applicable)', key: 'fsc2',       width: 22 },
      { header: 'Marks in MDCAT (If Applicable)',  key: 'mdcat',      width: 22 },
      { header: 'WhatsApp Number',                 key: 'whatsapp',   width: 16 },
      { header: 'Roll Number',                     key: 'roll',       width: 12 },
      { header: 'Merit',                           key: 'merit',      width: 10 },
      { header: 'Fee Paid (RS 100 EasyPaisa)',     key: 'feePaid',    width: 26 },
    ];

    const headerRow = sheet.getRow(1);
    headerRow.font      = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill      = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B5E20' } };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };
    headerRow.height    = 20;
  }

  sheet.addRow({
    timestamp:  new Date().toLocaleString('en-PK'),
    email:      'Via WhatsApp',
    fullName:   data.fullName,
    fatherName: data.fatherName,
    fatherOcc:  data.fatherOccupation,
    college:    data.collegeName,
    program:    data.program,
    matric:     data.matricMarks,
    fsc1:       data.fscPart1,
    fsc2:       data.fscPart2 !== 'skip' ? data.fscPart2 : 'N/A',
    mdcat:      data.mdcatMarks !== 'skip' ? data.mdcatMarks : 'N/A',
    whatsapp:   data.whatsappNumber,
    roll:       data.rollNumber,
    merit:      data.merit + '%',
    feePaid:    data.feePaid,
  });

  await workbook.xlsx.writeFile(EXCEL_FILE);
  console.log(`✅ Saved | ${data.fullName} | Roll: ${data.rollNumber} | Merit: ${data.merit}%`);
}

// ─────────────────────────────────────────
//  CONVERSATION STEPS
// ─────────────────────────────────────────
const PROGRAMS = [
  'Pre Medical - Boys',
  'Pre Medical - Girls',
  'Pre Engineering - Boys',
  'Pre Engineering - Girls',
  'ICS - Boys',
  'ICS - Girls',
];

const STEPS = [
  {
    key: 'fullName',
    ask: '👤 Please enter your *Full Name*:',
  },
  {
    key: 'fatherName',
    ask: "👨 Please enter your *Father's Name*:",
  },
  {
    key: 'fatherOccupation',
    ask: "💼 Please enter your *Father's Occupation*:",
  },
  {
    key: 'collegeName',
    ask: '🏫 Please enter your *Previous School / College Name*:',
  },
  {
    key: 'program',
    ask: '📚 Please select your *Program and Gender*\nReply with the number only:\n\n1️⃣  Pre Medical - Boys\n2️⃣  Pre Medical - Girls\n3️⃣  Pre Engineering - Boys\n4️⃣  Pre Engineering - Girls\n5️⃣  ICS - Boys\n6️⃣  ICS - Girls',
    validate: v => parseInt(v) >= 1 && parseInt(v) <= 6,
    parse:    v => PROGRAMS[parseInt(v) - 1],
    errorMsg: '⚠️ Please reply with a number from *1 to 6* only.',
  },
  {
    key: 'matricMarks',
    ask: '📝 Please enter your *Marks in Matric*:\n_(Format: Obtained/Total — e.g. *950/1200*)_',
    validate: v => {
      const p = v.split('/');
      return p.length === 2 && !isNaN(p[0].trim()) && !isNaN(p[1].trim()) && Number(p[1]) > 0;
    },
    errorMsg: '⚠️ Please use the format *Obtained/Total*\nExample: *950/1200*',
  },
  {
    key: 'fscPart1',
    ask: '📗 Please enter your *Marks in FSc Part I*:\n_(Format: Obtained/Total — e.g. *480/600*)_',
    validate: v => {
      const p = v.split('/');
      return p.length === 2 && !isNaN(p[0].trim()) && !isNaN(p[1].trim()) && Number(p[1]) > 0;
    },
    errorMsg: '⚠️ Please use the format *Obtained/Total*\nExample: *480/600*',
  },
  {
    key: 'fscPart2',
    ask: '📘 Please enter your *Marks in FSc Part II* (if applicable):\n_(Format: Obtained/Total — or type *Skip*)_',
    validate: v => {
      if (v.toLowerCase() === 'skip') return true;
      const p = v.split('/');
      return p.length === 2 && !isNaN(p[0].trim()) && !isNaN(p[1].trim()) && Number(p[1]) > 0;
    },
    parse: v => v.toLowerCase() === 'skip' ? 'skip' : v,
    errorMsg: '⚠️ Please use format *Obtained/Total* or type *Skip*',
  },
  {
    key: 'mdcatMarks',
    ask: '🩺 Please enter your *Marks in MDCAT* (if applicable):\n_(Format: Obtained/Total — or type *Skip*)_',
    validate: v => {
      if (v.toLowerCase() === 'skip') return true;
      const p = v.split('/');
      return p.length === 2 && !isNaN(p[0].trim()) && !isNaN(p[1].trim()) && Number(p[1]) > 0;
    },
    parse: v => v.toLowerCase() === 'skip' ? 'skip' : v,
    errorMsg: '⚠️ Please use format *Obtained/Total* or type *Skip*',
  },
  {
    key: 'whatsappNumber',
    ask: '📱 Please enter your *WhatsApp Number*:\n_(We will send you updates and results on this number)_',
  },
  {
    key: 'feePaid',
    ask: '💰 Have you submitted *RS 100/-* registration fee on EasyPaisa?\n📲 Account: *03436884574*\n\nReply with *Yes* or *No*',
    validate: v => ['yes', 'no'].includes(v.toLowerCase()),
    parse: v => v.charAt(0).toUpperCase() + v.slice(1).toLowerCase(),
    errorMsg: '⚠️ Please reply with *Yes* or *No* only.',
  },
];

// ─────────────────────────────────────────
//  RECEIPT
// ─────────────────────────────────────────
function buildReceipt(d) {
  return `✅ *AIMS Academy D.I.Khan*
━━━━━━━━━━━━━━━━━━━━━━━
🎓 *Application Confirmed!*
━━━━━━━━━━━━━━━━━━━━━━━
🔢 *Roll Number:*  ${d.rollNumber}
👤 *Full Name:*    ${d.fullName}
👨 *Father:*       ${d.fatherName}
💼 *Occupation:*   ${d.fatherOccupation}
🏫 *College:*      ${d.collegeName}
📚 *Program:*      ${d.program}
📝 *Matric:*       ${d.matricMarks}
📗 *FSc Part I:*   ${d.fscPart1}
📘 *FSc Part II:*  ${d.fscPart2 !== 'skip' ? d.fscPart2 : 'N/A'}
🩺 *MDCAT:*        ${d.mdcatMarks !== 'skip' ? d.mdcatMarks : 'N/A'}
📱 *WhatsApp:*     ${d.whatsappNumber}
💰 *Fee Paid:*     ${d.feePaid}
━━━━━━━━━━━━━━━━━━━━━━━
🏆 *Merit Score:*  ${d.merit}%
_(FSc 40% + MDCAT 50% + Matric 10%)_
━━━━━━━━━━━━━━━━━━━━━━━
Thank you for applying! Our team will contact you soon. 🎉
*AIMS Academy D.I.Khan*`;
}

// ─────────────────────────────────────────
//  SESSION STORE
// ─────────────────────────────────────────
const sessions = {};

function getSession(jid) {
  if (!sessions[jid]) sessions[jid] = { step: 'WELCOME', data: {} };
  return sessions[jid];
}

// ─────────────────────────────────────────
//  PROCESS INCOMING MESSAGE
// ─────────────────────────────────────────
async function handleMessage(sock, jid, text) {
  const session = getSession(jid);
  const send    = async (msg) => {
    await sock.sendMessage(jid, { text: msg });
  };

  // ── WELCOME ──
  if (session.step === 'WELCOME') {
    await send(
      `🌟 *Welcome to AIMS Academy D.I.Khan!* 🌟\n\n` +
      `📋 *Admission & Scholarship Portal*\n\n` +
      `This is the official registration desk for AIMS Academy.\n` +
      `Please provide accurate information to complete your admission and scholarship application.\n\n` +
      `Your unique *Roll Number* will be assigned immediately upon successful submission.\n\n` +
      `Type *Start* to begin your application ✨`
    );
    session.step = 'AWAIT_START';
    return;
  }

  // ── AWAIT START ──
  if (session.step === 'AWAIT_START') {
    if (text.toLowerCase() !== 'start') {
      await send('Please type *Start* to begin your application 👇');
      return;
    }
    session.step = 0;
    await send(STEPS[0].ask);
    return;
  }

  // ── COLLECTING ANSWERS ──
  if (typeof session.step === 'number') {
    const current = STEPS[session.step];

    // Validate
    if (current.validate && !current.validate(text)) {
      await send(current.errorMsg);
      return;
    }

    // Store
    const value = current.parse ? current.parse(text) : text;
    session.data[current.key] = value;

    // Confirm program selection
    if (current.key === 'program') {
      await send(`✅ Program selected: *${value}*`);
    }

    session.step++;

    // Next question
    if (session.step < STEPS.length) {
      await send(STEPS[session.step].ask);
      return;
    }

    // ── ALL DONE ──
    const d = session.data;

    await send('⏳ Processing your application, please wait...');

    // Parse marks for merit
    const [fsc1Obt, fsc1Total] = d.fscPart1.split('/').map(s => Number(s.trim()));
    const [matObt,  matTotal]  = d.matricMarks.split('/').map(s => Number(s.trim()));

    let mdcatObt = null, mdcatTotal = null;
    if (d.mdcatMarks && d.mdcatMarks !== 'skip') {
      [mdcatObt, mdcatTotal] = d.mdcatMarks.split('/').map(s => Number(s.trim()));
    }

    const merit      = calcMerit(fsc1Obt, fsc1Total, matObt, matTotal, mdcatObt, mdcatTotal);
    const rollNumber = assignRoll(d.program);
    const finalData  = { ...d, merit, rollNumber };

    await saveToExcel(finalData);
    await send(buildReceipt(finalData));

    // Clear session
    delete sessions[jid];
  }
}

// ─────────────────────────────────────────
//  START BOT
// ─────────────────────────────────────────
async function startBot() {
  const { state, saveCreds } = await useMultiFileAuthState('auth_info');
  const { version }          = await fetchLatestBaileysVersion();

  const sock = makeWASocket({
    version,
    auth:   state,
    logger: pino({ level: 'silent' }), // hide noise
    printQRInTerminal: false,          // we handle QR ourselves
    browser: ['AIMS Academy Bot', 'Chrome', '1.0.0'],
  });

  // ── QR CODE ──
  sock.ev.on('connection.update', async ({ connection, lastDisconnect, qr }) => {
    if (qr) {
      console.clear();
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.log('   AIMS Academy Bot — Scan QR to Connect');
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
      qrcode.generate(qr, { small: true });
      console.log('\n📱 Open WhatsApp → 3 dots → Linked Devices → Link a Device\n');
    }

    if (connection === 'open') {
      console.clear();
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
      console.log('  ✅ AIMS Academy Bot is LIVE!');
      console.log('  📋 Students can now apply via WhatsApp');
      console.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n');
    }

    if (connection === 'close') {
      const shouldReconnect =
        (lastDisconnect?.error instanceof Boom)
          ? lastDisconnect.error.output?.statusCode !== DisconnectReason.loggedOut
          : true;

      if (shouldReconnect) {
        console.log('🔄 Reconnecting...');
        startBot();
      } else {
        console.log('❌ Logged out. Delete auth_info folder and restart.');
      }
    }
  });

  // ── SAVE CREDENTIALS ──
  sock.ev.on('creds.update', saveCreds);

  // ── INCOMING MESSAGES ──
  sock.ev.on('messages.upsert', async ({ messages, type }) => {
    if (type !== 'notify') return;

    for (const msg of messages) {
      if (msg.key.fromMe)           continue; // ignore own messages
      if (msg.key.remoteJid.endsWith('@g.us')) continue; // ignore groups

      const jid  = msg.key.remoteJid;
      const text = (
        msg.message?.conversation ||
        msg.message?.extendedTextMessage?.text ||
        ''
      ).trim();

      if (!text) continue;

      console.log(`📩 Message from ${jid}: ${text}`);

      try {
        await handleMessage(sock, jid, text);
      } catch (err) {
        console.error('Error handling message:', err);
      }
    }
  });
}

startBot();

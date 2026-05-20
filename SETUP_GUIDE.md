# AIMS Academy Bot — Termux Setup Guide

---

## WHAT YOU NEED
- Any Android phone (even old one)
- Termux app installed from F-Droid
- Phone plugged into charger during admission period

---

## STEP 1 — Install Termux
Download from F-Droid (NOT Play Store):
👉 https://f-droid.org/packages/com.termux/

---

## STEP 2 — Setup Termux (run these one by one)

```bash
pkg update && pkg upgrade
```
(Press Y when asked, wait for finish)

```bash
pkg install nodejs git
```
(Press Y when asked)

---

## STEP 3 — Copy Bot Files to Phone

**Option A — USB Cable:**
- Connect phone to PC
- Copy the aims-baileys folder to phone storage
- In Termux:
```bash
cp -r /sdcard/aims-baileys ~/aims-baileys
cd ~/aims-baileys
```

**Option B — Clone from GitHub (if you pushed it):**
```bash
git clone YOUR_GITHUB_URL aims-baileys
cd aims-baileys
```

**Option C — Download ZIP:**
```bash
cd ~
mkdir aims-baileys
cd aims-baileys
```
Then manually copy each file using a file manager app.

---

## STEP 4 — Install Dependencies

```bash
cd ~/aims-baileys
npm install
```
Wait for it to finish (needs internet, ~2-3 minutes).

---

## STEP 5 — Start the Bot

```bash
npm start
```

A QR code will appear in Termux.

---

## STEP 6 — Scan QR Code
- Open WhatsApp on ANOTHER phone (or the same phone in split screen)
- Tap ⋮ (3 dots) → Linked Devices → Link a Device
- Scan the QR code shown in Termux

✅ Bot is now LIVE!

---

## STEP 7 — Keep Bot Running

**Important:** Enable these in Android settings:
1. Settings → Battery → AIMS/Termux → No restrictions
2. Settings → Apps → Termux → Disable battery optimization

This prevents Android from killing the bot.

Also install **Termux:Boot** from F-Droid so bot restarts automatically if phone reboots.

---

## SEND RECEIPTS TO EXISTING 31 STUDENTS

1. Put your Google Forms Excel export as: `existing_students.xlsx` in the aims-baileys folder
2. Open a NEW Termux session (swipe from left → New Session)
3. Run:
```bash
cd ~/aims-baileys
node send_receipts.js
```

---

## ROLL NUMBERS

| Program | Next Roll Number |
|---------|-----------------|
| Pre Medical - Boys | 26118 |
| Pre Medical - Girls | 27115 |
| Pre Engineering - Boys | 28101 |
| Pre Engineering - Girls | 29101 |
| ICS - Boys | 30103 |
| ICS - Girls | 31101 |

---

## MERIT FORMULA

```
Merit = FSc 40% + MDCAT 50% + Matric 10%
(No MDCAT: FSc 66.67% + Matric 33.33%)
```

---

## COMMON PROBLEMS

| Problem | Fix |
|---------|-----|
| Bot stops replying | Open Termux → run npm start again |
| QR expired | Just restart npm start, new QR appears |
| "Cannot find module" | Run npm install again |
| Phone restarted | Open Termux → cd ~/aims-baileys → npm start |
| Excel not saving | Check storage permission for Termux |

---

## GIVE TERMUX STORAGE ACCESS

Run this once:
```bash
termux-setup-storage
```
Allow when Android asks for permission.

---

## FILES EXPLAINED

| File | Purpose |
|------|---------|
| bot.js | Main bot — run this |
| send_receipts.js | Send receipts to existing students |
| package.json | Dependencies list |
| auth_info/ | Auto-created — WhatsApp session |
| roll_counters.json | Auto-created — tracks roll numbers |
| AIMS_Admissions.xlsx | Auto-created — all student data |
| existing_students.xlsx | YOUR FILE — Google Forms export |

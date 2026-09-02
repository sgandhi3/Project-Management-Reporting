// Opens a pre-filled draft addressed to EMAIL_TO with the report attached.
// The user reviews and sends it manually — this never sends anything itself.
// Used because this Deloitte O365 tenant has SMTP AUTH disabled tenant-wide
// (SmtpClientAuthentication is disabled for the Mailbox), so
// scripts/send-email-local.js's nodemailer/SMTP approach cannot work here at
// all, regardless of credentials.
//
// Platform behavior:
//   macOS   — AppleScript automates the classic Outlook desktop app directly,
//             attachment included.
//   Windows — tries PowerShell + Outlook COM automation first (attachment
//             included) — this needs the classic, COM-automatable Outlook
//             desktop app. If that's not installed (e.g. only the "new
//             Outlook for Windows" is present, which doesn't expose the same
//             COM interface), falls back to a `mailto:` link — this opens
//             whatever the OS's default mail handler is (new Outlook, classic
//             Outlook, or anything else) pre-filled with To/Subject/Body, and
//             separately opens File Explorer with the report file selected
//             so it can be dragged into the draft in one step. No known way
//             to attach a file via mailto/URI scheme on any platform — that's
//             blocked by every mail client/browser for security, not an
//             Outlook-specific limitation.
//
// Usage: node scripts/create-outlook-draft.js <path-to-report-file>
import '../env.js';
import fs from 'fs';
import path from 'path';
import { execFile } from 'child_process';
import { promisify } from 'util';

const execFileAsync = promisify(execFile);

const filePath = process.argv[2];
if (!filePath || !fs.existsSync(filePath)) {
  console.error(`Report file not found: ${filePath}`);
  process.exit(1);
}

const to = process.env.EMAIL_TO || process.env.OUTLOOK_USER;
if (!to) {
  console.error('EMAIL_TO (or OUTLOOK_USER) must be set in .env to address the draft.');
  process.exit(1);
}

const today = new Date().toLocaleDateString('en-US', {
  month: 'long', day: 'numeric', year: 'numeric',
});
const subject = `MMO SIT Report Ready — ${today}`;
const body = `The MMO SIT Execution & Defect Report for ${today} is attached.`;
const absPath = path.resolve(filePath);

async function openDraftMac() {
  // AppleScript string literals: escape backslashes and double quotes.
  const esc = s => s.replace(/\\/g, '\\\\').replace(/"/g, '\\"');

  const script = `
tell application "Microsoft Outlook"
  set newMessage to make new outgoing message with properties {subject:"${esc(subject)}", content:"${esc(body)}"}
  make new recipient at end of to recipients of newMessage with properties {email address:{address:"${esc(to)}"}}
  make new attachment at newMessage with properties {file:POSIX file "${esc(absPath)}"}
  open newMessage
  activate
end tell
`;
  await execFileAsync('osascript', ['-e', script]);
}

async function openDraftWindowsCom() {
  // PowerShell single-quoted string literals: escape embedded single quotes
  // by doubling them (PowerShell's escape convention, not a backslash).
  const esc = s => s.replace(/'/g, "''");

  const script = `
$outlook = New-Object -ComObject Outlook.Application
$mail = $outlook.CreateItem(0)
$mail.Subject = '${esc(subject)}'
$mail.Body = '${esc(body)}'
$mail.To = '${esc(to)}'
$mail.Attachments.Add('${esc(absPath)}')
$mail.Display()
`;
  await execFileAsync('powershell.exe', ['-NoProfile', '-NonInteractive', '-Command', script]);
}

// Fallback for when there's no COM-automatable Outlook (e.g. only "new
// Outlook for Windows" is installed, or macOS Outlook AppleScript fails for
// some other reason). Opens a mailto: draft (recipient/subject/body only —
// no attachment support exists on any platform via this mechanism) and
// reveals the report file in the OS file manager so it's a single
// drag-and-drop to finish.
async function openMailtoFallback() {
  const mailto = `mailto:${encodeURIComponent(to)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;

  if (process.platform === 'win32') {
    await execFileAsync('cmd.exe', ['/c', 'start', '""', mailto]);
    try {
      await execFileAsync('explorer.exe', [`/select,${absPath}`]);
    } catch { /* explorer.exe routinely exits non-zero even on success */ }
  } else {
    await execFileAsync('open', [mailto]);
    try {
      await execFileAsync('open', ['-R', absPath]);
    } catch { /* non-fatal — file manager reveal is a convenience, not required */ }
  }

  console.log(`Couldn't attach automatically — opened a draft addressed to ${to} instead, and revealed ${path.basename(absPath)} in the file manager. Drag it into the draft to finish.`);
}

try {
  if (process.platform === 'darwin') {
    await openDraftMac();
    console.log(`Draft opened in Outlook, addressed to ${to}, with ${path.basename(absPath)} attached. Review and send manually.`);
  } else if (process.platform === 'win32') {
    try {
      await openDraftWindowsCom();
      console.log(`Draft opened in Outlook, addressed to ${to}, with ${path.basename(absPath)} attached. Review and send manually.`);
    } catch (comError) {
      console.warn(`Outlook COM automation failed (${comError.message.split('\n')[0]}) — likely only "new Outlook for Windows" is installed, which doesn't support this. Falling back to a mailto draft.`);
      await openMailtoFallback();
    }
  } else {
    console.error(`Unsupported platform for draft automation: ${process.platform} (only macOS and Windows are supported).`);
    process.exit(1);
  }
} catch (e) {
  console.error(`Failed to open a draft: ${e.message}`);
  process.exit(1);
}

// Opens a pre-filled draft in the Microsoft Outlook desktop app (macOS) with
// the report attached — the user reviews and sends it manually. Never sends
// anything itself. Used because this Deloitte O365 tenant has SMTP AUTH
// disabled tenant-wide (SmtpClientAuthentication is disabled for the
// Mailbox), so scripts/send-email-local.js's nodemailer/SMTP approach cannot
// work here at all, regardless of credentials.
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

try {
  await execFileAsync('osascript', ['-e', script]);
  console.log(`Draft opened in Outlook, addressed to ${to}, with ${path.basename(absPath)} attached. Review and send manually.`);
} catch (e) {
  console.error(`Failed to open Outlook draft: ${e.message}`);
  process.exit(1);
}

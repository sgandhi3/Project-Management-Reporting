// Emails a generated report file as an attachment via Outlook/O365 SMTP.
// Usage: node scripts/send-email-local.js <path-to-report-file>
import nodemailer from 'nodemailer';
import fs from 'fs';
import path from 'path';

const filePath = process.argv[2];
if (!filePath || !fs.existsSync(filePath)) {
  console.error(`Report file not found: ${filePath}`);
  process.exit(1);
}

if (!process.env.OUTLOOK_USER || !process.env.OUTLOOK_PASS) {
  console.error('OUTLOOK_USER and OUTLOOK_PASS must be set in .env before sending email.');
  process.exit(1);
}

const today = new Date().toLocaleDateString('en-US', {
  month: 'long', day: 'numeric', year: 'numeric',
});

const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST || 'smtp.office365.com',
  port: 587,
  secure: false,
  auth: {
    user: process.env.OUTLOOK_USER,
    pass: process.env.OUTLOOK_PASS,
  },
});

await transporter.verify().catch(err => {
  console.error('SMTP connection failed:', err.message);
  process.exit(1);
});

const info = await transporter.sendMail({
  from: `"MMO Report" <${process.env.OUTLOOK_USER}>`,
  to: process.env.EMAIL_TO || process.env.OUTLOOK_USER,
  subject: `MMO SIT Report Ready — ${today}`,
  text: `The MMO SIT Execution & Defect Report for ${today} is attached.`,
  html: `<p>The MMO SIT Execution &amp; Defect Report for <strong>${today}</strong> is attached.</p>`,
  attachments: [{ filename: path.basename(filePath), path: filePath }],
});

console.log(`Report emailed → ${info.messageId}`);

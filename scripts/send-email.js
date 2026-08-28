import nodemailer from 'nodemailer';
import fs from 'fs';

const reportPath = process.env.REPORT_PATH || './report.pptx';

if (!fs.existsSync(reportPath)) {
  console.error(`Report not found: ${reportPath}`);
  process.exit(1);
}

const today = new Date().toLocaleDateString('en-US', {
  month: 'long', day: 'numeric', year: 'numeric', timeZone: 'America/New_York',
});
const dateSlug = new Date().toISOString().slice(0, 10);

const transporter = nodemailer.createTransport({
  host: 'smtp.office365.com',
  port: 587,
  secure: false,
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

await transporter.verify().catch(err => {
  console.error('SMTP connection failed:', err.message);
  process.exit(1);
});

const info = await transporter.sendMail({
  from: `"MMO Report" <${process.env.SMTP_USER}>`,
  to: process.env.EMAIL_TO,
  subject: `MMO SIT Execution & Defect Report — ${today}`,
  text: `Hi team,\n\nPlease find attached the MMO SIT Execution & Defect Report for ${today}.\n\nThis report was generated automatically from Azure DevOps.\n\nThanks`,
  html: `<p>Hi team,</p><p>Please find attached the MMO SIT Execution &amp; Defect Report for <strong>${today}</strong>.</p><p>This report was generated automatically from Azure DevOps.</p><p>Thanks</p>`,
  attachments: [{
    filename: `MMO_Report_${dateSlug}.pptx`,
    path: reportPath,
    contentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  }],
});

console.log(`Email sent → ${info.messageId}`);

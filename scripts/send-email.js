import nodemailer from 'nodemailer';

const runUrl = `${process.env.GITHUB_SERVER_URL}/${process.env.GITHUB_REPOSITORY}/actions/runs/${process.env.GITHUB_RUN_ID}`;
const today = new Date().toLocaleDateString('en-US', {
  month: 'long', day: 'numeric', year: 'numeric', timeZone: 'America/New_York',
});

const transporter = nodemailer.createTransport({
  host: 'smtp-mail.outlook.com',
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
  to: process.env.EMAIL_TO,
  subject: `MMO SIT Report Ready — ${today}`,
  text: `The MMO SIT Execution & Defect Report for ${today} is ready.\n\nDownload it here (sign in to GitHub first):\n${runUrl}\n\nThe file is available for 30 days.`,
  html: `
    <p>The MMO SIT Execution &amp; Defect Report for <strong>${today}</strong> is ready.</p>
    <p><a href="${runUrl}" style="background:#0078d4;color:#fff;padding:10px 20px;text-decoration:none;border-radius:4px;display:inline-block;">Download Report</a></p>
    <p style="color:#666;font-size:12px;">Link expires in 30 days. Sign in to GitHub to access.</p>
  `,
});

console.log(`Notification sent → ${info.messageId}`);

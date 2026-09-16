require('dotenv').config();
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Configuration
const CONFIG = {
  batchSize: 50,
  delayMs: 45000, // 45 seconds between each email
  excelFile: 'HR_Email_Automation.xlsx', // Your specified Excel file
  progressFile: 'progress.json',
  resumePath: path.join(__dirname, 'Twaran_Gupta_Backend_Software_Engineer.pdf'), // Your specified PDF
  // Test mode: run with `TEST_MODE=true node mailer.js` to send ONE email to
  // testEmail and exit, without reading the list or touching progress.json.
  testMode: process.env.TEST_MODE === 'true',
  testEmail: 'twarangupta01@gmail.com',
};

// Subject lines are rotated per contact so a few hundred identical subjects
// do not look like a bulk blast to spam filters. Add or edit freely.
const SUBJECTS = [
  'Application: Software Engineer (2 years experience) - Twaran Gupta',
  'Software Engineer - improved API response times by 60% at ZS Associates - Twaran Gupta',
  'Backend / Full-Stack Engineer open to new roles - Twaran Gupta',
  'Node.js / React Developer (2 yrs) looking to join your team - Twaran Gupta',
  'Software Engineer who loves building products - Twaran Gupta',
  'Experienced Software Engineer exploring opportunities - resume attached - Twaran Gupta',
  'Full-Stack Software Engineer application - Twaran Gupta',
];

// Setup Nodemailer transport using your environment variables
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.GMAIL_USER,
    pass: process.env.GMAIL_APP_PASS,
  },
});

// Helper to track progress
function getProgress() {
  if (!fs.existsSync(CONFIG.progressFile)) {
    return { lastIndex: 0 };
  }
  return JSON.parse(fs.readFileSync(CONFIG.progressFile, 'utf8'));
}

function saveProgress(lastIndex) {
  fs.writeFileSync(CONFIG.progressFile, JSON.stringify({ lastIndex }, null, 2));
}

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Plaintext version. Sent alongside the HTML as multipart/alternative:
// improves inbox placement and gives a fallback for clients that strip HTML.
function buildTextBody(hrName) {
  return `Hi ${hrName},

I am a Software Engineer who cut production API response times by 60% (about 500ms to 200ms) at ZS Associates, and I am now exploring new backend and full-stack roles.

At ZS Associates I built and shipped backend services for Pfizer and Merck: multi-tenant REST APIs secured with JWT and role-based access, plus an async re-architecture with Redis caching that drove the numbers above. Alongside my enterprise work, I designed and shipped SinkedIn, a live full-stack platform, end to end.

I am available to start immediately, and open to remote roles or on-site in Pune, Bangalore, Hyderabad, or Delhi NCR.

Core Tech Stack
Backend: Node.js, Express.js, PostgreSQL, Redis, REST APIs
Frontend: React.js, Next.js
Cloud & Tools: AWS, Docker, Git, CI/CD

Even a quick note on whether your team is hiring backend or full-stack engineers would help. My resume is attached for your review.

Twaran Gupta
+91 9718117060
linkedin.com/in/twarangupta | github.com/twarangupta | twarangupta.com`;
}

// HTML version.
function buildHtmlBody(hrName) {
  return `
  <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; background-color: #f4f6f9; padding: 30px 10px; width: 100%;">
    <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; padding: 30px; border-radius: 8px; border: 1px solid #e1e4e8; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">

      <!-- Header Section -->
      <div style="border-bottom: 2px solid #007bff; padding-bottom: 15px; margin-bottom: 20px;">
        <h1 style="margin: 0; font-size: 24px; color: #1a1a1a;">Twaran Gupta</h1>
        <p style="margin: 5px 0 0 0; font-size: 16px; color: #007bff; font-weight: 600;">Software Engineer | Backend &amp; Full-Stack</p>
      </div>

      <!-- Main Content -->
      <p style="color: #444444; font-size: 15px; line-height: 1.6; margin-top: 0;">Hi ${hrName},</p>

      <p style="color: #444444; font-size: 15px; line-height: 1.6;">
        I am a Software Engineer who <strong>cut production API response times by 60%</strong> (about 500ms to 200ms) at ZS Associates, and I am now exploring new backend and full-stack roles.
      </p>

      <p style="color: #444444; font-size: 15px; line-height: 1.6;">
        At ZS Associates I built and shipped backend services for Pfizer and Merck: multi-tenant REST APIs secured with JWT and role-based access, plus an async re-architecture with Redis caching that drove the number above. Alongside my enterprise work, I designed and shipped SinkedIn, a live full-stack platform, end to end.
      </p>

      <p style="color: #444444; font-size: 15px; line-height: 1.6;">
        I am available to <strong>start immediately</strong>, and open to remote roles or on-site in Pune, Bangalore, Hyderabad, or Delhi NCR.
      </p>

      <!-- Tech Stack (softened: neutral block, no marketing accent) -->
      <div style="background-color: #f8f9fa; border: 1px solid #e1e4e8; padding: 15px; margin: 25px 0; border-radius: 6px;">
        <h3 style="margin: 0 0 10px 0; color: #1a1a1a; font-size: 16px;">Core Tech Stack</h3>
        <p style="margin: 5px 0; font-size: 14px; color: #444444;">Backend: <strong>Node.js, Express.js, PostgreSQL, Redis, REST APIs</strong></p>
        <p style="margin: 5px 0; font-size: 14px; color: #444444;">Frontend: <strong>React.js, Next.js</strong></p>
        <p style="margin: 5px 0 0 0; font-size: 14px; color: #444444;">Cloud &amp; Tools: <strong>AWS, Docker, Git, CI/CD</strong></p>
      </div>

      <p style="color: #444444; font-size: 15px; line-height: 1.6;">
        Even a quick note on whether your team is hiring backend or full-stack engineers would help. My resume is attached for your review.
      </p>

      <!-- Footer with Buttons -->
      <div style="margin-top: 35px; padding-top: 20px; border-top: 1px solid #e1e4e8;">
        <p style="margin: 0; color: #1a1a1a; font-weight: bold; font-size: 15px;">Twaran Gupta</p>
        <p style="margin: 5px 0 15px 0; color: #666666; font-size: 14px;">+91 9718117060</p>

        <!-- Social Buttons -->
        <div>
          <a href="https://www.linkedin.com/in/twarangupta" style="display: inline-block; padding: 10px 16px; background-color: #0072b1; color: #ffffff; text-decoration: none; border-radius: 4px; font-size: 13px; font-weight: bold; margin-right: 8px;">LinkedIn</a>
          <a href="https://github.com/twarangupta" style="display: inline-block; padding: 10px 16px; background-color: #24292e; color: #ffffff; text-decoration: none; border-radius: 4px; font-size: 13px; font-weight: bold; margin-right: 8px;">GitHub</a>
          <a href="https://twarangupta.com" style="display: inline-block; padding: 10px 16px; background-color: #17a2b8; color: #ffffff; text-decoration: none; border-radius: 4px; font-size: 13px; font-weight: bold;">Portfolio</a>
        </div>
      </div>

    </div>
  </div>
  `;
}

async function sendTestEmail() {
  const hrName = 'Hiring Team';
  const mailOptions = {
    from: `"Twaran Gupta" <${process.env.GMAIL_USER}>`,
    replyTo: 'twarangupta01@gmail.com',
    to: CONFIG.testEmail,
    subject: `[TEST] ${SUBJECTS[0]}`,
    text: buildTextBody(hrName),
    html: buildHtmlBody(hrName),
    attachments: [
      {
        filename: 'Twaran_Gupta_Backend_Software_Engineer.pdf',
        path: CONFIG.resumePath,
      },
    ],
  };
  await transporter.sendMail(mailOptions);
  console.log(`Test email sent to ${CONFIG.testEmail}. Check how it renders, then run without TEST_MODE to send real batches.`);
}

async function runMailer() {
  // Read the Excel file
  const workbook = xlsx.readFile(CONFIG.excelFile);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const contacts = xlsx.utils.sheet_to_json(sheet);

  // Get current batch
  const { lastIndex } = getProgress();
  const batch = contacts.slice(lastIndex, lastIndex + CONFIG.batchSize);

  if (batch.length === 0) {
    console.log('All contacts have been processed!');
    return;
  }

  console.log(`Starting batch from index ${lastIndex} to ${lastIndex + batch.length}...`);

  // Process each contact in the batch
  for (let i = 0; i < batch.length; i++) {
    const contact = batch[i];
    const absoluteIndex = lastIndex + i;

    // Data mapping from your cleaned columns
    const firstName = contact['First Name'] ? String(contact['First Name']).trim() : '';
    const lastName = contact['Last Name'] ? String(contact['Last Name']).trim() : '';
    let hrName = `${firstName} ${lastName}`.trim();
    if (!hrName) hrName = 'Hiring Team';

    const hrEmail = contact['Email'];

    // Safeguard: skip blank emails (list is already verified)
    if (!hrEmail) {
      console.log(`Skipping index ${absoluteIndex}: No email found.`);
      continue;
    }

    // Rotate the subject line based on the contact's absolute position
    const subject = SUBJECTS[absoluteIndex % SUBJECTS.length];

    // Email options: send both plaintext and HTML (multipart/alternative)
    const mailOptions = {
      from: `"Twaran Gupta" <${process.env.GMAIL_USER}>`,
      replyTo: 'twarangupta01@gmail.com', // Ensures replies go to your main inbox
      to: hrEmail,
      subject: subject,
      text: buildTextBody(hrName),
      html: buildHtmlBody(hrName),
      attachments: [
        {
          filename: 'Twaran_Gupta_Backend_Software_Engineer.pdf',
          path: CONFIG.resumePath,
        },
      ],
    };

    // Send Email
    try {
      await transporter.sendMail(mailOptions);
      console.log(`[${i + 1}/${batch.length}] Sent to ${hrEmail}`);
    } catch (err) {
      console.error(`Failed sending to ${hrEmail}:`, err.message);
    }

    // Delay between emails to protect sender reputation. Runs whether the
    // send succeeded or failed, so a burst of failures cannot hammer Gmail
    // with no spacing. Skips the delay after the last email in the batch.
    if (i < batch.length - 1) {
      await sleep(CONFIG.delayMs);
    }
  }

  // Update tracking file
  saveProgress(lastIndex + batch.length);
  console.log(`Batch complete. Progress updated to index ${lastIndex + batch.length}.`);
}

if (CONFIG.testMode) {
  sendTestEmail().catch(console.error);
} else {
  runMailer().catch(console.error);
}

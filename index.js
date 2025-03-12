require('dotenv').config();
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const path = require('path');
const dns = require('dns').promises;

// Function to split and clean a string (for names or emails)
function splitAndClean(input) {
  if (!input) return [];
  input = input.trim();

  // Extract email addresses from the items
  const emailRegex = /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g;
  let matches = input.match(emailRegex) || [];
  matches = matches.map(email => email.replace(/^\d+[\.\)]/, '').replace(/[\s,;]+$/, ''));
  
  // Return the array of matched email addresses or an empty array if no matches are found
  console.log(matches)
  return matches || [];
}

// Function to validate an email's syntax and deliverability (via MX lookup)
async function isDeliverable(email) {
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!emailRegex.test(email)) return false;
  const domain = email.split('@')[1];
  try {
    const addresses = await dns.resolveMx(domain);
    return addresses && addresses.length > 0;
  } catch (err) {
    console.error(`DNS lookup failed for ${domain}:`, err.message);
    return false;
  }
}

// Function to read the Excel file and extract HR details
function getHRList(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  const hrList = [];
  data.forEach(row => {
    const company = row.Company || "Company Name Not Provided";
    const hrNamesRaw = row["Name of HR's"];
    const hrEmailsRaw = row["HR Email id "] || row["HR Email id"];
    if (!hrEmailsRaw) {
      console.warn("Skipping row (no HR Email id):", row);
      return;
    }
    let hrNames = splitAndClean(hrNamesRaw);
    let hrEmails = splitAndClean(hrEmailsRaw);
    while (hrNames.length < hrEmails.length) hrNames.push("HR");
    hrEmails.forEach((email, index) => {
      hrList.push({
        company,
        hrName: hrNames[index] || "HR",
        hrEmail: email
      });
    });
  });
  return hrList;
}

// Function to generate the email HTML body
function getEmailBody(hrName, companyName) {
  return `
    <p>Dear <span style="font-weight:bold;">${hrName}</span>,</p>
    <p>I'm <span style="font-weight:bold;">G. Dinesh Surya</span>, a final-year student at <span style="font-weight:bold;">NIT Kurukshetra</span>, and I'm highly interested in software opportunities (internships or full-time) at <span style="font-weight:bold;">${companyName}</span>.</p>
    <p><span style="font-weight:bold;">Key Highlights:</span></p>
    <ul>
      <li><span style="font-weight:bold;">Experience:</span> 6+ months of full-stack internship experience at <span style="font-weight:bold;">AutoRABIT</span> and <span style="font-weight:bold;">GrowthCraft</span>.</li>
      <li><span style="font-weight:bold;">Leadership:</span> Led the <span style="font-weight:bold;">Microbus Technical team (200+ members)</span> at NIT Kurukshetra.</li>
      <li><span style="font-weight:bold;">Coding Proficiency:</span> Top 3% on LeetCode (Knight rating: <span style="font-weight:bold;">1929</span>), University Rank 1 on InterviewBit (Top 0.7% globally).</li>
      <li><span style="font-weight:bold;">Tech Stack:</span> <span style="font-weight:bold;">C, C++, JavaScript, React.js, Node.js, Express.js, HTML/CSS</span>.</li>
    </ul>
    <p>My resume is attached for your review. I'd welcome the opportunity to discuss how my skills and experience align with your team's needs.</p>
    <p>Thank you for your time.</p>
    <p>Best regards,<br>
    <span style="font-weight:bold;">Dinesh Surya Gidijala</span><br>
    +91 8121400482<br>
    dineshsurya.2002@gmail.com</p>
    <p>LinkedIn: <a href="https://linkedin.com/in/dinesh21o9">linkedin.com/in/dinesh21o9</a> | GitHub: <a href="https://github.com/dinesh21o9">github.com/dinesh21o9</a></p>
  `;
}

// Function to send an email to a specific HR
async function sendEmail(companyName, hrName, hrEmail) {
  const mailOptions = {
    from: process.env.SMTP_USER,
    to: hrEmail,
    subject: `Software Developer Opportunity Inquiry at ${companyName}`,
    html: getEmailBody(hrName, companyName),
    attachments: [
      {
        filename: 'Dinesh_Surya_Gidijala_Resume.pdf',
        path: path.join(__dirname, 'Dinesh_Surya_Gidijala_Resume.pdf'),
      },
    ],
  };
  try {
    let info = await transporter.sendMail(mailOptions);
    console.log(`Email sent to ${hrEmail}: ${info.messageId}`);
    return true;
  } catch (error) {
    console.error(`Error sending email to ${hrEmail}:`, error.message);
    return false;
  }
}

// Main function to read the Excel file and send emails concurrently
async function main() {
  const { default: pLimit } = await import('p-limit'); // dynamic import for p-limit (ESM)
  const hrList = getHRList('./data/231-300.xlsx');
  const companyResults = {};
  let totalAttempts = 0, totalSuccesses = 0, totalFailures = 0;
  const limit = pLimit(20); // concurrency limit set to 20
  const tasks = hrList.map(hr => limit(async () => {
    totalAttempts++;
    if (!(await isDeliverable(hr.hrEmail))) {
      console.warn(`Skipping invalid/unreachable email: ${hr.hrEmail}`);
      totalFailures++;
      companyResults[hr.company] = companyResults[hr.company] || { attempts: 0, successes: 0, failures: 0 };
      companyResults[hr.company].attempts++;
      companyResults[hr.company].failures++;
      return;
    }
    console.log(`Sending email to ${hr.hrName} at ${hr.hrEmail} for ${hr.company}`);
    const success = await sendEmail(hr.company, hr.hrName, hr.hrEmail);
    companyResults[hr.company] = companyResults[hr.company] || { attempts: 0, successes: 0, failures: 0 };
    companyResults[hr.company].attempts++;
    if (success) {
      companyResults[hr.company].successes++;
      totalSuccesses++;
    } else {
      companyResults[hr.company].failures++;
      totalFailures++;
    }
  }));
  await Promise.all(tasks);
  const totalCompanies = Object.keys(companyResults).length;
  let companiesApplied = 0, companiesNotSent = 0;
  for (const company in companyResults) {
    if (companyResults[company].successes > 0) companiesApplied++;
    else companiesNotSent++;
  }
  console.log("\n====== Summary ======");
  console.log(`Total email attempts: ${totalAttempts}`);
  console.log(`Total successes: ${totalSuccesses}`);
  console.log(`Total failures (including invalid emails): ${totalFailures}`);
  console.log(`Total companies tried: ${totalCompanies}`);
  console.log(`Companies applied to (at least one email succeeded): ${companiesApplied}`);
  console.log(`Companies with no successful email: ${companiesNotSent}`);
}

main();

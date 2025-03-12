require('dotenv').config();
const nodemailer = require('nodemailer');
const XLSX = require('xlsx');
const path = require('path');

// Helper function to split and clean string entries
function splitAndClean(input) {
  if (!input) return [];
  input = input.trim();
  let items = [];
  // If the string starts with a numbered prefix like "1." or "1)"
  if (/^\d+[\.\)]/.test(input)) {
    // Split by any occurrence of digits followed by a period or closing parenthesis
    items = input.split(/\d+[\.\)]/).map(s => s.trim()).filter(Boolean);
  } else {
    // Otherwise, split on newline, comma, or semicolon.
    items = input.split(/[\n,;]+/).map(s => s.trim()).filter(Boolean);
  }
  return items;
}

// Function to read the Excel file and extract HR details
function getHRList(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  const hrList = [];

  data.forEach(row => {
    // Get company name; if missing, use a default.
    const company = row.Company || "Company Name Not Provided";

    // Extract HR names and emails from the respective columns.
    // The email column header might have a trailing space.
    const hrNamesRaw = row["Name of HR's"];
    const hrEmailsRaw = row["HR Email id "] || row["HR Email id"];

    // If no HR email is provided, skip this row.
    if (!hrEmailsRaw) {
      console.warn("Skipping row (no HR Email id):", row);
      return;
    }

    // Use the helper function to split and clean the names and emails.
    let hrNames = splitAndClean(hrNamesRaw);
    let hrEmails = splitAndClean(hrEmailsRaw);

    // If HR names array is smaller than emails array, fill missing names with default "HR"
    while (hrNames.length < hrEmails.length) {
      hrNames.push("HR");
    }

    // Create an entry for each HR email from this row.
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

// Create the email transporter using your SMTP configuration.
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: process.env.SMTP_PORT,
  secure: true, // true for port 465, false for port 587
  auth: {
    user: process.env.SMTP_USER,
    pass: process.env.SMTP_PASS,
  },
});

// Function to generate the email body using your updated template.
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

// Function to send an email to a specific HR individually.
// Returns true if the email was sent successfully, false otherwise.
async function sendEmail(companyName, hrName, hrEmail) {
  const mailOptions = {
    from: process.env.SMTP_USER,
    to: hrEmail, // Sending directly to HR's email.
    subject: `Software Opportunity Inquiry at ${companyName}`,
    html: getEmailBody(hrName, companyName), // Using HTML body
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
    console.error(`Error sending email to ${hrEmail}:`, error);
    return false;
  }
}


// Main function: read the Excel file and send emails to each HR.
async function main() {
  const hrList = getHRList('231-300.xlsx'); // Place your XLSX file in the same directory.
  
  // To keep track of results.
  const companyResults = {}; // { companyName: { attempts, successes, failures } }
  let totalAttempts = 0;
  let totalSuccesses = 0;
  let totalFailures = 0;
  
  // Loop through each HR entry and send an email individually.
  for (const hr of hrList) {
    totalAttempts++;
    console.log(`Sending email to ${hr.hrName} at ${hr.hrEmail} for ${hr.company}`);
    
    const success = await sendEmail(hr.company, hr.hrName, hr.hrEmail);
    
    // Initialize company results if not present.
    if (!companyResults[hr.company]) {
      companyResults[hr.company] = { attempts: 0, successes: 0, failures: 0 };
    }
    companyResults[hr.company].attempts++;
    if (success) {
      companyResults[hr.company].successes++;
      totalSuccesses++;
    } else {
      companyResults[hr.company].failures++;
      totalFailures++;
    }
  }
  
  // Calculate summary metrics.
  const totalCompanies = Object.keys(companyResults).length;
  let companiesApplied = 0; // Companies with at least one success.
  let companiesNotSent = 0; // Companies with zero success.
  
  for (const company in companyResults) {
    if (companyResults[company].successes > 0) {
      companiesApplied++;
    } else {
      companiesNotSent++;
    }
  }
  
  console.log("\n====== Summary ======");
  console.log(`Total email attempts: ${totalAttempts}`);
  console.log(`Total successes: ${totalSuccesses}`);
  console.log(`Total failures: ${totalFailures}`);
  console.log(`Total companies tried: ${totalCompanies}`);
  console.log(`Companies applied to (at least one email succeeded): ${companiesApplied}`);
  console.log(`Companies with no successful email: ${companiesNotSent}`);
}

main();

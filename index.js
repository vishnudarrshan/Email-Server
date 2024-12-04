import express from 'express';
import XLSX from 'xlsx';
import nodemailer from 'nodemailer';
import path from 'path';
import dotenv from 'dotenv';
import cron from 'node-cron';
import axios from 'axios';
import FormData from 'form-data';

// Load environment variables from .env file
dotenv.config();

const app = express();
const PORT = 5000;

// GitHub API Configuration
const GITHUB_TOKEN = process.env.GITHUB_TOKEN;
const GITHUB_OWNER = process.env.GITHUB_OWNER;
const GITHUB_REPO = process.env.GITHUB_REPO;
const GITHUB_FILE_PATH = process.env.GITHUB_FILE_PATH; // Path to the Excel file in the repository

// Gmail app password
const GMAIL_PASSWORD = process.env.GMAIL_PASSWORD; // Use environment variable for security

// Nodemailer Transport Configuration
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.GMAIL_USER, // Your Gmail address
    pass: GMAIL_PASSWORD, // App-specific password
  },
});

// Function to fetch the Excel file from GitHub
const fetchExcelFileFromGitHub = async () => {
  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE_PATH}`;
  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${GITHUB_TOKEN}`,
      },
    });
    // console.log(response);
    
    const fileContent = Buffer.from(response.data.content, 'base64');
    return XLSX.read(fileContent);
  } catch (error) {
    console.error('Error fetching Excel file from GitHub:', error);
    return null;
  }
};

// Function to write the updated Excel file back to GitHub
const writeExcelFileToGitHub = async (workbook) => {
  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE_PATH}`;
  
  // Prepare the file content and encode it to base64
  const fileBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });
  const base64File = fileBuffer.toString('base64');

  const data = {
    message: 'Update Excel file with new status',  // Commit message
    content: base64File,
    sha: await getFileSha(),  // Get the current file SHA for updating the file
  };

  try {
    const response = await axios.put(url, data, {
      headers: {
        Authorization: `Bearer ${GITHUB_TOKEN}`,
      },
    });
    console.log('Excel file successfully updated in GitHub');
  } catch (error) {
    console.error('Error updating Excel file to GitHub:', error);
  }
};

// Function to get the current file SHA from GitHub (required for update)
const getFileSha = async () => {
  const url = `https://api.github.com/repos/${GITHUB_OWNER}/${GITHUB_REPO}/contents/${GITHUB_FILE_PATH}`;
  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${GITHUB_TOKEN}`,
      },
    });
    return response.data.sha;
  } catch (error) {
    console.error('Error fetching file SHA from GitHub:', error);
    return null;
  }
};

// Function to send emails with status buttons
const sendEmail = async (to, cc, productName, timeRange) => {
  const serverUrl = `https://email-server-pearl.vercel.app/`; // Server URL
  const mailOptions = {
    from: process.env.GMAIL_USER, // Your Gmail address
    to: to,
    cc: cc.join(','),
    subject: `Scheduled Task for ${productName}`,
    html: `
      <p>The scheduled task for <b>${productName}</b> is planned at <b>${timeRange}</b> on tomorrow.</p>
      <p>Choose the task status:</p>
      <a href="${serverUrl}/update-status?product=${encodeURIComponent(productName)}&status=Completed">✔ Completed</a><br/>
      <a href="${serverUrl}/update-status?product=${encodeURIComponent(productName)}&status=Deferred">↺ Deferred</a><br/>
      <a href="${serverUrl}/update-status?product=${encodeURIComponent(productName)}&status=Not Completed">✖ Not Completed</a>
    `,
  };

  try {
    const info = await transporter.sendMail(mailOptions);
    console.log(`Email sent successfully to ${to}: ${info.response}`);
  } catch (error) {
    console.error(`Error sending email to ${to}:`, error);
  }
};

// Function to process the Excel data and send emails
const processAndSendEmails = () => {
  fetchExcelFileFromGitHub().then((workbook) => {
    if (!workbook) return;

    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
    if (!data) return;

    // Current date and time
    const now = new Date();

    data.forEach((row) => {
      const scheduledDateValue = row["Scheduled date"]; // Excel date (serial number)
      const scheduledDate = convertExcelDate(scheduledDateValue); // Convert to dd/mm/yyyy format
      const todayStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;
console.log(todayStr);

      if (scheduledDate === todayStr) {
        const emailAddresses = row.EmailAddresses.split(',').map((email) => email.trim());
        const to = emailAddresses[0]; // First email as 'to'
        const cc = emailAddresses.slice(1); // Remaining emails as 'cc'
        const timeRange = row.Time;
        const productName = row['Product name'];

        // Extract start time from the time range (e.g., "09:00 AM - 12:00 PM")
        const timeMatch = timeRange.match(/^(\d{1,2}:\d{2}\s*[APap][Mm])/);

        if (!timeMatch) {
          console.error(`Invalid or missing time format in row:`, row);
          return; // Skip this row
        }

        const startTime12Hour = timeMatch[1]; // Extracted start time in 12-hour format

        // Convert 12-hour format to 24-hour format
        const [time, period] = startTime12Hour.split(/\s+/);
        const [hour, minute] = time.split(':').map(Number);
        const startHour = period.toLowerCase() === 'pm' && hour < 12 ? hour + 12 : hour % 12;

        // Construct the scheduledTime
        const scheduledTime = new Date();
        scheduledTime.setHours(startHour, minute, 0, 0);

        // Calculate one hour before
        const oneHourBefore = new Date(scheduledTime.getTime() - 60 * 60 * 1000);

        console.log(`Now: ${now.toISOString()}`);
        console.log(`Scheduled Time: ${scheduledTime.toISOString()}`);
        console.log(`One Hour Before: ${oneHourBefore.toISOString()}`);

        // Send email 1 hour before the scheduled time
        if (now >= oneHourBefore && now < scheduledTime) {
          if (to) {
            sendEmail(to, cc, productName, timeRange);
          } else {
            console.error('No primary email address provided in row:', row);
          }
        } else {
          console.log("No email to be sent for row:", row);
        }
      } else {
        console.log("Scheduled date does not match today's date:", scheduledDate);

      }
    });
  });
};


// Helper function to convert Excel serial date to dd/mm/yyyy
const convertExcelDate = (serial) => {
  // Excel serial date starts from Jan 1, 1900
  const excelEpoch = new Date(1900, 0, 1);
  const adjustedDate = new Date(excelEpoch.getTime() + (serial - 2) * 24 * 60 * 60 * 1000); // Subtract 2 for Excel quirks
  return `${String(adjustedDate.getMonth() + 1).padStart(2, '0')}/${String(adjustedDate.getDate()).padStart(2, '0')}/${adjustedDate.getFullYear()}`;
};

// Endpoint to handle status updates
app.get('/update-status', (req, res) => {
  const { product, status } = req.query;

  if (!product || !status) {
    return res.status(400).send('Missing product or status parameters.');
  }

  console.log(`Received status update: ${product} -> ${status}`);

  fetchExcelFileFromGitHub().then((workbook) => {
    if (!workbook) return res.status(500).send('Error fetching Excel file from GitHub');

    const data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: "" });
    const updatedData = data.map((row) => {
      if (row['Product name'] === product) {
        row['Status'] = status; // Update the status in the sheet
      }
      return row;
    });

    // Create a new workbook and update it
    const updatedWorkbook = XLSX.utils.book_new();
    const updatedSheet = XLSX.utils.json_to_sheet(updatedData);
    XLSX.utils.book_append_sheet(updatedWorkbook, updatedSheet);

    writeExcelFileToGitHub(updatedWorkbook);

    res.send(`Status for ${product} updated to ${status}`);
  });
});

processAndSendEmails()
// Cron job to process and send emails every day at midnight (or any custom schedule)
cron.schedule('0 0 * * *', processAndSendEmails);

// Start the Express server
app.listen(PORT, () => {
  console.log(`Server running on https://email-server-pearl.vercel.app/`);
});

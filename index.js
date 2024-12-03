import express from 'express';
import XLSX from 'xlsx';
import nodemailer from 'nodemailer';
import path from 'path';
import dotenv from 'dotenv';
import cron from "node-cron";

dotenv.config();

const app = express();
const PORT = 5000;

// Environment Variables
const GMAIL_USER = process.env.GMAIL_USER; // Replace with your Gmail address
const GMAIL_PASSWORD = process.env.GMAIL_PASSWORD; // Replace with your App-specific password

// Nodemailer Transport Configuration
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: GMAIL_USER,
    pass: GMAIL_PASSWORD,
  },
});

// Path to the Excel file
const EXCEL_FILE_PATH = path.resolve('data', 'Patching.xlsx');

// Function to read and parse the Excel file
const readExcelFile = () => {
  try {
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    let data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

    // Filter out "__EMPTY" or similar keys
    data = data.map(row => {
      const filteredRow = {};
      for (const key in row) {
        if (key.trim() && !key.startsWith("__EMPTY")) {
          filteredRow[key] = row[key];
        }
      }
      return filteredRow;
    });

    return data;
  } catch (error) {
    console.error('Error reading Excel file:', error.message);
    return null;
  }
};


// Function to write data back to the Excel file
const writeExcelFile = (data) => {
  try {
    // Clean rows before writing
    const cleanedData = data.map(row => {
      const filteredRow = {};
      for (const key in row) {
        if (key.trim()) { // Exclude blank keys
          filteredRow[key] = row[key];
        }
      }
      return filteredRow;
    });

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(cleanedData);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, EXCEL_FILE_PATH);
    console.log('Excel file updated successfully.');
  } catch (error) {
    console.error('Error writing to Excel file:', error.message);
  }
};


// Function to send emails with status buttons
const sendEmail = async (to, cc, productName, timeRange) => {
  const serverUrl = `http://localhost:${PORT}`;
  const mailOptions = {
    from: GMAIL_USER,
    to: to,
    cc: cc.join(','),
    subject: `Scheduled Task for ${productName}`,
    html: `
      <p>The scheduled task for <b>${productName}</b> is planned at <b>${timeRange}</b> today.</p>
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
    console.error(`Error sending email to ${to}:`, error.message);
  }
};

// Function to process the Excel data and send emails
const processAndSendEmails = () => {
  const data = readExcelFile();
  if (!data) return;

  const now = new Date();
  const todayStr = `${String(now.getDate()).padStart(2, '0')}/${String(now.getMonth() + 1).padStart(2, '0')}/${now.getFullYear()}`;

  data.forEach((row) => {
    const scheduledDateValue = row["Scheduled date"];
    const scheduledDate = convertExcelDate(scheduledDateValue);

    if (scheduledDate === todayStr) {
      const emailAddresses = row.EmailAddresses.split(',').map((email) => email.trim());
      const to = emailAddresses[0];
      const cc = emailAddresses.slice(1);
      const productName = row['Product name'];

      // Extract the start time from the range (e.g., "09:00 AM - 12:00 PM")
      const timeRange = row.Time;
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

      if (now >= oneHourBefore && now < scheduledTime) {
        sendEmail(to, cc, productName, timeRange);
      }
    }
  });
};


// Function to convert Excel serial date to dd/mm/yyyy
const convertExcelDate = (serial) => {
  const excelEpoch = new Date(1900, 0, 1);
  const adjustedDate = new Date(excelEpoch.getTime() + (serial - 2) * 24 * 60 * 60 * 1000);
  return `${String(adjustedDate.getMonth() + 1).padStart(2, '0')}/${String(adjustedDate.getDate()).padStart(2, '0')}/${adjustedDate.getFullYear()}`;
};

// Endpoint to handle status updates
app.get('/update-status', (req, res) => {
  const { product, status } = req.query;

  if (!product || !status) {
    return res.status(400).send('Missing product or status parameters.');
  }

  const data = readExcelFile();
  if (!data) return res.status(500).send('Error reading the Excel file.');

  const updatedData = data.map((row) => {
    if (row["Product name"] === product) {
      row.Status = status;
    }
    return row;
  });

  writeExcelFile(updatedData);

  res.send(`Status for "${product}" updated to "${status}".`);
});

// Schedule the email processing task every minute to check for 1-hour logic
processAndSendEmails()


// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

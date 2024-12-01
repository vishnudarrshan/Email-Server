import express from 'express';
import XLSX from 'xlsx';
import nodemailer from 'nodemailer';
import path from 'path';
import dotenv from 'dotenv';
import cron from "node-cron"

// Load environment variables from .env file
dotenv.config();

const app = express();
const PORT = 5000;

// Gmail app password
const GMAIL_PASSWORD = "nmwd mapu jber bgas"; // Replace with your actual password

// Nodemailer Transport Configuration
const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: 'vishnudarrshanorp@gmail.com', // Your Gmail address
    pass: GMAIL_PASSWORD, // App-specific password
  },
});

// Path to the Excel file
const EXCEL_FILE_PATH = path.resolve('data', 'Book 2.xlsx'); // Adjust the filename and path if needed

// Function to read and parse the Excel file
const readExcelFile = () => {
  try {
    const workbook = XLSX.readFile(EXCEL_FILE_PATH);
    const sheetName = workbook.SheetNames[0]; // Assuming the first sheet
    const sheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(sheet, { defval: "" }); // Default empty values
    return jsonData;
  } catch (error) {
    console.error('Error reading Excel file:', error);
    return null;
  }
};

// Function to write data back to the Excel file
const writeExcelFile = (data) => {
  try {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    XLSX.writeFile(workbook, EXCEL_FILE_PATH);
    console.log('Excel file updated successfully.');
  } catch (error) {
    console.error('Error writing to Excel file:', error);
  }
};

// Function to send emails with status buttons
const sendEmail = async (to, cc, productName, timeRange) => {
  const serverUrl = `http://localhost:${PORT}`; // Server URL
  const mailOptions = {
    from: 'vishnudarrshanorp@gmail.com',
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
  const data = readExcelFile();
  if (!data) return;

  // Get tomorrow's date as dd/mm/yyyy
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = `${String(tomorrow.getDate()).padStart(2, '0')}/${String(tomorrow.getMonth() + 1).padStart(2, '0')}/${tomorrow.getFullYear()}`;

  console.log(`Tomorrow's Date: ${tomorrowStr}`);

  data.forEach((row) => {
    const scheduledDateValue = row["Scheduled date"]; // Excel date (serial number)
    const scheduledDate = convertExcelDate(scheduledDateValue); // Convert to dd/mm/yyyy format

    console.log(`Scheduled Date: ${scheduledDate}`);

    if (scheduledDate === tomorrowStr) {
      console.log("Match found for tomorrow:", scheduledDate);

      const emailAddresses = row.EmailAddresses.split(',').map((email) => email.trim());
      const to = emailAddresses[0]; // First email as 'to'
      const cc = emailAddresses.slice(1); // Remaining emails as 'cc'
      const timeRange = row.Time;
      const productName = row['Product name'];

      if (to) {
        sendEmail(to, cc, productName, timeRange);
      } else {
        console.error('No primary email address provided in row:', row);
      }
    } else {
      console.log("No match for tomorrow:", scheduledDate);
    }
  });
};

// Helper function to convert Excel serial date to dd/mm/yyyy
const convertExcelDate = (serial) => {
  // Excel serial date starts from Jan 1, 1900
  const excelEpoch = new Date(1900, 0, 1);
  const adjustedDate = new Date(excelEpoch.getTime() + (serial - 2) * 24 * 60 * 60 * 1000); // Subtract 2 for Excel quirks
  return `${String(adjustedDate.getDate()).padStart(2, '0')}/${String(adjustedDate.getMonth() + 1).padStart(2, '0')}/${adjustedDate.getFullYear()}`;
};












// Endpoint to handle status updates
app.get('/update-status', (req, res) => {
  const { product, status } = req.query;

  if (!product || !status) {
    return res.status(400).send('Missing product or status parameters.');
  }

  console.log(`Received status update: ${product} -> ${status}`);

  const data = readExcelFile();
  const updatedData = data.map((row) => {
    if (row["Product name"] === product) {
      row.Status = status; // Add or update "Status" column
    }
    return row;
  });

  writeExcelFile(updatedData);

  res.send(`Status for "${product}" updated to "${status}".`);
});

// Process emails immediately (for testing) and schedule the task daily at midnight
processAndSendEmails(); // Send emails on server start
// Schedule the task
cron.schedule('0 0 * * *', processAndSendEmails);

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});





import express from 'express';
import XLSX from 'xlsx';
import axios from 'axios';
import dotenv from 'dotenv';
import cron from 'node-cron';

dotenv.config();

const app = express();
const PORT = 5000;

// GitHub Configurations
const GITHUB_REPO = "https://github.com/vishnudarrshan/Email-Server.git";
const EXCEL_FILE_PATH = "data/Patching.xlsx"; // Example: data/Book2.xlsx
const GITHUB_TOKEN = process.env.GITHUB_TOKEN; // Set this in your environment or .env file

// Function to fetch Excel file from GitHub
const fetchExcelFile = async () => {
  try {
    const url = `https://raw.githubusercontent.com/${GITHUB_REPO}/main/${EXCEL_FILE_PATH}`;
    const response = await axios.get(url, { responseType: 'arraybuffer' });
    const workbook = XLSX.read(response.data, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    return XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
  } catch (error) {
    console.error("Error fetching Excel file:", error.message);
    throw error;
  }
};

// Function to update Excel file in GitHub
const updateExcelFile = async (data) => {
  try {
    // Convert JSON data to Excel workbook
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

    // Write Excel file to buffer
    const updatedExcelBuffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });

    // Get file's SHA for update
    const getUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${EXCEL_FILE_PATH}`;
    const getResponse = await axios.get(getUrl, {
      headers: {
        Authorization: `token ${GITHUB_TOKEN}`,
      },
    });

    const fileSha = getResponse.data.sha;

    // Update file using GitHub API
    const updateUrl = `https://api.github.com/repos/${GITHUB_REPO}/contents/${EXCEL_FILE_PATH}`;
    const updateResponse = await axios.put(
      updateUrl,
      {
        message: "Update Excel file via API",
        content: updatedExcelBuffer.toString("base64"),
        sha: fileSha,
      },
      {
        headers: {
          Authorization: `token ${GITHUB_TOKEN}`,
        },
      }
    );

    console.log("Excel file updated successfully:", updateResponse.data.commit.sha);
  } catch (error) {
    console.error("Error updating Excel file:", error.message);
    throw error;
  }
};

// Endpoint to handle status updates
app.get('/update-status', async (req, res) => {
  const { product, status } = req.query;

  if (!product || !status) {
    return res.status(400).send('Missing product or status parameters.');
  }

  try {
    const data = await fetchExcelFile();
    const updatedData = data.map((row) => {
      if (row["Product name"] === product) {
        row.Status = status; // Update the status column
      }
      return row;
    });

    await updateExcelFile(updatedData);

    res.send(`Status for "${product}" updated to "${status}".`);
  } catch (error) {
    console.error("Error in updating status:", error.message);
    res.status(500).send("Failed to update status.");
  }
});

// Schedule emails daily at midnight
cron.schedule('0 0 * * *', async () => {
  console.log("Scheduled task started...");
  try {
    const data = await fetchExcelFile();
    console.log("Fetched data for email processing:", data);
    // Add your email sending logic here...
  } catch (error) {
    console.error("Error during scheduled task:", error.message);
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});


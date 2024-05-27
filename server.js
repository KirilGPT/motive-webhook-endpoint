const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');

const app = express();
const port = 3000;

app.use(bodyParser.json());

app.get('/', (req, res) => {
  res.send('Webhook endpoint is running.');
});

app.post('/', (req, res) => {
  console.log('Received a POST request');
  console.log(req.body);

  // Extract data from the request body
  const data = req.body;

  // Check if data is valid
  if (!data || Object.keys(data).length === 0) {
    console.error('Received invalid data:', data);
    return res.status(400).send('Invalid data received');
  }

  try {
    // Create a new workbook and worksheet
    const workbook = xlsx.utils.book_new();
    const worksheet = xlsx.utils.json_to_sheet([data]);

    // Add worksheet to workbook
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Define the path where the file will be saved
    const filePath = 'webhook_data.xlsx';

    // Save workbook to file
    xlsx.writeFile(workbook, filePath);

    console.log('File saved to', filePath);
    res.send('POST request received and data saved to Excel file');
  } catch (err) {
    console.error('Error processing data:', err);
    res.status(500).send('Error saving data to Excel file');
  }
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

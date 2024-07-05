const express = require('express');
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');
const app = express();
const port = 3000;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Serve static files from the 'public' directory
app.use(express.static('public'));

// Handle form submission and save data to Excel
app.post('/submitForm', (req, res) => {
    const data = req.body;

    // Load existing data if the file exists
    let existingData = [];
    if (fs.existsSync('nurse_report.xlsx')) {
        const workbook = xlsx.readFile('nurse_report.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        existingData = xlsx.utils.sheet_to_json(worksheet);
    }

    // Add new data
    existingData.push(data);

    // Write data to Excel file
    const worksheet = xlsx.utils.json_to_sheet(existingData);
    const workbook = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Nurse Report');
    xlsx.writeFile(workbook, 'nurse_report.xlsx');

    res.redirect('/summary.html');
});

// Endpoint to handle Excel file download
app.get('/downloadExcel', (req, res) => {
    const file = 'nurse_report.xlsx';
    if (fs.existsSync(file)) {
        res.download(file);
    } else {
        res.status(404).send('File not found');
    }
});

// Serve the main form page
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});

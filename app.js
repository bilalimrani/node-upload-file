const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Set EJS as template engine (optional for rendering HTML)
app.set('view engine', 'ejs');

// Set up static folder to serve HTML files or styles
app.use(express.static('public'));

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  },
});

const upload = multer({ storage: storage });

// Render the file upload form
app.get('/', (req, res) => {
  res.render('index'); // Render index.ejs file from views folder
});

// Handle file upload and read XLSX content
app.post('/upload', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  const filePath = path.join(__dirname, 'uploads', req.file.filename);

  // Read XLSX file
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; // Get the first sheet
  const worksheet = workbook.Sheets[sheetName];

  // Convert the sheet data to JSON format (array of objects)
  const jsonData = xlsx.utils.sheet_to_json(worksheet);

  // Send the parsed data as JSON response
  res.json(jsonData);
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});

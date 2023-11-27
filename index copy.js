const express = require('express');
const multer = require('multer');
const path = require('path');
const app = express();
const port = 4002;

// Set up a storage engine to define where files will be stored
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'public/uploads/'); // Uploads will be stored in the 'uploads/' directory
  },
  filename: (req, file, cb) => {
    cb(null, file.originalname);
  },
});

// Create an upload instance with the defined storage engine
const upload = multer({ storage: storage });

// Serve static files from the 'uploads/' directory
app.use(express.static('public/uploads'));
app.use(express.static('public'));

// Define a route for the file upload form
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// Handle file upload
app.post('/upload', upload.single('file'), (req, res) => {
  res.send('File uploaded successfully');
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});

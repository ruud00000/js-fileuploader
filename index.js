const express = require('express')
const fileUpload = require('express-fileupload')
const path = require('path')

const app = express()
const port = 4002

// Serve static files from the 'uploads/' directory
app.use(express.static('public'))

app.use(fileUpload())

// Serve your HTML page
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'))
});

// File upload endpoint
app.post('/upload', (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).json({ message: 'Er zijn geen bestanden geupload.' })
    return
  }

  const uploadedFile = req.files.file;
  const uploadPath = path.join(__dirname, 'public/uploads', uploadedFile.name);

  uploadedFile.mv(uploadPath, (err) => {
    if (err) {
      return res.status(500).send(err)
    }

    // After successful file upload, update the message
    const newMessage = 'Het bestand is geupload!'

    res.json({ message: newMessage })
  });
});

const server = app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`)
});

module.exports = server
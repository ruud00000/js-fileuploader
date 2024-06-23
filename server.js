import express from 'express'
import fileUpload from 'express-fileupload'
import { fileURLToPath } from 'url'
import path from 'path'
import cors from 'cors'
import dotenv from 'dotenv'
// for convert:
import ExcelJS from 'exceljs'
import { PDFDocument, StandardFonts } from 'pdf-lib'
import fs from 'fs/promises'
import { sendMail } from 'wkb-utils'


const app = express()
// Get the file URL of the current module
const __filename = fileURLToPath(import.meta.url);
// Get the directory name of the current module
const __dirname = path.dirname(__filename);
const outputFolder = path.join(__dirname, 'public/pdfs')
let newMessage = ''

// init dotenv
dotenv.config()

// get environment variables
const PORT = process.env.PORT || 3002
const ORIGIN_URL_1 = process.env.ORIGIN_URL_1 || '' 
const ORIGIN_URL_2 = process.env.ORIGIN_URL_2 || '' 
const ORIGIN_URL_3 = process.env.ORIGIN_URL_3 || '' 

const allowedOrigins = [
  ORIGIN_URL_1, 
  ORIGIN_URL_2,
  ORIGIN_URL_3];

// CORS middleware configuration
const corsOptions = {    
  origin: (origin, callback) => {
  if (allowedOrigins.includes(origin) || !origin) {
      callback(null, true)
  } else {
      callback(new Error('Not allowed by CORS'))
  }
},
  methods: ['GET', 'POST'], // Add other HTTP methods as needed    
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: true
}

app.use(cors(corsOptions))

// init mail values
const to = process.env.MAILTO  
const subject = 'WKB upload'
let mailText = ''

// Serve static files from the 'uploads/' directory
app.use(express.static('public'))
app.use(fileUpload())

// Serve HTML page
app.get('/',   
  (req, res, next) => {
    next()
  }, 
  (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'))
})

// File download endpoint
app.get('/file/:filename',   
(req, res, next) => {
  next()
}, 
(req, res) => {
  const { filename } = req.params
  const filePath = path.join(outputFolder, filename)
  console.log(filePath)
  res.sendFile(filePath, (err) => {
    if (err) {
      console.error('Error sending file:', err);
      res.status(500).send('Internal Server Error');
    } 
  })
})

// File upload endpoint
app.post('/upload', (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).json({ message: 'Er zijn geen bestanden geupload.' })
    return
  }

  const uploadedFile = req.files.file
  const uploadPath = path.join(__dirname, 'public/uploads', uploadedFile.name)

  uploadedFile.mv(uploadPath, (err) => {
    if (err) {
      return res.status(500).send(err)
    }

    // After successful file upload, update the message
    newMessage = 'Het bestand is geupload!'

    res.json({ message: newMessage })
  })
})


// Convert Excel worksheets from workbook to pdfs
app.get('/convert',   
  (req, res, next) => {
    next()
  }, 
  async (req, res) => { 
    const file = path.join(__dirname, 'public/uploads', 'competitie2023.xlsx')
    newMessage = "Excel bestand wordt geconverteerd..."
    await excelToPdf(file, outputFolder, file)
      
    res.json({ message: newMessage })
  
})

// File upload endpoint
app.post('/uploadconvert', 
  (req, res, next) => {
    next()
  }, 
  async (req, res) => {
  if (!req.files || Object.keys(req.files).length === 0) {
    res.status(400).json({ message: 'Er zijn geen bestanden geupload.' })
    return
  }

  const uploadedFile = req.files.file
  // const uploadPath = path.join(__dirname, 'public/uploads', uploadedFile.name)
  const newName = 'competitie2023' + path.extname(uploadedFile.name);
  const uploadPath = path.join(__dirname, 'public/uploads', newName)

  const moveFile = () => {
    return new Promise((resolve, reject) => {
      uploadedFile.mv(uploadPath, (err) => {
        if (err) {
          reject(err)
        } else {
          resolve()
        }
      })
    })
  }

  try {
    // Move the uploaded file and wait for the operation to finish
    await moveFile();
  
    const excelFilePath = uploadPath  
    
    console.log('bestand ', excelFilePath, 'converteren...')
  
    // convert uploaded file
    await excelToPdf(excelFilePath, outputFolder, uploadedFile)
  
    // After successful file upload, update the message
    newMessage = 'Het bestand is geupload en geconverteerd!' 
    mailText += `Bestand ${excelFilePath} werd geconverteerd.\n`
  
    await sendMail(to, subject, mailText)
    mailText = ''
    // Send a response or perform other actions
    res.json({ message: newMessage })
  } catch (error) {
    // Handle errors
    mailText += `Fout bij converteren: ${error}.\n`
    await sendMail(to, subject, mailText)
    mailText = '' 
    console.error('Error:', error)
    res.status(500).send(error.message || 'Internal Server Error')
  }
})

const server = app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`)
})


export default server


async function excelToPdf(excelFilePath, outputFolder, uploadedFile) {
  const workbook = new ExcelJS.Workbook()
  
  try {
    await workbook.xlsx.readFile(excelFilePath) 
    console.log('Reading Excel file')
    mailText += `Bestand ${uploadedFile.name} wordt gelezen.\n`
  } catch (error) {
    console.error('Error reading Excel file:', error.message)  
    mailText += `Fout bij lezen Excel bestand: ${error.nmessage}.\n`
    return 
  }
  let result = true
  let wsCounter = 0
  // Kolombreedtes handmatig opgeven:
  const colWidths = [
    // individueel: 
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40], 
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    // teams:
    [40, 150, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],  
    [40, 150, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40],
    [40, 150, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40, 40]
  ]
  const wsNames = [
    'CHA', 'CHB', 'CHC1', 'CHC2', 
    'CDA', 'CDB', 'CDC1', 'CDC2', 
    'CHT1', 'CHT2', 'CDT1', 'CDT2'
  ]

  const worksheets = workbook.worksheets
  console.log('aantal werkbladen in ', excelFilePath, ': ', worksheets.length)
  mailText+= `Aantal werkbladen in ${excelFilePath}: ${worksheets.length}.\n`

  for (let i = 0; i < worksheets.length; i++) {
    if (!(worksheets[i].state == 'hidden') && i != 10) {  // sheet10 is invoersheet
      
      const worksheet = worksheets[i]
      
      const pdfDoc = await PDFDocument.create()

      const pdfPage = pdfDoc.addPage([841.89, 595.28])
      const { width, height } = pdfPage.getSize()
      const pdfFont = await pdfDoc.embedFont(StandardFonts.Helvetica)

      const cellHeight = 17

      let x = 20 // linkermarge
      let text = ''

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (!row.hidden) {
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const col = worksheet.getColumn(colNumber)
            if (!col.hidden) {
              if (cell.text && cell.style.numFmt === '0.00') {
                text = formatCell(cell.text)
              } else {
                text = cell.text ? cell.text : ''
              }
              const y = height - (rowNumber) * cellHeight
              pdfPage.drawText(text, { x, y, font: pdfFont, size: 12 })   
              x += colWidths[wsCounter][colNumber-1]   
            }
          })
        }
        x = 20
      })

      const pdfBytes = await pdfDoc.save()
      const pdfFileName = `${outputFolder}/${wsNames[wsCounter]}.pdf`

      try {
        await fs.writeFile(pdfFileName, pdfBytes)
        console.log(`PDF file created: ${pdfFileName}`)        
      } catch (error) {
        console.error(`Error writing PDF file ${pdfFileName}:`, error.message)        
        result = false
      }
    
      // After successful conversion, update the message
      if (result) {
        //newMessage = 'Excel bestand is geconverteerd naar pdf-bestand!'
        mailText += `PDF bestand ${pdfFileName} is gemaakt.\n`
      } else {
        //newMessage = 'Converteren naar pdf-bestand is niet gelukt.'
        mailText += `Fout bij maken PDF bestand ${pdfFileName}: ${error.message}.\n`
      }
      wsCounter++
    }
  }
  newMessage = 'Het bestand is geconverteerd!'
}

function formatCell(text) {
  //let result = text.substring(0,4)
  let position = text.indexOf('.')
  const result = position == -1 ? text + '.00' : text.substring(0,position+3)
  return result
}
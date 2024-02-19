const express = require('express')
const fileUpload = require('express-fileupload')
const path = require('path')
const cors = require('cors')
const dotenv = require('dotenv')
// for convert:
const ExcelJS = require('exceljs')
const { PDFDocument, StandardFonts } = require('pdf-lib')
const fs = require('fs/promises')

const app = express()
const port = 3002
const outputFolder = path.join(__dirname, 'public/pdfs')
let newMessage = ''

// init dotenv
dotenv.config()

// init mail values
const to = process.env.MAILTO  
const subject = 'WKB upload'
let mailText = ''

// Serve static files from the 'uploads/' directory
app.use(express.static('public'))
app.use(cors())
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
  (req, res) => { 
  excelToPdf(excelFilePath, outputFolder)

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

  const sendMail = (to, subject, mailtext) => {
    return new Promise(async (resolve, reject) => {
      try {
        const payload = {
          to: to,
          subject: subject,
          text: mailtext
      }
        const response = await fetch(`${process.env.MAIL_URL}/send-email`, {        
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify(payload)
        })
  
        if (!response.ok) {
          throw new Error('Mail sturen is niet gelukt.')
        }
  
        const data = await response.json()
    
      } catch (error) {
        console.error(error)
      }
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
    // newMessage = 'Het bestand is geupload en geconverteerd!' 
    // mailText += `Bestand ${excelFilePath} werd geconverteerd.\n`
  
    // Send a response or perform other actions
    res.json({ message: newMessage })
    sendMail(to, subject, mailText)
    mailText = ''

  } catch (error) {
    // Handle errors
    console.error('Error:', error)
    res.status(500).send(error.message || 'Internal Server Error')
    mailText += `Fout bij converteren: ${error}.\n`
    sendMail(to, subject, mailText)
    mailText = '' 
  }
})

const server = app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`)
})

module.exports = server



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
        mailText += `PDF bestand gemaakt: ${pdfFileName}.\n`
      } catch (error) {
        console.error(`Error writing PDF file ${pdfFileName}:`, error.message)
        mailText += `Fout bij maken PDF bestand ${pdfFileName}: ${error.message}.\n`
        result = false
      }
    
      // After successful conversion, update the message
      if (result) {
        newMessage = 'Excel bestand is geconverteerd naar pdf-bestanden!'
        mailText += `Excel bestand is geconverteerd naar pdf-bestanden.\n`
      } else {
        newMessage = 'Converteren naar pdf-bestanden is niet gelukt.'
        mailText += `Converteren naar pdf-bestanden is niet gelukt.\n`
      }
      wsCounter++
    }
  }
}

function formatCell(text) {
  //let result = text.substring(0,4)
  let position = text.indexOf('.')
  result = position == -1 ? text + '.00' : text.substring(0,position+3)
  return result
}
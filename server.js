const express = require('express')
const fileUpload = require('express-fileupload')
const path = require('path')
const cors = require('cors')
// for convert:
const ExcelJS = require('exceljs')
const { PDFDocument, StandardFonts } = require('pdf-lib')
const fs = require('fs/promises')

const app = express()
const port = 3002
/*
    add_header Access-Control-Allow-Origin *;
    add_header Access-Control-Allow-Methods 'POST, GET, OPTIONS';
    add_header Access-Control-Allow-Headers Content-Type;
*/
  // Example usage
  // const excelFilePath = path.join(__dirname, 'public/uploads', 'excel-file.xlsx')
  const outputFolder = path.join(__dirname, 'public/pdfs')

// Serve static files from the 'uploads/' directory
app.use(express.static('public'))
//app.use(cors())
app.use(fileUpload())

// Serve your HTML page
app.get('/',   
  (req, res, next) => {
    next()
  }, 
  (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'))
  })

  let newMessage = ''

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
  }
  )


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
  const uploadPath = path.join(__dirname, 'public/uploads', newName);

  const moveFile = () => {
    return new Promise((resolve, reject) => {
      uploadedFile.mv(uploadPath, (err) => {
        if (err) {
          reject(err);
        } else {
          resolve();
        }
      });
    });
  }

  try {
    // Move the uploaded file and wait for the operation to finish
    await moveFile();
  
    const excelFilePath = uploadPath;
  
    // Wait for 5 seconds
    // console.log('even wachten...');
    // await new Promise(resolve => setTimeout(resolve, 5000));
  
    // Log messages after waiting
    console.log('bestand ', excelFilePath, 'converteren...');
  
    // convert uploaded file
    await excelToPdf(excelFilePath, outputFolder);
  
    // After successful file upload, update the message
    newMessage = 'Het bestand is geupload en geconverteerd!' 
  
    // Send a response or perform other actions
    res.json({ message: newMessage });
  } catch (error) {
    // Handle errors
    console.error('Error:', error);
    res.status(500).send(error.message || 'Internal Server Error');
  }


})

const server = app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`)
})

module.exports = server



async function excelToPdf(excelFilePath, outputFolder) {
  const workbook = new ExcelJS.Workbook()
  
  try {
    await workbook.xlsx.readFile(excelFilePath) 
    console.log('Reading Excel file')
  } catch (error) {
    console.error('Error reading Excel file:', error.message)  
    return 
  }
  let result = true
  const worksheets = workbook.worksheets
  console.log('aantal werkbladen in ', excelFilePath, ': ', worksheets.length)

  for (let i = 0; i < worksheets.length; i++) {
    if (!(worksheets[i].state == 'hidden') && i != 10) {  // sheet10 is invoersheet
      const worksheet = worksheets[i]
      
      const pdfDoc = await PDFDocument.create()

      const pdfPage = pdfDoc.addPage([841.89, 595.28])
      const { width, height } = pdfPage.getSize()
      const pdfFont = await pdfDoc.embedFont(StandardFonts.Helvetica)

      const cellHeight = 17

      /*const colWidths = []
      let colWidth = 0
      worksheet.columns.forEach((column, colNumber) => {
        // console.log(`Column ${colNumber}: Width - ${column.width}`)
        if (typeof column.width == 'undefined') {
          colWidth = 75
        } else {
          colWidth = column.width
        }
        // const colWidth = typeof column.width == 'undefined' ? 75 : column.width
        colWidths.push(colWidth)
      })*/
      // Kolombreedtes handmatig opgeven:
      const colWidths = [ 
        40, 150, 40, 40, 100, 40, 40, 40, 40, 40, 
        40, 40, 40, 40, 40, 40, 40, 40, 40, 40,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
      let x = 0
      let text = ''

      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        if (!row.hidden) {
          //cellWidth = worksheet.getColumn(colNumber).width
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const col = worksheet.getColumn(colNumber)
            if (!col.hidden) {
              if (cell.text && cell.style.numFmt === '0.00') {
                // console.log('cell style: ', cell.style, 'value: ', cell.text, 'format: ', formatCell(cell.text))
                text = formatCell(cell.text)
              } else {
                text = cell.text ? cell.text : ''
              }
              
              //console.log('text: ', text, 'type: ', cell.type)
              // const x = (colNumber - 1) * cellWidth
              const y = height - (rowNumber) * cellHeight

              pdfPage.drawText(text, { x, y, font: pdfFont, size: 12 })
          
              x += colWidths[colNumber-1]
            }
          })
        }
        x = 0
      })

      const pdfBytes = await pdfDoc.save()
      const pdfFileName = `${outputFolder}/worksheet_${i + 1}.pdf`

      try {
        await fs.writeFile(pdfFileName, pdfBytes)
        console.log(`PDF file created: ${pdfFileName}`)
      } catch (error) {
        console.error(`Error writing PDF file ${pdfFileName}:`, error.message)
        result = false
      }
    
      // After successful conversion, update the message
      if (result) {
        newMessage = 'Excel bestand is geconverteerd naar pdf-bestanden!'
      } else {
        newMessage = 'Converteren naar pdf-bestanden is niet gelukt.'
      }
    }
  }
}

function isTwoDecimals(value) {
  return !isNaN(value) && value % 1 !== 0 && value.toString().split('.')[1].length === 2;
}

function formatCell(text) {
  //let result = text.substring(0,4)
  let position = text.indexOf('.')
  result = position == -1 ? text + '.00' : text.substring(0,position+3)
  return result
}
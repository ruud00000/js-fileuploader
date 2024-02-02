const express = require('express')
const fileUpload = require('express-fileupload')
const path = require('path')
// for convert:
const ExcelJS = require('exceljs')
const { PDFDocument, StandardFonts } = require('pdf-lib')
const fs = require('fs/promises')

const app = express()
const port = 3002

// Serve static files from the 'uploads/' directory
app.use(express.static('public'))

app.use(fileUpload())

// Serve your HTML page
app.get('/', (req, res) => {
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
})

// Convert Excel worksheets from workbook to pdfs
app.get('/convert', (req, res) => {
  
  async function excelToPdf(excelFilePath, outputFolder) {
    const workbook = new ExcelJS.Workbook()
  
    try {
      await workbook.xlsx.readFile(excelFilePath)
    } catch (error) {
      console.error('Error reading Excel file:', error.message)
      return
    }
    let result = true
    const worksheets = workbook.worksheets
  
    for (let i = 0; i < worksheets.length; i++) {
      const worksheet = worksheets[i]
      const pdfDoc = await PDFDocument.create()
  
      const pdfPage = pdfDoc.addPage([841.89, 595.28])
      const { width, height } = pdfPage.getSize()
      const pdfFont = await pdfDoc.embedFont(StandardFonts.Helvetica)
  
      const cellHeight = 20

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
        20, 20, 150, 40, 130, 20, 30, 30, 300, 40, 
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0,
        0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
      let x = 0
      
      worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
        //cellWidth = worksheet.getColumn(colNumber).width
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const text = cell.value ? cell.value.toString() : ''
          // const x = (colNumber - 1) * cellWidth
          const y = height - (rowNumber) * cellHeight
  
          pdfPage.drawText(text, { x, y, font: pdfFont, size: 12 })
      
          x += colWidths[colNumber-1]
        })
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
    }
    // After successful conversion, update the message
    if (result) {
      newMessage = 'Excel bestand is geconverteerd naar pdf-bestanden!'
    } else {
      newMessage = 'Converteren naar pdf-bestanden is niet gelukt.'
    }

    res.json({ message: newMessage })
  }
  
  // Example usage
  const excelFilePath = path.join(__dirname, 'public/uploads', 'excel-file.xlsx')
  const outputFolder = path.join(__dirname, 'public/pdfs')
  
  excelToPdf(excelFilePath, outputFolder)
  
})

const server = app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`)
})

module.exports = server
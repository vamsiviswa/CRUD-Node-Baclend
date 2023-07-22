const express = require('express');
const XlsxPopulate = require('xlsx-populate');
const bodyParser = require('body-parser');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 5000;
const EXCEL_FILE = 'data.xlsx'; 

app.use(cors());


app.use(bodyParser.json());


function readDataFromExcel() {
  return XlsxPopulate.fromFileAsync(EXCEL_FILE)
    .then(workbook => {
      const sheet = workbook.sheet(0);
      const rows = sheet.usedRange().value();
      return rows;
    })
    .catch(error => {
      console.error('Error reading data from Excel:', error);
      return [];
    });
}


function writeDataToExcel(data) {
  return XlsxPopulate.fromBlankAsync()
    .then(workbook => {
      const sheet = workbook.sheet(0);
      sheet.cell('A1').value(['ID', 'Name', 'Age', 'Email', 'Mobile']); 
      data.forEach((row, index) => {
        sheet.cell(`A${index + 2}`).value(row);
      });
      return workbook.toFileAsync(EXCEL_FILE);
    })
    .catch(error => {
      console.error('Error writing data to Excel:', error);
    });
}


app.get('/api/data', (req, res) => {
  readDataFromExcel()
    .then(data => {
      res.json(data);
    })
    .catch(error => {
      res.status(500).json({ error: 'Error reading data from Excel' });
    });
});


app.post('/api/data', (req, res) => {
  readDataFromExcel()
    .then(data => {
      const newData = req.body;
      data.push(newData);
      return writeDataToExcel(data);
    })
    .then(() => {
      res.json({ message: 'Data added successfully' });
    })
    .catch(error => {
      res.status(500).json({ error: 'Error adding data to Excel' });
    });
});


app.put('/api/data/:id', (req, res) => {
  const idToUpdate = parseInt(req.params.id);

  readDataFromExcel()
    .then(data => {
      const updatedData = req.body;
      data[idToUpdate - 1] = updatedData; 
      return writeDataToExcel(data);
    })
    .then(() => {
      res.json({ message: 'Data updated successfully' });
    })
    .catch(error => {
      res.status(500).json({ error: 'Error updating data in Excel' });
    });
});


app.delete('/api/data/:id', (req, res) => {
  const idToDelete = parseInt(req.params.id);

  readDataFromExcel()
    .then(data => {
      data.splice(idToDelete - 1, 1); 
      return writeDataToExcel(data);
    })
    .then(() => {
      res.json({ message: 'Data deleted successfully' });
    })
    .catch(error => {
      res.status(500).json({ error: 'Error deleting data from Excel' });
    });
});


app.listen(PORT, () => {
  console.log(`Server is running on port:${PORT}`);
});
const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const bodyParser = require('body-parser');
const multer = require('multer');
const XLSX = require('xlsx');
const WebSocket = require('ws');
const moment = require('moment');


const app = express();
app.use(cors()); // This enables CORS for all routes and origins
app.use(bodyParser.json());


mongoose.connect('mongodb://127.0.0.1:27017/excelData')
    .then(() => console.log('MongoDB connected'))
    .catch(err => console.log('MongoDB connection error:', err));
  

// Connect to MongoDB and other routes...
const excelFileSchema = new mongoose.Schema({
    fileName: { type: String, required: true },
    data: { type: Array, required: true } // Ensure this matches the data format
  });
  
  const ExcelFile = mongoose.model('ExcelFile', excelFileSchema);
  
  app.post('/save', async (req, res) => {
    try {
      console.log(req.body);  // Log the incoming request data for debugging
      const { data } = req.body;
    //   const fileName = `excel_data_${uuidv4()}.xlsx`;
    const fileName = `excel_data_${moment().format('YYYYMMDD_HHmmss')}.xlsx`; 
      const excelFile = new ExcelFile({
        fileName: fileName,
        data: data, // Assuming 'data' is an array
        
        
      });
      console.log(data)
      await excelFile.save();
      res.status(201).json({ message: 'Spreadsheet saved successfully!' });
    } catch (error) {
      console.error('Error saving spreadsheet:', error);
      res.status(500).json({ message: 'Error saving spreadsheet', error: error.message });
    }
  });
  
  

// Example CORS setup for specific origin (e.g., 'http://localhost:3000'):
app.use(cors({
  origin: 'http://localhost:3000',
  methods: 'GET,POST', // Specify allowed methods if needed
  credentials: true, // Allow credentials if needed
}));

// Set up WebSocket and other routes...
const upload = multer();

// app.post('/save', async (req, res) => {
//     try {
//       const { data } = req.body;
  
//       // Save data to MongoDB or handle it
//       const excelFile = new ExcelFile({
//         fileName: 'parsed_data.xlsx',  // You can customize this
//         data: Buffer.from(JSON.stringify(data)),  // Store as Buffer if needed
//       });
  
//       await excelFile.save();
//       res.status(200).json({ message: 'Data saved successfully!' });
//     } catch (error) {
//       res.status(500).json({ message: 'Error saving data', error });
//     }
//   });
  
  
  // Route to fetch all Excel file metadata
  app.get('/gallery', async (req, res) => {
    try {
      const files = await ExcelFile.find().select('fileName createdAt');
      res.status(200).json(files);
    } catch (error) {
      res.status(500).json({ message: 'Error fetching files', error });
    }
  });


//   app.get('/excel-files', async (req, res) => {
//     try {
//       const files = await ExcelFile.find().select('fileName createdAt');
//       res.status(200).json(files);
//       console.log(files)
//     } catch (error) {
//       res.status(500).json({ message: 'Error fetching files', error });
//     }
//   });
  
  // Route to fetch a specific Excel file by ID and return its content as JSON
  app.get('/excel-file/:id', async (req, res) => {
    try {
      const file = await ExcelFile.findById(req.params.id);
  
      if (!file) {
        return res.status(404).json({ message: 'File not found' });
      }
  
      // Parse the binary data into an Excel workbook using XLSX
      const workbook = XLSX.read(file.data, { type: 'buffer' });
      const firstSheetName = workbook.SheetNames[0];
      const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
  
      res.status(200).json({ fileName: file.fileName, data: worksheet });
    } catch (error) {
      res.status(500).json({ message: 'Error fetching file', error });
    }
  });

// Start the server
app.listen(5000, () => {
  console.log('Server running on port 5000');
});

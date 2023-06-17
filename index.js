const express = require('express');
const xlsx = require('xlsx');
const fs = require('fs');
const XlsxPopulate = require('xlsx-populate');
const _ = require('lodash');
const app = express();
const port = 3000;


const prepareData = async () => {
    try {
        // Read the Excel file
        const workbook = await XlsxPopulate.fromFileAsync("./input.xlsx");
    
        // Get the sheet names
        const sheetNames = workbook.sheets();
    
        // Process each sheet
        const jsonData = [];
        sheetNames.forEach((sheet) => {
          
    
          // Extract data from each row
          const data = [];
          sheet.usedRange().value().forEach((row, index) => {
            if (index !== 0) { // Exclude the header row
              const [employeeName, dateString, hrs, task] = row;
    
              // Derive the day from the date
              const date = new Date(dateString);
              const day = date.toLocaleDateString('en-US', { weekday: 'long' });
    
              // Create an object with the extracted data
              const item = { employeeName, date, day, hrs, task };
              data.push(item);
            }
          });
    
          // Group the data by employeeName
          const groupedData = _.groupBy(data, 'employeeName');
    
          // Add the grouped data to jsonData
          _.forEach(groupedData, (group, employeeName) => {
            jsonData.push({ sheetName: employeeName, data: group });
          });          
        });
    
        return jsonData;

      } catch (error) {
        console.error('Error:', error);        
      }
}

// Set EJS as the view engine
app.set('view engine', 'ejs');

// Serve static files from the "public" directory
app.use(express.static('public'));

// Render the landing page
app.get('/', (req, res) => {
  res.render('index');
});

// Render the preview page
app.get('/preview', (req, res) => {
  res.render('preview', { sheetsData: jsonData });
});

// Handle the generate request
app.get('/generate', (req, res) => {

    // Set the JSON data for each sheet
let jsonData =  prepareData();

  // Create a new workbook
  const workbook = xlsx.utils.book_new();

  // Loop through each sheet data
  jsonData.forEach((sheetData) => {
    // Create a new worksheet
    const worksheet = xlsx.utils.json_to_sheet(sheetData.data);

    // Add the worksheet to the workbook
    xlsx.utils.book_append_sheet(workbook, worksheet, sheetData.sheetName);
  });

  // Generate the XLSX file buffer
  const buffer = xlsx.write(workbook, { type: 'buffer', bookType: 'xlsx' });

  // Save the XLSX file
  const filename = 'my_spreadsheet.xlsx'; // Change the filename as needed
  fs.writeFileSync(filename, buffer);

  res.download(filename, (err) => {
    if (err) {
      console.error('Error:', err);
      res.status(500).send('An error occurred');
    }

    // Delete the XLSX file after download
    fs.unlinkSync(filename);
  });
});



// Start the server
app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});



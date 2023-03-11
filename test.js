const express = require('express');
const app = express();
const bodyParser = require('body-parser');
const xlsx = require('xlsx');
const fs = require('fs');

// Parse incoming request bodies in a middleware before your handlers, available under the req.body property.
app.use(bodyParser.urlencoded({ extended: true }));

// Define a route for the sign-up form
app.get('/signup', function(req, res) {
  res.sendFile(__dirname + '/signup.html');
});

// Define a route for the form submission
app.post('/signup', function(req, res) {
  // Read the existing data from the Excel file
  const workbook = xlsx.readFile('data.xlsx');
  const sheetName = 'Sheet1';
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);

  // Add the new data to the array
  data.push(req.body);

  // Write the updated data back to the Excel file
  const newWorkbook = xlsx.utils.book_new();
  const newWorksheet = xlsx.utils.json_to_sheet(data);
  xlsx.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);
  fs.writeFileSync('data.xlsx', xlsx.write(newWorkbook, { type: 'buffer' }), 'binary');

  // Send a response to the user
  res.send('Thank you for signing up!');
});

// Start the server
app.listen(3000, function() {
  console.log('Server started on port 3000');
});

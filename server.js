// server.js
// where your node app starts

// init project
const express = require('express');
const app = express();
const excel = require('node-excel-export');

// we've started you off with Express, 
// but feel free to use whatever libs or frameworks you'd like through `package.json`.

// http://expressjs.com/en/starter/static-files.html
app.use(express.static('public'));

// http://expressjs.com/en/starter/basic-routing.html
app.get('/', function(request, response) {
  response.sendFile(__dirname + '/views/index.html');
});



app.get('/xl', function(request, response) {
 
    // You can define styles as json object
    const styles = {
      headerDark: {
        fill: {
          fgColor: {
            rgb: 'FF000000'
          }
        },
        font: {
          color: {
            rgb: 'FFFFFFFF'
          },
          sz: 14,
          bold: true,
          underline: true
        }
      },
      cellPink: {
        fill: {
          fgColor: {
            rgb: 'FFFFFF'
          }
        }
      },
      cellGreen: {
        fill: {
          fgColor: {
            rgb: 'FFFFFF'
          }
        }
      }
    };

    //Array of objects representing heading rows (very top)
    const heading = [
      [{value: 'a1', style: styles.headerDark}, {value: 'b1', style: styles.headerDark}, {value: 'c1', style: styles.headerDark}],
      ['a2', 'b2', 'c2'] // <-- It can be only values
    ];

    //Here you specify the export structure
    const specification = {
      added: { // <- the key should match the actual data key
        displayName: 'Added', // <- Here you specify the column header
        headerStyle: styles.headerDark, // <- Header style
        cellStyle: function(value, row) { // <- style renderer function
          // if the status is 1 then color in green else color in red
          // Notice how we use another cell value to style the current one
          return (row.status_id == 1) ? styles.cellGreen : {fill: {fgColor: {rgb: 'FFFF0000'}}}; // <- Inline cell style is possible 
        },
        width: 120 // <- width in pixels
      },
      updated: {
        displayName: 'Updated',
        headerStyle: styles.headerDark,
        cellFormat: function(value, row) { // <- Renderer function, you can access also any row.property
          return (value == 1) ? 'Active' : 'Inactive';
        },
        width: '10' // <- width in chars (when the number is passed as string)
      },
      deleted: {
        displayName: 'Deleted',
        headerStyle: styles.headerDark,
        cellStyle: styles.cellPink, // <- Cell style
        width: 220 // <- width in pixels
      }
    }

    // The data set should have the following shape (Array of Objects)
    // The order of the keys is irrelevant, it is also irrelevant if the
    // dataset contains more fields as the report is build based on the
    // specification provided above. But you should have all the fields
    // that are listed in the report specification
    const dataset = [
      {added: 'IBM', status_id: 1, updated: 'some note', deleted: 'not shown'},
     
    ]

    // Define an array of merges. 1-1 = A:1
    // The merges are independent of the data.
    // A merge will overwrite all data _not_ in the top-left cell.
    const merges = [
      { start: { row: 1, column: 1 }, end: { row: 1, column: 10 } },
      { start: { row: 2, column: 1 }, end: { row: 2, column: 5 } },
      { start: { row: 2, column: 6 }, end: { row: 2, column: 10 } }
    ]

    // Create the excel report.
    // This function will return Buffer
    const report = excel.buildExport(
      [ // <- Notice that this is an array. Pass multiple sheets to create multi sheet report
        {
          name: 'Report', // <- Specify sheet name (optional)
          heading: heading, // <- Raw heading array (optional)
          merges: merges, // <- Merge cell ranges
          specification: specification, // <- Report specification
          data: dataset // <-- Report data
        }
      ]
    );

    // You can then return this straight
    response.attachment('report.xlsx'); // This is sails.js specific (in general you need to set headers)
    return response.send(report);

// OR you can save this buffer to the disk by creating a file.
});
// listen for requests :)
const listener = app.listen(process.env.PORT, function() {
  console.log('Your app is listening on port ' + listener.address().port);
});

/*

To run this script in local machine, make sure below tools were installed locally.
- nodejs 8 or above. Follow this documentation to download nodejs https://nodejs.org/en/download/

Setting up the project
- run 'npm i' to install the dependencies

Run the script
- run 'node ./index.js'

Note: Make sure you have the gtr excel file placed in this project (root project). 
This file later will be processed to populate new excel file desired by fyqa's logic

- you can use vscode editor to edit the script and use the terminal to run the script

*/


// xlsxFile - library used to read excel file 'gtr.xlsx'
// read https://www.npmjs.com/package/read-excel-file for more information
const xlsxFile = require('read-excel-file/node');

// library used to convert object to excell format
// read https://www.npmjs.com/package/json2xls for more information
const json2xls = require('json2xls');

// node js libary to populate new file for new constructed data 'data.xlsx'
// read https://www.w3schools.com/nodejs/nodejs_filesystem.asp for more information
const fs = require('fs')

// used ramda instead of javascript library for clean code
// read https://ramdajs.com/docs/ for more information
const R = require('ramda');

xlsxFile('./gtr.xlsx').then((rows) => {

  // get GTR headers
  const headers = R.pipe(
    (data) => data[1],
    R.filter((value) => value !== null)
  )(rows);

  // map data to headers in obj format
  const gtrData = R.pipe(
    R.slice(2, Infinity),
    R.map((data) => {
      const obj = {};
      const cleanData = R.slice(1, Infinity, data);
      cleanData.map((value, index) => {
        obj[headers[index]] = value;
      });

      return obj;
    })
  )(rows);

  // sort data based on user requirements
  const sortedData = gtrData.map(({ Order, Purchased, Date, ...others }) => {
    const CUST_NAME =  Order.split(' by ')[1];

    return {
      CUST_NAME,
      // more column will be added here based on fyqa's logic
    }
  });

  // convert array of object to excel
  const xls = json2xls(sortedData);

  // populate new excel file called data.xlsx
  fs.writeFileSync('data.xlsx', xls, 'binary');

}).catch(e => console.error(e));


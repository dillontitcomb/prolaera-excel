const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');
const builder = require('./components/buildReport');

const regulator = regulators[3];
const reportData = { regulator, profile, certificates };

// create workbook, add worksheet, configure margins for 8 1/2 x 11

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('User Compliance Report');
worksheet.pageSetup.margins = {
  left: 0.25,
  right: 0.25,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3
};

// BUILD REPORT OPTIONS:

// category       - single table for given category or year, input should be specific category name
// year           - single table for given year, input should be specific year
// cycle          - single table with all categories & years included
// annual         - multiple tables organized by year
// allcategories  - multiple tables organized by category
// complete       - cycle table, annual tables, then category tables

builder.buildReport(worksheet, reportData, 'complete');

// write xlsx file

workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('Report Written');
});

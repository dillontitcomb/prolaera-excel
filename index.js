const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');

//Format profile data for export

//Header Info

const pageNumber = 1;
const cycleYears = regulators[0].cycleYears;
const name = profile.first + ' ' + profile.last;
const regName = regulators[0].name;
const licenseNum = regulators[0].license_number;
const date = new Date(regulators[0].date);
const twoYearsPrior = new Date(date.getTime() - 31556952000 * cycleYears);

const cycleEnd = `${date.getMonth() +
  1}/${date.getDate()}/${date.getFullYear()}`;
const issueDate = cycleEnd;
const cycleStart = `${twoYearsPrior.getMonth() +
  1}/${twoYearsPrior.getDate()}/${twoYearsPrior.getFullYear()}`;
const reportingPeriod = `${cycleStart} - ${cycleEnd}`;
const cycleTotal = reportingPeriod;

//Table Body

const cols = ['DATE', 'TITLE', 'SPONSOR', 'DELIVERY METHOD', 'GENERAL'];

const dynamicColumns = [];
const keys = Object.keys(regulators[0].hour_categories);
keys.forEach(key => {
  if (key !== 'hours') {
    cols.push(key.replace('_', ' ').toUpperCase());
    dynamicColumns.push(key);
  }
});

const exDate = certificates[0].date;
const exTitle = certificates[0].cert;
const exDelMeth = certificates[0].delivery;
const exSponsor = certificates[0].sponsor || certificates.sponsors.name;
const generalHours = regulators[0].hour_categories['hours'].cycle.actual;
const catHours = [];
dynamicColumns.forEach(cat => {
  catHours.push(regulators[0].hour_categories[cat].cycle.actual);
});

//Summary

const totalCreditsEarned = [];
const totalCreditsApplied = [];
const totalCPEReq = [];

//Create Excel Data

// create workbook & add worksheet
var workbook = new Excel.Workbook();
var worksheet = workbook.addWorksheet('User Compliance Report');

// create 12 columns
worksheet.columns = [
  { header: ' ', key: 'one' },
  { header: ' ', key: 'two' },
  { header: ' ', key: 'three' },
  { header: ' ', key: 'four' },
  { header: ' ', key: 'five' },
  { header: ' ', key: 'six' },
  { header: ' ', key: 'seven' },
  { header: ' ', key: 'eight' },
  { header: ' ', key: 'nine' },
  { header: ' ', key: 'ten' },
  { header: ' ', key: 'eleven' },
  { header: ' ', key: 'twelve' }
];

//Create Header Infomation Rows
const rows = [];

for (let i = 0; i < 10; i++) {
  rows.push(new Array(12));
}

rows[0][0] = regName;
rows[0][11] = 'Page ' + pageNumber;
rows[2][0] = name;
rows[2][8] = 'License #: ' + licenseNum;
rows[3][9] = 'Issue Date: ' + issueDate;
rows[4][9] = 'Reporting Period: ' + reportingPeriod;
rows[5][0] = 'Cycle Total: ' + cycleTotal;

worksheet.addRows(rows);

worksheet.mergeCells('A2:D3');
worksheet.mergeCells('A4:F6');
worksheet.mergeCells('I4:L4');
worksheet.mergeCells('I5:L5');
worksheet.mergeCells('I6:L6');
worksheet.mergeCells('A7:D8');

console.log(rows);

// save workbook to disk
workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('File Written');
});

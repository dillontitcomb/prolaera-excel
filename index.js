const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');

const certificatesDict = certificates.reduce((obj, cert) => {
  obj[cert.cert_id] = cert;
  return obj;
}, {});

//Format profile data for export

//Header Info

const pageNumber = 1;
const cycleYears = regulators[0].cycleYears;
const name = profile.last + ', ' + profile.first;
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
const cols = ['DATE', 'TITLE', 'SPONSOR', 'DELIVERY METHOD'];

const dynamicColumns = [];
const keys = Object.keys(regulators[0].hour_categories);
keys.forEach(key => {
  dynamicColumns.push(key.replace('_', ' ').toUpperCase());
});

//Summary

const totalCreditsEarned = [];
const totalCreditsApplied = [];
const totalCPEReq = [];

//Create Excel Data

// create workbook & add worksheet

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

// create header rows

const headerRows = [];
for (let i = 0; i < 5; i++) {
  headerRows.push(new Array(10));
}

headerRows[0][0] = regName;
headerRows[0][9] = 'Page ' + pageNumber;
headerRows[2][0] = name;

//create subheader rows
const subHeaderRows = [new Array(10), new Array(10)];

//create table body rows
const tableHeaderRow = [new Array(10), new Array(10)];
for (let i = 0; i < cols.length; i++) {
  tableHeaderRow[0][i] = cols[i];
}
for (let i = 0; i < dynamicColumns.length; i++) {
  let start = 10 - dynamicColumns.length;
  tableHeaderRow[0][start + i] = dynamicColumns[i];
}

//get total number of certs
const yearKeys = Object.keys(regulators[0].years);

//get all certs
const allCerts = [];
yearKeys.forEach(key => {
  const tempAppliedCerts = regulators[0].years[key].certificates_applied;
  Object.keys(tempAppliedCerts).forEach(cert_id => {
    const { cert, date, sponsor, sponsors, delivery } = certificatesDict[
      cert_id
    ];
    const tempCert = {
      cert,
      cert_id,
      date,
      sponsor,
      sponsors,
      delivery,
      hours: tempAppliedCerts[cert_id]
    };

    allCerts.push(tempCert);
  });
});

console.log(allCerts);
const tableBodyRows = [];

allCerts.forEach(cert => {
  let tempRow = [new Array(10)];
  tempRow[0] = cert.date;
  tempRow[1] = cert.cert;
  tempRow[2] = cert.sponsor || cert.sponsors.name;
  tempRow[3] = cert.delivery;
  tempRow[11 - dynamicColumns.length] = cert.hours.hours || 0;
  tempRow[11 - dynamicColumns.length - 1] = cert.hours[dynamicColumns[0]] || 0;
  tempRow[11 - dynamicColumns.length - 2] = cert.hours[dynamicColumns[1]] || 0;
  tempRow[11 - dynamicColumns.length - 3] = cert.hours[dynamicColumns[2]] || 0;
  tableBodyRows.push(tempRow);
});

console.log(tableBodyRows);

//add all rows
const allRows = headerRows
  .concat(subHeaderRows)
  .concat(tableHeaderRow)
  .concat(tableBodyRows);
worksheet.addRows(allRows);

//header styles
worksheet.mergeCells('A1:D2');
worksheet.mergeCells('A3:F5');
worksheet.mergeCells('G3:J3');
worksheet.mergeCells('G4:J4');
worksheet.mergeCells('G5:J5');

worksheet.getCell('J1').alignment = {
  vertical: 'middle',
  horizontal: 'right'
};
worksheet.getCell('A1').font = {
  size: 12,
  bold: true
};
worksheet.getCell('A1').alignment = { vertical: 'middle', horizontal: 'left' };
worksheet.getCell('A3').font = {
  size: 16,
  bold: true
};
worksheet.getCell('A3').alignment = { vertical: 'middle', horizontal: 'left' };

worksheet.getCell('G3').value = {
  richText: [
    { font: { bold: true }, text: 'License #: ' },
    { font: { bold: false }, text: licenseNum }
  ]
};
worksheet.getCell('G3').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('G4').value = {
  richText: [
    { font: { bold: true }, text: 'Issue Date: ' },
    { font: { bold: false }, text: issueDate }
  ]
};
worksheet.getCell('G4').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('G5').value = {
  richText: [
    { font: { bold: true }, text: 'Reporting Period: ' },
    { font: { bold: false }, text: reportingPeriod }
  ]
};
worksheet.getCell('G5').alignment = { vertical: 'middle', horizontal: 'right' };

//subHeader styles
worksheet.mergeCells('A6:D7');
worksheet.getCell('A6').value = {
  richText: [
    { font: { bold: true }, text: 'Cycle Total: ' },
    { font: { bold: false }, text: cycleTotal }
  ]
};
worksheet.getCell('A6').alignment = { vertical: 'middle', horizontal: 'left' };

//body styles
worksheet.mergeCells('A8:A9');
worksheet.mergeCells('B8:B9');
worksheet.mergeCells('C8:C9');
worksheet.mergeCells('D8:D9');
worksheet.mergeCells('E8:E9');
worksheet.mergeCells('F8:F9');
worksheet.mergeCells('G8:G9');
worksheet.mergeCells('H8:H9');
worksheet.mergeCells('I8:I9');
worksheet.mergeCells('J8:J9');

workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('File Written');
});

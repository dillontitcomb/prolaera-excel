const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');
const data = require('./reportDataBuilder');

console.log(data.buildReportData(profile, regulators[0], certificates));

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

//Table Body Headers
const cols = [new Array(10), new Array(10)];
cols[0][0] = 'DATE';
cols[0][1] = 'TITLE';
cols[0][3] = 'SPONSOR';
cols[0][4] = 'DELIVERY METHOD';

const dynamicColumns = [];
const keys = Object.keys(regulators[0].hour_categories);
keys.forEach(key => {
  dynamicColumns.unshift(key);
});

let dynamicColsLength = dynamicColumns.length;

for (let i = 9; dynamicColsLength > 0; i--) {
  cols[0][i] = dynamicColumns[dynamicColsLength - 1]
    .replace('_', ' ')
    .toUpperCase();
  dynamicColsLength--;
}

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
    let newDateObj = new Date(date);
    let formattedDate = `${newDateObj.getMonth() +
      1}/${newDateObj.getDate()}/${newDateObj.getFullYear()}`;
    const tempCert = {
      cert,
      cert_id,
      formattedDate,
      sponsor,
      sponsors,
      delivery,
      hours: tempAppliedCerts[cert_id]
    };

    allCerts.push(tempCert);
  });
});

//get table body length

const tableBodyLength = allCerts.length;

//fill in table rows and store hour totals
const tableBodyRows = [];

const hourTotals = {};
dynamicColumns.forEach(col => {
  hourTotals[col] = 0;
});

allCerts.forEach(cert => {
  let tempRow = [new Array(10)];
  tempRow[0] = cert.formattedDate;
  tempRow[1] = cert.cert;
  tempRow[3] = cert.sponsor || cert.sponsors.name;
  tempRow[4] = cert.delivery;
  let dynColsLen = dynamicColumns.length;
  for (let i = 9; dynColsLen > 0; i--) {
    tempRow[i] = cert.hours[dynamicColumns[dynColsLen - 1]];
    let tempCategory = [dynamicColumns[dynColsLen - 1]];
    if (typeof cert.hours[dynamicColumns[dynColsLen - 1]] === 'number')
      hourTotals[tempCategory] += cert.hours[dynamicColumns[dynColsLen - 1]];
    dynColsLen--;
  }
  tableBodyRows.push(tempRow);
});

//build table summary
const tableSummaryRows = [
  new Array(10),
  new Array(10),
  new Array(10),
  new Array(10),
  new Array(10),
  new Array(10)
];

tableSummaryRows[2][0] = 'Total Credits Applied:';
tableSummaryRows[3][0] = 'Total Credits Earned:';
tableSummaryRows[4][0] = 'Continuing Education Requirement:';
tableSummaryRows[5][0] = 'Credits Remaining:';

let length = dynamicColumns.length;
for (let i = 9; length > 0; i--) {
  let tempCategoryName = [dynamicColumns[length - 1]];
  tableSummaryRows[0][i + 1] = [dynamicColumns[length - 1]][0]
    .replace('_', ' ')
    .toUpperCase();
  tableSummaryRows[2][i] = hourTotals[tempCategoryName];
  tableSummaryRows[3][i] = hourTotals[tempCategoryName];
  tableSummaryRows[4][i] =
    regulators[0].hour_categories[tempCategoryName].cycle.min;

  regulators[0].hour_categories[tempCategoryName].cycle.min -
    hourTotals[tempCategoryName] >
  0
    ? (tableSummaryRows[5][i] =
        regulators[0].hour_categories[tempCategoryName].cycle.min -
        hourTotals[tempCategoryName])
    : (tableSummaryRows[5][i] = 0);
  length--;
}

//add all rows
const {
  header,
  subHeader,
  tableHeader,
  tableBody,
  tableSummary
} = data.buildReportData(profile, regulators[0], certificates);
const allRows = header
  .concat(subHeader)
  .concat(tableHeader)
  .concat(tableBody)
  .concat(tableSummary);
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

//table header styles
worksheet.mergeCells('A8:A9');
worksheet.mergeCells('B8:C9');
worksheet.mergeCells('D8:D9');
worksheet.mergeCells('E8:E9');
worksheet.mergeCells('F8:F9');
worksheet.mergeCells('G8:G9');
worksheet.mergeCells('H8:H9');
worksheet.mergeCells('I8:I9');
worksheet.mergeCells('J8:J9');

//table styles B10, C10

for (let i = 0; i < tableBodyLength - 1; i++) {
  let rowNum = 10 + i;
  worksheet.mergeCells(`B${rowNum}:C${rowNum}`);
}

workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('File Written');
});

const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');
const data = require('./reportDataBuilder');
const headerBuilder = require('./components/header');
const subHeaderBuilder = require('./components/subHeader');
const tableHeaderBuilder = require('./components/tableHeader');

const regulator = regulators[3];
const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
const reportData = { regulator, profile, certificates };
// create workbook & add worksheet

const workbook = new Excel.Workbook();
const worksheet = workbook.addWorksheet('User Compliance Report');
const testSheet = workbook.addWorksheet('Components Test');

worksheet.pageSetup.margins = {
  left: 0.25,
  right: 0.25,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3
};

//get formatted report data

const {
  header,
  subHeader,
  tableHeader,
  tableBody,
  tableSummary,
  richTextData
} = data.buildReportData(profile, regulator, certificates);

//add all rows
const allRows = header
  .concat(subHeader)
  .concat(tableHeader)
  .concat(tableBody)
  .concat(tableSummary);
worksheet.addRows(allRows);

//header styles
// function to merge cells from header componeent and it takese a worksheet.
worksheet.mergeCells('A1:F2');
worksheet.mergeCells('A3:F5');
worksheet.mergeCells('G1:J2');
worksheet.mergeCells('G3:J3');
worksheet.mergeCells('G4:J4');
worksheet.mergeCells('G5:J5');

worksheet.getCell('G1').alignment = {
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
    { font: { bold: false }, text: richTextData.licenseNum }
  ]
};
worksheet.getCell('G3').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('G4').value = {
  richText: [
    { font: { bold: true }, text: 'Issue Date: ' },
    { font: { bold: false }, text: richTextData.issueDate }
  ]
};
worksheet.getCell('G4').alignment = { vertical: 'middle', horizontal: 'right' };
worksheet.getCell('G5').value = {
  richText: [
    { font: { bold: true }, text: 'Reporting Period: ' },
    { font: { bold: false }, text: richTextData.reportingPeriod }
  ]
};
worksheet.getCell('G5').alignment = { vertical: 'middle', horizontal: 'right' };

//subHeader styles
worksheet.mergeCells('A6:J7');
worksheet.getCell('A6').value = {
  richText: [
    { font: { bold: true }, text: 'Cycle Total: ' },
    { font: { bold: false }, text: richTextData.cycleTotal }
  ]
};
worksheet.getCell('A6').alignment = { vertical: 'middle', horizontal: 'left' };
for (let i = 0; i < 10; i++) {
  worksheet.getCell(`${alphabet[i]}6`).border = { top: { style: 'thick' } };
}
//table header styles

for (let i = 3; i < 10; i++) {
  worksheet.mergeCells(`${alphabet[i]}8:${alphabet[i]}9`);
}
worksheet.mergeCells('A8:A9');
worksheet.mergeCells('B8:C9');

for (let i = 0; i < 10; i++) {
  worksheet.getCell(`${alphabet[i]}8`).alignment = {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true
  };
  worksheet.getCell(`${alphabet[i]}8`).border = {
    bottom: { style: 'thick' }
  };
  worksheet.getCell(`${alphabet[i]}8`).font = {
    bold: true
  };
}

//table styles

for (let i = 0; i < tableBody.length - 1; i++) {
  let rowNum = 10 + i;
  worksheet.mergeCells(`B${rowNum}:C${rowNum}`);
}

//summary styles
let summaryStart = 10 + tableBody.length;
for (let i = 0; i < 10; i++) {
  worksheet.getCell(`${alphabet[i]}${summaryStart}`).border = {
    top: { style: 'thick' }
  };
  worksheet.getCell(`${alphabet[i]}${summaryStart}`).font = { bold: true };
  worksheet.getCell(`${alphabet[i]}${summaryStart}`).alignment = {
    vertical: 'middle',
    horizontal: 'center',
    wrapText: true
  };
}
for (let i = 6; i < 10; i++) {
  worksheet.mergeCells(
    `${alphabet[i]}${summaryStart}:${alphabet[i]}${summaryStart + 1}`
  );
}

for (let i = 2; i < 5; i++) {
  worksheet.mergeCells(`A${summaryStart + i}:F${summaryStart + i}`);
}
worksheet.mergeCells(`A${summaryStart}:F${summaryStart + 1}`);
worksheet.mergeCells(`A${summaryStart + 5}:F${summaryStart + 6}`);
worksheet.mergeCells(`G${summaryStart + 5}:G${summaryStart + 6}`);
worksheet.mergeCells(`H${summaryStart + 5}:H${summaryStart + 6}`);
worksheet.mergeCells(`I${summaryStart + 5}:I${summaryStart + 6}`);
worksheet.mergeCells(`J${summaryStart + 5}:J${summaryStart + 6}`);

worksheet.getCell(`A${summaryStart + 2}`).font = { bold: true };
worksheet.getCell(`A${summaryStart + 3}`).font = { bold: true };
worksheet.getCell(`A${summaryStart + 4}`).font = { bold: true };
worksheet.getCell(`A${summaryStart + 5}`).font = { bold: true, size: 14 };

for (let i = 0; i < 10; i++) {
  worksheet.getCell(`${alphabet[i]}${summaryStart + 4}`).border = {
    bottom: { style: 'dotted' }
  };
  worksheet.getCell(`${alphabet[i]}${summaryStart + 6}`).border = {
    bottom: { style: 'thick' }
  };
}

headerBuilder.buildHeader(testSheet, reportData);
subHeaderBuilder.buildSubHeader(testSheet, reportData, '2017');
tableHeaderBuilder.buildTableHeader(testSheet, reportData, '2017');
subHeaderBuilder.buildSubHeader(testSheet, reportData, 'ethics_state');
tableHeaderBuilder.buildTableHeader(testSheet, reportData, 'ethics_state');

workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('Report Written');
});

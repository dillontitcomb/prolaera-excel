const Excel = require('exceljs');
const certificates = require('./json/certificates.json');
const regulators = require('./json/regulators.json');
const profile = require('./json/profile.json');
const data = require('./reportDataBuilder');

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

//get formatted report data

const {
  header,
  subHeader,
  tableHeader,
  tableBody,
  tableSummary,
  richTextData
} = data.buildReportData(profile, regulators[0], certificates);

//add all rows
const allRows = header
  .concat(subHeader)
  .concat(tableHeader)
  .concat(tableBody)
  .concat(tableSummary);
worksheet.addRows(allRows);

//header styles
worksheet.mergeCells('A1:F2');
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
worksheet.mergeCells('A6:F7');
worksheet.getCell('A6').value = {
  richText: [
    { font: { bold: true }, text: 'Cycle Total: ' },
    { font: { bold: false }, text: richTextData.cycleTotal }
  ]
};
worksheet.getCell('A6').alignment = { vertical: 'middle', horizontal: 'left' };

//table header styles
const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

for (let i = 3; i < 10; i++) {
  worksheet.mergeCells(`${alphabet[i]}8:${alphabet[i]}9`);
}
worksheet.mergeCells('A8:A9');
worksheet.mergeCells('B8:C9');

//table styles

for (let i = 0; i < tableBody.length - 1; i++) {
  let rowNum = 10 + i;
  worksheet.mergeCells(`B${rowNum}:C${rowNum}`);
}

//summary styles
let summaryStart = 10 + tableBody.length;
for (let i = 0; i < 10; i++) {
  worksheet.mergeCells(
    `${alphabet[i]}${summaryStart}:${alphabet[i]}${summaryStart + 1}`
  );
}
for (let i = 2; i < 5; i++) {
  worksheet.mergeCells(`A${summaryStart + i}:F${summaryStart + i}`);
}
worksheet.mergeCells(`A${summaryStart + 5}:F${summaryStart + 6}`);

workbook.xlsx.writeFile('complianceReport.xlsx').then(function() {
  console.log('File Written');
});

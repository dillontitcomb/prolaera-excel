export default function buildHeader(worksheet, data) {
  const firstRow = worksheet.actualRowCount;
  //make all styling relative to firstRow
  const { profile, regulator, certificates } = data;

  const regulatorName = regulator.name;
  const pageNumber = 1;
  const name = profile.last + ', ' + profile.first;
  const headerRows = [];
  for (let i = 0; i < 5; i++) {
    headerRows.push(new Array(10));
  }

  headerRows[0][0] = regulatorName;
  headerRows[0][6] = 'Page ' + pageNumber;
  headerRows[2][0] = name;

  worksheet.addRows(headerRows);

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
  worksheet.getCell('A1').alignment = {
    vertical: 'middle',
    horizontal: 'left'
  };
  worksheet.getCell('A3').font = {
    size: 16,
    bold: true
  };
  worksheet.getCell('A3').alignment = {
    vertical: 'middle',
    horizontal: 'left'
  };

  worksheet.getCell('G3').value = {
    richText: [
      { font: { bold: true }, text: 'License #: ' },
      { font: { bold: false }, text: richTextData.licenseNum }
    ]
  };
  worksheet.getCell('G3').alignment = {
    vertical: 'middle',
    horizontal: 'right'
  };
  worksheet.getCell('G4').value = {
    richText: [
      { font: { bold: true }, text: 'Issue Date: ' },
      { font: { bold: false }, text: richTextData.issueDate }
    ]
  };
  worksheet.getCell('G4').alignment = {
    vertical: 'middle',
    horizontal: 'right'
  };
  worksheet.getCell('G5').value = {
    richText: [
      { font: { bold: true }, text: 'Reporting Period: ' },
      { font: { bold: false }, text: richTextData.reportingPeriod }
    ]
  };
  worksheet.getCell('G5').alignment = {
    vertical: 'middle',
    horizontal: 'right'
  };
}

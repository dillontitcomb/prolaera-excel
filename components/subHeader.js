const categoryHelper = require('../helpers/categoryHelper');

exports.buildSubHeader = function(worksheet, data, reportInput) {
  const { profile, regulator, certificates } = data;
  const firstRow = worksheet.actualRowCount + 1;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  const cycleYears = regulator.cycleYears;
  const date = new Date(regulator.date);
  const yearsPrior = new Date(date.getTime() - 31556952000 * cycleYears);
  const cycleEnd = `${date.getMonth() +
    1}/${date.getDate()}/${date.getFullYear()}`;
  const cycleStart = `${yearsPrior.getMonth() +
    1}/${yearsPrior.getDate()}/${yearsPrior.getFullYear()}`;

  let yearEnd;
  let yearStart;
  let yearEndFormatted;
  let yearStartFormatted;
  let reportingPeriod;
  let cycleYear;
  let cycleType;
  if (parseInt(reportInput) > 999 && parseInt(reportInput) < 2100) {
    cycleYear = parseInt(reportInput);
    yearEnd = new Date(`${cycleYear}-12-31`);
    yearStart = new Date(`${cycleYear}-1-1`);
    yearEndFormatted = `${yearEnd.getMonth() +
      1}/${yearEnd.getDate()}/${yearEnd.getFullYear()}`;
    yearStartFormatted = `${yearStart.getMonth() +
      1}/${yearStart.getDate()}/${yearStart.getFullYear()}`;
    reportingPeriod = `${yearStartFormatted} - ${yearEndFormatted}`;
    cycleType = 'Annual';
  } else if (categoryHelper.categoryReadable[reportInput]) {
    cycleType = categoryHelper.getCategory(reportInput);
    reportingPeriod = `${cycleStart} - ${cycleEnd}`;
  } else {
    cycleType = 'Cycle';
    reportingPeriod = `${cycleStart} - ${cycleEnd}`;
  }

  // cell styling and rich text

  worksheet.mergeCells(`A${firstRow}}:J${firstRow + 1}`);
  worksheet.getCell(`A${firstRow}`).value = {
    richText: [
      { font: { bold: true }, text: `${cycleType} Total: ` },
      { font: { bold: false }, text: reportingPeriod }
    ]
  };
  worksheet.getCell(`A${firstRow}`).alignment = {
    vertical: 'middle',
    horizontal: 'left'
  };
  for (let i = 0; i < 10; i++) {
    worksheet.getCell(`${alphabet[i]}${firstRow}`).border = {
      top: { style: 'thick' }
    };
  }
};

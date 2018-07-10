const categoryHelper = require('../helpers/categoryHelper');
const certsByCategory = require('../helpers/certsByCategory');
const certsByYear = require('../helpers/certsByYear');
const certsAll = require('../helpers/allCerts');

function getHourTotals(certificates, hourCategories) {
  const hourTotals = {};
  certificates.forEach(cert => {
    hourCategories.forEach(category => {
      if (!hourTotals[category]) hourTotals[category] = 0;
      if (cert.hours[category]) hourTotals[category] += cert.hours[category];
    });
  });
  return hourTotals;
}

exports.buildTableSummary = function(worksheet, data, reportInput) {
  const { profile, regulator, certificates } = data;
  const firstRow = worksheet.actualRowCount + 1;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let allCerts;
  let certHourTotals;

  let inputType;
  let cycleType;
  let categories = [];
  if (parseInt(reportInput) > 999 && parseInt(reportInput) < 2100) {
    inputType = 'date';
    cycleType = 'Annual';
    const keys = Object.keys(regulator.hour_categories);
    keys.forEach(key => {
      categories.unshift(key);
    });
    const certsObject3 = certsByYear.getCertsByYear(regulator, certificates);
    allCerts = certsObject3[reportInput];
  } else if (categoryHelper.categoryReadable[reportInput]) {
    inputType = 'category';
    cycleType = categoryHelper.getCategory(reportInput);
    categories = [reportInput];
    const certsObject1 = certsByCategory.getCertsByCategory(
      regulator,
      certificates
    );
    allCerts = certsObject1[reportInput];
  } else {
    const keys = Object.keys(regulator.hour_categories);
    keys.forEach(key => {
      categories.unshift(key);
    });
    inputType = 'default';
    cycleType = 'Cycle';
    allCerts = certsAll.getAllCerts(regulator, certificates);
  }
  certHourTotals = getHourTotals(allCerts, categories);
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
  let length = categories.length;
  for (let i = 9; length > 0; i--) {
    let tempCategoryName = [categories[length - 1]];
    tableSummaryRows[0][i + 1] = [categories[length - 1]][0]
      .replace('_', ' ')
      .toUpperCase();
    tableSummaryRows[2][i] = certHourTotals[tempCategoryName];
    tableSummaryRows[3][i] = certHourTotals[tempCategoryName];
    tableSummaryRows[4][i] =
      regulator.hour_categories[tempCategoryName].cycle.min;

    regulator.hour_categories[tempCategoryName].cycle.min -
      certHourTotals[tempCategoryName] >
    0
      ? (tableSummaryRows[5][i] =
          regulator.hour_categories[tempCategoryName].cycle.min -
          certHourTotals[tempCategoryName])
      : (tableSummaryRows[5][i] = 0);
    length--;
  }
  worksheet.addRows(tableSummaryRows);

  //cell styling and rich text

  for (let i = 0; i < 10; i++) {
    worksheet.getCell(`${alphabet[i]}${firstRow}`).border = {
      top: { style: 'thick' }
    };
    worksheet.getCell(`${alphabet[i]}${firstRow}`).font = { bold: true };
    worksheet.getCell(`${alphabet[i]}${firstRow}`).alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true
    };
  }
  for (let i = 6; i < 10; i++) {
    worksheet.mergeCells(
      `${alphabet[i]}${firstRow}:${alphabet[i]}${firstRow + 1}`
    );
  }

  for (let i = 2; i < 5; i++) {
    worksheet.mergeCells(`A${firstRow + i}:F${firstRow + i}`);
  }
  worksheet.mergeCells(`A${firstRow}:F${firstRow + 1}`);
  worksheet.mergeCells(`A${firstRow + 5}:F${firstRow + 6}`);
  worksheet.mergeCells(`G${firstRow + 5}:G${firstRow + 6}`);
  worksheet.mergeCells(`H${firstRow + 5}:H${firstRow + 6}`);
  worksheet.mergeCells(`I${firstRow + 5}:I${firstRow + 6}`);
  worksheet.mergeCells(`J${firstRow + 5}:J${firstRow + 6}`);

  worksheet.getCell(`A${firstRow + 2}`).font = { bold: true };
  worksheet.getCell(`A${firstRow + 3}`).font = { bold: true };
  worksheet.getCell(`A${firstRow + 4}`).font = { bold: true };
  worksheet.getCell(`A${firstRow + 5}`).font = { bold: true, size: 14 };

  for (let i = 0; i < 10; i++) {
    worksheet.getCell(`${alphabet[i]}${firstRow + 4}`).border = {
      bottom: { style: 'dotted' }
    };
    worksheet.getCell(`${alphabet[i]}${firstRow + 6}`).border = {
      bottom: { style: 'thick' }
    };
  }
};

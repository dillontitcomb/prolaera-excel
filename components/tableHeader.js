const categoryHelper = require('../helpers/categoryHelper');

exports.buildTableHeader = function(worksheet, data, reportInput) {
  const { profile, regulator, certificates } = data;
  const firstRow = worksheet.actualRowCount + 1;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

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
  } else if (categoryHelper.categoryReadable[reportInput]) {
    inputType = 'category';
    cycleType = categoryHelper.getCategory(reportInput);
    categories = [reportInput];
  } else {
    const keys = Object.keys(regulator.hour_categories);
    keys.forEach(key => {
      categories.unshift(key);
    });
    inputType = 'default';
    cycleType = 'Cycle';
  }

  const tableHeaderRows = [new Array(10), new Array(10)];
  tableHeaderRows[0][0] = 'DATE';
  tableHeaderRows[0][1] = 'TITLE';
  tableHeaderRows[0][3] = 'SPONSOR';
  tableHeaderRows[0][4] = 'DELIVERY METHOD';
  let categoryNumber = categories.length;
  for (let i = 9; categoryNumber > 0; i--) {
    tableHeaderRows[0][i] = categories[categoryNumber - 1]
      .replace('_', ' ')
      .toUpperCase();
    categoryNumber--;
  }

  worksheet.addRows(tableHeaderRows);

  // cell styling and rich text

  for (let i = 3; i < 10; i++) {
    worksheet.mergeCells(
      `${alphabet[i]}${firstRow}:${alphabet[i]}${firstRow + 1}`
    );
  }
  worksheet.mergeCells(`A${firstRow}:A${firstRow + 1}`);
  worksheet.mergeCells(`B${firstRow}:C${firstRow + 1}`);

  for (let i = 0; i < 10; i++) {
    worksheet.getCell(`${alphabet[i]}${firstRow}`).alignment = {
      vertical: 'middle',
      horizontal: 'center',
      wrapText: true
    };
    worksheet.getCell(`${alphabet[i]}${firstRow}`).border = {
      bottom: { style: 'thick' }
    };
    worksheet.getCell(`${alphabet[i]}${firstRow}`).font = {
      bold: true
    };
  }
};

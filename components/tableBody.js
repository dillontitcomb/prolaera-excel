const categoryHelper = require('../helpers/categoryHelper');
const certsByCategory = require('../helpers/certsByCategory');
const certsByYear = require('../helpers/certsByYear');
const certsAll = require('../helpers/allCerts');

function getSponsor(cert) {
  let sponsor;
  if (cert.sponsor) {
    sponsor = cert.sponsor;
  } else if (cert.sponsors) {
    sponsor = cert.sponsors.name || Object.values(cert.sponsors)[0];
  } else {
    sponsor = 'N/A';
  }
  return sponsor;
}

exports.buildTableBody = function(worksheet, data, reportInput) {
  const { profile, regulator, certificates } = data;
  const firstRow = worksheet.actualRowCount + 1;
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let allCerts;

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

  const tableBodyRows = [];
  allCerts.forEach(cert => {
    let tempRow = [new Array(10)];
    tempRow[0] = cert.formattedDate;
    tempRow[1] = cert.cert;
    tempRow[3] = getSponsor(cert);
    tempRow[4] = cert.delivery;
    let categoryNumber = categories.length;
    for (let i = 9; categoryNumber > 0; i--) {
      tempRow[i] = cert.hours[categories[categoryNumber - 1]];
      let tempCategory = [categories[categoryNumber - 1]];
      categoryNumber--;
    }
    tableBodyRows.push(tempRow);
  });

  worksheet.addRows(tableBodyRows);

  for (let i = 0; i < tableBodyRows.length - 1; i++) {
    let rowNum = firstRow + i;
    worksheet.mergeCells(`B${rowNum}:C${rowNum}`);
  }
};

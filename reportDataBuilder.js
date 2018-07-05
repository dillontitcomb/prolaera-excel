//Multiple-use data
exports.buildReportData = function(profile, regulator, certificates) {
  const certificatesDict = certificates.reduce((obj, cert) => {
    obj[cert.cert_id] = cert;
    return obj;
  }, {});

  const dynamicCategories = [];
  const keys = Object.keys(regulator.hour_categories);
  keys.forEach(key => {
    dynamicCategories.unshift(key);
  });

  const hourTotals = {};
  dynamicCategories.forEach(col => {
    hourTotals[col] = 0;
  });

  const allCerts = [];
  const yearKeys = Object.keys(regulator.years);
  yearKeys.forEach(key => {
    const tempAppliedCerts = regulator.years[key].certificates_applied;
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

  function buildHeader() {
    const regulatorName = regulator.name;
    const pageNumber = 1;
    const name = profile.last + ', ' + profile.first;
    const headerRows = [];
    for (let i = 0; i < 5; i++) {
      headerRows.push(new Array(10));
    }

    headerRows[0][0] = regulatorName;
    headerRows[0][9] = 'Page ' + pageNumber;
    headerRows[2][0] = name;
    return headerRows;
  }
  function buildSubHeader() {
    const subHeaderRows = [new Array(10), new Array(10)];
    return subHeaderRows;
  }
  function buildTableHeader() {
    const tableHeaderRows = [new Array(10), new Array(10)];
    tableHeaderRows[0][0] = 'DATE';
    tableHeaderRows[0][1] = 'TITLE';
    tableHeaderRows[0][3] = 'SPONSOR';
    tableHeaderRows[0][4] = 'DELIVERY METHOD';
    let numberOfDynamicCategories = dynamicCategories.length;
    for (let i = 9; numberOfDynamicCategories > 0; i--) {
      tableHeaderRows[0][i] = dynamicCategories[numberOfDynamicCategories - 1]
        .replace('_', ' ')
        .toUpperCase();
      numberOfDynamicCategories--;
    }
    return tableHeaderRows;
  }
  function buildTableBody() {
    const tableBodyRows = [];
    allCerts.forEach(cert => {
      let tempRow = [new Array(10)];
      tempRow[0] = cert.formattedDate;
      tempRow[1] = cert.cert;
      tempRow[3] = cert.sponsor || cert.sponsors.name;
      tempRow[4] = cert.delivery;
      let numberOfDynamicCategories = dynamicCategories.length;
      for (let i = 9; numberOfDynamicCategories > 0; i--) {
        tempRow[i] =
          cert.hours[dynamicCategories[numberOfDynamicCategories - 1]];
        let tempCategory = [dynamicCategories[numberOfDynamicCategories - 1]];
        if (
          typeof cert.hours[
            dynamicCategories[numberOfDynamicCategories - 1]
          ] === 'number'
        )
          hourTotals[tempCategory] +=
            cert.hours[dynamicCategories[numberOfDynamicCategories - 1]];
        numberOfDynamicCategories--;
      }
      tableBodyRows.push(tempRow);
    });
    return tableBodyRows;
  }
  function buildTableSummary() {
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
    let length = dynamicCategories.length;
    for (let i = 9; length > 0; i--) {
      let tempCategoryName = [dynamicCategories[length - 1]];
      tableSummaryRows[0][i + 1] = [dynamicCategories[length - 1]][0]
        .replace('_', ' ')
        .toUpperCase();
      tableSummaryRows[2][i] = hourTotals[tempCategoryName];
      tableSummaryRows[3][i] = hourTotals[tempCategoryName];
      tableSummaryRows[4][i] =
        regulator.hour_categories[tempCategoryName].cycle.min;

      regulator.hour_categories[tempCategoryName].cycle.min -
        hourTotals[tempCategoryName] >
      0
        ? (tableSummaryRows[5][i] =
            regulator.hour_categories[tempCategoryName].cycle.min -
            hourTotals[tempCategoryName])
        : (tableSummaryRows[5][i] = 0);
      length--;
    }
    return tableSummaryRows;
  }
  const headerRows = buildHeader();
  const subHeaderRows = buildSubHeader();
  const tableHeaderRows = buildTableHeader();
  const tableBodyRows = buildTableBody();
  const tableSummaryRows = buildTableSummary();

  const reportData = {
    header: headerRows,
    subHeader: subHeaderRows,
    tableHeader: tableHeaderRows,
    tableBody: tableBodyRows,
    tableSummary: tableSummaryRows
  };
  return reportData;
};

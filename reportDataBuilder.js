module.exports = function buildReportData(regulators, profile, certificates) {
  function buildHeader() {
    console.log(regulators);
  }
  function buildSubHeader() {
    console.log(profile);
  }
  function buildTableHeader() {
    console.log(certificates);
  }
  function buildTableBody() {}
  function buildTableSummary() {}

  const headerRows = buildHeader();
  const subHeaderRows = buildSubHeader();
  const tableHeaderRows = buildTableHeader();
  const tableBodyRows = buildTableBody();
  const tableSummaryRows = buildTableSummary();

  //   const reportData = headerRows
  //     .concat(subHeaderRows)
  //     .concat(tableHeaderRows)
  //     .concat(tableBodyRows)
  //     .concat(tableSummaryRows);

  //   return reportData;
};

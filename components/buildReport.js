const headerBuilder = require('./header');
const subHeaderBuilder = require('./subHeader');
const tableHeaderBuilder = require('./tableHeader');
const tableBodyBuilder = require('./tableBody');
const tableSummaryBuilder = require('./tableSummary');

exports.buildReport = function(worksheet, reportData, reportType) {
  const { profile, regulator, certificates } = reportData;
  const years = Object.keys(regulator.years);
  let categories;
  headerBuilder.buildHeader(worksheet, reportData);
  switch (reportType) {
    case 'cycle':
      subHeaderBuilder.buildSubHeader(worksheet, reportData, 'cycle');
      tableHeaderBuilder.buildTableHeader(worksheet, reportData, 'cycle');
      tableBodyBuilder.buildTableBody(worksheet, reportData, 'cycle');
      tableSummaryBuilder.buildTableSummary(worksheet, reportData, 'cycle');
      break;
    case 'annual':
      years.forEach(year => {
        subHeaderBuilder.buildSubHeader(worksheet, reportData, year);
        tableHeaderBuilder.buildTableHeader(worksheet, reportData, year);
        tableBodyBuilder.buildTableBody(worksheet, reportData, year);
        tableSummaryBuilder.buildTableSummary(worksheet, reportData, year);
      });
      break;
    case 'allcategories':
      categories = Object.keys(regulator.hour_categories);
      categories.forEach(cat => {
        subHeaderBuilder.buildSubHeader(worksheet, reportData, cat);
        tableHeaderBuilder.buildTableHeader(worksheet, reportData, cat);
        tableBodyBuilder.buildTableBody(worksheet, reportData, cat);
        tableSummaryBuilder.buildTableSummary(worksheet, reportData, cat);
      });
      break;
    case 'complete':
      subHeaderBuilder.buildSubHeader(worksheet, reportData, 'cycle');
      tableHeaderBuilder.buildTableHeader(worksheet, reportData, 'cycle');
      tableBodyBuilder.buildTableBody(worksheet, reportData, 'cycle');
      tableSummaryBuilder.buildTableSummary(worksheet, reportData, 'cycle');
      years.forEach(year => {
        subHeaderBuilder.buildSubHeader(worksheet, reportData, year);
        tableHeaderBuilder.buildTableHeader(worksheet, reportData, year);
        tableBodyBuilder.buildTableBody(worksheet, reportData, year);
        tableSummaryBuilder.buildTableSummary(worksheet, reportData, year);
      });
      categories = Object.keys(regulator.hour_categories);
      categories.forEach(cat => {
        subHeaderBuilder.buildSubHeader(worksheet, reportData, cat);
        tableHeaderBuilder.buildTableHeader(worksheet, reportData, cat);
        tableBodyBuilder.buildTableBody(worksheet, reportData, cat);
        tableSummaryBuilder.buildTableSummary(worksheet, reportData, cat);
      });
      break;
    // default case builds single table for year or category reportType (e.g. '2017' or 'ethics_state')
    default:
      subHeaderBuilder.buildSubHeader(worksheet, reportData, reportType);
      tableHeaderBuilder.buildTableHeader(worksheet, reportData, reportType);
      tableBodyBuilder.buildTableBody(worksheet, reportData, reportType);
      tableSummaryBuilder.buildTableSummary(worksheet, reportData, reportType);
      break;
  }
};

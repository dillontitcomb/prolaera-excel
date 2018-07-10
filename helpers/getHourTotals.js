exports.getHourTotals = function(certificates, hourCategories) {
  const hourTotals = {};
  certificates.forEach(cert => {
    hourCategories.forEach(category => {
      if (!hourTotals[category]) hourTotals[category] = 0;
      if (cert.hours[category]) hourTotals[category] += cert.hours[category];
    });
  });
  return hourTotals;
};

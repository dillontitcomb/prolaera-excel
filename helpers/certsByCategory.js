const categoryHelper = require('../helpers/categoryHelper');

exports.getCertsByCategory = function(regulator, certificates) {
  const certificatesDict = certificates.reduce((obj, cert) => {
    obj[cert.cert_id] = cert;
    return obj;
  }, {});
  const certsByCategory = {};
  certsByCategory['totals'] = {};
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
      let categoryKeys = Object.keys(categoryHelper.categoryReadable);
      categoryKeys.forEach(category => {
        if (!certsByCategory[category] && tempCert.hours[category] > 0)
          certsByCategory[category] = [];
        if (tempCert.hours[category]) certsByCategory[category].push(tempCert);
      });
    });
  });
  return certsByCategory;
};

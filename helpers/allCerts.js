exports.getAllCerts = function(regulator, certificates) {
  const certificatesDict = certificates.reduce((obj, cert) => {
    obj[cert.cert_id] = cert;
    return obj;
  }, {});
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
  return allCerts;
};

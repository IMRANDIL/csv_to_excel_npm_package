const csvToExcel = require('../lib/index')


 // Add header row
 const headers = [
    { header: 'Time', key: 'Time', width: 10 },
    { header: 'V1', key: 'V1', width: 10 },
    { header: 'V2', key: 'V2', width: 10 },
    { header: 'V3', key: 'V3', width: 10 },
    { header: 'V4', key: 'V4', width: 10 },
    { header: 'V5', key: 'V5', width: 10 },
    { header: 'V6', key: 'V6', width: 10 },
    { header: 'V7', key: 'V7', width: 10 },
    { header: 'V8', key: 'V8', width: 10 },
    { header: 'V9', key: 'V9', width: 10 },
    { header: 'V10', key: 'V10', width: 10 },
    { header: 'V11', key: 'V11', width: 10 },
    { header: 'V12', key: 'V12', width: 10 },
    { header: 'V13', key: 'V13', width: 10 },
    { header: 'V14', key: 'V14', width: 10 },
    { header: 'V15', key: 'V15', width: 10 },
    { header: 'V16', key: 'V16', width: 10 },
    { header: 'V17', key: 'V17', width: 10 },
    { header: 'V18', key: 'V18', width: 10 },
    { header: 'V19', key: 'V19', width: 10 },
    { header: 'V20', key: 'V20', width: 10 },
    { header: 'V21', key: 'V21', width: 10 },
    { header: 'V22', key: 'V22', width: 10 },
    { header: 'V23', key: 'V23', width: 10 },
    { header: 'V24', key: 'V24', width: 10 },
    { header: 'V25', key: 'V25', width: 10 },
    { header: 'V26', key: 'V26', width: 10 },
    { header: 'V27', key: 'V27', width: 10 },
    { header: 'V28', key: 'V28', width: 10 },
    { header: 'Amount', key: 'Amount', width: 10 },
    { header: 'Class', key: 'Class', width: 10 }
  ];
csvToExcel(headers, 'creditcard.csv', 'output.xlsx').catch(err => console.error(err));
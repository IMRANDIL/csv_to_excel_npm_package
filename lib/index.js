const fs = require('fs');
const csv = require('csvtojson');
const { Transform, pipeline } = require('stream');
const ExcelJS = require('exceljs');

const main = async (header, inputFilePath, outPutFilePath) => {
    const readStream = fs.createReadStream(inputFilePath);
   

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');

    // Add header row
    ws.columns = header.map(header => ({ header: header.header, key: header.key, width: header.width }));

    const myTransform = new Transform({
        objectMode: true,
        transform(chunk, enc, cb) {
           // Extract values from the chunk object
        const rowValues = Object.values(chunk);
        
        // Add rowValues as a row to the worksheet
        ws.addRow(rowValues);
        
        // Call the callback function to indicate that processing is complete for this chunk
        cb();
        }
    });

    try {
    // Handle pipeline properly
   pipeline(
        readStream,
        csv({
            delimiter: ';'
        },{objectMode: true}),
        myTransform,
        async (err) => {
            if (err) {
                console.error('Pipeline failed.', err);
            } else {
                const writeStream = fs.createWriteStream(outPutFilePath);
                await wb.xlsx.write(writeStream);
                console.log('Pipeline succeeded.');
            }
        }
    );
   
    } catch (error) {
        console.error(error)
    }


};

module.exports = main;

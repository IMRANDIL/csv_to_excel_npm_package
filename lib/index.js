const fs = require('fs');
const csv = require('csvtojson');
const { Transform, pipeline } = require('stream');
const ExcelJS = require('exceljs');

const main = async (header, inputFilePath, outPutFilePath) => {
    const readStream = fs.createReadStream(inputFilePath);
    const writeStream = fs.createWriteStream(outPutFilePath);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');

    // Add header row
    ws.columns = header.map(header => ({ header: header.header, key: header.key, width: header.width }));

    const myTransform = new Transform({
        objectMode: true,
        transform(chunk, enc, cb) {
            // Assuming chunk is an object containing data
            // Modify this part according to your data structure
            console.log('chunk>>>>>>>>>>', chunk)
            ws.addRow(chunk);
            cb(null, chunk);
        }
    });

    try {
    // Handle pipeline properly
    await pipeline(
        readStream,
        csv({
            delimiter: ';'
        },{objectMode: true}),
        myTransform,
        writeStream
    );
    } catch (error) {
        console.error(error)
    }


};

module.exports = main;

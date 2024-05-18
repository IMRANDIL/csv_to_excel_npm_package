const fs = require('fs');
const csv = require('csv-parser');
const ExcelJS = require('exceljs');

const main = async (header, inputFilePath, outPutFilePath) => {
    try {
        const csvStream = fs.createReadStream(inputFilePath);
        const writeStream = fs.createWriteStream(outPutFilePath);
        
        // Construct a new workbook
        const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
            stream: writeStream,
            useStyles: true,
            useSharedStrings: true
        });

        const worksheet = workbook.addWorksheet('Data', {
            views: [{ state: 'frozen', xSplit: 0, ySplit: 1 }]
        });

        // Add header row
        worksheet.addRow(header.map(h => h.header)).commit();

        const csvParserStream = csv()
            .on('data', (data) => {
                const rowData = Object.values(data);
                worksheet.addRow(rowData).commit();
            })
            .on('end', async () => {
                console.log('CSV parsing completed.');

                try {
                    // Commit the worksheet
                    await worksheet.commit();

                    // Finish the workbook to complete the XLSX document
                    await workbook.commit();
                    console.log('Excel file created successfully.');
                } catch (error) {
                    console.error('Error writing Excel file:', error);
                }
            })
            .on('error', (err) => {
                console.error('Error reading CSV file:', err);
            });

        csvStream.pipe(csvParserStream);

    } catch (error) {
        console.error('Error processing CSV to Excel:', error);
    }
};

module.exports = main;

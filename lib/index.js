const fs = require('fs');
const csv = require('csv-parser');
const ExcelJS = require('exceljs');
const path = require('path');

const validateFilePath = (filePath, fileType) => {
    if (typeof filePath !== 'string' || !filePath.trim()) {
        throw new Error(`${fileType} path must be a non-empty string.`);
    }

    const ext = path.extname(filePath).toLowerCase();
    if ((fileType === 'Input' && ext !== '.csv') || (fileType === 'Output' && ext !== '.xlsx')) {
        throw new Error(`${fileType} file must have a valid ${fileType === 'Input' ? '.csv' : '.xlsx'} extension.`);
    }

    if (fileType === 'Input' && !fs.existsSync(filePath)) {
        throw new Error(`${fileType} file does not exist.`);
    }
};

const main = async (header, inputFilePath, outPutFilePath) => {
    try {
        validateFilePath(inputFilePath, 'Input');
        validateFilePath(outPutFilePath, 'Output');

        if ((header && !Array.isArray(header)) || header.length == 0) {
            throw new Error('Header must be an array or null.');
        }

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

        let headerWritten = false;

        const csvParserStream = csv()
            .on('data', (data) => {
                if (!headerWritten) {
                    if (!header) {
                        header = Object.keys(data).map((key) => ({ header: key }));
                    }
                    worksheet.addRow(header.map(h => h.header)).commit();
                    headerWritten = true;
                }
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

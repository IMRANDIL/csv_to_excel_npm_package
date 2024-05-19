# CSV to Excel Converter

A simple and efficient Node.js module for converting CSV files to Excel (XLSX) format using streams. This package leverages the power of `csv-parser` and `exceljs` libraries to handle large datasets efficiently without consuming too much memory.

## Features

- **Stream-based processing**: Efficiently handles large CSV files.
- **Optional header support**: Use your own header or extract from CSV data.
- **Easy-to-use API**: Simple function call to convert your CSV to Excel.

## Installation

```bash
npm install @imrandil/csv-to-excel-converter
```

## Usage

### Basic Usage

```javascript
const csvToExcel = require('@imrandil/csv-to-excel-converter');

const header = null; // Use null to extract headers from the CSV file
const inputFilePath = './path/to/input.csv';
const outPutFilePath = './path/to/output.xlsx';

csvToExcel(header, inputFilePath, outPutFilePath);
```

### With Custom Headers

```javascript
const csvToExcel = require('@imrandil/csv-to-excel-converter');

const header = [
  { header: 'id', key: 'id', width: 15 },
  { header: 'is_active', key: 'is_active', width: 10 },
  { header: 'created_date', key: 'created_date', width: 25 },
  { header: 'last_modified_date', key: 'last_modified_date', width: 25 },
  { header: 'unique_id', key: 'unique_id', width: 20 },
    // add more headers as needed
];
const inputFilePath = './path/to/input.csv';
const outPutFilePath = './path/to/output.xlsx';

csvToExcel(header, inputFilePath, outPutFilePath);
```

## Parameters

- `header`: An array of objects containing headers (optional). If not provided, headers will be extracted from the CSV file.
- `inputFilePath`: The path to the input CSV file (required).
- `outPutFilePath`: The path to the output Excel file (required).

## Example Use Cases

### 1. Data Migration

Easily migrate data from legacy systems that export data in CSV format to systems that accept Excel files.

### 2. Data Reporting

Generate Excel reports from CSV data for better readability and data manipulation.

### 3. Data Analysis

Convert CSV files from data analysis tools into Excel format for further analysis with tools like Excel or Google Sheets.

## Error Handling

The module performs validation on input and output paths and provides meaningful error messages for the user.

```javascript
try {
    csvToExcel(header, inputFilePath, outPutFilePath);
} catch (error) {
    console.error('Error processing CSV to Excel:', error.message);
}
```

## Contributions

Contributions are welcome! Please submit a pull request or open an issue to discuss improvements or features.

## License

This project is licensed under the MIT License.

---

With this package, converting CSV files to Excel format is a breeze. It efficiently handles large datasets and offers flexibility with optional headers. Perfect for data migration, reporting, and analysis!

---

### Repository

For more information, issues, or contributions, visit our [GitHub repository](https://github.com/IMRANDIL/csv_to_excel_npm_package).

### Author

Developed by [Ali Imran Adil](https://github.com/IMRANDIL).

---

## Support

If you have any questions or need support, feel free to [open an issue](https://github.com/IMRANDIL/csv_to_excel_npm_package/issues) on GitHub.
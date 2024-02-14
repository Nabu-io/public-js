import ExcelJS, { Worksheet } from 'exceljs'

const worksheetNameToColumns: { [key: string]: string[] } = {
    'Organizations': [
        'Organization Name',
    ],
    'Groups': [
        'Group Name',
        'Group Type',
        'Parent',
        'Organization Name',
        'Company Registration Number',
        'Country',
        'Address - Given Name',
        'Address - Surname',
        'Address - Company Name',
        'Address - Street No',
        'Address - Street Type',
        'Address - Street Name',
        'Address - Floor',
        'Address - Town',
        'Address - Region',
        'Address - Postcode',
        'Address - Country',
        'Address - Tag',
    ],
    'Users': [
        'Email',
        'Role',
        'Active',
        'Group Name',
        'First Name',
        'Last Name',
        'Position',
        'Language',
        'Password'
    ],
}

function validateSheet(worksheet: ExcelJS.Worksheet, expectedColumns: string[]) {
    let actualColumns: string[] = [];
    worksheet.eachRow({ includeEmpty: false }, function (row, rowNumber) {
        if (rowNumber === 1) { // Assuming the first row contains column headers
            row.eachCell({ includeEmpty: false }, function (cell, colNumber) {
                actualColumns.push(cell.value as string);
            });
        }
    });

    for (let col of expectedColumns) {
        if (!actualColumns.includes(col)) {
            throw new Error(`Sheet [${worksheet.name}] is missing column: ${col}`);
        }
    }
    return true;
}


export default function validateExcelFile(file: ArrayBuffer): Promise<ExcelJS.Workbook> {
    const workbook = new ExcelJS.Workbook();
    return workbook.xlsx.load(file).then(() => {

        for (let sheetName of workbook.worksheets.map(w => w.name.trim())) {
            const columns = worksheetNameToColumns[sheetName];
            if (!columns) {
                throw new Error(`Columns for worksheet '${sheetName}' not found.`);
            }

            const valid = validateSheet(workbook.getWorksheet(sheetName) as Worksheet, columns);
            if (!valid) {
                throw new Error(`Validation failed for worksheet '${sheetName}'.`);
            }
        }

        return workbook
    })
}
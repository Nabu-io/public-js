import ExcelJS, { Worksheet } from 'exceljs'

const worksheetNameToColumns: { [key: string]: string[] } = {
    'Organization & Groups': [
        'Group Name',
        'Group Type',
        'Parent'
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
    'Company Registration Numbers': [
        'Value',
        'Country',
        'Group Name'
    ],
    'Addresses': [
        'Given Name',
        'Surname',
        'Company Name',
        'Street No',
        'Street Name',
        'Floor',
        'Town',
        'Region',
        'PostCode',
        'Country',
        'Tag',
        'Group Name'
    ],
    'Schemas': [
        'Name',
        'Content'
    ],
    'Export Data Templates': [
        'Name',
        'Schema Name',
        'New'
    ],
    'Transports': [
        'Type',
        'Name',
        'Address',
        'Port',
        'Username',
        'Password',
        'Path',
        'File Name Template'
    ],
    'Configurations': [
        'Name',
        'ExportDataTemplate Name',
        'Transport Name',
        'Group Name',
        'Preset'
    ],
    'Settings': [
        'Key',
        'Value',
        'Group Name'
    ]
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


export default function validateExcelFile(file: File): Promise<any> {
    const workbook = new ExcelJS.Workbook();
    return workbook.xlsx.load(file)
        .then(() => {

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
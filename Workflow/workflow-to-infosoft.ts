// @ts-nocheck
function main(workbook: ExcelScript.Workbook) {
    const sheet = workbook.getWorksheet('default');
    const range = sheet.getUsedRange();
    const values = range.getValues();


    for (let i = 1; i < values.length; i++) {
        const author = values[i][0];
        const leave_type = values[i][1];
        const start_date = values[i][2];
        const total_days = values[i][3];

        writeNew(author.toString(), leave_type.toString(), +start_date, +total_days, workbook)
    }


}

function writeNew(author: string, leave_type:string, start_date: number, total_days: number, workbook: ExcelScript.Workbook) {
    // create a new worksheet called "readyToImport"
    const sheetName = 'readyToImport';
    let readyToImportSheet = workbook.getWorksheet(sheetName);

    if (!readyToImportSheet) {
        readyToImportSheet = workbook.addWorksheet(sheetName);
        readyToImportSheet.getRange('A1:E1').setValues([['Name"', 'Date"', 'IN"','OUT"','']]);
    }


    //get last cell
    const sheet = workbook.getWorksheet('readyToImport');
    const range = sheet.getUsedRange();
    const values = range.getValues();
    let lastCell:number = values.length;
    //insert bottom of last
    let cuti = '';

    if (leave_type === "One_Day_Leave") {
      cuti = "CUTI"
    } else if (leave_type === "AM_Leave") {
      cuti = "CUTI AM 0.5"
    } else if (leave_type === "PM_Leave"){
      cuti = "CUTI PM 0.5"
    }
    if (total_days > 1) {
        for (let j = 0; j < total_days; j++) {
            const modifiedStartDate:number = start_date + j;
            readyToImportSheet.getCell(lastCell, 0).setValues(author.toUpperCase());
            readyToImportSheet.getCell(lastCell, 1).setValues(modifiedStartDate);
            readyToImportSheet.getCell(lastCell, 1).setNumberFormat('mm/dd/yyy');
            readyToImportSheet.getCell(lastCell, 2).setValues('');
            readyToImportSheet.getCell(lastCell, 3).setValues('');
            readyToImportSheet.getCell(lastCell, 4).setValues(cuti);
            lastCell +=1;
        }
    } else {
      // copy the row to the "readyToImport" sheet
      readyToImportSheet.getCell(lastCell, 0).setValues(author.toUpperCase());
      readyToImportSheet.getCell(lastCell, 1).setValues(start_date);
      readyToImportSheet.getCell(lastCell, 1).setNumberFormat('mm/dd/yyy');
      readyToImportSheet.getCell(lastCell, 2).setValues('');
      readyToImportSheet.getCell(lastCell, 3).setValues('');
      readyToImportSheet.getCell(lastCell, 4).setValues(cuti);
      lastCell +=1;
    }
}
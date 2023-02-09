// @ts-nocheck
function main(workbook: ExcelScript.Workbook) {

    generate_totalAssets_by_Status(workbook);
    generate_totalByDepartmentAndTypeAssets(workbook);
  }
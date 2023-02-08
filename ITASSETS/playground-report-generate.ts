// @ts-nocheck
function main(workbook: ExcelScript.Workbook) {

    generate_totalAssets_by_Status(workbook);
    generate_totalByDepartmentAndTypeAssets(workbook);
  }


  //GENERATE TOTAL ASSETS BY STATUS
  function generate_totalAssets_by_Status(workbook: ExcelScript.Workbook) {
    let database = workbook.getWorksheet("database");
    const database_range = database.getUsedRange();
    const database_values = database_range.getValues();
    const database_lastCell = database_values.length;
    const selectedColumns = database_values.map((row) => {
      return [row[1], row[8]];
    });

    // count by status
    let count = {};
    for (let i = 0; i < selectedColumns.length; i++) {
      let item = selectedColumns[i][0];
      let status = selectedColumns[i][1];

      if (!(JSON.stringify(item) in count)) {
        count[JSON.stringify(item)] = { Borrowed: 0, Available: 0, Missing: 0, Broken: 0 };
      }
      count[JSON.stringify(item)][status]++;
    }
    let finalArray: { item: string, Borrowed: number, Available: number, Missing: number, Broken: number }[] = []
    for (let item in count) {
      let obj = {
        item: JSON.parse(item),
        Borrowed: count[item]['Borrowed'],
        Available: count[item]['Available'],
        Missing: count[item]['Missing'],
        Broken: count[item]['Broken']
      }
      finalArray.push(obj);
    }
    finalArray.shift();


    // add to new data to  worksheet
    workbook.getWorksheet("Status by Type Assets")?.delete();
    const report = workbook.addWorksheet("Status by Type Assets");
    report.getRange("C1:G1").setValues([['Type Asset', 'Borrowed', 'Available', 'Missing', 'Broken']]);
    const reportL = report.getUsedRange();
    const reportV = reportL.getValues();
    let reportTR = reportV.length;
    finalArray.forEach((data) => {
      reportTR += 1;
      report.getRange(`C${reportTR}:G${reportTR}`).setValues([[data.item, data.Borrowed, data.Available, data.Missing, data.Broken]]);
    })


    //Draw graprh
    let lastCol = "G";
    let firstA = "C1:"
    let firstB = lastCol.concat(reportTR.toString());
    let finalcol = firstA.concat(firstB);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range 1:1 on selectedSheet
    report.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("1:1").getFormat().setIndentLevel(0);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range B:E on selectedSheet
    report.getRange("C:G").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("C:G").getFormat().setIndentLevel(0);
    // Set fill color to B4C6E7 for range A1:E1 on selectedSheet
    report.getRange("C1:G1").getFormat().getFill().setColor("B4C6E7");

    let chartAssetsByStatus = drawCart(workbook, "Status by Type Assets", finalcol);
    chartAssetsByStatus.setPosition(`B${reportTR + 3}`);
    chartAssetsByStatus.getTitle().setText("Status by Type Assets");
  }

  //GENERATE TOTAL ASSET BY TYPE AND DEPARTMENT TEMPORARY
  function generate_totalByDepartmentAndTypeAssetsTmp(workbook: ExcelScript.Workbook) {
    let database = workbook.getWorksheet("database");
    const database_range = database.getUsedRange();
    const database_values = database_range.getValues();
    const database_lastCell = database_values.length;

    let selectedColumns = database_values.map((row) => {
      return [row[1], row[6] === '' ? '-' : row[6], row[8]];
    });

    selectedColumns = selectedColumns.filter((row) => {
      return row[2] === "Borrowed";
    });

    let finalArray = selectedColumns.reduce((acc, curr) => {
      const [typeAsset, department, status] = curr;
      const found = acc.find(
        (obj) => obj.TypeAsset === typeAsset && obj.Department === department
      );

      if (found) {
        found.Total++;
      } else {
        acc.push({ TypeAsset: typeAsset, Department: department, Total: 1 });
      }
      return acc;
    }, []);


    workbook.getWorksheet("Type Asset by Department Tmp")?.delete();
    const report = workbook.addWorksheet("Type Asset by Department Tmp");
    report.getRange("A1:C1").setValues([['TypeAssets', 'Department', 'Total Device']]);

    const reportL = report.getUsedRange();
    const reportV = reportL.getValues();
    let reportTR = reportV.length;

    finalArray.forEach((data) => {
      reportTR += 1;
      report.getRange(`A${reportTR}:C${reportTR}`).setValues([[data.TypeAsset, data.Department, data.Total]]);
    });

    generate_totalByDepartmentAndTypeAssets(workbook)
  }

  //GENERATE TOTAL ASSET BY TYPE AND DEPARTMENT FINAL
  function generate_totalByDepartmentAndTypeAssets(workbook: ExcelScript.Workbook) {
    let report = workbook.getWorksheet("Type Asset by Department Tmp");
    const report_range = report.getUsedRange();
    const report_values = report_range.getValues();
    const headerRow = report_values.shift();

    let dataMap = new Map<string, Map<string, number>>();

    for (const row of report_values) {
      let department = row[1];
      let typeAsset = row[0];
      let totalDevice = row[2];

      if (!dataMap.has(department)) {
        dataMap.set(department, new Map<string, number>());
      }

      let typeAssetMap = dataMap.get(department)!;
      typeAssetMap.set(typeAsset, totalDevice);
    }

    let departments = Array.from(dataMap.keys());
    let typeAssets = Array.from(dataMap.values()).map((map) => Array.from(map.keys()));

    let transposedData = [
      [headerRow[1], ...Array.from(new Set(typeAssets.flat()))],
    ];

    for (const department of departments) {
      let typeAssetMap = dataMap.get(department)!;
      let rowData = [department];

      for (const typeAsset of transposedData[0].slice(1)) {
        let totalDevice = typeAssetMap.get(typeAsset) || 0;
        rowData.push(totalDevice);
      }

      transposedData.push(rowData);
    }

    //add to new table
    workbook.getWorksheet("Type Asset by Department")?.delete();
    let transposedSheet = workbook.addWorksheet("Type Asset by Department");
    totalColumn = transposedData[0].length;

    for (i = 0; i < transposedData.length; i++) {
      transposedSheet.getRange(`C${i+1}:${convertNumToAlphabet(totalColumn)}${i+1}`).setValues([transposedData[i]])
    }
    //delete temporary
    workbook.getWorksheet("Type Asset by Department Tmp")?.delete();

    //Draw graprh
    let lastCol = "G";
    let firstA = "C1:"
    let firstB = lastCol.concat(reportTR.toString());
    let finalcol = firstA.concat(firstB);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range 1:1 on selectedSheet
    report.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("1:1").getFormat().setIndentLevel(0);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range B:E on selectedSheet
    report.getRange("C:G").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("C:G").getFormat().setIndentLevel(0);
    // Set fill color to B4C6E7 for range A1:E1 on selectedSheet
    report.getRange("C1:G1").getFormat().getFill().setColor("B4C6E7");

    let chartAssetsByStatus = drawCart(workbook, "Status by Type Assets", finalcol);
    chartAssetsByStatus.setPosition(`B${reportTR + 3}`);
    chartAssetsByStatus.getTitle().setText("Status by Type Assets");

  }

  //ADDONS FUNCTION
  function drawCart(workbook: ExcelScript.Workbook, sheet: string, col: string) {
    const selectedSheet = workbook.getWorksheet(sheet);
    let chart = selectedSheet.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange(col));
    return chart;
  }

  function convertNumToAlphabet(num: number) {
    let str = '';
    while (num > 0) {
      let remainder = (num - 1) % 26;
      str = String.fromCharCode(65 + remainder) + str;
      num = (num - remainder - 1) / 26;
    }
    return str;
  }
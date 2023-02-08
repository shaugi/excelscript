// @ts-nocheck
interface items {
    Title: string,
    Category: {
      Value: string
    },
    Manufacturer: {
      Value: string
    },
    ModelNo_x002e_: string,
    Condition: {
      Value: string
    },
    Location: {
      Value: string
    },
    Status: {
      Value: string
    },
    price: number,
    Dept: string
  }
  interface Category {
    Value: string
  }

  function main(workbook: ExcelScript.Workbook, items: Array<items>){

      let sheet_database = workbook.getWorksheet('database');
      let sheet_report = workbook.getWorksheet('report');
      if(sheet_database){
          sheet_database.delete();
          sheet_database = workbook.addWorksheet("database")
      }

    sheet_database.getRange("A1:I1").setValues([['name', 'type assets', 'manufacture', 'model', 'condition', 'price', 'department', 'location', 'status']]);
    let sheet_database_lastCell = 0;

    //insert data from automate to excel
    items.forEach((data) => {
      sheet_database_lastCell = sheet_database_lastCell+1;
      const name = data.Title;
      const department: string = data.Dept;
      const price = data.price;
      const model = data.ModelNo_x002e_;
      const typeAssetObj = data.Category;
      const typeAsset = typeAssetObj.Value;
      const manufacture = data.Manufacturer.Value;
      const condition = data.Condition.Value;
      const location = data.Location.Value;
      const status = data.Status.Value;

      sheet_database.getCell(sheet_database_lastCell, 0).setValue(name);
      sheet_database.getCell(sheet_database_lastCell, 1).setValue(typeAsset);
      sheet_database.getCell(sheet_database_lastCell, 2).setValue(manufacture);
      sheet_database.getCell(sheet_database_lastCell, 3).setValue(model);
      sheet_database.getCell(sheet_database_lastCell, 4).setValue(condition);
      sheet_database.getCell(sheet_database_lastCell, 5).setValue(price);
      sheet_database.getCell(sheet_database_lastCell, 6).setValue(department);
      sheet_database.getCell(sheet_database_lastCell, 7).setValue(location);
      sheet_database.getCell(sheet_database_lastCell, 8).setValue(status);
    });

    //generate chart
    generate_totalAssets_by_Status(workbook);


  }



  function generate_totalAssets_by_Status(workbook: ExcelScript.Workbook) {
    let database = workbook.getWorksheet("database");
    const database_range = database.getUsedRange();
    const database_values = database_range.getValues();
    const database_lastCell = database_values.length;

    const selectedColumns = database_values.map((row) => {
      return [row[1], row[8]];
    });


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

    workbook.getWorksheet("Status by Type Assets")?.delete();
    const report = workbook.addWorksheet("Status by Type Assets");
    report.getRange("A1:E1").setValues([['Type Asset', 'Borrowed', 'Available', 'Missing', 'Broken']]);

    const reportL = report.getUsedRange();
    const reportV = reportL.getValues();
    let reportTR = reportV.length;

    finalArray.forEach((data) => {
      reportTR += 1;
      report.getRange(`A${reportTR}:E${reportTR}`).setValues([[data.item, data.Borrowed, data.Available, data.Missing, data.Broken]]);
    })

    let lastCol = "E";
    let firstA = "A1:"
    let firstB = lastCol.concat(reportTR.toString());
    let finalcol = firstA.concat(firstB);

    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range 1:1 on selectedSheet
    report.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("1:1").getFormat().setIndentLevel(0);
    // Set horizontal alignment to ExcelScript.HorizontalAlignment.center for range B:E on selectedSheet
    report.getRange("B:E").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
    report.getRange("B:E").getFormat().setIndentLevel(0);
    // Set fill color to B4C6E7 for range A1:E1 on selectedSheet
    report.getRange("A1:E1").getFormat().getFill().setColor("B4C6E7");

    let chartAssetsByStatus = drawCart(workbook, "Status by Type Assets", finalcol);
    chartAssetsByStatus.setPosition(`B${reportTR+3}`);
    chartAssetsByStatus.getTitle().setText("Status by Type Assets");
  }

  function updateRange(sheet: ExcelScript.Worksheet, data: string[][]): ExcelScript.Range {
    const targetRange = sheet.getRange('A1').getResizedRange(data.length - 1, data[0].length - 1);
    targetRange.setValues(data);
    return targetRange;
  }

  function drawCart(workbook: ExcelScript.Workbook, sheet: string, col: string) {
    const selectedSheet = workbook.getWorksheet(sheet);
    console.log(col);
    let chart = selectedSheet.addChart(ExcelScript.ChartType.columnClustered, selectedSheet.getRange(col));
    return chart;
  }


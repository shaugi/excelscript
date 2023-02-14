// @ts-nocheck
//GENERATE TOTAL ASSET BY TYPE AND DEPARTMENT TEMPORARY
function main(workbook: ExcelScript.Workbook) {
	let database = workbook.getWorksheet("database");
	const database_range = database.getUsedRange();
	const database_values = database_range.getValues();

	let selectedColumns = database_values.map((row) => {
		return [row[0], row[1], row[2], row[3]];
	});
	selectedColumns.shift();
	let data = {};

	for (let i = 0; i < selectedColumns.length; i++) {
		let typeAssets = selectedColumns[i][1];
		let department = selectedColumns[i][2];
		let username = selectedColumns[i][3];

		if (!data[department]) {
			data[department] = {};
		}

		if (!data[department][username]) {
			data[department][username] = {};
		}

		if (!data[department][username][typeAssets]) {
			data[department][username][typeAssets] = 0;
		}

		data[department][username][typeAssets] += 1;
	}


	let dataArray = [{}];
	for (let department in data) {
		for (let username in data[department]) {
			let obj = {};
			obj['Department'] = department;
			obj['Username'] = username;
			obj['PC'] = data[department][username]['PC'] || 0;
			obj['Laptop'] = data[department][username]['Laptop'] || 0;
			dataArray.push(obj);
		}
	}
	dataArray.shift();

	const groupedData = [{}];
	const departments = new Set(dataArray.map(item => item.Department));

	for (const department of departments) {
		const accounts = dataArray.filter(item => item.Department === department)
			.map(item => {
				return {
					Username: item.Username,
					PC: item.PC,
					Laptop: item.Laptop
				};
			});
		groupedData.push({ Department: department, accounts });
	}
	groupedData.shift()

	//insert into HOME
	workbook.getWorksheet("Home")?.delete();
	let HomeSheet = workbook.addWorksheet("Home");
	HomeSheet.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	HomeSheet.getRange("2:2").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	let theMostRow = 0;
	let DeptColFirstNum = 1;

	let summary = [{}];
	let totalCountPC = 0;
	let totalCountLaptop = 0;
	//lopping to insert
	for (i = 0; i < groupedData.length; i++) {
		//selec column 4 each department
		let DeptColLastNum = DeptColFirstNum + 2;
		let DeptColFirstAL = convertNumToAlphabet(DeptColFirstNum);
		let DeptColLastAL = convertNumToAlphabet(DeptColLastNum);

		//insert header
		HomeSheet.getRange(`${DeptColFirstAL}1:${DeptColLastAL}1`).setValue([[groupedData[i].Department, null, null,]]);
		HomeSheet.getRange(`${DeptColFirstAL}1:${DeptColLastAL}1`).merge(false);
		HomeSheet.getRange(`${DeptColFirstAL}2:${DeptColLastAL}2`).setValue([["User Name", "Laptop", "PC"]]);

		//config table
		HomeSheet.getRange(`${DeptColFirstNum}:${DeptColFirstNum}`).getFormat().autofitColumns();
		HomeSheet.getRange("1:2").getFormat().getFill().setColor("9BC2E6");

		//insert values

		let indexColValue = 3;
		for (j = 0; j < groupedData[i].accounts.length; j++) {
			HomeSheet.getRange(`${DeptColFirstAL}${indexColValue}:${DeptColLastAL}${indexColValue}`).setValue([[groupedData[i].accounts[j].Username, groupedData[i].accounts[j].PC, groupedData[i].accounts[j].Laptop]]);
			indexColValue += 1;
			totalCountPC = totalCountPC + groupedData[i].accounts[j].PC;
			totalCountLaptop = totalCountLaptop + groupedData[i].accounts[j].Laptop;

		}
		if (j = groupedData[i].accounts.length) {

			summary.push({ Dept: groupedData[i].Department, totalPC: totalCountPC, totalLaptop: totalCountLaptop });
			totalCountPC = 0;
			totalCountLaptop = 0;
		}

		DeptColFirstNum = DeptColLastNum + 1
		if (theMostRow < indexColValue) {
			theMostRow = indexColValue + 1;
		} else {
			theMostRow = theMostRow;
		}
	}
	summary.shift();
	generateSummary(workbook, summary);
}

function generateSummary(workbook: ExcelScript.Workbook, data = [{}]) {
	workbook.getWorksheet("Summary")?.delete();
	let SummarySheet = workbook.addWorksheet("Summary");

	first = 2;
	SummarySheet.getRange("A1:C1").setValue([["Department", "Total PC", "Total Laptop"]]);
	SummarySheet.getRange("A1:C1").getFormat().getFill().setColor("9BC2E6");

	let totalLaptop = 0;
	let totalPC = 0;
	for (i = 0; i < data.length; i++) {
		SummarySheet.getRange(`A${first}:C${first}`).setValue([[data[i].Dept, data[i].totalPC, data[i].totalLaptop]]);

		totalLaptop = totalLaptop + data[i].totalLaptop;
		totalPC = totalPC + data[i].totalPC;

		first += 1
	}
	if (first > data.length) {
		SummarySheet.getRange(`A${first + 1}:C${first + 1}`).setValue([["Total", totalLaptop, totalPC]]);
	}

	SummarySheet.getRange("A:A").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	SummarySheet.getRange("A:A").getFormat().autofitColumns();

	SummarySheet.getRange("B:B").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	SummarySheet.getRange("B:B").getFormat().setColumnWidth(80);

	SummarySheet.getRange("C:C").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	SummarySheet.getRange("C:C").getFormat().setColumnWidth(80);

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
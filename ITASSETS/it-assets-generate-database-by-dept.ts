// @ts-nocheck
interface items {
	Title: string,
	Category: {
		Value: string
	},
	BorrowedBy: string,
	Status: {
		Value: string
	},
	Dept: string
}
interface Category {
	Value: string
}

function main(workbook: ExcelScript.Workbook, items: Array<items>) {

	let sheet_database = workbook.getWorksheet('database');
	if (sheet_database) {
		sheet_database.delete();
		sheet_database = workbook.addWorksheet("database")
	}

	sheet_database.getRange("A1:E1").setValues([['Assets Name', 'Type Assets', 'Department', 'User Name', 'Status']]);
	let sheet_database_lastCell = 0;

	//insert data from automate to excel
	items.forEach((data) => {
		const name = data.Title;
		const department: string = data.Dept;
		const status = data.Status.Value;
		const typeAssetObj = data.Category;
		const typeAsset = typeAssetObj.Value;
		const username = data.BorrowedBy;

		if (status === "Borrowed" && (typeAsset === "Laptop" || typeAsset === "PC")) {
			sheet_database_lastCell = sheet_database_lastCell + 1;
			sheet_database.getCell(sheet_database_lastCell, 0).setValue(name);
			sheet_database.getCell(sheet_database_lastCell, 1).setValue(typeAsset);
			sheet_database.getCell(sheet_database_lastCell, 2).setValue(department);
			sheet_database.getCell(sheet_database_lastCell, 3).setValue(username);
			sheet_database.getCell(sheet_database_lastCell, 4).setValue(status);
		}
	});

	createTable(workbook);
}

function createTable(workbook: ExcelScript.Workbook) {
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
	// const groupedData: { Department: string, accounts: { Username: string, PC: number, Laptop: number }[] }[] = [];
	// const departments = new Set(dataArray.map(item => item.Department));
	const departments: string[] = [];
	for (const item of dataArray) {
		if (item.Department && !departments.includes(item.Department)) {
			departments.push(item.Department);
		}
	}


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
	let homeSheet = workbook.addWorksheet("Home");
	homeSheet.getRange("1:1").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	homeSheet.getRange("2:2").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
	let theMostRow = 0;
	let deptColFirstNum = 1;

	let summary = [{}];
	let totalCountPC = 0;
	let totalCountLaptop = 0;
	//lopping to
	for (i = 0; i < groupedData.length; i++) {
		//selec column 4 each department
		let deptColLastNum = deptColFirstNum + 2;
		let deptColFirstAL = convertNumToAlphabet(deptColFirstNum);
		let deptColLastAL = convertNumToAlphabet(deptColLastNum);

		//insert header
		homeSheet.getRange(`${deptColFirstAL}1:${deptColLastAL}1`).setValue([[groupedData[i].Department, null, null,]]);
		homeSheet.getRange(`${deptColFirstAL}1:${deptColLastAL}1`).merge(false);
		homeSheet.getRange(`${deptColFirstAL}2:${deptColLastAL}2`).setValue([["User Name", "PC", "Laptop"]]);

		//config table
		homeSheet.getRange(`${deptColFirstNum}:${deptColFirstNum}`).getFormat().autofitColumns();
		homeSheet.getRange("1:2").getFormat().getFill().setColor("9BC2E6");
		//insert values
		let indexColValue = 3;
		for (j = 0; j < groupedData[i].accounts.length; j++) {

			homeSheet.getRange(`${deptColFirstAL}${indexColValue}:${deptColLastAL}${indexColValue}`).setValue([[groupedData[i].accounts[j].Username, groupedData[i].accounts[j].PC, groupedData[i].accounts[j].Laptop]]);
			indexColValue += 1;
			totalCountPC = totalCountPC + groupedData[i].accounts[j].PC;
			totalCountLaptop = totalCountLaptop + groupedData[i].accounts[j].Laptop;

		}
		if (j = groupedData[i].accounts.length) {

			summary.push({ Dept: groupedData[i].Department, totalPC: totalCountPC, totalLaptop: totalCountLaptop });
			totalCountPC = 0;
			totalCountLaptop = 0;
		}

		deptColFirstNum = deptColLastNum + 1
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
		SummarySheet.getRange(`A${first + 1}:C${first + 1}`).setValue([["Total", totalPC, totalLaptop]]);
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
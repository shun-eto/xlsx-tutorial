let xlsx = require("xlsx");

/**
 * エクセルファイルの読み込み
 * @param {string} pathname
 * @param {number} sheetIndex
 * @param {Object} JSON
 */
function readToJson(pathname, sheetIndex) {
	const workbook = xlsx.readFile(pathname);
	const sheetNames = workbook.SheetNames;
	const sheet = workbook.Sheets[sheetNames[sheetIndex]];
	const json = xlsx.utils.sheet_to_json(sheet);
	return json;
}

/**
 *
 * @param {*} pathname
 */
function updateWithJson(pathname, sheetIndex, data) {
	const workbook = xlsx.readFile(pathname);
	const sheetNames = workbook.SheetNames;
	const sheet = xlsx.utils.json_to_sheet(data);
	console.log("sheet", sheet);
	xlsx.utils.book_append_sheet(workbook, sheet, "test");
	xlsx.writeFile(workbook, "test.xlsx");
}

const json = readToJson("sample.xlsx", 0);
updateWithJson("sample.xlsx", 0, [
	...json,
	{ id: "006", Name: "Name6", Age: 26, Location: "Tiba" },
]);

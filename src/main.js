const xlsx = require("xlsx");

function csvToExcel(inputs) {
	const wb = xlsx.utils.book_new();
	const ws = xlsx.utils.aoa_to_sheet(inputs);
	xlsx.utils.book_append_sheet(wb, ws, "sheet");
	xlsx.writeFile(wb, "sample1.xlsx");
}

const inputs = [
	["a", "b", "c"],
	[1, 2, 3],
	[1, 2, 3],
];
csvToExcel(inputs);

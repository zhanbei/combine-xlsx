'use strict';

const fs = require('fs');
const XLSX = require('xlsx');

// 给定 Excel 的格式
// 给定的 Excel 第一行是标题，此处定义各个列。
const mTitleKeys = [{from: '学号', to: 'id'}, {from: '姓名', to: 'name'}, {from: '成绩', to: 'score'}];
// 导出的 CSV 文件的格式
// 定义合并后CSV文件的各个列。
const mTargetCsvTitles = [{from: 'id'}, {from: 'name'}, {from: 'score'}];

/* ---------- 列取目标文件 ---------- */
const mTargetFolder = process.argv.length >= 3 ? process.argv[2] || '.' : '.';

const mTargetFiles = fs.readdirSync(mTargetFolder);
const xlsxFiles = mTargetFiles.filter(file => {
	if (file.endsWith('.xlsx')) {return true;}
	console.log('过滤掉非"XLSX"文件：', file);
	return false;
});

if (xlsxFiles.length <= 0) {
	console.error(`目标文件夹"${mTargetFolder}"未发现目标文件！`);
	process.exit(1);
}

console.log();
console.log(`在目标文件夹"${mTargetFolder}"中发现${xlsxFiles.length}个文件，过滤掉了${mTargetFiles.length - xlsxFiles.length}个文件！`);
console.log();

/* ---------- 解析目标文件 ---------- */
const mRawXlsxSheets = [];
const mParsedXlsxSheets = [];
const mCombinedXlsxRows = [];

// 规范给定参数
const mTitleKeysMap = {};
mTitleKeys.map(key => mTitleKeysMap[key.to] = key.from);
mTargetCsvTitles.map(title => title.to ? undefined : title.to = mTitleKeysMap[title.from]);

xlsxFiles.map((file, index) => {
	const workbook = XLSX.readFile(file);
	console.log(`读取并解析第[${index + 1}]个文件：`, file, workbook.SheetNames);
	if (workbook.SheetNames.length < 1) {
		console.error(`当前 XLSX 文件无数据[${index + 1}]：`, file);
		return;
	}
	const worksheet = workbook.Sheets[workbook.SheetNames[0]];
	// console.log(JSON.stringify(worksheet), worksheet['!cols']);
	const xlsx = XLSX.utils.sheet_to_json(worksheet);
	// console.log(JSON.stringify(xlsx));
	const rows = xlsx.map(row => {
		const json = {};
		mTitleKeys.map(key => {
			json[key.to] = row[key.from];
		});
		return json;
	});
	mRawXlsxSheets.push({
		file: file,
		sheetNames: workbook.SheetNames,
		xlsx: xlsx,
	});
	mParsedXlsxSheets.push({
		file: file,
		sheetNames: workbook.SheetNames,
		rows: rows,
	});
	mCombinedXlsxRows.push(...rows);
});

/* ---------- 渲染 CSV 文件 ---------- */
const _cell = (value) => `"${value}"`;
const mCsvSeparator = ', ';
const csvLines = [];
csvLines.push(mTargetCsvTitles.map(title => _cell(title.to)).join(mCsvSeparator));
mCombinedXlsxRows.map(row => {
	csvLines.push(mTargetCsvTitles.map(title => _cell(row[title.from])).join(mCsvSeparator));
});

/* ---------- 输入结果 ---------- */
// console.log('resolved xlsx:', JSON.stringify(mParsedXlsxSheets), JSON.stringify(mCombinedXlsxRows));
// console.log('csv:', csv);
fs.writeFileSync('./output.csv', csvLines.join('\n'));
fs.writeFileSync('./output.raw.json', JSON.stringify(mRawXlsxSheets));
fs.writeFileSync('./output.parsed.json', JSON.stringify(mParsedXlsxSheets));
fs.writeFileSync('./output.mixed.json', JSON.stringify(mCombinedXlsxRows));

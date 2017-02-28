//以下是辅助函数
function datenum(v, date1904) {
	if(date1904) v += 1462;
	var epoch = Date.parse(v);
	return(epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
}

function s2ab(s) {
	var buf = new ArrayBuffer(s.length);
	var view = new Uint8Array(buf);
	for(var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
	return buf;
}

function Workbook() {
	if(!(this instanceof Workbook)) return new Workbook();
	this.SheetNames = [];
	this.Sheets = {};
}

function sheet_from_array_of_arrays(data, opts) {
	var ws = {};
	var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
	for(var R = 0; R != data.length; ++R) {
		for(var C = 0; C != data[R].length; ++C) {
			if(range.s.r > R) range.s.r = R;
			if(range.s.c > C) range.s.c = C;
			if(range.e.r < R) range.e.r = R;
			if(range.e.c < C) range.e.c = C;
			var cell = { v: data[R][C] };
			if(cell.v == null) continue;
			var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

			if(typeof cell.v === 'number') cell.t = 'n';
			else if(typeof cell.v === 'boolean') cell.t = 'b';
			else if(cell.v instanceof Date) {
				cell.t = 'n';
				cell.z = XLSX.SSF._table[14];
				cell.v = datenum(cell.v);
			} else cell.t = 's';
			ws[cell_ref] = cell;
		}
	}

	if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
	return ws;

}

/*
 * 保存数据
 * filename: String, 可选项，文件名，默认为test.xlsx
 * data: Array, 必须项， 为二维数组结构,如： [['名称', '年龄'], ['张三', '12'], ['李四', '23']];
 * sheetName: String, 可选， 为表的名称，暂时只支持一张表
 * wscols: Array, 可选， 为每一列的宽度： 如[12, 10],
 * 	
 */
function saveToExcel(config) {

	var data = config.data;
	var filename = config.filename || 'test.xlsx';
	var sheetName = config.sheetName || 'sheet';

	var wscols = [];
	wscols = config.wscols ? wscols.map(function(cols) { return { wch: cols }; }) : config.data[0].map(function(colname) { return { wch: colname.length * 10 || 10 } });

	var ws = sheet_from_array_of_arrays(data);
	var wb = new Workbook();
	wb.SheetNames.push(sheetName);
	wb.Sheets[sheetName] = ws;

	ws['!cols'] = wscols;

	/* 写入文件 */
	var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' };
	var wbout = XLSX.write(wb, wopts);
	saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), filename);
}

/*
 * 这里是入口
 */
document.getElementById('saveBtn').onclick = function() {
	saveToExcel({
		filename: 'test.xlsx',
		data: [
			['名称', '年龄'],
			['张三', '12'],
			['李四', '23']
		],
		wscols: [20, 20]
	});
};
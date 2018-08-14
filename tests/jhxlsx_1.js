/* 
 * ####################################################################################################
 * https://www.npmjs.com/package/xlsx-style
 * ####################################################################################################
 */
var Jhxlsx = {
    config: {
        fileName: "report",
        extension: ".xlsx",
        sheetName: "Report",
        fileFullName: "report.xlsx",
        header: true,
        createEmptyRow: true,
        maxCellWidth: 20
    },
    tableData: [],
    rowCount: 0,
    wscols: [],
    ws: {},
    range: {s: {c: 10000000, r: 10000000}, e: {c: 0, r: 0}},
    init: function (tableData, options) {
        this.reset();
        this.tableData = tableData;
        for (var key in this.config) {
            if (options.hasOwnProperty(key)) {
                this.config[key] = options[key];
            }
        }
        this.config['fileFullName'] = this.config.fileName + this.config.extension;
    },
    reset: function () {
        this.tableData = [];
        this.rowCount = 0;
        this.wscols = [];
        this.ws = {};
    },
    cellWidth: function (cellText, pos) {
        var max = (cellText.length * 1.3);
        if (this.wscols[pos]) {
            if (max > this.wscols[pos].wch) {
                this.wscols[pos] = {wch: max};
            }
        } else {
            this.wscols[pos] = {wch: max};
        }
    },
    cellWidthValidate: function () {
        for (var i in this.wscols) {
            if (this.wscols[i].wch > this.config.maxCellWidth) {
                this.wscols[i].wch = this.config.maxCellWidth;
            }
        }
    },
    datenum: function (v, date1904) {
        if (date1904)
            v += 1462;
        var epoch = Date.parse(v);
        return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
    },
    jhAddRow: function (rows, style, calColWidth) {

        for (var C = 0; C != rows.length; ++C) {

            if (this.range.s.r > this.rowCount)
                this.range.s.r = this.rowCount;
            if (this.range.s.c > C)
                this.range.s.c = C;
            if (this.range.e.r < this.rowCount)
                this.range.e.r = this.rowCount;
            if (this.range.e.c < C)
                this.range.e.c = C;
            var cellText = rows[C];
            var cell = {v: cellText};
            if (calColWidth) {
                this.cellWidth(cellText, C);
            }
            if (cell.v == null)
                continue;
            var cell_ref = XLSX.utils.encode_cell({c: C, r: this.rowCount});
            if (typeof cell.v === 'number') {
                cell.t = 'n';
            } else if (typeof cell.v === 'boolean') {
                cell.t = 'b';
            } else if (cell.v instanceof Date) {
                cell.t = 'n';
                cell.z = XLSX.SSF._table[14];
                cell.v = this.datenum(cell.v);
            } else {
                cell.t = 's';
            }

            if (style) {
                cell.s = style;
            }

            this.ws[cell_ref] = cell;
        }
        this.rowCount++;
    },
    createWorkSheet: function () {
        var jh = this;
        var merges = [];
        if (jh.tableData[0].header) {
            for (var j in jh.tableData[0].header) {
                var wch = (jh.tableData[0].header[j].length * 1.3);
                jh.wscols[j] = {wch: wch};
            }
        }
        for (var i in jh.tableData) {
            var table = jh.tableData[i];
            if (table.title) {
                for (var j in table.title) {
                    merges.push({s: {r: jh.rowCount, c: 0}, e: {r: jh.rowCount, c: table.header.length - 1}});
                    jh.jhAddRow([table.title[j]], table.styles.title, false);
                }
            }
            if (jh.config.header) {
                jh.jhAddRow(table.header, table.styles.header, false);
            }
            if (table.subHeader) {
                for (var j in table.subHeader) {
                    jh.jhAddRow(table.subHeader[j], table.styles.header, true);
                }
            }
            for (var j in table.body) {
                jh.jhAddRow(table.body[j], table.styles.body, true);
            }
            if (jh.config.createEmptyRow) {
                jh.rowCount++;
            }
        }
        this.cellWidthValidate();
        jh.ws['!merges'] = merges;
        jh.ws['!cols'] = this.wscols;
        if (jh.range.s.c < 10000000)
            jh.ws['!ref'] = XLSX.utils.encode_range(jh.range);
        return jh.ws;
    },
    s2ab: function (s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i != s.length; ++i)
            view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    },
    export: function (tableData, options) {
        this.init(tableData, options);
        var wb = new Workbook();
        this.createWorkSheet();
        /* add worksheet to workbook */
        wb.SheetNames.push(this.config.sheetName);
        wb.Sheets[this.config.sheetName] = this.ws;
        var wbout = XLSX.write(wb, {bookType: 'xlsx', bookSST: true, type: 'binary'});
        saveAs(new Blob([this.s2ab(wbout)], {type: "application/octet-stream"}), this.config.fileFullName)
    }
}

function Workbook() {
    if (!(this instanceof Workbook))
        return new Workbook();
    this.SheetNames = [];
    this.Sheets = {};
}

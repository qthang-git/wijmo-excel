import * as wjcCore from '@grapecity/wijmo';
import * as wjcGrid from '@grapecity/wijmo.grid';
import * as wjcGridXlsx from '@grapecity/wijmo.grid.xlsx';
import * as wjcXlsx from '@grapecity/wijmo.xlsx';
//
const ExcelExportDocName = 'TEMPLATE-DEFAULT.xlsx';
//
export class ExportService {
    constructor() {
        this._wb = null;
        this._ws = [];
        this._wscolumns = [];
        this._wsrows = [];
        this._wscount = 0;
        this._wsname = 'テスト仕様書';
        this._excelExportDocName = ExcelExportDocName;
        this._ws_font_family = 'Meiryo UI';
        this._arr_wsname = [];
        this._arr_wsindex = [];
        this._objGroup = {};
    }
    startExcelExport(flex, ctx) {
        if (ctx.preparing || ctx.exporting) {
            return;
        }
        ctx.exporting = false;
        ctx.progress = 0;
        ctx.preparing = true;
        if (flex.xlsx_name != '') {
            this._excelExportDocName = flex.xlsx_name + '.xlsx';
        }
        if (flex.worksheet_name != '') {
            this._wsname = flex.worksheet_name + '_' + this._wsname;
        }
        this._wscount = flex.worksheet_count;
        this._arr_wsindex = flex.worksheet_index;
        this._objGroup = flex._objGroup;
        this._createOtherWSName(flex);
        this._wb = wjcGridXlsx.FlexGridXlsxConverter.saveAsync(flex, {
            includeColumnHeaders: true,
            includeStyles: false,
            includeColumns: this._includeColumns.bind(this),
            formatItem: this._formatItemExcel.bind(this)
        })
        this._formatWorksheet(flex);
        this._wb.saveAsync(this._excelExportDocName, () => {
            console.log('Export to Excel completed');
            this._resetExcelContext(ctx);
        }, err => {
            console.error(`Export to Excel failed: ${err}`);
            this._resetExcelContext(ctx);
        }, prg => {
            if (ctx.preparing) {
                ctx.exporting = true;
                ctx.preparing = false;
            }
            ctx.progress = prg / 100.;
        }, true);
        console.log('Export to Excel started');
    }
    cancelExcelExport(ctx) {
        wjcGridXlsx.FlexGridXlsxConverter.cancelAsync(() => {
            console.log('Export to Excel canceled');
            this._resetExcelContext(ctx);
        });
    }
    _createOtherWSName(flex) {
        let arr_wsname = [];
        let keys = Object.keys(this._objGroup);
        for (let i = 0; i < this._arr_wsindex.length; i++) {
            let sheetChildNo = this._arr_wsindex[i];
            if (keys.length != 0) {
                for (let j = 0; j < keys.length; j++) {
                    let Group = this._objGroup[keys[j]];
                    if (Group.findIndex(gr => gr == sheetChildNo) != -1) {
                        let fromSheet = Group[0];
                        let toSheet = Group[Group.length - 1];
                        if (fromSheet != toSheet) {
                            sheetChildNo = fromSheet + "~" + toSheet;
                            break;
                        }
                    }
                }
            }
            let sheetChildName = 'エビデンス（No ' + sheetChildNo + '. ' + flex.worksheet_name + '）';
            arr_wsname.push(sheetChildName);
        }
        this._arr_wsname = [...new Set(arr_wsname)];
    }
    _formatItemExcel(e) {
    }
    _formatWorksheet(flex) {
        this._ws = this._wb.sheets;
        this._ws[0].name = this._wsname;
        this._wscolumns = this._ws[0].columns;
        this._wsrows = this._ws[0].rows;
        // disable freeze row & column
        this._ws[0].frozenPane.columns = 0;
        this._ws[0].frozenPane.rows = 0;
        // show all column
        this._visibleAllColumns(this._wscolumns, flex);
        this._wscolumns.unshift(this._newEmptyColumn(18));
        // template main
        let countrows = this._wsrows.length;
        let idxFomula = 0;
        for (let i = 0; i < countrows; i++) {
            let cells = this._wsrows[i].cells;
            cells.unshift(this._newEmptyCell());
            this._wsrows[i].height = i == 0 ? this._convertToPixel(14.25) : this._convertToPixel(18.75);
            let countcells = cells.length;
            let status_row = 1;
            if (i != 0) {
                status_row = flex.rows[i - 1].dataItem.status;
            }
            for (let j = 0; j < countcells; j++) {
                let cell = cells[j];
                let style = cell.style;
                style.font = {};
                style.fill = {};
                style.borders = {
                    top: {
                        color: '#000',
                        style: 0
                    },
                    right: {
                        color: '#000',
                        style: 1
                    },
                    bottom: {
                        color: '#000',
                        style: 1
                    },
                    left: {
                        color: '#000',
                        style: 0
                    }
                };
                style.font.size = this._convertToPixel(10);
                style.font.family = this._ws_font_family;
                if (i == 0) { // header
                    style.hAlign = wjcXlsx.HAlign.Center;
                    style.borders.top.style = 1;
                    if (j == 0) {
                        cell.colSpan = 2;
                        cell.value = 'No';
                        style.borders.left.style = 1;
                    }
                    // fill background header column
                    if (j <= 5) {
                        style.fill.color = 'rgb(226, 239, 218)';
                    }
                    else {
                        style.fill.color = 'rgb(217, 225, 242)';
                    }
                }
                else {
                    let previous_cell = this._wsrows[i - 1].cells[j];
                    if (status_row == 0) {
                        this._wsrows[i].height = 19;
                        style.fill.color = 'rgb(255, 242, 204)';
                        cell.colSpan = countcells;
                        if (j < countcells - 1) {
                            style.borders.right.style = 0;
                        }
                        if (j == 0) {
                            cell.value = cells[3].value;
                            cells[3].value = '';
                        }
                    }
                    if (j == 0) {
                        style.fill.color = 'rgb(255, 242, 204)';
                        style.borders.top.style = 0;
                        style.borders.left.style = 1;
                        if (i < countrows - 1) {
                            if (i == 1) {
                                style.borders.right.style = 0;
                            }
                            style.borders.bottom.style = 0;
                        }
                    }
                    if (i > 1 && j == 1) {
                        if (cell.value != '' && cell.value != "'") {
                            style.format = 'General';
                            style.font.underline = true;
                            style.font.color = '#4F81BD';
                            if (i != 2) {
                                let minus_row = idxFomula != 0 ? ('-' + idxFomula) : '';
                                cell.formula = '=OFFSET(INDIRECT(ADDRESS(ROW()' + minus_row + ',COLUMN())), -1, 0)+1';
                            }
                            cell.link = "#'" + this._arr_wsname[cell.value - 1] + "'!B1";
                            if (previous_cell.value == '' || previous_cell.value == "'") {
                                idxFomula = 0;
                            }
                        } else {
                            idxFomula += 1;
                        }
                    }
                }
            }
        }
        // insert row to excel
        for (var idx = 0; idx < 5; idx++) {
            this._wsrows.unshift(this._newEmptyRow(this._wscolumns.length));
        }
        // insert column empty to excel
        this._wscolumns.unshift(this._newEmptyColumn(12));
        countrows = this._wsrows.length;
        for (let i = 0; i < countrows; i++) {
            let cells = this._wsrows[i].cells;
            cells.splice(0, 0, this._newEmptyCell());
            if (i < 5) {
                let value = '';
                let bgcolor = '';
                let colSpan = 0;
                let borders = {};
                if (i == 1 || i == 2) {
                    borders = {
                        top: {
                            color: '#000',
                            style: 1
                        },
                        right: {
                            color: '#000',
                            style: 1
                        },
                        bottom: {
                            color: '#000',
                            style: 1
                        },
                        left: {
                            color: '#000',
                            style: 1
                        }
                    };
                    value = i == 1 ? '利用ブラウザ' : '環境';
                    bgcolor = 'rgb(226, 239, 218)';
                    colSpan = 3;
                }
                [1, 2, 3, 4, 5, 6].forEach(item => {
                    cells[item].style.borders = borders;
                    cells[item].style.font.size = this._convertToPixel(10);
                    cells[item].style.font.family = this._ws_font_family;
                    if (item == 1) {
                        cells[item].colSpan = colSpan;
                        cells[item].value = value;
                        cells[item].style.hAlign = wjcXlsx.HAlign.Center;
                        cells[item].style.fill.color = bgcolor;
                    }
                    else if (item == 4) {
                        cells[item].colSpan = colSpan;
                        cells[item].value = i == 1 ? 'FireFox 117.0 (64 ビット)' : '';
                    }
                })
            }
        }
        this._wsrows[4].cells[3].formula = '="総数「" & MAX(C5:C49653) & "」　OK「" & COUNTIF(J5:J49653,"OK") & "」　NG「" & COUNTIF(J5:J49653,"NG") & "」　未「" & MAX(C5:C49653)-(COUNTIF(J5:J49653,"OK")+COUNTIF(J5:J49653,"NG")) & "」　進捗「" & ROUND((COUNTIF(J5:J49653,"OK")+COUNTIF(J5:J49653,"NG"))/MAX(C5:C49653)*100,1) & "%」"';
        this._addNewWorkSheet(flex);
    }
    _newEmptyCell() {
        let cellEmpty = new wjcXlsx.WorkbookCell();
        cellEmpty.style = new wjcXlsx.WorkbookStyle();
        cellEmpty.style.fill = {};
        cellEmpty.style.font = {};
        cellEmpty.style.borders = {};
        cellEmpty.colSpan = 0;
        cellEmpty.rowSpan = 0;
        cellEmpty.HAlign = wjcXlsx.HAlign.Left;
        cellEmpty.value = '';
        return cellEmpty;
    }
    _newEmptyRow(countcells) {
        let rowEmpty = new wjcXlsx.WorkbookRow();
        rowEmpty.visible = true;
        rowEmpty.height = 19;
        for (var i = 0; i < countcells; i++) {
            let _tempCell = this._newEmptyCell();
            _tempCell.style.font.family = this._ws_font_family;
            _tempCell.style.font.size = this._convertToPixel(10);
            rowEmpty.cells.push(_tempCell);
        }
        return rowEmpty;
    }
    _newEmptyColumn(width) {
        let columnEmpty = new wjcXlsx.WorkbookColumn();
        columnEmpty.autoWidth = true;
        columnEmpty.style = {
            format: '',
            hAlign: 1
        };
        columnEmpty.visible = true;
        if (typeof width !== 'undefined') {
            columnEmpty.width = this._convertToPixel(width);
        }
        return columnEmpty;
    }
    _newEmptyWorkSheet(sheetname, row = 4, column = 20) {
        let worksheet = {};
        worksheet.name = sheetname;
        worksheet.visible = true;
        worksheet.frozenPane = {};
        worksheet.frozenPane.rows = row;
        worksheet.frozenPane.columns = 0;
        worksheet.rows = [];
        worksheet.columns = [];
        for (let i = 0; i < row; i++) {
            worksheet.rows.push(this._newEmptyRow(column));
        }
        for (let i = 0; i < column; i++) {
            worksheet.columns.push(this._newEmptyColumn());
        }
        return worksheet;
    }
    _addNewWorkSheet(flex) {
        for (let i = 0; i < this._arr_wsname.length; i++) {
            let sheetname = this._arr_wsname[i];
            let hyperlink_cell = flex.itemsSource.items.findIndex(item => item.no == this._arr_wsindex[i]) + 7;
            let worksheet = this._createContentForWS(this._newEmptyWorkSheet(sheetname), hyperlink_cell)
            this._ws.push(worksheet);
        }
    }
    _createContentForWS(worksheet, hyperlink_cell) {
        // hyperlink back to main sheet
        worksheet.rows[0].cells[1].style.font.underline = true;
        worksheet.rows[0].cells[1].style.font.color = '#4F81BD';
        worksheet.rows[0].cells[1].value = '戻る';
        worksheet.rows[0].cells[1].link = "#'" + this._wsname + "'!C" + hyperlink_cell;
        // header with light green background
        // get text '操作' from main sheet
        worksheet.rows[1].cells[1].formula = "='" + this._wsname + "'!E6";
        worksheet.rows[1].cells[1].colSpan = 10;
        worksheet.rows[1].cells[1].style.fill.color = 'rgb(226, 239, 218)';
        // get text '確認事項' from main sheet
        worksheet.rows[1].cells[11].formula = "='" + this._wsname + "'!G6";
        worksheet.rows[1].cells[11].style.fill.color = 'rgb(226, 239, 218)';
        worksheet.rows[1].cells[11].colSpan = 9;
        // 
        worksheet.rows[2].cells[1].formula = "='" + this._wsname + "'!E" + hyperlink_cell;
        worksheet.rows[2].cells[1].colSpan = 10;
        worksheet.rows[2].cells[11].formula = "='" + this._wsname + "'!G" + hyperlink_cell;
        worksheet.rows[2].cells[11].colSpan = 9;

        // add border for cell
        for (let i = 1; i < 4 - 1; i++) {
            for (let j = 1; j < 20; j++) {
                worksheet.rows[i].cells[j].style.borders = {
                    top: {
                        color: '#000',
                        style: 1
                    },
                    right: {
                        color: '#000',
                        style: 1
                    },
                    bottom: {
                        color: '#000',
                        style: 1
                    },
                    left: {
                        color: '#000',
                        style: 1
                    }
                }
            }
        }
        return worksheet;
    }
    _visibleAllColumns(columns, flex) {
        columns.forEach((col, idx) => {
            if (!col.visible) {
                col.width = 50;
            }
            if (['operation', 'checklist'].includes(flex.columns[idx].binding)) {
                col.width = 520;
            }
            col.visible = true;
        });
    }
    _includeColumns(column) {
        // remove 3 columns button
        return !(column.binding.includes('btn') || column.binding.includes('group'));
    }
    // ポイントをピクセルに変換する 1 pixel = 0.75 point
    _convertToPixel(point) {
        return point * 4 / 3;
    }
    _resetExcelContext(ctx) {
        ctx.exporting = false;
        ctx.progress = 0;
        ctx.preparing = false;
    }
}

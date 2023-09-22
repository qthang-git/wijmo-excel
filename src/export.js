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
    }
    startExcelExport(flex, ctx) {
        if (ctx.preparing || ctx.exporting) {
            return;
        }
        ctx.exporting = false;
        ctx.progress = 0;
        ctx.preparing = true;
        const ExcelDocName = flex.FileName == '' ? ExcelExportDocName : flex.FileName + '.xlsx';
        this._wb = wjcGridXlsx.FlexGridXlsxConverter.saveAsync(flex, {
            includeColumnHeaders: true,
            includeStyles: false,
            includeColumns: this._includeColumns.bind(this),
            formatItem: this._formatItemExcel.bind(this)
        })
        this._formatWorksheet(flex);
        // return;
        this._wb.saveAsync(ExcelDocName, () => {
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
    _formatItemExcel(e) {
    }
    _formatWorksheet(flex) {
        this._ws = this._wb.sheets;
        this._ws[0].name = flex.SheetName + '_テスト仕様書';
        this._wscolumns = this._ws[0].columns;
        this._wsrows = this._ws[0].rows;
        // disable freeze row & column
        this._ws[0].frozenPane.columns = 0;
        this._ws[0].frozenPane.rows = 0;
        // show all column
        this._visibleAllColumns(this._wscolumns, flex);
        let columnEmpty = new wjcXlsx.WorkbookColumn();
        columnEmpty.width = this._convertToPixel(18);
        this._wscolumns.unshift(columnEmpty);
        // template main
        let countrows = this._wsrows.length;
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
                style.font.family = 'Meiryo UI';
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
                    if (status_row == 0) {
                        // cell.colSpan = countcells;
                        this._wsrows[i].height = 19;
                        style.fill.color = 'rgb(255, 242, 204)';
                        if (j < countcells - 1) {
                            style.borders.right.style = 0;
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
                }
            }
        }
        // insert row to excel
        for (var idx = 0; idx < 5; idx++) {
            this._wsrows.unshift(this._newEmptyRow(this._wscolumns.length));
        }
        // insert column empty to excel
        columnEmpty.width = this._convertToPixel(12);
        this._wscolumns.unshift(columnEmpty);
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
                    cells[item].style.font.family = 'Meiryo UI';
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
        return cellEmpty;
    }
    _newEmptyRow(countcells) {
        let rowEmpty = new wjcXlsx.WorkbookRow();
        rowEmpty.visible = true;
        rowEmpty.height = 19;
        for (var i = 0; i < countcells; i++) {
            rowEmpty.cells.push(this._newEmptyCell());
        }
        return rowEmpty;
    }
    _addNewWorkSheet() {
        // this._ws.push()
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
        return !column.binding.includes('btn');
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

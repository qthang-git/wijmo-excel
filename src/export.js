import * as wjcCore from '@grapecity/wijmo';
import * as wjcGrid from '@grapecity/wijmo.grid';
import * as wjcGridXlsx from '@grapecity/wijmo.grid.xlsx';
import * as wjcXlsx from '@grapecity/wijmo.xlsx';
//
const ExcelExportDocName = 'TEMPLATE-DEFAULT.xlsx';
const ExcelExportSheetName = 'テスト仕様書';
//
export class ExportService {
    constructor() {
        this._wb = null;
        this._ws = [];
        this._wscolumns = [];
        this._wsrows = [];
        this._wscount = 0;
        this._wsname = ExcelExportSheetName;
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
            this._wsname = '';
            this._wsname = flex.worksheet_name + '_' + ExcelExportSheetName;
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
        this._arr_wsname = [];
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
            this._arr_wsname.push(sheetChildName);
        }
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
            this._wsrows[i].height = i == 0 ? this._convertToPixel(14.25) : this._wsrows[i].height > this._convertToPixel(18.75) ? this._wsrows[i].height : this._convertToPixel(18.75);
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
                style.vAlign = wjcXlsx.VAlign.Center;
                if (i == 0) { // header
                    style.hAlign = wjcXlsx.HAlign.Center;
                    style.borders.top.style = 1;
                    style.font.bold = true;
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
                        // cell.colSpan = countcells;
                        if (j < countcells - 1) {
                            style.borders.right.style = 0;
                            // style.borders.bottom.style = 0;
                        }
                        cell.value = undefined;
                        if (j == 0) {
                            cell.value = cells[3].value;
                            style.font.bold = true;

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
                    // begin from row 2
                    if (i > 1) {
                        // remove value "'" when if cell in grid is empty
                        if (cell.value == "'") {
                            cell.value = undefined;
                        }
                        style.format = 'General';
                        // column [no]
                        if (j == 1) {
                            if (cell.value != undefined) {
                                if (i != 2) {
                                    let minus_row = idxFomula != 0 ? ('-' + idxFomula) : '';
                                    cell.formula = '=OFFSET(INDIRECT(ADDRESS(ROW()' + minus_row + ',COLUMN())), -1, 0)+1';
                                }
                                if (flex._isCreateSheetChild) {
                                    style.font.underline = true;
                                    style.font.color = '#4F81BD';
                                    cell.link = "#'" + this._arr_wsname[cell.value - 1] + "'!B1";
                                }
                                if (previous_cell.value == undefined) {
                                    idxFomula = 0;
                                }
                            } else {
                                idxFomula += 1;
                            }
                        }
                        // column [date]
                        else if (j == 6) {
                            style.format = 'yyyy-mm-dd';
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
        if (flex._isCreateSheetChild) {
            this._addNewWorkSheet(flex);
        }
    }
    // initializes empty cell
    _newEmptyCell() {
        let cellEmpty = new wjcXlsx.WorkbookCell();
        cellEmpty.style = new wjcXlsx.WorkbookStyle();
        cellEmpty.style.fill = {};
        cellEmpty.style.font = {};
        cellEmpty.style.borders = {};
        cellEmpty.colSpan = 0;
        cellEmpty.rowSpan = 0;
        cellEmpty.HAlign = wjcXlsx.HAlign.Left;
        cellEmpty.value = undefined;
        return cellEmpty;
    }
    // initializes empty row
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
    // initializes empty column
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
    // initializes empty worksheet
    _newEmptyWorkSheet(sheetname, i, row = 4, column = 20) {
        if (this._ws.filter(m => m.name == sheetname).length != 0) {
            return;
        }
        let worksheet = {};
        worksheet.name = sheetname;
        worksheet.visible = true;

        worksheet.rows = [];
        worksheet.columns = [];
        for (let i = 0; i < row; i++) {
            worksheet.rows.push(this._newEmptyRow(column));
        }
        // when exists group row
        const objGroup = this._objGroup;
        for (let key in objGroup) {
            if (typeof objGroup[key].find(obj => obj == i) !== 'undefined') {
                const GroupRow = objGroup[key];
                for (let j = 1; j < GroupRow.length; j++) {
                    worksheet.rows.splice((j + 2), 0, this._newEmptyRow(column));
                }
            }
        }
        for (let i = 0; i < column; i++) {
            worksheet.columns.push(this._newEmptyColumn());
        }
        worksheet.frozenPane = {};
        worksheet.frozenPane.rows = worksheet.rows.length;
        worksheet.frozenPane.columns = 0;
        return worksheet;
    }
    // add child worksheet into main worksheet
    _addNewWorkSheet(flex) {
        for (let i = 0; i < this._arr_wsindex.length; i++) {
            let sheetname = this._arr_wsname[i];
            let row_no = this._arr_wsindex[i];
            let row_item = flex.itemsSource.items.find(item => item.no == row_no);
            let hyperlink_cell = flex.itemsSource.items.findIndex(item => item.no == row_no) + 7;
            let worksheet = this._newEmptyWorkSheet(sheetname, this._arr_wsindex[i]);
            if (typeof worksheet !== 'undefined') {
                this._ws.push(this._createContentForWS(flex, worksheet, hyperlink_cell, row_item));
            }
        }
    }
    // create content for worksheet child
    _createContentForWS(flex, worksheet, hyperlink_cell, row_item) {
        // header with light green background
        // get text '操作' from main sheet
        worksheet.rows[1].cells[1].formula = "='" + this._wsname + "'!E6";
        worksheet.rows[1].cells[1].colSpan = 10;
        worksheet.rows[1].cells[1].style.fill.color = 'rgb(226, 239, 218)';
        // get text '確認事項' from main sheet
        worksheet.rows[1].cells[11].formula = "='" + this._wsname + "'!G6";
        worksheet.rows[1].cells[11].style.fill.color = 'rgb(226, 239, 218)';
        worksheet.rows[1].cells[11].colSpan = 9;
        // hyperlink back to main sheet
        worksheet.rows[0].cells[1].style.font.underline = true;
        worksheet.rows[0].cells[1].style.font.color = '#4F81BD';
        worksheet.rows[0].cells[1].value = '戻る';
        worksheet.rows[0].cells[1].link = "#'" + this._wsname + "'!C" + hyperlink_cell;
        // let item = flex.itemsSource.items.find(item => item.no == this._arr_wsindex[indexSheet]);
        let idxBorder = 4;
        const keys = Object.keys(this._objGroup);
        if (keys.length > 0) {
            if (typeof row_item !== 'undefined' && row_item.group != '') {
                let len = flex.itemsSource.items.filter(flr => flr.group == row_item.group).length;
                idxBorder += len - 1;
                for (let i = 0; i < len; i++) {
                    // get test case main -> child
                    worksheet.rows[i + 2].cells[1].formula = "='" + this._wsname + "'!E" + (hyperlink_cell + i);
                    worksheet.rows[i + 2].cells[1].colSpan = 10;
                    worksheet.rows[i + 2].cells[11].formula = "='" + this._wsname + "'!G" + (hyperlink_cell + i);
                    worksheet.rows[i + 2].cells[11].colSpan = 9;
                }
            } else {
                // get test case main -> child
                worksheet.rows[2].cells[1].formula = "='" + this._wsname + "'!E" + hyperlink_cell;
                worksheet.rows[2].cells[1].colSpan = 10;
                worksheet.rows[2].cells[11].formula = "='" + this._wsname + "'!G" + hyperlink_cell;
                worksheet.rows[2].cells[11].colSpan = 9;
            }
        } else {
            // get test case main -> child
            worksheet.rows[2].cells[1].formula = "='" + this._wsname + "'!E" + hyperlink_cell;
            worksheet.rows[2].cells[1].colSpan = 10;
            worksheet.rows[2].cells[11].formula = "='" + this._wsname + "'!G" + hyperlink_cell;
            worksheet.rows[2].cells[11].colSpan = 9;
        }
        // add border for cell
        for (let i = 1; i < idxBorder - 1; i++) {
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
    // show all columns on the grid
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
    // handle event include column to excel
    _includeColumns(column) {
        // remove 3 columns button
        return !(column.binding.includes('btn') || column.binding.includes('group'));
    }
    // convert pixel to point (1 pixel = 0.75 point)
    _convertToPixel(point) {
        return point * 4 / 3;
    }
    _resetExcelContext(ctx) {
        ctx.exporting = false;
        ctx.progress = 0;
        ctx.preparing = false;
    }
}

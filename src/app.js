import 'bootstrap.css';
import '@grapecity/wijmo.styles/wijmo.css';
import './styles.css';
//
import '@grapecity/wijmo.touch';
import * as wjcCore from '@grapecity/wijmo';
import * as wjcGrid from '@grapecity/wijmo.grid';
import { CellMaker } from '@grapecity/wijmo.grid.cellmaker';
import { DataService } from './data';
import { ExportService } from './export';
import 'https://cdnjs.cloudflare.com/ajax/libs/jquery/3.7.1/jquery.min.js';
//
class App {
    constructor(dataSvc, exportSvc) {
        this._lastId = 5;
        this._dataSvc = dataSvc;
        this._exportSvc = exportSvc;
        // initializes export
        const btnExportToExcel = document.getElementById('btnExportToExcel');
        this._excelExportContext = new ExcelExportContext(btnExportToExcel);
        btnExportToExcel.addEventListener('click', () => {
            this._exportToExcel();
        });
        // initializes the grid
        this._initializeGrid();
        this._formatItem();
        // initializes items source
        this._itemsSource = this._createItemsSource();
        this._theGrid.itemsSource = this._itemsSource;
    }
    close() {
        const ctx = this._excelExportContext;
        this._exportSvc.cancelExcelExport(ctx);
    }
    _initializeGrid() {
        // creates columns
        this._columns = [
            { binding: 'no', header: 'No', width: 50, isReadOnly: true, dataType: "String", align: "right" },
            { binding: 'user', header: 'user', width: 30, visible: false, dataType: "String", align: "left" },
            { binding: 'operation', header: '操作', width: '*', wordWrap: true, dataType: "String", align: "left" },
            { binding: 'tag', header: 'tag', width: 30, visible: false, dataType: "String", align: "left" },
            { binding: 'checklist', header: '確認事項', width: '*', wordWrap: true, dataType: "String", align: "left" },
            { binding: 'date', header: '日付', width: 50, visible: false, dataType: "Date", align: "left" },
            { binding: 'verifier', header: '検証者', width: 50, visible: false, dataType: "String", align: "left" },
            { binding: 'result', header: '結果', width: 50, visible: false, dataType: "String", align: "left" },
            { binding: 'notes', header: '備考', width: 50, visible: false, dataType: "String", align: "left" },
            {
                binding: 'btn-add', header: ' ', width: 30, minWidth: 30, maxWidth: 30,
                cellTemplate: CellMaker.makeLink({
                    text: [
                        '<div style="text-align: center;display: table;width: 100%;height: 100%;">',
                        '<div style="display: table-cell;vertical-align: middle;">',
                        `<?xml version="1.0" ?><!DOCTYPE svg  PUBLIC '-//W3C//DTD SVG 1.1//EN'  'http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd'><svg enable-background="new 0 0 512 512" height="512px" id="Layer_1" version="1.1" viewBox="0 0 512 512" width="512px" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><path d="M256,512C114.625,512,0,397.391,0,256C0,114.609,114.625,0,256,0c141.391,0,256,114.609,256,256  C512,397.391,397.391,512,256,512z M256,64C149.969,64,64,149.969,64,256s85.969,192,192,192c106.047,0,192-85.969,192-192  S362.047,64,256,64z M288,384h-64v-96h-96v-64h96v-96h64v96h96v64h-96V384z"/></svg>`,
                        '</div>',
                        '</div>'
                    ].join(""),
                    click: (e, ctx) => {
                        this._addRow(ctx.row.index);
                    }
                })
            },
            {
                binding: 'btn-del', header: ' ', width: 30, minWidth: 30, maxWidth: 30,
                cellTemplate: CellMaker.makeLink({
                    text: [
                        '<div style="text-align: center;display: table;width: 100%;height: 100%;">',
                        '<div style="display: table-cell;vertical-align: middle;">',
                        `<?xml version="1.0" ?><!DOCTYPE svg  PUBLIC '-//W3C//DTD SVG 1.1//EN'  'http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd'><svg enable-background="new 0 0 512 512" height="512px" id="Layer_1" version="1.1" viewBox="0 0 512 512" width="512px" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><g><g><path d="M256,0C114.625,0,0,114.625,0,256c0,141.391,114.625,256,256,256c141.391,0,256-114.609,256-256    C512,114.625,397.391,0,256,0z M256,448c-106.031,0-192-85.969-192-192S149.969,64,256,64c106.047,0,192,85.969,192,192    S362.047,448,256,448z M128,288h256v-64H128V288z"/></g></g></svg>`,
                        '</div>',
                        '</div>'
                    ].join(""),
                    click: (e, ctx) => {
                        this._delRow(ctx.row.index);
                    }
                })
            },
            {
                binding: 'btn-copy', header: ' ', width: 30, minWidth: 30, maxWidth: 30,
                cellTemplate: CellMaker.makeLink({
                    text: [
                        '<div style="text-align: center;display: table;width: 100%;height: 100%;">',
                        '<div style="display: table-cell;vertical-align: middle;">',
                        `<?xml version="1.0" ?><!DOCTYPE svg  PUBLIC '-//W3C//DTD SVG 1.1//EN'  'http://www.w3.org/Graphics/SVG/1.1/DTD/svg11.dtd'><svg enable-background="new 0 0 512 512" height="512px" id="Layer_1" version="1.1" viewBox="0 0 512 512" width="512px" xml:space="preserve" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink"><g><g><path d="M480,0H224c-17.688,0-32,14.312-32,32v256c0,17.688,14.312,32,32,32h256c17.688,0,32-14.312,32-32V32    C512,14.312,497.688,0,480,0z M448,256H256V64h192V256z M256,448H64V256h64v-64H32c-17.688,0-32,14.312-32,32v256    c0,17.688,14.312,32,32,32h256c17.688,0,32-14.312,32-32v-96h-64V448z"/></g></g></svg>`,
                        '</div>',
                        '</div>'
                    ].join(""),
                    click: (e, ctx) => {
                        this._copyRow(ctx.row.index);
                    }
                })
            }
        ]
        // creates the grid
        this._theGrid = new wjcGrid.FlexGrid('#theGrid', {
            autoRowHeights: true,
            autoGenerateColumns: false,
            showMarquee: true,
            columns: this._columns
        });
        this._theGrid.select(-1, -1);
    }
    _addRow(index) {
        const data = this._itemsSource.sourceCollection;
        const obj = {};
        obj.json = data;
        obj.json.splice(index + 1, 0, {
            no: index + 1,
            testcase: '',
            result: '',
            status: 1
        });
        for (var i = index + 1; i < obj.json.length; i++) {
            obj.json[i].no = i + 1;
        }
        this._theGrid.itemsSource.sourceCollection = obj.json;
        this._theGrid.select(-1, -1);
        this._theGrid.itemsSource.refresh();
    }
    _delRow(index) {
        const data = this._itemsSource.sourceCollection;
        const rowcount = data.length;
        const obj = {};
        obj.json = data;
        if (rowcount > 1) {
            obj.json.splice(index, 1);
            for (var i = index; i < obj.json.length; i++) {
                obj.json[i].no = i + 1;
            }
            this._theGrid.itemsSource.sourceCollection = obj.json;
            this._theGrid.select(-1, -1);
            this._theGrid.itemsSource.refresh();
        }
    }
    _copyRow(index) {
        const data = this._itemsSource.sourceCollection;
        const obj = {};
        obj.json = data;
        obj.json.splice(index + 1, 0, {
            no: index + 1,
            testcase: obj.json[index].testcase,
            result: obj.json[index].result,
            status: obj.json[index].status
        });
        for (var i = index + 1; i < obj.json.length; i++) {
            obj.json[i].no = i + 1;
        }
        this._theGrid.itemsSource.sourceCollection = obj.json;
        this._theGrid.select(-1, -1);
        this._theGrid.itemsSource.refresh();
    }
    _exportToExcel() {
        const ctx = this._excelExportContext;
        if (!ctx.exporting) {
            // this._customizeGridForExcel();
            this._exportSvc.startExcelExport(this._theGrid, ctx);
        }
        else {
            this._exportSvc.cancelExcelExport(ctx);
        }
    }
    _customizeGridForExcel() {
        // remove 3 columns contains button add/del/copy
        let columns = this._theGrid.columns;
        let columnsCount = columns.length;
        columns.splice(columnsCount - 3, 3);
    }
    _createItemsSource() {
        const data = this._dataSvc.getData(5);
        const view = new wjcCore.CollectionView(data);
        view.collectionChanged.addHandler((s, e) => {
        });
        return view;
    }
    _formatItem() {
        this._theGrid.formatItem.addHandler(function (s, e) {
            if (e.panel == s.columnHeaders) {
                if (e.row == 0) {
                    e.cell.style.textAlign = 'center';
                }
                s.countTestCase = s.itemsSource.sourceCollection.filter(item => item.status != 0).length;
            }
            if (e.panel == s.rowHeaders) {
                const _id = e.row;
                const data = s.itemsSource.sourceCollection;
                let item = data[_id];
                e.cell.innerHTML = '<input class="row-checkbox" id="chk_' + _id + '" type="checkbox" ' + (item.status != 1 ? 'checked' : '') + '>';
                if (e.row == 0 && e.col == 0) {
                    e.cell.classList.add('wj-state-disabled');
                }
                $('#chk_' + _id).off('click').on('click', function (event) {
                    item.no = '';
                    item.status = event.target.checked ? 0 : 1;
                    let index = event.target.checked ? _id + 1 : _id;
                    for (var i = index; i < data.length; i++) {
                        data[i].no = event.target.checked ? i : i + 1;
                    }
                    s.refresh();
                });
            }
        });
    }
}
//
class ExcelExportContext {
    constructor(btn) {
        this._exporting = false;
        this._progress = 0;
        this._preparing = false;
        this._btn = btn;
    }
    get exporting() {
        return this._exporting;
    }
    set exporting(value) {
        if (value !== this._exporting) {
            this._exporting = value;
            this._onPropertyChanged();
        }
    }
    get progress() {
        return this._progress;
    }
    set progress(value) {
        if (value !== this._progress) {
            this._progress = value;
            this._onPropertyChanged();
        }
    }
    get preparing() {
        return this._preparing;
    }
    set preparing(value) {
        if (value !== this._preparing) {
            this._preparing = value;
            this._onPropertyChanged();
        }
    }
    _onPropertyChanged() {
        wjcCore.enable(this._btn, !this._preparing);
        if (this._exporting) {
            const percent = wjcCore.Globalize.formatNumber(this._progress, 'p0');
            this._btn.textContent = `Cancel (${percent} done)`;
        }
        else {
            this._btn.textContent = 'Export To Excel';
        }
    }
}
//
document.readyState === 'complete' ? init() : window.onload = init;
//
function init() {
    const dataSvc = new DataService();
    const exportSvc = new ExportService();
    const app = new App(dataSvc, exportSvc);
    // console.log(app);
    window.addEventListener('unload', () => {
        app.close();
    });
}

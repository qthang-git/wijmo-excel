import * as wjcCore from '@grapecity/wijmo';
//
export class DataService {
    constructor() {
        this._no = 1;
        this._user = '';
        this._operation = '';
        this._tag = '';
        this._checklist = '';
        this._date = '';
        this._verifier = '';
        this._result = '';
        this._notes = '';
        this.group = '';
        this._status = 1;
    }
    getData(count) {
        const data = [];
        const itemsCount = Math.max(count, 2);
        // add items
        for (let i = 0; i < itemsCount; i++) {
            const item = this._getItem(i);
            data.push(item);
        }
        return data;
    }
    _getItem(index) {
        const item = {
            no: index == 0 ? '' : index,
            user: this._user,
            operation: this._operation,
            tag: this._tag,
            checklist: this._checklist,
            date: this._date,
            verifier: this._verifier,
            result: this._result,
            notes: this._notes,
            group: '',
            status: index == 0 ? 0 : this._status
        }
        return item;
    }
}

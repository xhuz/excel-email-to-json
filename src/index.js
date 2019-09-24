const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

class XLSXEmailToJson {
    get result() {
        return this._result;
    }
    constructor(filename) {
        this._result = [];
        this._workBook = XLSX.readFile(path.resolve(__dirname, filename));
        this._Sheets = this._workBook.Sheets;
        this._SheetNames = this._workBook.SheetNames;
        this._toJson();
    }

    _extractEmail(obj) {
        const result = [];
        const reg = /([A-Za-z0-9_\-.]+@[A-Za-z0-9_\-]+\.[A-Za-z._\-]+)/g;
        for(const i in obj) {
            const match = obj[i].match(reg);
            result.push(...match);
        }
        return result;
    }

    _toJson() {
        this._result.length = 0;
        const sheetNames = this._SheetNames;
        for (let i = 0; i < sheetNames.length; i++) {
            const data = XLSX.utils.sheet_to_json(this._Sheets[sheetNames[i]]);
            for(const item of data) {
                const result = this._extractEmail(item);
                this._result.push(...result);
            }
        }
    }

    writeFile(filename, row = 0) {
        if(typeof window !== 'undefined') {
            throw new Error('method only use in node');
        }else {
            const t = filename.split(path.sep);
            const last = t.pop();
            const tmp = last.split('.');
            const ext = tmp.pop();
            const name = tmp.join('.');
            const dir = t.join(path.sep);
            mkdir(dir);
            const length = this._result.length;
            if(row === 0) {
                fs.writeFileSync(path.join(dir, `${name}.${ext}`), JSON.stringify(this._result), 'utf8');
            } else {
                for(let i = 0; i < Math.ceil(length / row); i++) {
                    const data = this._result.slice(i * row, (i + 1) * row);
                    fs.writeFileSync(path.join(dir, `${name}-${i + 1}.${ext}`), JSON.stringify(data), 'utf8');
                }
            }
        }
    }
}

function mkdir(dirname) {
    if (fs.existsSync(dirname)) {
        return true;
    } else {
        if (mkdir(path.dirname(dirname))) {
            fs.mkdirSync(dirname);
            return true;
        }
    } 
}

module.exports = XLSXEmailToJson;
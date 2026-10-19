/**
 * Code.gs をローカル（Node / ブラウザ）で動かすための最小限の GAS モック。
 * 本物と同様、セルの値は getDisplayValues で文字列として返す。
 */
(function (root) {
  'use strict';

  function MockSheet(name) {
    this.name = name;
    this.cells = [];        // 2 次元配列（行 → 列）
    this.formats = {};      // "row,col" -> numberFormat
    this.frozenRows = 0;
  }
  MockSheet.prototype.getName = function () { return this.name; };
  MockSheet.prototype.getLastRow = function () {
    for (var r = this.cells.length; r > 0; r--) {
      if (this.cells[r - 1].some(function (v) { return v !== ''; })) return r;
    }
    return 0;
  };
  MockSheet.prototype.getLastColumn = function () {
    var max = 0;
    this.cells.forEach(function (row) {
      for (var c = row.length; c > max; c--) {
        if (row[c - 1] !== '') { max = c; break; }
      }
    });
    return max;
  };
  MockSheet.prototype.getMaxRows = function () { return Math.max(1000, this.cells.length); };
  MockSheet.prototype.setFrozenRows = function (n) { this.frozenRows = n; };
  MockSheet.prototype.getDataRange = function () {
    return new MockRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()));
  };
  MockSheet.prototype.getRange = function (row, col, numRows, numCols) {
    return new MockRange(this, row, col, numRows || 1, numCols || 1);
  };
  MockSheet.prototype.cell = function (row, col) {
    var r = this.cells[row - 1];
    return r && r[col - 1] !== undefined ? r[col - 1] : '';
  };
  MockSheet.prototype.write = function (row, col, value) {
    while (this.cells.length < row) this.cells.push([]);
    var r = this.cells[row - 1];
    while (r.length < col) r.push('');
    r[col - 1] = value;
  };
  /** テスト用：シートの中身を直接セットする */
  MockSheet.prototype.load = function (rows) {
    var self = this;
    this.cells = [];
    rows.forEach(function (row, i) {
      row.forEach(function (v, j) { self.write(i + 1, j + 1, v); });
    });
  };

  function MockRange(sheet, row, col, numRows, numCols) {
    this.sheet = sheet; this.row = row; this.col = col; this.numRows = numRows; this.numCols = numCols;
  }
  MockRange.prototype.forEachCell = function (fn) {
    for (var i = 0; i < this.numRows; i++) {
      for (var j = 0; j < this.numCols; j++) fn(this.row + i, this.col + j, i, j);
    }
  };
  MockRange.prototype.getDisplayValues = function () {
    var out = [];
    for (var i = 0; i < this.numRows; i++) {
      var line = [];
      for (var j = 0; j < this.numCols; j++) line.push(String(this.sheet.cell(this.row + i, this.col + j)));
      out.push(line);
    }
    return out;
  };
  MockRange.prototype.getValues = MockRange.prototype.getDisplayValues;
  MockRange.prototype.setValues = function (values) {
    if (values.length !== this.numRows || values.some(function (r) { return r.length !== this.numCols; }, this)) {
      throw new Error('The number of rows or columns in the data does not match the range.');
    }
    var sheet = this.sheet;
    this.forEachCell(function (r, c, i, j) { sheet.write(r, c, values[i][j]); });
    return this;
  };
  MockRange.prototype.setValue = function (value) {
    var sheet = this.sheet;
    this.forEachCell(function (r, c) { sheet.write(r, c, value); });
    return this;
  };
  MockRange.prototype.setNumberFormat = function (format) {
    var formats = this.sheet.formats;
    this.forEachCell(function (r, c) { formats[r + ',' + c] = format; });
    return this;
  };

  function MockSpreadsheet() { this.sheets = {}; }
  MockSpreadsheet.prototype.getSheetByName = function (name) { return this.sheets[name] || null; };
  MockSpreadsheet.prototype.insertSheet = function (name) {
    this.sheets[name] = new MockSheet(name);
    return this.sheets[name];
  };

  /** GAS のグローバル（SpreadsheetApp など）一式を作る */
  function createGasGlobals() {
    var ss = new MockSpreadsheet();
    var lock = { held: false, fail: false, acquired: 0 };
    return {
      spreadsheet: ss,
      lock: lock,
      globals: {
        SpreadsheetApp: {
          getActiveSpreadsheet: function () { return ss; },
          openById: function () { return ss; },
          flush: function () {},
          getUi: function () { throw new Error('UI is not available in mock'); }
        },
        PropertiesService: {
          getScriptProperties: function () { return { getProperty: function () { return null; } }; }
        },
        LockService: {
          getScriptLock: function () {
            return {
              tryLock: function () {
                if (lock.fail || lock.held) return false;
                lock.held = true;
                lock.acquired++;
                return true;
              },
              releaseLock: function () { lock.held = false; }
            };
          }
        }
      }
    };
  }

  var api = { createGasGlobals: createGasGlobals, MockSheet: MockSheet };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.GasMock = api;
})(this);

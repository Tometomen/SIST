// ローカル検証用。Google の認証・シート・HTMLサービスの実動作は代替しない。
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const appDirectory = path.join(__dirname, '..', 'apps-script');

class Sheet {
  constructor(rows = []) { this.rows = structuredClone(rows); }
  getLastRow() { return this.rows.length; }
  getDataRange() {
    return { getDisplayValues: () => this.rows.map(row => row.map(value =>
      typeof value === 'boolean' ? String(value).toUpperCase() : String(value ?? ''))) };
  }
  getRange(row, column, height, width) {
    const sheet = this;
    return {
      setValues(values) {
        if (values.length !== height || values.some(value => value.length !== width)) throw new Error('Range dimensions');
        values.forEach((cells, i) => {
          sheet.rows[row - 1 + i] ||= [];
          cells.forEach((value, j) => { sheet.rows[row - 1 + i][column - 1 + j] = value; });
        });
        return this;
      },
      insertCheckboxes() { return this.setValues(Array.from({ length: height }, () => Array(width).fill(false))); },
      setFontWeight() { return this; },
      setBackground() { return this; }
    };
  }
  setFrozenRows() {}
  setColumnWidth() {}
}

function createRuntime({ seedSheets = true, initialized = true } = {}) {
  const sheets = new Map();
  const properties = new Map(initialized ? [['LINKS_SPREADSHEET_ID', 'local-sheet']] : []);
  const errors = [];
  const logs = [];
  const spreadsheet = {
    getId: () => 'local-sheet',
    getSheetByName: name => sheets.get(name),
    insertSheet(name) { const sheet = new Sheet(); sheets.set(name, sheet); return sheet; }
  };
  const context = vm.createContext({
    console: { error: error => errors.push(error), log: message => logs.push(message) },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => spreadsheet,
      openById(id) { if (id !== 'local-sheet') throw new Error('Unknown sheet'); return spreadsheet; },
      flush() {}
    },
    PropertiesService: { getScriptProperties: () => ({
      getProperty: name => properties.get(name),
      setProperty: (name, value) => properties.set(name, value)
    }) },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    HtmlService: {
      createHtmlOutput(html) {
        return { html, setTitle(title) { this.title = title; return this; }, addMetaTag() { return this; } };
      },
      createHtmlOutputFromFile(name) {
        return { getContent: () => fs.readFileSync(path.join(appDirectory, name + '.html'), 'utf8') };
      },
    }
  });
  for (const name of ['Code.gs']) {
    vm.runInContext(fs.readFileSync(path.join(appDirectory, name), 'utf8'), context, { filename: name });
  }
  const seed = JSON.parse(JSON.stringify(context.initialData_()));
  const headers = JSON.parse(vm.runInContext('JSON.stringify(LINK_HEADERS_)', context));
  if (seedSheets) {
    sheets.set('リンク', new Sheet([headers, ...seed.links]));
    sheets.set('設定', new Sheet([['項目', '値'], ...seed.settings]));
  }
  return { context, sheets, properties, errors, logs, seed, headers };
}

module.exports = { createRuntime };

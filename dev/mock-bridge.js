// Stand-in for Excel so the panel can be tried in an ordinary browser.
// Reads the same sample rows as Sample Sales.xlsx and reports what Excel
// would have been asked to do.
import { buildModel, rowsKeptBySpec } from '../js/engine.js';

export async function createMockBridge({ onFilter = () => {} } = {}) {
  const data = await (await fetch('./sample-data.json')).json();
  const model = buildModel(data);
  let settings = {};
  try {
    settings = JSON.parse(localStorage.getItem('excel-bridge-dev') || '{}');
  } catch {
    settings = {};
  }
  return {
    name: 'preview',
    theme: () => null,
    listTables: async () => [{ name: data.tableName, sheet: 'Sales' }],
    readTable: async () => ({ headers: data.headers, rows: data.rows, formats: data.formats, use1904: false }),
    applyFilters: async (table, spec, previous) => {
      onFilter({ action: 'apply', table, spec, previous, kept: rowsKeptBySpec(model, spec), total: model.n });
    },
    clearFilters: async (table, names) => {
      onFilter({ action: 'clear', table, names, kept: model.n, total: model.n });
    },
    makeTableFromSelection: async () => data.tableName,
    watchTable: () => () => {},
    loadSettings: () => settings,
    saveSettings: async (value) => {
      settings = value;
      try {
        localStorage.setItem('excel-bridge-dev', JSON.stringify(value));
      } catch {
        // Preview only; nothing to do if storage is blocked.
      }
    },
  };
}

/* global Excel, Office */
// The only file that talks to Excel. Everything else works on plain data, so
// the panel can also run in a browser preview with dev/mock-bridge.js.

const CHUNK = 5000;
const SETTINGS_KEY = 'excelBridge';

export function createExcelBridge() {
  return {
    name: 'excel',

    theme() {
      try {
        const colour = Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
        if (!colour) return null;
        const hex = colour.replace('#', '');
        const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
        return 0.2126 * r + 0.7152 * g + 0.0722 * b < 0.5 ? 'dark' : 'light';
      } catch {
        return null;
      }
    },

    async listTables() {
      return Excel.run(async (ctx) => {
        const tables = ctx.workbook.tables.load('items/name');
        await ctx.sync();
        const sheets = tables.items.map((t) => t.worksheet.load('name'));
        await ctx.sync();
        return tables.items.map((t, i) => ({ name: t.name, sheet: sheets[i].name }));
      });
    },

    async readTable(name) {
      return Excel.run(async (ctx) => {
        const table = ctx.workbook.tables.getItem(name);
        const header = table.getHeaderRowRange().load('values');
        const body = table.getDataBodyRange().load('rowCount,columnCount');
        const has1904 = Office.context.requirements.isSetSupported('ExcelApi', '1.16');
        if (has1904) ctx.workbook.load('use1904DateSystem');
        await ctx.sync();
        const first = body.getRow(0).load('numberFormat');
        await ctx.sync();
        const rows = [];
        for (let start = 0; start < body.rowCount; start += CHUNK) {
          const count = Math.min(CHUNK, body.rowCount - start);
          const part = body.getCell(start, 0).getResizedRange(count - 1, body.columnCount - 1).load('values');
          await ctx.sync();
          for (const row of part.values) rows.push(row);
        }
        return {
          headers: header.values[0],
          formats: first.numberFormat[0],
          rows,
          use1904: has1904 ? !!ctx.workbook.use1904DateSystem : false,
        };
      });
    },

    // spec: { [column]: { kind: 'values', values, blanks } | { kind: 'months', months: ['2025-04'] } }
    // previous: columns this add-in filtered last time; any not in spec are cleared.
    async applyFilters(tableName, spec, previous) {
      const withBlanks = [];
      await Excel.run(async (ctx) => {
        const table = ctx.workbook.tables.getItem(tableName);
        const names = [...new Set([...previous, ...Object.keys(spec)])];
        const columns = await columnsByName(ctx, table, names);
        for (const name of previous) {
          if (!spec[name] && columns[name]) columns[name].filter.clear();
        }
        for (const [name, f] of Object.entries(spec)) {
          const column = columns[name];
          if (!column) continue;
          if (f.kind === 'months') {
            const specificity = Excel.FilterDatetimeSpecificity ? Excel.FilterDatetimeSpecificity.month : 'Month';
            column.filter.applyValuesFilter(f.months.map((m) => ({ date: `${m}-01`, specificity })));
          } else if (f.blanks) {
            withBlanks.push([name, f.values]);
          } else {
            column.filter.applyValuesFilter(f.values);
          }
        }
        await ctx.sync();
      });
      // FilterCriteria has no "blanks" flag. An empty string in the values list
      // keeps blank cells; if this Excel build refuses it, keep the named values only.
      for (const [name, values] of withBlanks) {
        const run = (list) => Excel.run(async (ctx) => {
          ctx.workbook.tables.getItem(tableName).columns.getItem(name).filter.applyValuesFilter(list);
          await ctx.sync();
        });
        try {
          await run([...values, '']);
        } catch (err) {
          console.warn('Blank cells could not be included in the sheet filter', err);
          if (values.length) await run(values);
        }
      }
    },

    async clearFilters(tableName, names) {
      await Excel.run(async (ctx) => {
        const table = ctx.workbook.tables.getItemOrNullObject(tableName);
        await ctx.sync();
        if (table.isNullObject) return;
        const columns = await columnsByName(ctx, table, names);
        for (const column of Object.values(columns)) column.filter.clear();
        await ctx.sync();
      });
    },

    async makeTableFromSelection() {
      return Excel.run(async (ctx) => {
        let range = ctx.workbook.getSelectedRange().load('rowCount,columnCount');
        await ctx.sync();
        if (range.rowCount === 1 && range.columnCount === 1) range = range.getSurroundingRegion();
        const table = ctx.workbook.worksheets.getActiveWorksheet().tables.add(range, true);
        table.load('name');
        await ctx.sync();
        return table.name;
      });
    },

    // Calls onChange when the table's cells change. Returns a function that stops watching.
    watchTable(name, onChange) {
      let handle = null;
      let stopped = false;
      Excel.run(async (ctx) => {
        handle = ctx.workbook.tables.getItem(name).onChanged.add(async () => { onChange(); });
        await ctx.sync();
      }).catch((err) => console.warn('Could not watch table', err));
      return () => {
        if (stopped || !handle) return;
        stopped = true;
        Excel.run(handle.context, async (ctx) => {
          handle.remove();
          await ctx.sync();
        }).catch(() => {});
      };
    },

    loadSettings() {
      return Office.context.document.settings.get(SETTINGS_KEY) || {};
    },

    saveSettings(value) {
      const settings = Office.context.document.settings;
      settings.set(SETTINGS_KEY, value);
      return new Promise((resolve) => settings.saveAsync(() => resolve()));
    },
  };
}

async function columnsByName(ctx, table, names) {
  const found = names.map((n) => [n, table.columns.getItemOrNullObject(n)]);
  await ctx.sync();
  const out = {};
  for (const [n, column] of found) if (!column.isNullObject) out[n] = column;
  return out;
}

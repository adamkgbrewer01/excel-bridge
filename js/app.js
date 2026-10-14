// Excel Bridge task pane. Holds the dashboard state, draws it, and keeps the
// sheet's table filters in step when "Filter the sheet too" is on.
// `bridge` is either the real Excel connection (excel-bridge.js) or the
// browser preview's stand-in (dev/mock-bridge.js).
import * as E from './engine.js';
import { el, barList, columnChart, dataTable, barScale, columnScale, hideTip } from './charts.js';

const DOUBLE_CLICK_MS = 400;
const isMac = /Mac|iPhone|iPad/.test(navigator.platform || navigator.userAgent);

export async function start(bridge, root = document.getElementById('app')) {
  const state = {
    tables: [],
    tableName: null,
    data: null,
    model: null,
    dash: null,
    error: null,
    applied: new Set(),
    settings: {},
    unwatch: null,
  };
  let lastClick = null;
  let syncTimer = null;
  let saveTimer = null;
  let reloadTimer = null;
  let syncChain = Promise.resolve();

  applyTheme();
  try {
    state.settings = (await bridge.loadSettings()) || {};
    await refreshTables();
    const preferred = state.tables.find((t) => t.name === state.settings.tableName) || state.tables[0];
    if (preferred) {
      await openTable(preferred.name);
      // Filters left on the sheet by an earlier session no longer match the
      // panel, which always starts unfiltered, so take them off.
      const stale = state.settings.applied?.[preferred.name];
      if (stale && stale.length) {
        await bridge.clearFilters(preferred.name, stale);
        rememberApplied();
      }
    }
  } catch (err) {
    fail(err);
  }
  render();

  // ---------- Loading ----------

  function applyTheme() {
    const theme = bridge.theme && bridge.theme();
    if (theme) document.documentElement.dataset.theme = theme;
  }

  async function refreshTables() {
    state.tables = await bridge.listTables();
  }

  async function openTable(name, keepState = false) {
    const previous = keepState ? state.dash : null;
    const data = await bridge.readTable(name);
    const model = E.buildModel(data, { use1904: !!data.use1904 });
    let dash = previous ? E.restoreDashboard(model, E.serialiseDashboard(previous)) : null;
    if (dash && previous) {
      dash.charts.forEach((c, i) => {
        const old = previous.charts[i];
        if (old && old.kind === c.kind && E.chartFields(old).join() === E.chartFields(c).join()) {
          c.drill = old.drill;
          c.selection = old.selection;
          c.showAll = old.showAll;
        }
      });
    }
    if (!dash) dash = E.restoreDashboard(model, state.settings.dashboards?.[name]) || E.defaultDashboard(model);
    if (state.tableName !== name) {
      if (state.unwatch) state.unwatch();
      state.unwatch = bridge.watchTable(name, () => {
        clearTimeout(reloadTimer);
        reloadTimer = setTimeout(() => reload(), 600);
      });
      state.applied = new Set();
    }
    Object.assign(state, { tableName: name, data, model, dash, error: null });
  }

  async function reload() {
    root.classList.add('loading');
    try {
      await refreshTables();
      if (state.tables.some((t) => t.name === state.tableName)) await openTable(state.tableName, true);
      else if (state.tables[0]) await openTable(state.tables[0].name);
      else Object.assign(state, { tableName: null, model: null, dash: null });
    } catch (err) {
      fail(err);
    }
    root.classList.remove('loading');
    render();
    scheduleSync();
  }

  function fail(err) {
    console.error(err);
    state.error = err && err.message ? err.message : String(err);
  }

  // ---------- Changes ----------

  function changed({ sync = true } = {}) {
    render();
    scheduleSave();
    if (sync) scheduleSync();
  }

  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      if (!state.dash) return;
      const dashboards = { ...(state.settings.dashboards || {}), [state.tableName]: E.serialiseDashboard(state.dash) };
      state.settings = { ...state.settings, tableName: state.tableName, dashboards };
      Promise.resolve(bridge.saveSettings(state.settings)).catch((err) => console.warn('Could not save settings', err));
    }, 500);
  }

  function scheduleSync() {
    clearTimeout(syncTimer);
    syncTimer = setTimeout(() => {
      syncChain = syncChain.then(syncSheet).catch((err) => { fail(err); render(); });
    }, 250);
  }

  async function syncSheet() {
    if (!state.model) return;
    const table = state.tableName;
    if (!state.dash.filterSheet) {
      if (state.applied.size) {
        await bridge.clearFilters(table, [...state.applied]);
        state.applied = new Set();
        rememberApplied();
      }
      return;
    }
    const spec = E.sheetFilterSpec(state.model, state.dash);
    await bridge.applyFilters(table, spec, [...state.applied]);
    state.applied = new Set(Object.keys(spec));
    rememberApplied();
  }

  function rememberApplied() {
    state.settings = { ...state.settings, applied: { ...(state.settings.applied || {}), [state.tableName]: [...state.applied] } };
    scheduleSave();
  }

  function pick(chart, result, key, evt) {
    const now = Date.now();
    const repeat = lastClick && lastClick.chart === chart.id && lastClick.key === key && now - lastClick.time < DOUBLE_CLICK_MS;
    lastClick = { chart: chart.id, key, time: now };
    if (repeat && result.canDrill && !(evt.metaKey || evt.ctrlKey || evt.shiftKey)) {
      lastClick = null;
      drill(chart, result, key);
      return;
    }
    const sel = chart.selection;
    if (evt.metaKey || evt.ctrlKey || evt.shiftKey) {
      if (sel.has(key)) sel.delete(key);
      else sel.add(key);
    } else if (sel.size === 1 && sel.has(key)) {
      sel.clear();
    } else {
      sel.clear();
      sel.add(key);
    }
    changed();
  }

  function drill(chart, result, key) {
    if (!result.canDrill) return;
    const item = result.items.find((i) => i.key === key);
    chart.drill = [...chart.drill.slice(0, result.depth), { key, label: item ? item.label : key }];
    chart.selection = new Set();
    chart.showAll = false;
    changed();
  }

  function goUp(chart, depth) {
    chart.drill = chart.drill.slice(0, depth);
    chart.selection = new Set();
    changed();
  }

  function clearAll() {
    for (const c of state.dash.charts) {
      c.drill = [];
      c.selection = new Set();
    }
    changed();
  }

  // ---------- Drawing ----------

  function render() {
    hideTip();
    const focusId = document.activeElement && document.activeElement.dataset ? document.activeElement.dataset.fid : null;
    root.replaceChildren();
    if (state.error) {
      const box = el('div', 'notice error');
      box.append(el('h2', '', 'Something went wrong'), el('p', '', state.error));
      const again = el('button', 'btn', 'Try again');
      again.addEventListener('click', () => { state.error = null; reload(); });
      box.append(again);
      root.append(box);
    }
    if (!state.tables.length) {
      root.append(noTables());
    } else if (state.model) {
      root.append(setup(), headline(), filterRow(), cards(), addChart(), about());
    }
    if (focusId) {
      const target = root.querySelector(`[data-fid="${CSS.escape(focusId)}"]`);
      if (target) target.focus({ preventScroll: true });
    }
  }

  function noTables() {
    const box = el('div', 'notice');
    box.append(
      el('h2', '', 'No tables in this workbook yet'),
      el('p', '', 'Excel Bridge works from Excel tables. Select your data, including the header row, then make it a table.'),
    );
    const row = el('div', 'row');
    const make = el('button', 'btn primary', 'Make a table from my selection');
    make.addEventListener('click', async () => {
      try {
        const name = await bridge.makeTableFromSelection();
        await refreshTables();
        await openTable(name);
      } catch (err) {
        fail(err);
      }
      render();
    });
    const again = el('button', 'btn', 'Check again');
    again.addEventListener('click', reload);
    row.append(make, again);
    box.append(row);
    return box;
  }

  function select(options, value, onChange, label) {
    const s = el('select');
    for (const o of options) {
      const opt = el('option', '', o.label);
      opt.value = o.value;
      if (o.value === value) opt.selected = true;
      s.append(opt);
    }
    if (label) s.setAttribute('aria-label', label);
    s.addEventListener('change', () => onChange(s.value));
    return s;
  }

  function labelled(text, control, cls = '') {
    const l = el('label', cls);
    l.append(text, control);
    return l;
  }

  function setup() {
    const { model, dash } = state;
    const wrap = el('div');
    const grid = el('div', 'setup');
    const tableSelect = select(
      state.tables.map((t) => ({ value: t.name, label: state.tables.length > 1 ? `${t.name} (${t.sheet})` : t.name })),
      state.tableName,
      async (name) => {
        root.classList.add('loading');
        try {
          if (state.applied.size) {
            await bridge.clearFilters(state.tableName, [...state.applied]);
            state.applied = new Set();
            rememberApplied();
          }
          await openTable(name);
        } catch (err) {
          fail(err);
        }
        root.classList.remove('loading');
        changed();
      },
    );
    grid.append(labelled('Table', tableSelect, 'wide'));
    const measures = model.fields.filter((f) => f.kind === 'measure');
    const measureSelect = select(
      measures.length ? measures.map((f) => ({ value: f.name, label: f.name })) : [{ value: '', label: 'No number columns' }],
      dash.measure,
      (v) => { dash.measure = v; if (dash.agg === 'count') dash.agg = 'sum'; changed({ sync: false }); },
    );
    measureSelect.disabled = !measures.length || dash.agg === 'count';
    const aggOptions = [{ value: 'sum', label: 'Total' }, { value: 'avg', label: 'Average' }, { value: 'count', label: 'Number of rows' }];
    const aggSelect = select(measures.length ? aggOptions : aggOptions.slice(2), dash.agg, (v) => {
      dash.agg = v;
      if (v !== 'count' && !dash.measure) dash.measure = measures[0]?.name ?? null;
      changed({ sync: false });
    });
    grid.append(labelled('Measure', measureSelect), labelled('Show as', aggSelect));
    wrap.append(grid);

    const row = el('div', 'row');
    row.style.marginTop = '8px';
    const toggle = el('label', 'toggle');
    const box = el('input');
    box.type = 'checkbox';
    box.checked = dash.filterSheet;
    box.addEventListener('change', () => { dash.filterSheet = box.checked; changed(); });
    toggle.append(box, 'Filter the sheet too');
    toggle.title = 'Apply the same filters to the table on the sheet, so its totals and charts follow your clicks';
    const refresh = el('button', 'btn', 'Refresh');
    refresh.addEventListener('click', reload);
    row.append(toggle, el('span', 'spacer'), refresh);
    wrap.append(row);
    return wrap;
  }

  function measureLabel() {
    const { dash } = state;
    if (dash.agg === 'count') return 'Rows';
    return dash.agg === 'avg' ? `Average ${dash.measure}` : dash.measure;
  }

  function measureInfo() {
    const { dash, model } = state;
    return dash.agg === 'count' ? {} : model.byName.get(dash.measure) || {};
  }

  function headline() {
    const { kpi } = state.result = E.evaluate(state.model, state.dash);
    const info = measureInfo();
    const card = el('div', 'kpi');
    card.append(el('div', 'label', measureLabel()));
    const value = el('div', 'value', E.formatValue(kpi.value, info));
    value.title = E.formatValue(kpi.value, info, 'full');
    card.append(value);
    const rows = `${kpi.rows.toLocaleString('en-GB')} of ${kpi.totalRows.toLocaleString('en-GB')} rows`;
    if (!kpi.filtered) {
      card.append(el('div', 'sub', `All ${kpi.totalRows.toLocaleString('en-GB')} rows`));
    } else if (state.dash.agg === 'avg') {
      card.append(el('div', 'sub', `All rows: ${E.formatValue(kpi.total, info)} · ${rows}`));
    } else {
      const share = kpi.total ? kpi.value / kpi.total : 0;
      card.append(el('div', 'sub', `${Math.round(share * 100)}% of ${E.formatValue(kpi.total, info)} · ${rows}`));
      if (share >= 0 && share <= 1) {
        const meter = el('div', 'meter');
        const fill = el('div');
        fill.style.width = `${share * 100}%`;
        meter.append(fill);
        meter.setAttribute('aria-hidden', 'true');
        card.append(meter);
      }
    }
    return card;
  }

  function filterRow() {
    const chips = E.activeFilters(state.model, state.dash);
    const row = el('div', 'filters');
    row.setAttribute('aria-label', 'Active filters');
    if (!chips.length) {
      const key = isMac ? 'Cmd' : 'Ctrl';
      row.append(el('p', 'hint', `Click a bar to filter everything else. Double-click to drill down. ${key}-click to pick more than one.`));
      return row;
    }
    for (const chip of chips) {
      const c = el('span', 'chip');
      c.append(el('span', 'field', `${chip.field}`), el('span', 'text', chip.text));
      const x = el('button', '', '×');
      x.setAttribute('aria-label', `Remove filter ${chip.field} ${chip.text}`);
      x.dataset.fid = `chip:${chip.chartId}:${chip.type}:${chip.depth ?? ''}`;
      x.addEventListener('click', () => {
        const chart = state.dash.charts.find((ch) => ch.id === chip.chartId);
        if (chip.type === 'drill') goUp(chart, chip.depth);
        else { chart.selection = new Set(); changed(); }
      });
      c.append(x);
      row.append(c);
    }
    const clear = el('button', 'btn link', 'Clear all');
    clear.addEventListener('click', clearAll);
    row.append(clear);
    return row;
  }

  function cards() {
    const grid = el('div', 'cards');
    state.result.charts.forEach((result) => grid.append(card(result)));
    return grid;
  }

  function card(result) {
    const { chart } = result;
    const info = measureInfo();
    const c = el('section', `card${chart.splitBy ? ' wide' : ''}`);
    const title = `${measureLabel()} by ${result.level}`;
    c.setAttribute('aria-label', title);
    const head = el('div', 'head');
    const remove = el('button', 'icon', '×');
    remove.setAttribute('aria-label', `Remove chart ${title}`);
    remove.title = 'Remove this chart';
    remove.addEventListener('click', () => {
      state.dash.charts = state.dash.charts.filter((x) => x !== chart);
      changed();
    });
    const other = chart.view === 'table' ? 'chart' : 'table';
    const viewToggle = el('button', 'btn link', other === 'table' ? 'Table' : 'Chart');
    viewToggle.title = other === 'table' ? 'Show these figures as a table' : 'Show the chart again';
    viewToggle.dataset.fid = `view:${chart.id}`;
    viewToggle.addEventListener('click', () => { chart.view = other; changed({ sync: false }); });
    head.append(el('h2', '', title), viewToggle, remove);
    c.append(head, crumbs(result));

    const tools = el('div', 'tools');
    // Always present when drilling is possible, so selecting a bar never moves
    // the chart between the two clicks of a double-click.
    if (result.canDrill) {
      const key = chart.selection.size === 1 ? [...chart.selection][0] : null;
      const item = key != null && result.items.find((i) => i.key === key);
      const b = el('button', 'btn', 'Drill down ↓');
      b.dataset.fid = `drill:${chart.id}`;
      b.disabled = key == null;
      b.title = key == null ? 'Click one bar first, then drill down into it' : `Drill into ${item ? item.label : key}`;
      b.addEventListener('click', () => drill(chart, result, key));
      tools.append(b);
    }
    if (result.depth > 0) {
      const up = el('button', 'btn', 'Up a level ↑');
      up.dataset.fid = `up:${chart.id}`;
      up.addEventListener('click', () => goUp(chart, result.depth - 1));
      tools.append(up);
    }
    const splits = E.splitOptions(state.model, chart);
    if (splits.length) {
      const s = select(
        [{ value: '', label: 'No split' }, ...splits.map((f) => ({ value: f, label: `Split by ${f}` }))],
        chart.splitBy || '',
        (v) => { chart.splitBy = v || null; changed({ sync: false }); },
        'Small multiples',
      );
      s.dataset.fid = `split:${chart.id}`;
      s.title = 'Repeat this chart once for each value of another field';
      tools.append(s);
    }
    if (tools.childNodes.length) c.append(tools);

    const opts = {
      info,
      agg: state.dash.agg,
      fid: `mark:${chart.id}`,
      onPick: (key, evt) => pick(chart, result, key, evt),
      help: () => (result.canDrill ? 'Click to filter. Double-click to drill down.' : 'Click to filter.'),
    };
    const body = el('div', chart.selection.size ? 'has-selection' : '');
    if (chart.view === 'table') body.append(tableView(result, opts));
    else if (result.panels) body.append(multiples(result, opts));
    else if (chart.kind === 'date') body.append(columnChart(result.items, opts));
    else body.append(barList(result.items, opts));
    c.append(body);

    const foot = el('div', 'foot');
    if (result.hidden) {
      foot.append(`Showing the top 10 of ${result.items.length + result.hidden}.`);
      const more = el('button', 'btn link', 'Show all');
      more.addEventListener('click', () => { chart.showAll = true; changed({ sync: false }); });
      foot.append(more);
    } else if (chart.showAll && result.items.length > 10) {
      const less = el('button', 'btn link', 'Show top 10');
      less.addEventListener('click', () => { chart.showAll = false; changed({ sync: false }); });
      foot.append(less);
    }
    if (result.panels) {
      foot.append(`One panel per ${chart.splitBy}, all on the same scale.`);
      if (result.hiddenPanels) foot.append(` ${result.hiddenPanels} smaller ones not shown.`);
    }
    if (foot.childNodes.length) c.append(foot);
    return c;
  }

  function crumbs(result) {
    const { chart, levels, depth } = result;
    const wrap = el('div', 'crumbs');
    if (levels.length < 2) return wrap;
    const sep = () => el('span', 'sep', '›');
    if (!depth) {
      levels.forEach((level, i) => {
        if (i) wrap.append(sep());
        wrap.append(el('span', i === 0 ? 'here' : '', capitalise(level)));
      });
      wrap.title = 'Double-click a bar to drill down a level';
      return wrap;
    }
    const all = el('button', '', chart.kind === 'date' ? 'All years' : 'All');
    all.dataset.fid = `crumb:${chart.id}:0`;
    all.addEventListener('click', () => goUp(chart, 0));
    wrap.append(all);
    chart.drill.slice(0, depth).forEach((d, i) => {
      wrap.append(sep());
      if (i === depth - 1) {
        wrap.append(el('span', 'here', d.label));
      } else {
        const b = el('button', '', d.label);
        b.dataset.fid = `crumb:${chart.id}:${i + 1}`;
        b.addEventListener('click', () => goUp(chart, i + 1));
        wrap.append(b);
      }
    });
    return wrap;
  }

  function multiples(result, opts) {
    const grid = el('div', 'multiples');
    const all = result.panels.flatMap((p) => p.items);
    const scale = result.chart.kind === 'date' ? columnScale(all) : barScale(all);
    for (const panel of result.panels) {
      const p = el('div', 'panel compact');
      p.append(el('h3', '', panel.key));
      const panelOpts = {
        ...opts,
        scale,
        compact: true,
        fid: `${opts.fid}:${panel.key}`,
        titleFor: (item) => `${panel.key} · ${item.label}`,
      };
      p.append(result.chart.kind === 'date' ? columnChart(panel.items, panelOpts) : barList(panel.items, panelOpts));
      grid.append(p);
    }
    return grid;
  }

  function tableView(result, opts) {
    const filtered = result.items.some((i) => i.fg !== i.bg);
    const head = capitalise(result.level);
    if (result.panels) {
      const rows = result.panels.flatMap((panel) => panel.items.map((i) => ({
        key: i.key, selected: i.selected,
        cells: [panel.key, i.label, E.formatValue(i.fg, opts.info, 'full')],
      })));
      return dataTable([{ label: result.chart.splitBy }, { label: head }, { label: measureLabel(), num: true }], rows, opts);
    }
    const columns = [{ label: head }, { label: measureLabel(), num: true }];
    if (filtered) columns.push({ label: 'Before filters', num: true });
    const rows = result.items.map((i) => ({
      key: i.key,
      selected: i.selected,
      cells: [i.label, E.formatValue(i.fg, opts.info, 'full'), ...(filtered ? [E.formatValue(i.bg, opts.info, 'full')] : [])],
    }));
    return dataTable(columns, rows, opts);
  }

  function addChart() {
    const options = E.availableCharts(state.model, state.dash);
    const wrap = el('div', 'add');
    if (!options.length) return wrap;
    const s = select(options.map((o, i) => ({ value: String(i), label: o.label })), '0', () => {}, 'Field for the new chart');
    const b = el('button', 'btn', 'Add chart');
    b.addEventListener('click', () => {
      const o = options[+s.value];
      state.dash.charts.push(E.makeChart(o.spec));
      changed({ sync: false });
    });
    wrap.append(s, b);
    return wrap;
  }

  function about() {
    const { model } = state;
    const skipped = model.fields.filter((f) => f.kind === 'id').map((f) => f.name);
    let text = `${model.n.toLocaleString('en-GB')} rows from ${state.tableName}.`;
    if (skipped.length) text += ` Not charted: ${skipped.join(', ')} (a different value on almost every row).`;
    return el('p', 'hint', text);
  }
}

function capitalise(s) {
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : s;
}

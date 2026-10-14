// Excel Bridge engine: pure functions, no DOM and no Office.js.
// Turns a table (headers, rows, number formats) into a column model, detects
// date fields, measures and drill hierarchies, and evaluates a dashboard with
// cross-filtering, drill-down and small multiples.

export const BLANK = '(Blank)';
export const DATE_LEVELS = ['year', 'quarter', 'month'];
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
const DAY_MS = 86400000;
const ID_NAME = /(^|[\s_-])(id|code|key|ref|no\.?|number|#)$/i;
const YEAR_NAME = /\b(year|yr|fy)\b/i;
const MEASURE_HINT = /revenue|sales|amount|value|total|income|turnover|spend|cost/i;

// ---------- Dates ----------

export function serialToMonthIndex(serial, use1904 = false) {
  const d = new Date(Math.round(((use1904 ? serial + 1462 : serial) - 25569) * DAY_MS));
  return d.getUTCFullYear() * 12 + d.getUTCMonth();
}

export function textToMonthIndex(text) {
  const s = String(text).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})(?:-(\d{1,2}))?(?:[T ].*)?$/);
  if (m) return valid(+m[1], +m[2]);
  m = s.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/); // British day/month/year
  if (m) return +m[1] <= 31 ? valid(+m[3], +m[2]) : null;
  return null;
  function valid(y, mo) {
    return y > 1800 && y < 2300 && mo >= 1 && mo <= 12 ? y * 12 + mo - 1 : null;
  }
}

export function isDateFormat(fmt) {
  if (!fmt || typeof fmt !== 'string') return false;
  const f = fmt.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '');
  if (/^general$/i.test(f.trim())) return false;
  return /[dy]/i.test(f);
}

export function monthIndexToIso(mi) {
  return `${Math.floor(mi / 12)}-${String((mi % 12) + 1).padStart(2, '0')}`;
}

export function dateKey(level, mi) {
  if (mi == null || mi < 0) return BLANK;
  const y = Math.floor(mi / 12);
  const m = mi % 12;
  if (level === 'year') return String(y);
  if (level === 'quarter') return `${y}-Q${Math.floor(m / 3) + 1}`;
  return monthIndexToIso(mi);
}

export function dateLabel(level, key, insideYear) {
  if (key === BLANK) return BLANK;
  if (level === 'year') return key;
  const [y, rest] = key.split('-');
  if (level === 'quarter') return insideYear ? rest : `${rest} ${y}`;
  const name = MONTHS[+rest - 1];
  return insideYear ? name : `${name} ${y}`;
}

// ---------- Number formats ----------

export function formatInfo(fmt) {
  const info = { currency: '', percent: false };
  if (!fmt || typeof fmt !== 'string') return info;
  const bracket = fmt.match(/\[\$([^\]-]+)/);
  const plain = fmt.replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1');
  const symbol = plain.match(/[£$€¥₹]/);
  info.currency = bracket ? bracket[1] : symbol ? symbol[0] : '';
  info.percent = /%/.test(plain);
  return info;
}

const full = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 2 });
const one = new Intl.NumberFormat('en-GB', { maximumFractionDigits: 1 });

function compact(abs) {
  for (const [size, suffix] of [[1e12, 'T'], [1e9, 'B'], [1e6, 'M'], [1e3, 'K']]) {
    if (abs >= size * 0.9995) return one.format(abs / size) + suffix;
  }
  return full.format(abs);
}

export function formatValue(v, info = {}, mode = 'compact') {
  if (v == null || !Number.isFinite(v)) return '–';
  if (info.percent) return `${(v * 100).toFixed(1)}%`;
  const sign = v < 0 ? '-' : '';
  const abs = Math.abs(v);
  const body = mode === 'full' ? full.format(abs >= 1000 ? Math.round(abs) : abs) : abs < 1000 ? full.format(+abs.toFixed(abs < 10 ? 2 : 0)) : compact(abs);
  return `${sign}${info.currency || ''}${body}`;
}

export function niceMax(v) {
  if (!(v > 0)) return 1;
  const exp = Math.pow(10, Math.floor(Math.log10(v)));
  for (const step of [1, 1.2, 1.5, 2, 2.5, 3, 4, 5, 6, 8, 10]) {
    if (step * exp >= v) return step * exp;
  }
  return 10 * exp;
}

// ---------- Model ----------

export function buildModel({ headers, rows, formats = [] }, { use1904 = false } = {}) {
  const n = rows.length;
  const fields = headers.map((rawName, j) => {
    const name = String(rawName ?? `Column ${j + 1}`).trim() || `Column ${j + 1}`;
    const values = rows.map((r) => r[j]);
    const filled = values.filter((v) => v !== '' && v != null);
    const numeric = filled.filter((v) => typeof v === 'number').length;
    const fmt = formats[j];
    const base = { name, index: j, fmt, ...formatInfo(fmt) };
    if (filled.length && numeric >= 0.9 * filled.length) {
      if (isDateFormat(fmt)) {
        return { ...base, kind: 'date', source: 'serial',
          col: values.map((v) => (typeof v === 'number' ? serialToMonthIndex(v, use1904) : -1)) };
      }
      if (!(ID_NAME.test(name) || YEAR_NAME.test(name))) {
        return { ...base, kind: 'measure', col: values.map((v) => (typeof v === 'number' ? v : null)) };
      }
    } else if (filled.length) {
      const asDates = filled.map((v) => (typeof v === 'string' ? textToMonthIndex(v) : null));
      if (asDates.filter((v) => v != null).length >= 0.9 * filled.length) {
        return { ...base, kind: 'date', source: 'text', raw: values.map((v) => (v == null ? '' : String(v))),
          col: values.map((v) => (v === '' || v == null ? -1 : textToMonthIndex(String(v)) ?? -1)) };
      }
    }
    const col = values.map((v) => (v === '' || v == null ? BLANK : String(v)));
    const distinct = new Set(col).size;
    const isId = distinct > 100 || (n >= 50 && distinct > 0.5 * n);
    return { ...base, kind: isId ? 'id' : 'dimension', col, distinct };
  });
  const model = { n, fields, byName: new Map(fields.map((f) => [f.name, f])) };
  model.hierarchies = detectHierarchies(model);
  return model;
}

// A field B drills down from A when every B value sits under exactly one A
// value and B has more values than A (Region > Country, Category > Product).
export function detectHierarchies(model) {
  const dims = model.fields.filter((f) => f.kind === 'dimension' && f.distinct >= 2);
  const parentOf = new Map();
  for (const child of dims) {
    let best = null;
    for (const parent of dims) {
      if (parent === child || parent.distinct >= child.distinct) continue;
      if (!determines(child, parent)) continue;
      if (!best || parent.distinct > best.distinct) best = parent;
    }
    if (best) parentOf.set(child.name, best.name);
  }
  const children = new Map();
  for (const [c, p] of parentOf) {
    if (!children.has(p)) children.set(p, []);
    children.get(p).push(model.byName.get(c));
  }
  const used = new Set();
  const chains = [];
  const roots = dims.filter((d) => !parentOf.has(d.name));
  const queue = [...roots];
  while (queue.length) {
    const root = queue.shift();
    if (used.has(root.name)) continue;
    const chain = [root.name];
    used.add(root.name);
    let current = root.name;
    for (;;) {
      const kids = (children.get(current) || []).filter((k) => !used.has(k.name)).sort((a, b) => a.distinct - b.distinct);
      if (!kids.length) break;
      chain.push(kids[0].name);
      used.add(kids[0].name);
      queue.push(...kids.slice(1));
      current = kids[0].name;
    }
    chains.push(chain);
  }
  for (const d of dims) if (!used.has(d.name)) chains.push([d.name]);
  return chains;

  function determines(child, parent) {
    const seen = new Map();
    for (let r = 0; r < model.n; r++) {
      const c = child.col[r];
      const p = parent.col[r];
      const prev = seen.get(c);
      if (prev === undefined) seen.set(c, p);
      else if (prev !== p) return false;
    }
    return true;
  }
}

export function dateStartLevel(model, fieldName) {
  const col = model.byName.get(fieldName).col;
  for (const level of DATE_LEVELS) {
    const keys = new Set();
    for (const mi of col) if (mi >= 0) keys.add(dateKey(level, mi));
    if (keys.size >= 2) return level;
  }
  return 'month';
}

// ---------- Dashboard ----------

let nextId = 1;
export function makeChart(spec) {
  return { id: `c${nextId++}`, splitBy: null, view: 'chart', ...spec, drill: [], selection: new Set(), showAll: false };
}

export function chartLevels(model, chart) {
  if (chart.kind === 'date') {
    const start = DATE_LEVELS.indexOf(dateStartLevel(model, chart.field));
    return DATE_LEVELS.slice(start);
  }
  return chart.levels;
}

export function chartFields(chart) {
  return chart.kind === 'date' ? [chart.field] : chart.levels;
}

export function defaultMeasure(model) {
  const measures = model.fields.filter((f) => f.kind === 'measure');
  const hinted = measures.find((f) => MEASURE_HINT.test(f.name));
  const m = hinted || measures[0];
  return m ? { measure: m.name, agg: m.percent ? 'avg' : 'sum' } : { measure: null, agg: 'count' };
}

export function defaultDashboard(model) {
  const charts = [];
  const date = model.fields.find((f) => f.kind === 'date');
  if (date) charts.push(makeChart({ kind: 'date', field: date.name }));
  const chains = [...model.hierarchies].sort((a, b) =>
    model.byName.get(a[0]).distinct - model.byName.get(b[0]).distinct);
  for (const chain of chains.slice(0, 3)) charts.push(makeChart({ kind: 'dim', levels: chain }));
  return { ...defaultMeasure(model), charts, filterSheet: false };
}

// Rebuild a saved dashboard against the current model, dropping anything stale.
export function restoreDashboard(model, saved) {
  if (!saved || !Array.isArray(saved.charts)) return null;
  const has = (name, kinds) => model.byName.has(name) && kinds.includes(model.byName.get(name).kind);
  const dash = defaultDashboard(model);
  if (saved.agg === 'count' || has(saved.measure, ['measure'])) {
    dash.measure = saved.agg === 'count' ? dash.measure : saved.measure;
    dash.agg = ['sum', 'avg', 'count'].includes(saved.agg) ? saved.agg : dash.agg;
  }
  dash.filterSheet = !!saved.filterSheet;
  const charts = [];
  for (const c of saved.charts) {
    if (c.kind === 'date' && has(c.field, ['date'])) charts.push(makeChart({ kind: 'date', field: c.field }));
    else if (c.kind === 'dim' && Array.isArray(c.levels) && c.levels.length && c.levels.every((l) => has(l, ['dimension']))) {
      charts.push(makeChart({ kind: 'dim', levels: c.levels }));
    } else continue;
    const chart = charts[charts.length - 1];
    if (c.splitBy && has(c.splitBy, ['dimension']) && !chartFields(chart).includes(c.splitBy)) chart.splitBy = c.splitBy;
    if (c.view === 'table') chart.view = 'table';
  }
  dash.charts = charts;
  return dash;
}

export function serialiseDashboard(dash) {
  return {
    measure: dash.measure,
    agg: dash.agg,
    filterSheet: dash.filterSheet,
    charts: dash.charts.map((c) => ({ kind: c.kind, field: c.field, levels: c.levels, splitBy: c.splitBy, view: c.view })),
  };
}

// Fields that can still get their own chart.
export function availableCharts(model, dash) {
  const used = new Set(dash.charts.flatMap(chartFields));
  const out = [];
  for (const f of model.fields) {
    if (f.kind === 'date' && !used.has(f.name)) out.push({ label: f.name, spec: { kind: 'date', field: f.name } });
  }
  for (const chain of model.hierarchies) {
    const free = chain.filter((name) => !used.has(name));
    if (free.length) out.push({ label: free.join(' › '), spec: { kind: 'dim', levels: free } });
  }
  return out;
}

export function splitOptions(model, chart) {
  const own = new Set(chartFields(chart));
  return model.fields.filter((f) => f.kind === 'dimension' && f.distinct >= 2 && !own.has(f.name)).map((f) => f.name);
}

// ---------- Evaluation ----------

function keyArray(model, chart, level) {
  const field = model.byName.get(chart.kind === 'date' ? chart.field : level);
  if (chart.kind === 'date') return field.col.map((mi) => dateKey(level, mi));
  return field.col;
}

function newAcc() {
  return { bs: 0, bn: 0, bc: 0, fs: 0, fn: 0, fc: 0 };
}

function aggValue(sum, numeric, rows, agg) {
  if (agg === 'count') return rows;
  if (agg === 'avg') return numeric ? sum / numeric : null;
  return numeric ? sum : rows ? 0 : null;
}

export function evaluate(model, dash) {
  const n = model.n;
  const measure = dash.agg === 'count' ? null : model.byName.get(dash.measure);
  const mcol = measure ? measure.col : null;

  const prepared = dash.charts.map((chart) => {
    const levels = chartLevels(model, chart);
    const depth = Math.min(chart.drill.length, levels.length - 1);
    const level = levels[depth];
    return { chart, levels, depth, level, keys: keyArray(model, chart, level) };
  });

  // Drill paths are hard filters on every chart.
  const context = new Uint8Array(n).fill(1);
  for (const p of prepared) {
    for (let i = 0; i < p.depth; i++) {
      const keys = keyArray(model, p.chart, p.levels[i]);
      const want = p.chart.drill[i].key;
      for (let r = 0; r < n; r++) if (keys[r] !== want) context[r] = 0;
    }
  }
  // Selections filter every chart except the one they were made in.
  const fails = new Uint8Array(n);
  const excluded = prepared.map((p) => {
    if (!p.chart.selection.size) return null;
    const out = new Uint8Array(n);
    for (let r = 0; r < n; r++) if (!p.chart.selection.has(p.keys[r])) { out[r] = 1; fails[r]++; }
    return out;
  });

  const charts = prepared.map((p, ci) => {
    const own = excluded[ci];
    const pass = (r) => fails[r] - (own ? own[r] : 0) === 0;
    const groups = new Map();
    const splitCol = p.chart.splitBy ? model.byName.get(p.chart.splitBy).col : null;
    const panels = splitCol ? new Map() : null;
    for (let r = 0; r < n; r++) {
      if (!context[r]) continue;
      const k = p.keys[r];
      const v = mcol ? mcol[r] : null;
      const fg = pass(r);
      add(groups, k, v, fg);
      if (panels) {
        const s = splitCol[r];
        if (!panels.has(s)) panels.set(s, new Map());
        add(panels.get(s), k, v, fg);
      }
    }
    const toItems = (map) => [...map].map(([key, a]) => ({
      key,
      label: p.chart.kind === 'date' ? dateLabel(p.level, key, p.depth > 0 && p.levels[0] === 'year') : key,
      bg: aggValue(a.bs, a.bn, a.bc, dash.agg),
      fg: aggValue(a.fs, a.fn, a.fc, dash.agg),
      rows: a.fc,
      selected: p.chart.selection.has(key),
    }));
    let items = toItems(groups);
    if (p.chart.kind === 'date') items.sort((a, b) => (a.key === BLANK) - (b.key === BLANK) || (a.key < b.key ? -1 : 1));
    else items.sort((a, b) => (b.bg ?? -Infinity) - (a.bg ?? -Infinity));
    let hidden = 0;
    const limit = p.chart.kind === 'dim' && !p.chart.showAll ? 10 : Infinity;
    if (items.length > limit) {
      const keep = items.filter((it, i) => i < limit || it.selected);
      hidden = items.length - keep.length;
      items = keep;
    }
    const out = {
      chart: p.chart, level: p.level, levels: p.levels, depth: p.depth,
      levelName: p.chart.kind === 'date' ? `${p.chart.field} (${p.level})` : p.level,
      canDrill: p.depth < p.levels.length - 1,
      items, hidden,
    };
    if (panels) {
      const order = items.map((it) => it.key);
      const index = new Map(order.map((k, i) => [k, i]));
      const all = [...panels].map(([key, map]) => {
        const byKey = new Map(toItems(map).map((it) => [it.key, it]));
        const panelItems = order.map((k) => byKey.get(k) || { key: k, label: items[index.get(k)].label, bg: null, fg: null, rows: 0, selected: p.chart.selection.has(k) });
        const total = [...map.values()].reduce((s, a) => s + a.bc, 0);
        return { key, items: p.chart.kind === 'dim' ? panelItems.slice(0, 5) : panelItems, rows: total };
      });
      all.sort((a, b) => b.rows - a.rows);
      out.panels = all.slice(0, 9);
      out.hiddenPanels = Math.max(0, all.length - 9);
    }
    return out;
  });

  // Headline number: everything applied, against the unfiltered total.
  let sAll = 0; let nAll = 0; let cAll = 0; let sTot = 0; let nTot = 0;
  for (let r = 0; r < n; r++) {
    const v = mcol ? mcol[r] : null;
    if (v != null) { sTot += v; nTot++; }
    if (context[r] && fails[r] === 0) {
      cAll++;
      if (v != null) { sAll += v; nAll++; }
    }
  }
  const kpi = {
    value: aggValue(sAll, nAll, cAll, dash.agg),
    total: aggValue(sTot, nTot, n, dash.agg),
    rows: cAll,
    totalRows: n,
    filtered: cAll !== n,
  };
  return { charts, kpi };

  function add(map, k, v, fg) {
    let a = map.get(k);
    if (!a) map.set(k, (a = newAcc()));
    a.bc++;
    if (v != null) { a.bs += v; a.bn++; }
    if (fg) {
      a.fc++;
      if (v != null) { a.fs += v; a.fn++; }
    }
  }
}

// ---------- Filters as words and as sheet filters ----------

export function activeFilters(model, dash) {
  const chips = [];
  for (const chart of dash.charts) {
    const levels = chartLevels(model, chart);
    chart.drill.forEach((d, i) => {
      chips.push({ chartId: chart.id, type: 'drill', depth: i, field: fieldLabel(chart, levels[i]), text: d.label });
    });
    if (chart.selection.size) {
      const level = levels[Math.min(chart.drill.length, levels.length - 1)];
      const insideYear = chart.kind === 'date' && chart.drill.length > 0 && levels[0] === 'year';
      const labels = [...chart.selection].map((k) => (chart.kind === 'date' ? dateLabel(level, k, insideYear) : k));
      chips.push({ chartId: chart.id, type: 'selection', field: fieldLabel(chart, level), text: labels.join(', ') });
    }
  }
  return chips;
  function fieldLabel(chart, level) {
    if (chart.kind !== 'date') return level;
    const word = level.charAt(0).toUpperCase() + level.slice(1);
    return model.fields.filter((f) => f.kind === 'date').length > 1 ? `${chart.field} ${level}` : word;
  }
}

// What the sheet's table filters should be, field by field.
//   { [field]: { kind: 'values', values: [...], blanks } | { kind: 'months', months: ['2025-04', ...] } }
export function sheetFilterSpec(model, dash) {
  const spec = {};
  for (const chart of dash.charts) {
    const levels = chartLevels(model, chart);
    const depth = Math.min(chart.drill.length, levels.length - 1);
    if (chart.kind === 'dim') {
      for (let i = 0; i < depth; i++) restrict(levels[i], new Set([chart.drill[i].key]));
      if (chart.selection.size) restrict(levels[depth], chart.selection);
      continue;
    }
    if (!depth && !chart.selection.size) continue;
    const field = model.byName.get(chart.field);
    const months = new Set();
    for (const mi of new Set(field.col)) {
      if (mi < 0) continue;
      let ok = true;
      for (let i = 0; i < depth && ok; i++) ok = dateKey(levels[i], mi) === chart.drill[i].key;
      if (ok && chart.selection.size) ok = chart.selection.has(dateKey(levels[depth], mi));
      if (ok) months.add(mi);
    }
    if (field.source === 'serial') {
      spec[field.name] = { kind: 'months', months: [...months].sort((a, b) => a - b).map(monthIndexToIso) };
    } else {
      const values = new Set();
      field.col.forEach((mi, r) => { if (months.has(mi)) values.add(field.raw[r]); });
      spec[field.name] = { kind: 'values', values: [...values], blanks: false };
    }
  }
  for (const [name, f] of Object.entries(spec)) {
    if (f.set) spec[name] = { kind: 'values', values: [...f.set].filter((v) => v !== BLANK), blanks: f.set.has(BLANK) };
  }
  return spec;

  function restrict(name, keys) {
    const prev = spec[name]?.set;
    spec[name] = { set: prev ? new Set([...prev].filter((k) => keys.has(k))) : new Set(keys) };
  }
}

// Rows that the sheet filter spec keeps; used by the preview and the tests.
export function rowsKeptBySpec(model, spec) {
  let kept = 0;
  for (let r = 0; r < model.n; r++) {
    let ok = true;
    for (const [name, f] of Object.entries(spec)) {
      const field = model.byName.get(name);
      if (f.kind === 'months') ok = f.months.includes(monthIndexToIso(field.col[r]));
      else if (field.kind === 'date') ok = f.values.includes(field.raw[r]);
      else ok = field.col[r] === BLANK ? f.blanks : f.values.includes(field.col[r]);
      if (!ok) break;
    }
    if (ok) kept++;
  }
  return kept;
}

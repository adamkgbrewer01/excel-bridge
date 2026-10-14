// Run with:  node --test tests/   (from the Add-in folder)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import {
  buildModel, defaultDashboard, evaluate, sheetFilterSpec, rowsKeptBySpec, activeFilters,
  formatValue, formatInfo, isDateFormat, textToMonthIndex, makeChart, restoreDashboard,
  serialiseDashboard, availableCharts, BLANK,
} from '../js/engine.js';

const sample = JSON.parse(readFileSync(new URL('../dev/sample-data.json', import.meta.url)));
const sum = (xs) => xs.reduce((s, x) => s + (x ?? 0), 0);
const close = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.01, `${msg}: ${a} vs ${b}`);
const revenueIdx = sample.headers.indexOf('Revenue');
const totalRevenue = sum(sample.rows.map((r) => r[revenueIdx]));

function fresh() {
  const model = buildModel(sample);
  return { model, dash: defaultDashboard(model) };
}
const chartFor = (dash, field) => dash.charts.find((c) => (c.kind === 'date' ? c.field : c.levels[0]) === field);

test('field kinds are detected', () => {
  const { model } = fresh();
  const kinds = Object.fromEntries(model.fields.map((f) => [f.name, f.kind]));
  assert.deepEqual(kinds, {
    'Order ID': 'id', 'Order Date': 'date', Region: 'dimension', Country: 'dimension',
    Category: 'dimension', Product: 'dimension', Channel: 'dimension',
    Units: 'measure', Revenue: 'measure', Profit: 'measure',
  });
  assert.equal(model.byName.get('Revenue').currency, '£');
});

test('drill hierarchies are found from the data', () => {
  const { model } = fresh();
  const chains = model.hierarchies.map((c) => c.join(' > ')).sort();
  assert.deepEqual(chains, ['Category > Product', 'Channel', 'Region > Country']);
});

test('default dashboard picks Revenue and a chart per hierarchy', () => {
  const { dash } = fresh();
  assert.equal(dash.measure, 'Revenue');
  assert.equal(dash.agg, 'sum');
  assert.deepEqual(dash.charts.map((c) => c.field || c.levels[0]).sort(), ['Category', 'Channel', 'Order Date', 'Region']);
});

test('unfiltered charts add up to the table total', () => {
  const { model, dash } = fresh();
  const out = evaluate(model, dash);
  for (const c of out.charts) {
    close(sum(c.items.map((i) => i.bg)), totalRevenue, c.levelName);
    close(sum(c.items.map((i) => i.fg)), totalRevenue, c.levelName);
  }
  assert.deepEqual(out.charts[0].items.map((i) => i.label), ['2023', '2024', '2025']);
  close(out.kpi.value, totalRevenue, 'kpi');
  assert.equal(out.kpi.filtered, false);
});

test('clicking a bar cross-filters the other charts but not its own', () => {
  const { model, dash } = fresh();
  const region = chartFor(dash, 'Region');
  region.selection.add('Europe');
  const out = evaluate(model, dash);
  const europe = sum(sample.rows.filter((r) => r[2] === 'Europe').map((r) => r[revenueIdx]));
  const own = out.charts.find((c) => c.chart === region);
  close(sum(own.items.map((i) => i.fg)), totalRevenue, 'own chart keeps every bar');
  assert.ok(own.items.find((i) => i.key === 'Europe').selected);
  for (const c of out.charts.filter((c) => c.chart !== region)) {
    close(sum(c.items.map((i) => i.fg)), europe, `${c.levelName} highlight`);
    close(sum(c.items.map((i) => i.bg)), totalRevenue, `${c.levelName} background`);
  }
  close(out.kpi.value, europe, 'kpi');
});

test('drilling into a year filters everything and shows quarters', () => {
  const { model, dash } = fresh();
  const date = chartFor(dash, 'Order Date');
  date.drill.push({ key: '2025', label: '2025' });
  const out = evaluate(model, dash);
  const dateOut = out.charts.find((c) => c.chart === date);
  assert.deepEqual(dateOut.items.map((i) => i.label), ['Q1', 'Q2', 'Q3', 'Q4']);
  const y2025 = sum(sample.rows.filter((r) => new Date((r[1] - 25569) * 86400000).getUTCFullYear() === 2025).map((r) => r[revenueIdx]));
  for (const c of out.charts) close(sum(c.items.map((i) => i.bg)), y2025, c.levelName);
  date.drill.push({ key: '2025-Q4', label: 'Q4' });
  const months = evaluate(model, dash).charts.find((c) => c.chart === date);
  assert.deepEqual(months.items.map((i) => i.label), ['Oct', 'Nov', 'Dec']);
  assert.equal(months.canDrill, false);
});

test('drilling a category hierarchy shows its children only', () => {
  const { model, dash } = fresh();
  const region = chartFor(dash, 'Region');
  region.drill.push({ key: 'North America', label: 'North America' });
  const out = evaluate(model, dash).charts.find((c) => c.chart === region);
  assert.deepEqual(out.items.map((i) => i.key).sort(), ['Canada', 'United States']);
  assert.equal(out.level, 'Country');
});

test('sheet filter spec keeps the same rows as the headline number', () => {
  const { model, dash } = fresh();
  chartFor(dash, 'Order Date').drill.push({ key: '2024', label: '2024' });
  chartFor(dash, 'Order Date').selection.add('2024-Q4');
  chartFor(dash, 'Region').selection.add('Europe');
  chartFor(dash, 'Region').selection.add('Asia Pacific');
  chartFor(dash, 'Category').drill.push({ key: 'Technology', label: 'Technology' });
  const spec = sheetFilterSpec(model, dash);
  assert.deepEqual(spec['Order Date'], { kind: 'months', months: ['2024-10', '2024-11', '2024-12'] });
  assert.deepEqual(spec.Region.values.sort(), ['Asia Pacific', 'Europe']);
  assert.deepEqual(spec.Category, { kind: 'values', values: ['Technology'], blanks: false });
  assert.equal(spec.Channel, undefined);
  const out = evaluate(model, dash);
  assert.equal(rowsKeptBySpec(model, spec), out.kpi.rows);
  assert.ok(out.kpi.rows > 0 && out.kpi.rows < model.n);
});

test('filter chips describe drills and selections', () => {
  const { model, dash } = fresh();
  chartFor(dash, 'Order Date').drill.push({ key: '2025', label: '2025' });
  chartFor(dash, 'Order Date').selection.add('2025-Q2');
  chartFor(dash, 'Channel').selection.add('Online');
  const chips = activeFilters(model, dash).map((c) => `${c.field}: ${c.text}`);
  assert.deepEqual(chips, ['Year: 2025', 'Quarter: Q2', 'Channel: Online']);
});

test('small multiples share categories across panels', () => {
  const { model, dash } = fresh();
  const date = chartFor(dash, 'Order Date');
  date.splitBy = 'Region';
  const out = evaluate(model, dash).charts.find((c) => c.chart === date);
  assert.equal(out.panels.length, 3);
  for (const p of out.panels) assert.deepEqual(p.items.map((i) => i.label), ['2023', '2024', '2025']);
  close(sum(out.panels.flatMap((p) => p.items.map((i) => i.bg))), totalRevenue, 'panels add up');
});

test('average and row count aggregations', () => {
  const { model, dash } = fresh();
  dash.agg = 'count';
  let out = evaluate(model, dash);
  assert.equal(out.kpi.value, model.n);
  dash.agg = 'avg';
  out = evaluate(model, dash);
  close(out.kpi.value, totalRevenue / model.n, 'average');
});

test('saved dashboards round-trip and drop stale fields', () => {
  const { model, dash } = fresh();
  chartFor(dash, 'Order Date').splitBy = 'Channel';
  const saved = JSON.parse(JSON.stringify(serialiseDashboard(dash)));
  saved.charts.push({ kind: 'dim', levels: ['Missing'] });
  const back = restoreDashboard(model, saved);
  assert.equal(back.charts.length, 4);
  assert.equal(chartFor(back, 'Order Date').splitBy, 'Channel');
  dash.charts = dash.charts.filter((c) => c.kind === 'date');
  assert.deepEqual(availableCharts(model, dash).map((a) => a.label).sort(), ['Category › Product', 'Channel', 'Region › Country']);
});

test('text dates, blanks and unrelated fields', () => {
  const table = {
    headers: ['Date', 'Team', 'Colour', 'Score'],
    formats: ['General', 'General', 'General', '0.0%'],
    rows: [
      ['2025-01-03', 'A', 'Red', 0.5], ['2025-02-10', 'A', 'Blue', 0.25],
      ['03/03/2025', 'B', 'Red', 0.75], ['', '', 'Blue', 0.1],
    ],
  };
  const model = buildModel(table);
  const date = model.byName.get('Date');
  assert.equal(date.kind, 'date');
  assert.equal(date.source, 'text');
  assert.equal(model.byName.get('Score').percent, true);
  assert.deepEqual(model.hierarchies.map((c) => c.join('>')).sort(), ['Colour', 'Team']);
  const dash = defaultDashboard(model);
  assert.equal(dash.agg, 'avg');
  const team = chartFor(dash, 'Team');
  team.selection.add(BLANK);
  team.selection.add('B');
  const spec = sheetFilterSpec(model, dash);
  assert.deepEqual(spec.Team, { kind: 'values', values: ['B'], blanks: true });
  chartFor(dash, 'Date').selection.add('2025-03');
  const spec2 = sheetFilterSpec(model, dash);
  assert.deepEqual(spec2.Date, { kind: 'values', values: ['03/03/2025'], blanks: false });
  assert.equal(rowsKeptBySpec(model, spec2), evaluate(model, dash).kpi.rows);
});

test('formatting helpers', () => {
  assert.equal(formatValue(2256239.6, { currency: '£' }), '£2.3M');
  assert.equal(formatValue(-845.4, { currency: '£' }), '-£845');
  assert.equal(formatValue(12345, {}), '12.3K');
  assert.equal(formatValue(0.256, { percent: true }), '25.6%');
  assert.equal(formatValue(null), '–');
  assert.equal(formatValue(1555766.5, { currency: '£' }, 'full'), '£1,555,767');
  assert.equal(formatValue(12.5, { currency: '£' }, 'full'), '£12.5');
  assert.deepEqual(formatInfo('[$€-x-euro2] #,##0.00'), { currency: '€', percent: false });
  assert.equal(formatInfo('\\£#,##0.00').currency, '£');
  assert.equal(isDateFormat('dd/mm/yyyy'), true);
  assert.equal(isDateFormat('h:mm'), false);
  assert.equal(isDateFormat('#,##0 "days"'), false);
  assert.equal(textToMonthIndex('2025-04-30'), 2025 * 12 + 3);
  assert.equal(textToMonthIndex('hello'), null);
});

test('chart ids stay unique', () => {
  assert.notEqual(makeChart({ kind: 'dim', levels: ['A'] }).id, makeChart({ kind: 'dim', levels: ['A'] }).id);
});

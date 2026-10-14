// Chart drawing for Excel Bridge. Plain DOM, one button per mark so every bar
// can be clicked, focused with the keyboard and read by a screen reader.
import { formatValue, niceMax } from './engine.js';

export function el(tag, cls, text) {
  const e = document.createElement(tag);
  if (cls) e.className = cls;
  if (text != null) e.textContent = text;
  return e;
}

// ---------- Tooltip ----------

let tip = null;
function tipEl() {
  if (!tip) {
    tip = el('div', 'tip');
    tip.setAttribute('role', 'tooltip');
    document.body.append(tip);
  }
  return tip;
}

export function hideTip() {
  if (tip) tip.classList.remove('on');
}

function showTip(target, item, opts, evt) {
  if (!target.isConnected) return;
  const t = tipEl();
  t.replaceChildren();
  t.append(el('div', 'k', opts.titleFor ? opts.titleFor(item) : item.label));
  const main = el('div', 'line');
  main.append(el('span', 'key'), el('span', 'v', formatValue(item.fg, opts.info, 'full')));
  t.append(main);
  if (item.bg != null && item.fg !== item.bg) {
    const share = item.bg && opts.agg !== 'avg' && item.fg != null ? ` (${Math.round((item.fg / item.bg) * 100)}%)` : '';
    const line = el('div', 'line');
    line.append(el('span', 'key faded'), el('span', 'k', `${formatValue(item.bg, opts.info, 'full')} before filters${share}`));
    t.append(line);
  }
  if (opts.help) t.append(el('div', 'help', opts.help(item)));
  t.classList.add('on');
  const r = target.getBoundingClientRect();
  const w = t.offsetWidth;
  const h = t.offsetHeight;
  const x = evt && evt.clientX ? evt.clientX : r.left + r.width / 2;
  let top = r.top - h - 8;
  if (top < 4) top = r.bottom + 8;
  const left = Math.max(4, Math.min(window.innerWidth - w - 4, x - w / 2));
  t.style.left = `${left}px`;
  t.style.top = `${top}px`;
}

function wireMark(button, item, opts) {
  button.type = 'button';
  button.dataset.fid = `${opts.fid}:${item.key}`;
  button.setAttribute('aria-pressed', String(!!item.selected));
  button.setAttribute('aria-label', `${item.label}, ${formatValue(item.fg, opts.info, 'full')}`);
  if (opts.onPick) button.addEventListener('click', (e) => { hideTip(); opts.onPick(item.key, e); });
  button.addEventListener('pointermove', (e) => showTip(button, item, opts, e));
  button.addEventListener('pointerleave', hideTip);
  button.addEventListener('focus', () => showTip(button, item, opts));
  button.addEventListener('blur', hideTip);
}

// ---------- Geometry ----------

function extent(items) {
  let min = 0;
  let max = 0;
  for (const it of items) {
    for (const v of [it.bg, it.fg]) {
      if (v == null || !Number.isFinite(v)) continue;
      if (v < min) min = v;
      if (v > max) max = v;
    }
  }
  return { min, max };
}

export function barScale(items) {
  const { min, max } = extent(items);
  return { min, max: max === min ? min + 1 : max };
}

export function columnScale(items) {
  const { min, max } = extent(items);
  const top = max > 0 ? niceMax(max) : 0;
  const bottom = min < 0 ? -niceMax(-min) : 0;
  return { min: bottom, max: top === bottom ? bottom + 1 : top };
}

function place(v, s) {
  const range = s.max - s.min;
  const base = ((0 - s.min) / range) * 100;
  if (v == null || !Number.isFinite(v)) return { start: base, size: 0, cls: 'pos' };
  const size = (Math.abs(v) / range) * 100;
  return v >= 0 ? { start: base, size, cls: 'pos' } : { start: base - size, size, cls: 'neg' };
}

// ---------- Horizontal bars ----------

export function barList(items, opts) {
  const s = opts.scale || barScale(items);
  const wrap = el('div', 'bars');
  for (const item of items) {
    const b = el('button', `mark${item.selected ? ' selected' : ''}`);
    const name = el('span', 'name', item.label);
    name.title = item.label;
    const track = el('span', 'track');
    if (s.min < 0) {
      const zero = el('span', 'zero');
      zero.style.left = `${place(0, s).start}%`;
      track.append(zero);
    }
    for (const [cls, v] of [['bg', item.bg], ['fg', item.fg]]) {
      const p = place(v, s);
      const bar = el('span', `${cls} ${p.cls}`);
      bar.style.left = `${p.start}%`;
      bar.style.width = `${p.size}%`;
      track.append(bar);
    }
    b.append(name, track, el('span', 'val', formatValue(item.fg, opts.info)));
    wireMark(b, item, opts);
    wrap.append(b);
  }
  return wrap;
}

// ---------- Columns ----------

export function columnChart(items, opts) {
  const s = opts.scale || columnScale(items);
  const height = opts.compact ? 64 : 132;
  const root = el('div', 'colchart');
  const cols = el('div', 'cols');
  cols.style.height = `${height}px`;
  const plot = el('div', 'plot');
  const ticks = s.min < 0 ? [s.min, 0, s.max] : [0, s.max / 2, s.max];
  if (!opts.compact) {
    const axis = el('div', 'yaxis');
    for (const t of ticks) {
      const label = el('span', '', formatValue(t, opts.info));
      label.style.bottom = `${place(t, s).start + (t > 0 ? place(t, s).size : 0)}%`;
      axis.append(label);
    }
    cols.append(axis);
  }
  for (const t of ticks) {
    const line = el('div', `gridline${t === 0 ? ' base' : ''}`);
    const p = place(t, s);
    line.style.bottom = `${t > 0 ? p.start + p.size : p.start}%`;
    plot.append(line);
  }
  const slots = el('div', 'slots');
  for (const item of items) {
    const b = el('button', `mark${item.selected ? ' selected' : ''}`);
    for (const [cls, v] of [['bg', item.bg], ['fg', item.fg]]) {
      const p = place(v, s);
      const bar = el('span', `${cls} ${p.cls}`);
      bar.style.bottom = `${p.start}%`;
      bar.style.height = `${p.size}%`;
      b.append(bar);
    }
    if (item.selected && !opts.compact && item.fg != null) {
      const p = place(item.fg, s);
      const cap = el('span', 'cap', formatValue(item.fg, opts.info));
      cap.style.bottom = `${item.fg >= 0 ? p.start + p.size : p.start}%`;
      if (item.fg < 0) cap.style.transform = 'translate(-50%, 100%)';
      b.append(cap);
    }
    wireMark(b, item, opts);
    slots.append(b);
  }
  plot.append(slots);
  cols.append(plot);
  root.append(cols);

  const labels = el('div', `xlabels${opts.compact ? ' ends' : ''}`);
  if (opts.compact) {
    if (items.length) labels.append(el('span', '', items[0].label));
    if (items.length > 1) labels.append(el('span', '', items[items.length - 1].label));
  } else {
    const every = items.length > 24 ? 3 : items.length > 12 ? 2 : 1;
    items.forEach((item, i) => {
      const span = el('span', item.selected ? 'selected' : '', i % every === 0 ? item.label : '');
      labels.append(span);
    });
  }
  root.append(labels);
  return root;
}

// ---------- Table view ----------

export function dataTable(columns, rows, opts = {}) {
  const table = el('table', 'data');
  const thead = el('thead');
  const hr = el('tr');
  for (const c of columns) {
    const th = el('th', c.num ? 'num' : '', c.label);
    th.scope = 'col';
    hr.append(th);
  }
  thead.append(hr);
  const tbody = el('tbody');
  for (const row of rows) {
    const tr = el('tr', row.selected ? 'selected' : '');
    row.cells.forEach((cell, i) => tr.append(el('td', columns[i].num ? 'num' : '', cell)));
    if (opts.onPick && row.key != null) {
      tr.addEventListener('click', (e) => opts.onPick(row.key, e));
      tr.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); opts.onPick(row.key, e); }
      });
      tr.tabIndex = 0;
      tr.dataset.fid = `${opts.fid}:${row.key}`;
      tr.title = 'Click to filter';
    }
    tbody.append(tr);
  }
  table.append(thead, tbody);
  return table;
}

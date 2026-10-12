'use strict';

// Renders each [data-chart] as paired horizontal bars (ExcelJS, excel-rs) per
// row, scaled to the row's ExcelJS value, with a hover/focus tooltip. The
// same numbers are in the tables below each chart.
(function () {
  const tip = document.createElement('div');
  tip.className = 'tip';
  tip.setAttribute('role', 'tooltip');
  document.body.appendChild(tip);

  function fmt(v, unit) {
    return unit === 's' ? `${v.toFixed(2)} s` : `${Math.round(v).toLocaleString('en')} MB`;
  }

  function bar(cls, value, base, unit, label, note) {
    const el = document.createElement('div');
    el.className = `bar ${cls}`;
    el.tabIndex = 0;
    el.setAttribute('aria-label', label);
    el.dataset.tip = label;
    const fill = document.createElement('div');
    fill.className = 'fill';
    fill.style.width = `${Math.max(0.5, (value / base) * 85)}%`;
    const val = document.createElement('span');
    val.className = 'val';
    val.textContent = note ? `${fmt(value, unit)} · ${note}` : fmt(value, unit);
    el.append(fill, val);
    return el;
  }

  for (const host of document.querySelectorAll('[data-chart]')) {
    const spec = JSON.parse(host.dataset.chart);
    host.innerHTML = '';
    const title = document.createElement('p');
    title.className = 'chart-title';
    title.textContent = spec.title;
    const sub = document.createElement('p');
    sub.className = 'chart-sub';
    sub.textContent = spec.sub;
    const legend = document.createElement('div');
    legend.className = 'legend';
    legend.innerHTML =
      '<span><i style="background:var(--series-2)"></i>ExcelJS 4.4.0</span>' +
      '<span><i style="background:var(--series-1)"></i>excel-rs</span>';
    host.append(title, sub, legend);
    for (const [name, detail, a, b] of spec.rows) {
      const row = document.createElement('div');
      row.className = 'row';
      const label = document.createElement('div');
      label.className = 'name';
      label.textContent = name;
      if (detail) {
        const small = document.createElement('small');
        small.textContent = detail;
        label.appendChild(small);
      }
      const bars = document.createElement('div');
      bars.className = 'bars';
      const ratio = spec.unit === 's' ? `${(a / b).toFixed(1)}× faster` : `${Math.round((1 - b / a) * 100)}% less`;
      bars.append(
        bar('a', a, a, spec.unit, `${name}: ExcelJS ${fmt(a, spec.unit)}`),
        bar('b', b, a, spec.unit, `${name}: excel-rs ${fmt(b, spec.unit)} (${ratio})`, ratio)
      );
      row.append(label, bars);
      host.appendChild(row);
    }
  }

  function show(el, x, y) {
    tip.textContent = el.dataset.tip;
    tip.style.display = 'block';
    const w = tip.offsetWidth;
    tip.style.left = `${Math.min(x + 12, window.innerWidth - w - 8)}px`;
    tip.style.top = `${y + 14}px`;
  }
  document.addEventListener('pointermove', e => {
    const el = e.target.closest && e.target.closest('.bar');
    if (el) show(el, e.clientX, e.clientY);
    else tip.style.display = 'none';
  });
  document.addEventListener('focusin', e => {
    const el = e.target.closest && e.target.closest('.bar');
    if (!el) return;
    const r = el.getBoundingClientRect();
    show(el, r.left, r.bottom - 10);
  });
  document.addEventListener('focusout', () => {
    tip.style.display = 'none';
  });
})();

/* 总览界面直接接收 app.js 的 currentReport，不另行读文件或保存经营数据。 */
(function () {
  'use strict';
  const $ = id => document.getElementById(id);
  let report = null, result = null, detailPage = 0;
  const PAGE_SIZE = 31;
  const controls = { start: 'overview-start', end: 'overview-end', source: 'overview-source', business: 'overview-business', carrier: 'overview-carrier', period: 'overview-period', account: 'overview-account', spend: 'overview-spend', refund: 'overview-refund' };
  function el(tag, text, className) {
    const node = document.createElement(tag);
    if (text !== undefined) node.textContent = text;
    if (className) node.className = className;
    return node;
  }
  function options(id, entries, selected) {
    const select = $(id); select.replaceChildren();
    for (const entry of entries) {
      const item = typeof entry === 'string' ? { id: entry, label: entry } : entry;
      const option = el('option', item.label); option.value = item.id; select.append(option);
    }
    select.value = selected;
  }
  function filters() { return Object.fromEntries(Object.entries(controls).map(([key, id]) => [key, $(id).value])); }
  function localTime(value) {
    if (!value || Number.isNaN(new Date(value).getTime())) return '未记录';
    return new Intl.DateTimeFormat('zh-CN', { dateStyle: 'medium', timeStyle: 'medium', timeZoneName: undefined }).format(new Date(value)) + '（浏览器本地时间）';
  }
  function setReport(value) {
    report = value;
    const c = Overview.choices(report), f = Overview.defaults(report);
    options('overview-business', c.business, f.business);
    options('overview-carrier', c.carrier, f.carrier);
    options('overview-period', c.period, f.period);
    options('overview-account', c.accounts, f.account);
    const sourceEntries = [];
    if (report.sheets.some(sheet => sheet.name === '收支分析')) sourceEntries.push({ id: 'store', label: '店铺维度（收支分析）' });
    if (c.accounts.length) sourceEntries.push({ id: 'account', label: '单个账号（账号收支详情）' });
    options('overview-source', sourceEntries, 'store');
    for (const [key, id] of Object.entries(controls)) if (f[key] !== undefined) $(id).value = f[key];
    $('overview-file').textContent = report.fileName;
    $('overview-file-range').textContent = report.dates[0] + ' 至 ' + report.dates.at(-1);
    $('overview-imported-at').textContent = localTime(report.importedAt);
    $('overview-platform-time').textContent = '报表未提供，无法确认；导入时间不等于平台数据截至时间';
    $('overview-empty').hidden = true;
    $('overview-content').hidden = false;
    refresh();
  }
  function refresh() {
    if (!report) return;
    const f = filters();
    const accountMode = f.source === 'account';
    $('overview-account-wrap').hidden = !accountMode;
    for (const key of ['business', 'carrier', 'period']) $(controls[key]).disabled = accountMode;
    detailPage = 0;
    result = Overview.calculate(report, f);
    $('overview-scope').textContent = '当前范围：' + result.scope;
    const notice = $('overview-notices'); notice.replaceChildren();
    result.notices.forEach(text => notice.append(el('p', text)));
    notice.hidden = !result.notices.length;
    $('overview-empty-selection').hidden = result.hasRows;
    renderMetrics(); renderTrend(); renderDaily(); renderHints();
  }
  function renderMetrics() {
    const container = $('overview-metrics'); container.replaceChildren();
    for (const metric of result.metrics) {
      const card = el('article', undefined, 'metric overview-metric'); card.dataset.metric = metric.id;
      card.append(el('h3', metric.label));
      const value = el('p', metric.value === null ? '无法计算' : metric.value, 'metric-value');
      value.dataset.value = metric.value ?? 'unavailable'; card.append(value);
      card.append(el('p', '单位：' + metric.unit, 'metric-unit'));
      card.append(el('p', '公式：' + metric.formula, 'metric-formula'));
      if (metric.reason) card.append(el('p', metric.reason, 'metric-reason'));
      card.append(el('p', metric.notes, 'metric-notes'));
      const details = el('details'); details.append(el('summary', '来源与汇总规则'));
      details.append(el('p', '来源工作表：' + metric.source));
      details.append(el('p', '使用字段：' + metric.fields.join('、')));
      details.append(el('p', '筛选条件：' + metric.scope));
      details.append(el('p', '汇总规则：' + metric.aggregation));
      details.append(el('p', metric.snapshot));
      card.append(details); container.append(card);
    }
  }
  function svgNode(tag, attributes, text) {
    const node = document.createElementNS('http://www.w3.org/2000/svg', tag);
    for (const [key, value] of Object.entries(attributes || {})) node.setAttribute(key, String(value));
    if (text !== undefined) node.textContent = text;
    return node;
  }
  function renderTrend() {
    const container = $('overview-chart'); container.replaceChildren();
    const series = [
      { key: 'sales', name: '成交金额', color: '#24634c' },
      { key: 'refund', name: result.refundField, color: '#ac4c25' },
      { key: 'spend', name: result.spendField, color: '#385faa' }
    ];
    const legend = $('overview-legend'); legend.replaceChildren();
    series.forEach(item => { const label = el('span', item.name); label.style.borderColor = item.color; legend.append(label); });
    const finite = result.daily.flatMap(day => series.map(item => day[item.key])).filter(value => value !== null);
    if (!finite.length) { container.append(el('p', '所选范围没有可绘制的数据；缺失不画成 0。', 'helper')); return; }
    // 只有绘图坐标使用近似 Number；所有金额及比率计算保留 BigInt。
    let min = 0, max = 0;
    finite.forEach(value => { const n = Number(value) / 100; min = Math.min(min, n); max = Math.max(max, n); });
    if (min === max) max = min + 1;
    const width = 960, height = 310, left = 100, right = 24, top = 20, bottom = 52;
    const x = i => left + (result.daily.length > 1 ? i / (result.daily.length - 1) : 0.5) * (width - left - right);
    const y = value => top + (max - Number(value) / 100) / (max - min) * (height - top - bottom);
    const svg = svgNode('svg', { viewBox: '0 0 ' + width + ' ' + height, role: 'img', 'aria-labelledby': 'trend-title trend-description' });
    svg.append(svgNode('title', { id: 'trend-title' }, '每日成交、成交退款及投放消耗趋势'));
    svg.append(svgNode('desc', { id: 'trend-description' }, '单位为元。缺失记录时曲线断开，真实零值位于零线。点可聚焦查看日期和金额，精确值见每日明细。'));
    for (let i = 0; i <= 4; i++) {
      const amount = min + (max - min) * i / 4;
      const yy = y(amount * 100);
      svg.append(svgNode('line', { x1: left, y1: yy, x2: width - right, y2: yy, stroke: '#dfe6e1' }));
      svg.append(svgNode('text', { x: left - 8, y: yy + 5, 'text-anchor': 'end', fill: '#66746e', 'font-size': 13 }, amount.toLocaleString('zh-CN', { maximumFractionDigits: 2 })));
    }
    const ticks = new Set([0, result.daily.length - 1]);
    for (let n = 1; n < 5; n++) ticks.add(Math.floor((result.daily.length - 1) * n / 5));
    for (const i of ticks) svg.append(svgNode('text', { x: x(i), y: height - 20, 'text-anchor': 'middle', fill: '#66746e', 'font-size': 13 }, result.daily[i].date.slice(5)));
    for (const item of series) {
      let segment = [];
      const flush = () => {
        if (segment.length > 1) svg.append(svgNode('polyline', { points: segment.join(' '), fill: 'none', stroke: item.color, 'stroke-width': 2, 'data-series': item.key }));
        segment = [];
      };
      result.daily.forEach((day, i) => {
        if (day[item.key] === null) { flush(); return; }
        const cx = x(i), cy = y(day[item.key]); segment.push(cx + ',' + cy);
        const label = day.date + ' ' + item.name + ' ' + ReportImport.formatMoney(day[item.key]) + ' 元';
        const dot = svgNode('circle', { cx, cy, r: 3.5, fill: item.color, tabindex: 0, 'aria-label': label, 'data-series': item.key, 'data-date': day.date });
        dot.append(svgNode('title', {}, label)); svg.append(dot);
      });
      flush();
    }
    container.append(svg);
  }
  function displayMoney(value) { return value === null ? '缺失 / 无法计算' : ReportImport.formatMoney(value); }
  function renderDaily() {
    const head = $('daily-table').querySelector('thead'), body = $('daily-table').querySelector('tbody');
    head.replaceChildren(); body.replaceChildren();
    const heading = el('tr');
    for (const label of ['日期', '成交金额（元）', '直播成交金额（元）', result.refundField + '（元）', result.spendField + '（元）', '退款金额占比（支付时间）', '整体成交消耗比', '原文件行号', '直播子集原行号', '数据说明']) {
      const th = el('th', label); th.scope = 'col'; heading.append(th);
    }
    head.append(heading);
    const start = detailPage * PAGE_SIZE;
    for (const day of result.daily.slice(start, start + PAGE_SIZE)) {
      const tr = el('tr'); tr.dataset.date = day.date;
      const status = day.issues.length ? day.issues.join('；') : '有效记录；数值 0 表示原表实际零值';
      const values = [day.date, displayMoney(day.sales), displayMoney(day.live), displayMoney(day.refund), displayMoney(day.spend), day.refundRatio === null ? '无法计算' : day.refundRatio + '%', day.overallRatio === null ? '无法计算' : day.overallRatio + ' 倍', day.rows.join('、') || '无匹配记录', day.liveRows.join('、') || '不适用 / 无匹配记录', status];
      values.forEach((value, i) => { const cell = el(i === 0 ? 'th' : 'td', value); if (i === 0) cell.scope = 'row'; tr.append(cell); });
      body.append(tr);
    }
    const pages = Math.max(1, Math.ceil(result.daily.length / PAGE_SIZE));
    $('daily-page-info').textContent = '第 ' + (detailPage + 1) + ' / ' + pages + ' 页，共 ' + result.daily.length + ' 个日历日（包含缺失日）';
    $('daily-previous').disabled = detailPage === 0;
    $('daily-next').disabled = detailPage + 1 >= pages;
    $('daily-caption').textContent = '每日明细 — 来源：' + result.source + '；金额单位：元';
    $('daily-table').closest('.table-scroll').scrollTop = 0;
  }
  function renderHints() {
    const container = $('overview-hints'); container.replaceChildren();
    if (!result.hints.length) { container.append(el('p', result.hasRows ? '当前范围未触发可核实的规则提示；不代表经营正常或没有风险。' : '暂无可用于经营提示的数据。', 'helper')); return; }
    for (const hint of result.hints) {
      const card = el('article', undefined, 'hint-card'); card.append(el('h4', hint.title));
      for (const [label, text] of [['数据事实', hint.fact], ['待核实的问题', hint.question], ['建议动作', hint.action], ['后续验证指标', hint.verify]]) {
        const p = el('p'); p.append(el('strong', label + '：'), document.createTextNode(text)); card.append(p);
      }
      container.append(card);
    }
  }
  $('overview-filter-form').addEventListener('submit', event => event.preventDefault());
  Object.values(controls).forEach(id => $(id).addEventListener('change', refresh));
  $('overview-reset').addEventListener('click', () => { if (report) setReport(report); });
  $('daily-previous').addEventListener('click', () => { if (detailPage > 0) { detailPage--; renderDaily(); } });
  $('daily-next').addEventListener('click', () => { if ((detailPage + 1) * PAGE_SIZE < result.daily.length) { detailPage++; renderDaily(); } });
  window.OverviewUI = { setReport };
})();

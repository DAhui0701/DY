'use strict';

const navigationLinks = document.querySelectorAll('nav a[href^="#"]');
function updateNavigation() {
  const currentHash = window.location.hash || '#overview';
  navigationLinks.forEach(link => {
    if (link.getAttribute('href') === currentHash) link.setAttribute('aria-current', 'location');
    else link.removeAttribute('aria-current');
  });
}
window.addEventListener('hashchange', updateNavigation);
updateNavigation();

const element = id => document.getElementById(id);
const input = element('report-file');
const chooseButton = element('choose-report');
chooseButton.addEventListener('click', () => input.click());
const PAGE_SIZE = 50;
let currentReport = null;
let currentBytes = null;
let sheetIndex = 0;
let pageIndex = 0;
let importing = false;

// 导入文件的所有文本均通过 textContent 展示，不作为 HTML 执行。
function node(tag, content, className) {
  const result = document.createElement(tag);
  if (content !== undefined) result.textContent = content;
  if (className) result.className = className;
  return result;
}
function sameBytes(left, right) {
  if (!left || left.length !== right.length) return false;
  for (let i = 0; i < left.length; i++) if (left[i] !== right[i]) return false;
  return true;
}
function dateRange(dates) {
  return dates.length ? dates[0] + ' 至 ' + dates.at(-1) : '未识别（仅原样展示）';
}

function renderSummary() {
  const report = currentReport;
  const summary = element('file-summary');
  summary.replaceChildren();
  summary.append(node('h3', '当前有效文件'));
  const facts = node('dl', undefined, 'file-facts');
  for (const [label, value] of [
    ['文件名', report.fileName], ['记录日期', dateRange(report.dates)],
    ['工作表', report.sheets.length + ' 个'], ['记录数', report.totalRows + ' 条（不含表头和全空行）']
  ]) facts.append(node('dt', label), node('dd', value));
  summary.append(facts);
  const list = node('ul', undefined, 'sheet-list');
  report.sheets.forEach(sheet => list.append(node('li', sheet.name + '：' + sheet.rows.length + ' 条 / ' + sheet.headers.length + ' 个字段；日期：' + dateRange(sheet.dates))));
  summary.append(list);
  const details = node('details');
  details.append(node('summary', '查看导入金额核对与汇总口径'));
  details.append(node('p', '只核对“收支分析”中“全店＋全部＋不限”的 ' + report.selectedCount + ' 条每日汇总记录。不同层级、两个工作表不能直接相加。', 'helper'));
  const totals = node('dl', undefined, 'file-facts');
  for (const [field, value] of Object.entries(report.totals)) totals.append(node('dt', field + '（元）'), node('dd', ReportImport.formatMoney(value)));
  details.append(totals);
  details.append(node('p', '金额使用整数“分”核对。费比、效率、退款率仅保留原值，不求和或简单平均；不推断平台公式，不将收支差额称为净利润。', 'helper'));
  if (report.missingRatios) details.append(node('p', '原表共有 ' + report.missingRatios + ' 个比率/效率单元格为“-”，表示未提供值，已原样保留。', 'helper'));
  summary.append(details);
  summary.hidden = false;
  const warnings = element('import-warnings');
  warnings.replaceChildren();
  report.warnings.forEach(message => warnings.append(node('p', message)));
  warnings.hidden = !report.warnings.length;
  element('file-badge').textContent = '报表已导入 · 仅当前页面';
  element('data-count').textContent = report.totalRows + ' 条原始记录';
  const select = element('sheet-select');
  select.replaceChildren();
  report.sheets.forEach((sheet, index) => {
    const option = node('option', sheet.name + '（' + sheet.rows.length + ' 条）');
    option.value = String(index);
    select.append(option);
  });
  element('data-empty').hidden = true;
  element('data-browser').hidden = false;
  sheetIndex = 0;
  pageIndex = 0;
  renderTable();
}

function renderTable() {
  const sheet = currentReport.sheets[sheetIndex];
  const table = element('raw-table');
  const head = table.querySelector('thead');
  const body = table.querySelector('tbody');
  head.replaceChildren(); body.replaceChildren();
  const heading = node('tr');
  for (const name of ['原文件行号', ...sheet.headers]) {
    const cell = node('th', name); cell.scope = 'col'; heading.append(cell);
  }
  head.append(heading);
  const start = pageIndex * PAGE_SIZE;
  const end = Math.min(start + PAGE_SIZE, sheet.rows.length);
  const fragment = document.createDocumentFragment();
  for (const row of sheet.rows.slice(start, end)) {
    const tr = node('tr');
    const rowHeader = node('th', row.sourceRow); rowHeader.scope = 'row'; tr.append(rowHeader);
    for (const value of row.values) tr.append(node('td', value === '' ? '（空）' : value));
    fragment.append(tr);
  }
  body.append(fragment);
  const pages = Math.max(1, Math.ceil(sheet.rows.length / PAGE_SIZE));
  element('table-caption').textContent = sheet.name + ' — 原始字段与记录';
  element('sheet-info').textContent = sheet.headers.length + ' 个字段 / ' + sheet.rows.length + ' 条记录 / ' + dateRange(sheet.dates);
  element('page-info').textContent = '第 ' + (pageIndex + 1) + ' / ' + pages + ' 页 · 当前 ' + (sheet.rows.length ? start + 1 : 0) + '–' + end + ' 条，共 ' + sheet.rows.length + ' 条';
  element('previous-page').disabled = pageIndex === 0;
  element('next-page').disabled = pageIndex + 1 >= pages;
  element('raw-table').closest('.table-scroll').scrollTop = 0;
}

element('sheet-select').addEventListener('change', event => {
  sheetIndex = Number(event.target.value); pageIndex = 0; renderTable();
});
element('previous-page').addEventListener('click', () => { if (pageIndex > 0) { pageIndex--; renderTable(); } });
element('next-page').addEventListener('click', () => {
  if ((pageIndex + 1) * PAGE_SIZE < currentReport.sheets[sheetIndex].rows.length) { pageIndex++; renderTable(); }
});

input.addEventListener('change', async () => {
  const file = input.files[0];
  if (!file || importing) return;
  importing = true;
  input.disabled = true;
  chooseButton.disabled = true;
  const status = element('import-status');
  const errorBox = element('import-error');
  errorBox.hidden = true;
  errorBox.textContent = '';
  status.textContent = '正在读取和校验，请稍候；当前有效数据暂不变更。';
  try {
    if (!window.Overview || !window.OverviewUI) throw new Error('经营总览组件未加载，请确认 overview.js 和 overview-ui.js 文件完整后重新打开。');
    if (!window.ReportImport || !window.ExcelJS) throw new Error('Excel 读取组件未加载，请确认 report.js 和 vendor 文件夹完整，然后重新打开页面。');
    if (!/\.xlsx$/i.test(file.name)) throw new Error('文件类型不支持，请选择 .xlsx 文件；不支持 .xls、.xlsm、CSV 或加密工作簿。');
    if (!file.size) throw new Error('文件为空，请重新导出 Excel。');
    if (file.size > ReportImport.LIMITS.bytes) throw new Error('文件超过 10 MB，请缩小导出日期范围。');
    const buffer = await file.arrayBuffer();
    const bytes = new Uint8Array(buffer);
    if (sameBytes(currentBytes, bytes)) {
      status.textContent = '该文件内容与当前文件完全相同，已跳过，不会重复累计（即使改了文件名也不会重复）。';
      return;
    }
    if (currentReport && !window.confirm('导入新文件将替换当前“' + currentReport.fileName + '”的全部工作表，不会追加或累计。\n新文件校验失败时，仍保留当前有效数据。是否继续？')) {
      status.textContent = '已取消替换，当前有效数据保持不变。';
      return;
    }
    ReportImport.inspectArchive(buffer);
    if (!window.ExcelJS) throw new Error('Excel 读取组件未加载，请确认 vendor 文件夹完整，然后重新打开页面。');
    const workbook = new ExcelJS.Workbook();
    try { await workbook.xlsx.load(buffer); }
    catch { throw new Error('Excel 解析失败：文件可能损坏、加密或不是标准 .xlsx，请重新导出。'); }
    const candidate = ReportImport.parseWorkbook(workbook, file.name);
    candidate.importedAt = new Date().toISOString();
    // 所有工作表均校验成功后才替换，避免失败时丢失旧数据。
    currentReport = candidate;
    currentBytes = bytes;
    renderSummary();
    OverviewUI.setReport(currentReport);
    status.textContent = '导入成功：' + candidate.sheets.length + ' 个工作表、' + candidate.totalRows + ' 条记录。当前数据已替换，不累计。';
  } catch (error) {
    errorBox.textContent = (error instanceof Error ? error.message : '发生未知错误，请重新导出报表后重试。') + '\n' + (currentReport ? '上一份有效数据已保留，可继续查看。' : '尚无有效数据，页面保持空状态。');
    errorBox.hidden = false;
    status.textContent = '导入失败，未替换当前数据。';
  } finally {
    input.value = '';
    input.disabled = false;
    chooseButton.disabled = false;
    importing = false;
  }
});

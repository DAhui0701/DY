// 仅含人工构造的测试输入，不包含用户报表、账号或真实经营金额。
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs-4.4.0.min.js');
const report = require('../report.js');
const metrics = [
  '结算金额', '7日结算金额', '14日结算金额', '结算金额（退款时间）',
  '投放消耗（店铺被投）', '投放消耗（店铺绑定）',
  '投放费比（店铺被投）', '投放费比（剔除退款、店铺被投）',
  '投放费比（店铺绑定）', '投放费比（剔除退款、店铺绑定）',
  '综合费比（店铺被投）', '综合费比（剔除退款、店铺被投）',
  '综合费比（店铺绑定）', '综合费比（剔除退款、店铺绑定）',
  '投放效率（店铺被投）', '投放效率（店铺绑定）',
  '全店效率（店铺被投）', '全店效率（店铺绑定）',
  '成交金额', '用户支付金额', '预售定金', '智能优惠券金额',
  '电商平台补贴金额', '达人补贴金额', '成交退款金额（支付时间）',
  '退款金额（支付时间）', '智能优惠券退款（支付时间）',
  '平台补贴退款（支付时间）', '达人补贴退款（支付时间）',
  '成交退款金额（退款时间）', '退款金额（退款时间）',
  '智能优惠券退款+平台补贴退款（退款时间）', '达人补贴退款（退款时间）',
  '支出金额', '平台佣金（预估口径）', '达人佣金（预估口径）',
  '退款率（支付时间）', '免佣金额汇总', '免佣金额明细'
];
function fixture() {
  const workbook = new ExcelJS.Workbook();
  const main = workbook.addWorksheet('收支分析');
  main.addRow(['日期', '经营类型', '载体类型', '推广时段', ...metrics]);
  main.addRow(['20240101', '全店', '全部', '不限', ...metrics.map(() => '0.00')]);
  const account = workbook.addWorksheet('账号收支详情');
  account.addRow(['日期', '账号名称', '账号类型', '抖音号', ...metrics]);
  account.addRow(['20240101', '测试账号（非真实）', '自营', 'TEST_ONLY', ...metrics.map(() => '0.00')]);
  return workbook;
}
function cell(workbook, field, row = 2) {
  const sheet = workbook.getWorksheet('收支分析');
  const column = sheet.getRow(1).values.indexOf(field);
  return sheet.getCell(row, column);
}
test('金额按分计算，保留负数并拒绝无效金额', () => {
  assert.equal(report.cents('0.10') + report.cents('0.20'), 30n);
  assert.equal(report.formatMoney(report.cents('-1234.56')), '-1,234.56');
  for (const value of ['', '-', 'NaN', '1.234', '=1+1', '1,23', 'Infinity']) assert.throws(() => report.cents(value));
});
test('日期验证实际日历，不接受未来日期', () => {
  assert.equal(report.date('20240229'), '2024-02-29');
  assert.equal(report.date(new Date('2024-02-29T00:00:00Z')), '2024-02-29');
  for (const value of ['20230229', '20241301', '20240431', '', '20990101']) assert.throws(() => report.date(value));
});
test('完整工作簿可读；汇总不叠加细分行和账号表', () => {
  const w = fixture();
  cell(w, '成交金额').value = '0.10';
  cell(w, '结算金额').value = '-0.20';
  const values = w.getWorksheet('收支分析').getRow(2).values.slice(1);
  values[1] = '自营'; values[22] = '999.99';
  w.getWorksheet('收支分析').addRow(values);
  cell(w, '投放效率（店铺被投）').value = '-';
  const result = report.parseWorkbook(w, '测试.xlsx');
  assert.equal(result.totalRows, 3);
  assert.equal(result.selectedCount, 1);
  assert.equal(result.totals['成交金额'], 10n);
  assert.equal(result.totals['结算金额'], -20n);
  assert.equal(result.missingRatios, 1);
});
test('缺少工作表、必需字段及表头重复时提供中文提示', () => {
  let w = fixture(); w.removeWorksheet('账号收支详情');
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /缺少工作表/);
  w = fixture(); cell(w, '成交金额', 1).value = '错误字段';
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /缺少字段：成交金额/);
  w = fixture(); cell(w, '成交金额', 1).value = '结算金额';
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /重复表头/);
});
test('金额、比率、日期、空维度和重复粒度错误会阻止导入', () => {
  const cases = [
    ['成交金额', '无效金额', /金额无效/],
    ['日期', '20240230', /日期异常/],
    ['经营类型', '', /不能为空/],
    ['经营类型', '不识别的经营类型', /取值无法识别/],
    ['退款率（支付时间）', 'abc', /不是有效比率/],
    ['成交金额', { formula: '1+1', result: 2 }, /金额无效/]
  ];
  for (const [field, value, pattern] of cases) {
    const w = fixture(); cell(w, field).value = value;
    assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), pattern);
  }
  const w = fixture(); w.getWorksheet('收支分析').addRow(w.getWorksheet('收支分析').getRow(2).values.slice(1));
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /重复日期与维度组合/);
});
test('空工作簿、只有表头、缺少全店汇总行均拒绝', () => {
  assert.throws(() => report.parseWorkbook(new ExcelJS.Workbook(), '测试.xlsx'), /工作簿为空/);
  let w = fixture(); w.getWorksheet('收支分析').spliceRows(2, 1);
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /没有有效记录/);
  w = fixture(); cell(w, '经营类型').value = '自营';
  assert.throws(() => report.parseWorkbook(w, '测试.xlsx'), /缺少部分日期/);
});
test('新增工作表和字段保留，空行不计入记录，原始值不丢失', () => {
  const w = fixture(), sheet = w.getWorksheet('收支分析');
  sheet.getCell(1, 44).value = '新增字段'; sheet.getCell(2, 44).value = '<img src=x onerror=alert(1)>';
  sheet.getCell(5, 1).value = null;
  const extra = w.addWorksheet('附加表'); extra.addRow(['备注']); extra.addRow(['仅测试']);
  const result = report.parseWorkbook(w, '测试.xlsx');
  assert.equal(result.sheets.length, 3);
  assert.equal(result.totalRows, 3);
  assert.equal(result.sheets[0].rows[0].values[43], '<img src=x onerror=alert(1)>');
  assert.equal(result.sheets[0].rows[0].sourceRow, 2);
  assert.equal(result.warnings.length, 2);
});
test('文件名日期不符会提示，使用实际记录日期', () => {
  const result = report.parseWorkbook(fixture(), '报表-20240102-20240103.xlsx');
  assert.deepEqual(result.dates, ['2024-01-01']);
  assert.match(result.warnings[0], /文件名日期与记录日期不一致/);
});
test('ZIP 结构预检查接受标准 xlsx，拒绝伪装及截断文件', async () => {
  const buffer = await fixture().xlsx.writeBuffer();
  const raw = Uint8Array.from(buffer).buffer;
  assert.doesNotThrow(() => report.inspectArchive(raw));
  assert.throws(() => report.inspectArchive(new TextEncoder().encode('not excel').buffer), /不是有效/);
  assert.throws(() => report.inspectArchive(raw.slice(0, raw.byteLength - 10)), /不完整或已损坏/);
});

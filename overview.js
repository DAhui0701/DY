/* 总览计算只读取同一份导入结果；金额/比率用 BigInt 计算，图形坐标才转 Number。 */
(function (root) {
  'use strict';
  const R = typeof module !== 'undefined' && module.exports ? require('./report.js') : root.ReportImport;
  const FIELD = { sales: '成交金额', live: '成交金额' };
  const dateOf = value => { try { return R.date(value); } catch { return null; } };
  const indexOf = sheet => Object.fromEntries((sheet?.headers || []).map((name, i) => [name, i]));
  const valueOf = (row, indices, field) => String(row.values[indices[field]] ?? '').trim();
  function choices(report) {
    const main = report.sheets.find(s => s.name === '收支分析');
    const account = report.sheets.find(s => s.name === '账号收支详情');
    const i = indexOf(main), a = indexOf(account);
    const distinct = field => [...new Set((main?.rows || []).map(row => valueOf(row, i, field)).filter(Boolean))];
    const accounts = new Map();
    for (const row of account?.rows || []) {
      const id = valueOf(row, a, '抖音号');
      if (id) accounts.set(id, valueOf(row, a, '账号名称') + '（' + id + '）');
    }
    return { business: distinct('经营类型'), carrier: distinct('载体类型'), period: distinct('推广时段'), accounts: [...accounts].map(([id, label]) => ({ id, label })) };
  }
  function defaults(report) {
    const c = choices(report);
    const dates = report.sheets.find(s => s.name === '收支分析')?.dates || [];
    return { start: dates[0] || '', end: dates.at(-1) || '', source: 'store', business: c.business.includes('全店') ? '全店' : c.business[0], carrier: c.carrier.includes('全部') ? '全部' : c.carrier[0], period: c.period.includes('不限') ? '不限' : c.period[0], account: c.accounts[0]?.id || '', spend: '店铺被投', refund: '支付时间' };
  }
  function calendar(start, end) {
    const first = dateOf(start), last = dateOf(end);
    if (!first || !last || first > last) return [];
    const days = [];
    for (let time = Date.parse(first); time <= Date.parse(last); time += 86400000) days.push(new Date(time).toISOString().slice(0, 10));
    return days;
  }
  function ratio(numerator, denominator, percent = false) {
    if (numerator === null || denominator === null || denominator <= 0n) return null;
    const negative = numerator < 0n;
    const absolute = negative ? -numerator : numerator;
    const scale = percent ? 10000n : 100n;
    const rounded = (absolute * scale * 2n + denominator) / (2n * denominator);
    return (negative ? '-' : '') + (rounded / 100n) + '.' + (rounded % 100n).toString().padStart(2, '0');
  }
  function groupRows(sheet, filters, live = false) {
    const indices = indexOf(sheet), groups = new Map();
    const source = filters.source === 'account' ? 'account' : 'store';
    for (const row of sheet?.rows || []) {
      if (source === 'account') {
        if (valueOf(row, indices, '抖音号') !== filters.account) continue;
      } else if (valueOf(row, indices, '经营类型') !== filters.business || valueOf(row, indices, '载体类型') !== (live ? '直播' : filters.carrier) || valueOf(row, indices, '推广时段') !== filters.period) continue;
      const day = dateOf(row.values[indices['日期']]);
      if (!day || day < filters.start || day > filters.end) continue;
      if (!groups.has(day)) groups.set(day, []);
      groups.get(day).push(row);
    }
    return { indices, groups };
  }
  function read(group, day, field) {
    const rows = group.groups.get(day) || [];
    if (!rows.length) return { value: null, reason: '该日无匹配记录（缺失，不是 0）' };
    if (rows.length !== 1) return { value: null, reason: '同日同范围存在多条记录，无法排除重复累计' };
    if (!(field in group.indices)) return { value: null, reason: '缺少字段：' + field };
    try { return { value: R.cents(rows[0].values[group.indices[field]]), reason: '' }; }
    catch { return { value: null, reason: '字段缺失或无效：' + field + '（不当作 0）' }; }
  }
  function total(days, key) {
    if (!days.length) return { value: null, reason: '没有匹配的日期或记录' };
    const missing = days.filter(day => day[key] === null);
    if (missing.length) return { value: null, reason: missing.length + ' 天数据缺失、重复或字段无效，不能作为完整区间合计' };
    return { value: days.reduce((sum, day) => sum + day[key], 0n), reason: '' };
  }
  function calculate(report, filters) {
    const f = { ...defaults(report), ...filters };
    const days = calendar(f.start, f.end);
    const sourceName = f.source === 'account' ? '账号收支详情' : '收支分析';
    const sheet = report.sheets.find(s => s.name === sourceName);
    const spendField = '投放消耗（' + f.spend + '）';
    const refundField = '成交退款金额（' + f.refund + '）';
    const scope = f.source === 'account' ? '账号=' + f.account + '；账号报表全部载体口径，不继承店铺维度筛选' : '经营类型=' + f.business + '；载体类型=' + f.carrier + '；推广时段=' + f.period;
    const scopeText = f.start + ' 至 ' + f.end + '；' + scope + '；消耗口径=' + f.spend + '；退款口径=' + f.refund;
    const group = groupRows(sheet, f);
    const liveGroup = f.source === 'store' ? groupRows(sheet, f, true) : null;
    const liveIssue = f.source === 'account' ? '账号表没有载体类型字段，不能识别直播成交，未使用全店数据代替' : !['全部', '直播'].includes(f.carrier) ? '当前选择非直播载体，直播指标与当前范围不匹配' : '';
    const daily = days.map(day => {
      const result = { date: day, rows: (group.groups.get(day) || []).map(row => row.sourceRow) };
      const issues = [];
      for (const [key, field] of [['sales', FIELD.sales], ['spend', spendField], ['refund', refundField]]) {
        const parsed = read(group, day, field); result[key] = parsed.value;
        if (parsed.reason) issues.push(field + '：' + parsed.reason);
      }
      const live = liveIssue ? { value: null, reason: liveIssue } : read(liveGroup, day, FIELD.live);
      result.live = live.value;
      result.liveRows = liveGroup?.groups.get(day)?.map(row => row.sourceRow) || [];
      if (live.reason) issues.push('直播成交：' + live.reason);
      result.issues = [...new Set(issues)];
      result.refundRatio = f.refund === '支付时间' ? ratio(result.refund, result.sales, true) : null;
      result.overallRatio = ratio(result.sales, result.spend);
      if (result.sales !== null && result.sales <= 0n) result.issues.push('退款金额占比无法计算：成交金额分母为 0 或负数');
      if (result.spend !== null && result.spend <= 0n) result.issues.push('整体成交消耗比无法计算：消耗分母为 0 或负数');
      if (f.refund !== '支付时间') result.issues.push('当前为退款发生时间口径，不计算本期退款率');
      return result;
    });
    const summed = Object.fromEntries(['sales', 'live', 'spend', 'refund'].map(key => [key, total(daily, key)]));
    if (liveIssue) summed.live.reason = liveIssue;
    const dateError = !days.length ? '日期范围无效：请选择有效日期，且开始日期不能晚于结束日期' : '';
    const hasRows = daily.some(day => day.rows.length);
    const snapshot = '平台数据截至时间：报表未提供；仅代表本次导出快照。后续退款可能改变支付时间口径占比。本地导入时间不是平台数据截至时间。';
    const common = { source: sourceName, scope: scopeText, aggregation: '每个日期只取一条所选维度记录，跨日期汇总；不叠加汇总与细分、不跨工作表累计。金额按整数分计算，显示时四舍五入到两位。', snapshot };
    const metric = (id, label, unit, fields, formula, raw, notes, overrideReason) => ({ ...common, id, label, unit, fields, formula, value: raw === null ? null : R.formatMoney(raw), reason: raw === null ? (overrideReason || summed[id]?.reason || '缺少可用数据') : '', notes });
    const metrics = [
      metric('sales', '成交金额', '元', ['成交金额'], '同范围每日成交金额之和', summed.sales.value, '含平台原表成交口径，不代表利润。'),
      metric('live', '直播成交金额', '元', ['成交金额', '载体类型'], '同日期、经营类型、推广时段，载体类型=直播的成交金额之和', summed.live.value, '载体选“全部”时单独读取其直播子集；选“直播”时等于当前成交金额。非直播载体或账号视图无法按同一范围计算。', liveIssue),
      metric('spend', '投放消耗（' + f.spend + '）', '元', [spendField], '同范围每日“' + spendField + '”之和', summed.spend.value, '仅使用已选消耗口径，不将店铺被投与店铺绑定相加。'),
      metric('refund', '成交退款金额（' + f.refund + '）', '元', [refundField], '同范围每日“' + refundField + '”之和', summed.refund.value, f.refund === '支付时间' ? snapshot : '按退款发生时间统计，可能对应更早支付的订单，不与本期成交额相除称为本期退款率。')
    ];
    metrics[1].scope = scopeText + (f.source === 'store' && !liveIssue ? '；本项载体固定=直播' : '');
    const refundRatio = f.refund === '支付时间' ? ratio(summed.refund.value, summed.sales.value, true) : null;
    const overallRatio = ratio(summed.sales.value, summed.spend.value);
    function ratioMetric(id, label, unit, fields, formula, value, reason, notes) {
      return { ...common, id, label, unit, fields, formula, value, reason: value === null ? reason : '', notes, aggregation: ['conversion', 'perCustomer', 'adRatio'].includes(id) ? '本报表缺少可靠口径，不执行人数或广告归因汇总；不跨日累加人数替代区间去重。' : common.aggregation + '先汇总同范围金额，再求比值；不平均每日比率。' };
    }
    metrics.push(
      ratioMetric('refundRatio', '成交退款金额占比（支付时间）', '%', ['成交退款金额（支付时间）', '成交金额'], '同范围成交退款金额（支付时间）÷成交金额×100%', refundRatio, f.refund !== '支付时间' ? '当前选择退款时间口径，不能除以本期成交额称为本期退款率' : summed.sales.value !== null && summed.sales.value <= 0n ? '成交金额分母为 0 或负数' : '成交或退款数据缺失，无法给出完整期间占比', snapshot + '；不使用原表每日退款率的简单平均。'),
      ratioMetric('overallRatio', '整体成交消耗比', '倍', ['成交金额', spendField], '同范围成交金额÷' + spendField, overallRatio, summed.spend.value !== null && summed.spend.value <= 0n ? '投放消耗分母为 0 或负数' : '成交或消耗数据缺失，无法计算', '包含自然成交，不等于广告归因投产比；不是利润。'),
      ratioMetric('conversion', '支付转化率', '%', ['支付人数（缺失）', '观看人数或访问人数（缺失）'], '同范围去重支付人数÷对应去重观看或访问人数×100%', null, '缺少支付人数、观看/访问人数及区间去重依据；不能用订单数、人次或跨日人数相加替代', '本文件不能确定观看转化或访问转化来源，故不推断具体转化口径。'),
      ratioMetric('perCustomer', '人均客单价', '元/人', ['用户支付金额', '支付人数（缺失）'], '同范围支付金额÷区间去重支付人数', null, '缺少支付人数及区间去重依据，无法计算；不将订单均价当作人均客单价', '即使将来补充逐日人数，也不能直接相加作为区间去重人数。'),
      ratioMetric('adRatio', '广告归因投产比', '倍', ['广告归因成交金额（缺失）', '对应广告消耗（匹配关系未提供）'], '广告归因成交金额÷对应广告消耗', null, '缺少广告归因成交金额和对应消耗的匹配口径，不能使用全店成交或原表“投放效率”替代', '整体成交消耗比包含自然成交，与广告归因投产比不同。')
    );
    const notices = [];
    if (dateError) notices.push(dateError);
    else if (!hasRows) notices.push('所选范围没有匹配记录；缺失不是零，请调整日期或筛选条件。');
    const missingDays = daily.filter(day => !day.rows.length).length;
    if (missingDays && hasRows) notices.push(missingDays + ' 天无匹配记录，趋势断开，不补零；受影响的区间指标显示“无法计算”。');
    if (f.source === 'account') notices.push('当前为单个账号口径，不能按经营类型、载体或推广时段筛选，也不回退到全店数据。');
    if (liveIssue) notices.push(liveIssue);
    return { filters: f, scope: scopeText, source: sourceName, spendField, refundField, metrics, daily, notices, hasRows, snapshot, hints: hints(daily, summed, f, refundRatio) };
  }
  function hints(daily, totals, filters, refundRatio) {
    const result = [];
    const add = (title, fact, question, action, verify) => result.push({ title, fact, question, action, verify });
    if (totals.refund.value !== null && totals.sales.value > 0n && filters.refund === '支付时间' && totals.refund.value > 0n) {
      const high = totals.refund.value * 2n >= totals.sales.value;
      add('核实支付时间口径退款', '所选范围支付时间口径成交退款金额占比为 ' + refundRatio + '%' + (high ? '，达到本页排查提示线 50%（并非行业标准）。' : '；仅为本次导出快照。'), '退款原因、商品或场次是否集中？平台截至时间未提供，后续退款会使占比变化。', '补充订单、退款原因、商品和场次明细，按同一支付批次核实；不要直接归因于主播、质量或履约。', '同支付批次更新后的退款金额占比、退款原因分布及商品/场次明细。');
    }
    const zeroLive = daily.filter(day => day.live === 0n);
    if (zeroLive.length) add('核实直播零成交日期', zeroLive.length + ' 天存在有效直播记录且成交金额为 0，例如 ' + zeroLive.slice(0, 3).map(day => day.date).join('、') + '。', '这些日期是否停播，或存在平台数据缺失、统计范围变化？不能直接判断直播效果下降。', '核对开播记录、导出时间与当前筛选，补充场次和时长数据。', '核实后的开播场次、直播时长及对应场次成交金额。');
    const noSpend = daily.filter(day => day.spend === 0n && day.sales > 0n);
    if (noSpend.length) add('核实零消耗有成交', noSpend.length + ' 天成交金额大于 0，但“投放消耗（' + filters.spend + '）”为 0；这些日期的成交消耗比无法计算。', '是否为自然成交，或广告账户归属、消耗口径不匹配？', '对照广告账户和平台消耗明细核实，不据此判定广告投产无限大。', '同口径广告消耗及具有匹配归因范围的成交金额。');
    const missing = daily.filter(day => !day.rows.length);
    if (missing.length) add('补齐缺失日期', missing.length + ' 天没有匹配记录，未将其当作 0。', '是漏导、当前维度无数据，还是日期范围选到了文件之外？', '核对导出范围和当前筛选，必要时重新导出完整报表。', '所选日期覆盖率与完整期间金额核对结果。');
    if (filters.refund === '退款时间' && totals.refund.value !== null && totals.refund.value > 0n) add('核实退款发生时间', '当前展示的是退款时间口径金额，而非本期支付订单的最终退款。', '这些退款对应哪些支付日期？', '补充退款与原支付订单的关联记录，区分支付批次。', '原支付批次的退款占比及退款发生时间分布。');
    return result;
  }
  const api = { choices, defaults, calculate, ratio, calendar };
  root.Overview = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window === 'undefined' ? globalThis : window);

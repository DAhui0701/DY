/* 报表读取与校验规则；不访问网络、不保存文件、不推测平台指标公式。 */
(function (root) {
  'use strict';
  const MONEY_FIELDS = [
    '结算金额', '7日结算金额', '14日结算金额', '结算金额（退款时间）',
    '投放消耗（店铺被投）', '投放消耗（店铺绑定）', '成交金额', '用户支付金额',
    '预售定金', '智能优惠券金额', '电商平台补贴金额', '达人补贴金额',
    '成交退款金额（支付时间）', '退款金额（支付时间）', '智能优惠券退款（支付时间）',
    '平台补贴退款（支付时间）', '达人补贴退款（支付时间）', '成交退款金额（退款时间）',
    '退款金额（退款时间）', '智能优惠券退款+平台补贴退款（退款时间）',
    '达人补贴退款（退款时间）', '支出金额', '平台佣金（预估口径）',
    '达人佣金（预估口径）', '免佣金额汇总', '免佣金额明细'
  ];
  const RATE_FIELDS = [
    '投放费比（店铺被投）', '投放费比（剔除退款、店铺被投）',
    '投放费比（店铺绑定）', '投放费比（剔除退款、店铺绑定）',
    '综合费比（店铺被投）', '综合费比（剔除退款、店铺被投）',
    '综合费比（店铺绑定）', '综合费比（剔除退款、店铺绑定）',
    '投放效率（店铺被投）', '投放效率（店铺绑定）',
    '全店效率（店铺被投）', '全店效率（店铺绑定）', '退款率（支付时间）'
  ];
  const PROFILES = {
    '收支分析': ['日期', '经营类型', '载体类型', '推广时段'],
    '账号收支详情': ['日期', '账号名称', '账号类型', '抖音号']
  };
  const LIMITS = { bytes: 10 * 1024 * 1024, expanded: 40 * 1024 * 1024, sheets: 20, rows: 25000, columns: 200, cells: 500000 };

  // 在解压前检查 ZIP 中央目录，拒绝超大文件、加密包及 ZIP64。
  function inspectArchive(buffer) {
    const view = new DataView(buffer);
    if (view.byteLength < 22 || view.getUint32(0, true) !== 0x04034b50) {
      throw new Error('文件不是有效的 .xlsx 工作簿，请从平台重新导出，不要只修改文件后缀。');
    }
    let end = -1;
    for (let i = view.byteLength - 22; i >= Math.max(0, view.byteLength - 65557); i--) {
      if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === view.byteLength) { end = i; break; }
    }
    if (end < 0) throw new Error('Excel 文件不完整或已损坏，请重新导出。');
    const count = view.getUint16(end + 10, true);
    const directorySize = view.getUint32(end + 12, true);
    let offset = view.getUint32(end + 16, true);
    const boundary = offset + directorySize;
    if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) || count === 65535 || boundary !== end) {
      throw new Error('暂不支持分卷或 ZIP64 工作簿，请导出较小的标准 .xlsx 文件。');
    }
    let expanded = 0;
    for (let n = 0; n < count; n++) {
      if (offset + 46 > boundary || view.getUint32(offset, true) !== 0x02014b50) throw new Error('Excel 压缩结构损坏。');
      const size = view.getUint32(offset + 24, true);
      if ((view.getUint16(offset + 8, true) & 1) || size === 0xffffffff) throw new Error('暂不支持加密或 ZIP64 工作簿。');
      expanded += size;
      if (expanded > LIMITS.expanded) throw new Error('Excel 解压后超过 40 MB，请缩小导出日期范围。');
      offset += 46 + view.getUint16(offset + 28, true) + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    }
    if (offset !== boundary) throw new Error('Excel 压缩结构损坏。');
  }

  function text(value) {
    if (value === null || value === undefined) return '';
    if (value instanceof Date) return Number.isNaN(value.getTime()) ? '无效日期' : value.toISOString().slice(0, 10);
    if (typeof value !== 'object') return String(value);
    if (value.formula || value.sharedFormula) return '=' + (value.formula || value.sharedFormula);
    if (value.richText) return value.richText.map(part => part.text).join('');
    if (value.text !== undefined) return String(value.text);
    if (value.error) return value.error;
    return '[不支持的单元格类型]';
  }

  function date(value) {
    let source = text(value).trim();
    if (/^\d{8}$/.test(source)) source = source.slice(0, 4) + '-' + source.slice(4, 6) + '-' + source.slice(6);
    const match = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(source);
    if (!match) throw new Error('日期格式异常，请使用 YYYYMMDD 或 YYYY-MM-DD。');
    const year = Number(match[1]), month = Number(match[2]), day = Number(match[3]);
    const parsed = new Date(Date.UTC(year, month - 1, day));
    const result = parsed.toISOString().slice(0, 10);
    const today = new Date();
    const todayText = [today.getFullYear(), String(today.getMonth() + 1).padStart(2, '0'), String(today.getDate()).padStart(2, '0')].join('-');
    if (year < 2000 || parsed.getUTCFullYear() !== year || parsed.getUTCMonth() !== month - 1 || parsed.getUTCDate() !== day || result > todayText) {
      throw new Error('日期异常：不是有效的经营日期，或晚于今天。');
    }
    return result;
  }

  function cents(value) {
    const source = text(value).trim();
    if (!/^[+-]?\d+(?:\.\d{1,2})?$/.test(source)) throw new Error('金额无效或缺失，应为最多两位小数的数字；“-”不能视为 0。');
    const negative = source.startsWith('-');
    const [whole, fraction = ''] = source.replace(/^[+-]/, '').split('.');
    const result = BigInt(whole) * 100n + BigInt(fraction.padEnd(2, '0'));
    if (result > 9007199254740991n) throw new Error('金额过大，超出本版本可安全处理的范围。');
    return negative ? -result : result;
  }

  function formatMoney(value) {
    const negative = value < 0n;
    const absolute = negative ? -value : value;
    return (negative ? '-' : '') + (absolute / 100n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',') + '.' + (absolute % 100n).toString().padStart(2, '0');
  }

  function parseWorkbook(workbook, fileName) {
    if (!workbook.worksheets.length) throw new Error('工作簿为空，没有可读取的工作表。');
    if (workbook.worksheets.length > LIMITS.sheets) throw new Error('工作表超过 20 个，请减少后重试。');
    const missingSheets = Object.keys(PROFILES).filter(name => !workbook.getWorksheet(name));
    if (missingSheets.length) throw new Error('缺少工作表：' + missingSheets.join('、') + '。请使用完整的罗盘收支分析报表。');
    const errors = [], warnings = [], sheets = [];
    let issueCount = 0, missingRatios = 0, totalRows = 0, totalCells = 0;
    const addError = message => { issueCount++; if (errors.length < 12) errors.push(message); };
    for (const worksheet of workbook.worksheets) {
      if (worksheet.rowCount > LIMITS.rows || worksheet.columnCount > LIMITS.columns) throw new Error('工作表超过 25000 行或 200 列，请缩小导出范围。');
      totalCells += worksheet.rowCount * worksheet.columnCount;
      if (totalCells > LIMITS.cells) throw new Error('单元格超过 50 万个，请缩小导出范围。');
      const width = worksheet.columnCount;
      const headers = Array.from({ length: width }, (_, i) => text(worksheet.getRow(1).getCell(i + 1).value).trim());
      const known = PROFILES[worksheet.name];
      if (!width) {
        if (known) addError('“' + worksheet.name + '”是空工作表。');
        sheets.push({ name: worksheet.name, headers: [], rows: [], dates: [], known: false });
        continue;
      }
      if (headers.some(header => !header) || new Set(headers).size !== headers.length) {
        addError('“' + worksheet.name + '”第一行有空白或重复表头；本版本要求第一行为单行完整表头。');
        continue;
      }
      const required = known ? [...known, ...MONEY_FIELDS, ...RATE_FIELDS] : [];
      const missing = required.filter(field => !headers.includes(field));
      if (missing.length) { addError('“' + worksheet.name + '”缺少字段：' + missing.join('、')); continue; }
      if (!known) warnings.push('“' + worksheet.name + '”不是已识别的报表类型，仅原样展示，不做金额核对。');
      if (known && headers.some(header => !required.includes(header))) warnings.push('“' + worksheet.name + '”包含新增字段，已保留展示；新增字段不参与校验和核对。');
      const indices = Object.fromEntries(headers.map((header, i) => [header, i]));
      const rows = [], dates = new Set(), grains = new Set();
      for (let r = 2; r <= worksheet.rowCount; r++) {
        const row = worksheet.getRow(r);
        const values = headers.map((_, i) => text(row.getCell(i + 1).value));
        if (!values.some(value => value.trim())) continue;
        const location = '“' + worksheet.name + '”第 ' + r + ' 行';
        rows.push({ sourceRow: r, values });
        if (!known) continue;
        let normalizedDate;
        try { normalizedDate = date(row.getCell(indices['日期'] + 1).value); dates.add(normalizedDate); }
        catch (error) { addError(location + '，日期：' + error.message); }
        for (const field of known.slice(1)) if (!values[indices[field]].trim()) addError(location + '，“' + field + '”不能为空。');
        if (worksheet.name === '收支分析') {
          const allowed = {
            '经营类型': ['全店', '自营', '合作'],
            '载体类型': ['全部', '直播', '短视频', '商品卡', '图文', '其他'],
            '推广时段': ['不限', '全域+乘方投放时段', '标准+品牌投放', '非投放时段']
          };
          for (const [field, options] of Object.entries(allowed)) {
            if (!options.includes(values[indices[field]].trim())) addError(location + '，“' + field + '”取值无法识别，请确认报表版本。');
          }
        }
        const grain = JSON.stringify(known.map(field => field === '日期' ? normalizedDate : values[indices[field]].trim()));
        if (grains.has(grain)) addError(location + '，存在重复日期与维度组合，无法确定唯一记录。');
        grains.add(grain);
        for (const field of MONEY_FIELDS) {
          try { cents(values[indices[field]]); } catch (error) { addError(location + '，“' + field + '”：' + error.message); }
        }
        for (const field of RATE_FIELDS) {
          const value = values[indices[field]].trim();
          if (value === '-') { missingRatios++; continue; }
          if (!/^[+-]?\d+(?:\.\d+)?$/.test(value) || !Number.isFinite(Number(value))) addError(location + '，“' + field + '”不是有效比率或效率值；未提供值请保留“-”。');
        }
      }
      if (known && !rows.length) addError('“' + worksheet.name + '”只有表头，没有有效记录。');
      totalRows += rows.length;
      sheets.push({ name: worksheet.name, headers, rows, dates: [...dates].sort(), known: Boolean(known) });
    }
    if (issueCount) throw new Error('导入未完成，共发现 ' + issueCount + ' 处问题（最多展示前 12 处）：\n' + errors.join('\n'));
    const main = sheets.find(sheet => sheet.name === '收支分析');
    const mainIndices = Object.fromEntries(main.headers.map((header, i) => [header, i]));
    const selected = main.rows.filter(row => row.values[mainIndices['经营类型']].trim() === '全店' && row.values[mainIndices['载体类型']].trim() === '全部' && row.values[mainIndices['推广时段']].trim() === '不限');
    if (selected.length !== main.dates.length) throw new Error('“收支分析”缺少部分日期的“全店＋全部＋不限”汇总行，无法安全核对金额。');
    const dateSets = sheets.filter(sheet => sheet.known).map(sheet => sheet.dates.join(','));
    if (new Set(dateSets).size > 1) warnings.push('两个已识别工作表的日期范围或日期覆盖不一致，请确认是否为同次完整导出。');
    const dates = [...new Set(sheets.flatMap(sheet => sheet.dates))].sort();
    if (dates.length && (Date.parse(dates.at(-1)) - Date.parse(dates[0])) / 86400000 + 1 !== dates.length) warnings.push('记录日期不连续；缺少的日期不补零、不推测。');
    const fileDates = /(\d{8})[-_](\d{8})/.exec(fileName);
    if (fileDates) {
      try { if (date(fileDates[1]) !== dates[0] || date(fileDates[2]) !== dates.at(-1)) warnings.push('文件名日期与记录日期不一致；本页以实际记录为准。'); }
      catch { warnings.push('文件名包含异常日期；本页以实际记录为准。'); }
    }
    const totals = {};
    for (const field of ['成交金额', '结算金额', '支出金额']) totals[field] = selected.reduce((sum, row) => sum + cents(row.values[mainIndices[field]]), 0n);
    return { fileName, sheets, totalRows, dates, selectedCount: selected.length, totals, missingRatios, warnings };
  }
  const api = { parseWorkbook, inspectArchive, cents, formatMoney, date, LIMITS };
  root.ReportImport = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window === 'undefined' ? globalThis : window);

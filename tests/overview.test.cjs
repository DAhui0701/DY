// 全部为人工测试输入，不包含真实报表或经营数据，不用于页面指标填充。
const test = require('node:test');
const assert = require('node:assert/strict');
const O = require('../overview.js');
const headers = ['日期','经营类型','载体类型','推广时段','成交金额','投放消耗（店铺被投）','投放消耗（店铺绑定）','成交退款金额（支付时间）','成交退款金额（退款时间）'];
function fixture() {
  const rows = [
    ['20240101','全店','全部','不限','100.10','10.00','0.00','60.00','120.00'],
    ['20240101','全店','直播','不限','80.10','8.00','0.00','40.00','90.00'],
    ['20240101','自营','全部','不限','90.10','9.00','0.00','50.00','100.00'],
    ['20240101','自营','直播','不限','70.10','7.00','0.00','30.00','80.00'],
    ['20240101','全店','商品卡','不限','20.00','2.00','0.00','20.00','30.00'],
    ['20240102','全店','全部','不限','0.00','0.00','0.00','0.00','0.00'],
    ['20240102','全店','直播','不限','0.00','0.00','0.00','0.00','0.00']
  ].map((values,i)=>({values,sourceRow:i+2}));
  const accountHeaders = ['日期','账号名称','账号类型','抖音号',...headers.slice(4)];
  return {fileName:'人工测试.xlsx',dates:['2024-01-01','2024-01-02'],sheets:[
    {name:'收支分析',headers:[...headers],dates:['2024-01-01','2024-01-02'],rows},
    {name:'账号收支详情',headers:accountHeaders,dates:['2024-01-01','2024-01-02'],rows:[
      {sourceRow:2,values:['20240101','测试账号','自营','TEST_ONLY','33.33','3.00','0.00','10.00','15.00']},
      {sourceRow:3,values:['20240102','测试账号','自营','TEST_ONLY','0.00','0.00','0.00','0.00','0.00']}
    ]}
  ]};
}
const card = (r,id)=>r.metrics.find(m=>m.id===id);
test('只读取目标维度，不把汇总与细分或账号记录重复相加',()=>{
 const r=O.calculate(fixture());
 assert.equal(card(r,'sales').value,'100.10');
 assert.equal(card(r,'live').value,'80.10');
 assert.equal(card(r,'spend').value,'10.00');
 assert.equal(card(r,'refundRatio').value,'59.94');
 assert.equal(card(r,'overallRatio').value,'10.01');
 assert.deepEqual(r.daily[0].rows,[2]);
 assert.deepEqual(r.daily[0].liveRows,[3]);
});
test('单日、经营类型、直播、非直播筛选遵循同范围口径',()=>{
 let r=O.calculate(fixture(),{start:'2024-01-01',end:'2024-01-01',business:'自营',carrier:'直播'});
 assert.equal(card(r,'sales').value,'70.10'); assert.equal(card(r,'live').value,'70.10');
 r=O.calculate(fixture(),{start:'2024-01-01',end:'2024-01-01',carrier:'商品卡'});
 assert.equal(card(r,'sales').value,'20.00'); assert.equal(card(r,'live').value,null);
 assert.match(card(r,'live').reason,/非直播/);
});
test('单账号不回退全店，不虚构载体划分或跨账号合计',()=>{
 const r=O.calculate(fixture(),{source:'account',account:'TEST_ONLY',carrier:'直播',business:'合作'});
 assert.equal(card(r,'sales').value,'33.33'); assert.equal(card(r,'live').value,null);
 assert.equal(card(r,'sales').source,'账号收支详情'); assert.match(card(r,'live').reason,/没有载体/);
});
test('消耗分母为零时无法计算，不将自然成交当广告归因投产',()=>{
 const r=O.calculate(fixture(),{spend:'店铺绑定'});
 assert.equal(card(r,'spend').value,'0.00'); assert.equal(card(r,'overallRatio').value,null);
 assert.match(card(r,'overallRatio').reason,/分母为 0/);
 assert.equal(card(r,'adRatio').value,null); assert.match(card(r,'adRatio').reason,/缺少广告归因/);
 assert.match(card(r,'overallRatio').notes,/包含自然成交/);
});
test('退款发生时间金额可展示，但不能冒充本期支付批次退款率',()=>{
 const r=O.calculate(fixture(),{refund:'退款时间'});
 assert.equal(card(r,'refund').value,'120.00'); assert.equal(card(r,'refundRatio').value,null);
 assert.match(card(r,'refundRatio').reason,/退款时间口径/);
 assert.equal(r.daily[0].refundRatio,null);
});
test('真实零值保留；缺日不填零、不输出部分区间合计',()=>{
 const report=fixture();report.sheets[0].rows=report.sheets[0].rows.filter(row=>row.values[0]!=='20240102');
 const r=O.calculate(report,{end:'2024-01-03'});
 assert.equal(r.daily.length,3);assert.equal(r.daily[1].sales,null);assert.equal(card(r,'sales').value,null);
 const zero=O.calculate(fixture(),{start:'2024-01-02',end:'2024-01-02'});
 assert.equal(zero.daily[0].sales,0n);assert.equal(card(zero,'sales').value,'0.00');
 assert.equal(card(zero,'refundRatio').value,null);
});
test('缺字段、缺值、非法值或重复日期不默认为零',()=>{
 for (const kind of ['field','value','invalid','duplicate']) {
  const report=fixture(),sheet=report.sheets[0];
  if(kind==='field')sheet.headers[4]='不是成交金额';
  if(kind==='value')sheet.rows[0].values[4]='';
  if(kind==='invalid')sheet.rows[0].values[4]='NaN';
  if(kind==='duplicate')sheet.rows.push({...sheet.rows[0]});
  const r=O.calculate(report);assert.equal(card(r,'sales').value,null);assert.equal(r.daily[0].sales,null);
 }
});
test('没有人数/区间去重依据时，两项人数指标均不可计算',()=>{
 const r=O.calculate(fixture());
 for(const id of ['conversion','perCustomer']) {assert.equal(card(r,id).value,null);assert.match(card(r,id).reason,/去重/);}
 assert.match(r.snapshot,/截至时间：报表未提供/);
});
test('无匹配、日期倒置时空状态；替换输入不残留旧结果',()=>{
 const report=fixture();
 const empty=O.calculate(report,{account:'MISSING',source:'account'});assert.equal(empty.hasRows,false);assert.equal(card(empty,'sales').value,null);
 const reversed=O.calculate(report,{start:'2024-01-02',end:'2024-01-01'});assert.equal(reversed.daily.length,0);assert.match(reversed.notices[0],/日期范围无效/);
 assert.equal(card(O.calculate(report),'sales').value,'100.10');
 report.sheets[0].rows[0].values[4]='0.10';assert.equal(card(O.calculate(report),'sales').value,'0.10');
});
test('比率按总分子总分母计算并可靠四舍五入',()=>{
 assert.equal(O.ratio(1n,6n),'0.17');assert.equal(O.ratio(1n,6n,true),'16.67');
 assert.equal(O.ratio(1n,200n),'0.01');assert.equal(O.ratio(-1n,6n),'-0.17');
 assert.equal(O.ratio(100n,0n),null);assert.equal(O.ratio(null,10n),null);
});
test('规则提示包含事实、待核实问题、动作和验证项，不直接判断原因',()=>{
 const r=O.calculate(fixture());assert.ok(r.hints.length>=2);
 for(const h of r.hints) for(const key of ['fact','question','action','verify']) assert.ok(h[key]);
 assert.ok(r.hints.some(h=>/是否停播/.test(h.question)));
 assert.ok(r.hints.some(h=>/不要直接归因/.test(h.action)));
});

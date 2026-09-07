// 生成 HTML 周报对比页（数据源与图表配置见 report_data.js）
const fs = require('fs');
const { userRows, trafRows, totals, lastDayUsers, meta, optUsers, optTraf } = require('./report_data');

console.log('=== 用户数（周日均，万）===');
console.table(userRows);
console.log('全省: ', totals.prevU, '->', totals.currU,
  ' 增量', +(totals.currU - totals.prevU).toFixed(2),
  ' 增幅', ((totals.currU - totals.prevU) / totals.prevU * 100).toFixed(2) + '%');
console.log('=== 期末总用户数（两周最后一天，万）===');
console.log(`${lastDayUsers.prevDate}: ${lastDayUsers.prev} 万 -> ${lastDayUsers.currDate}: ${lastDayUsers.curr} 万` +
  `  增减额 ${lastDayUsers.delta >= 0 ? '+' : ''}${lastDayUsers.delta} 万`);
console.log('=== 流量（日均T）===');
console.table(trafRows);
console.log('全省: ', Math.round(totals.prevT), '->', Math.round(totals.currT),
  ' 增量', (totals.currT - totals.prevT).toFixed(1),
  ' 增幅', ((totals.currT - totals.prevT) / totals.prevT * 100).toFixed(2) + '%');

const echartsJs = fs.readFileSync('node_modules/echarts/dist/echarts.min.js', 'utf8');

const ld = lastDayUsers;
const ldCls = ld.delta >= 0 ? 'up' : 'down';
const ldSign = ld.delta >= 0 ? '+' : '';
const lastDayHtml = `<h2>附表三：期末总用户数（两周最后一天，单位：万）</h2>
<table><thead><tr><th>项目</th><th>日期</th><th>总用户数(万)</th></tr></thead><tbody>
<tr><td>上周最后一天</td><td>${ld.prevDate}</td><td>${ld.prev.toLocaleString()}</td></tr>
<tr><td>本周最后一天</td><td>${ld.currDate}</td><td>${ld.curr.toLocaleString()}</td></tr>
<tr><td>增减额</td><td>—</td><td class="${ldCls}">${ldSign}${ld.delta.toLocaleString()}</td></tr>
</tbody></table>`;

function tableHtml(title, rows, unit, extra) {
  const trs = rows.map(r => {
    const cls = r.delta >= 0 ? 'up' : 'down';
    const sign = r.delta >= 0 ? '+' : '';
    return `<tr><td>${r.city}</td><td>${r.prev.toLocaleString()}</td><td>${r.curr.toLocaleString()}</td>` +
      `<td class="${cls}">${sign}${r.delta.toLocaleString()}</td>` +
      `<td class="${cls}">${sign}${r.pct}%</td><td>${r.share}%</td>` +
      (extra ? `<td>${r[extra.key].toLocaleString()}</td>` : '') + '</tr>';
  }).join('');
  return `<h2>${title}</h2><table><thead><tr><th>地市</th><th>上周(${unit})</th><th>本周(${unit})</th>` +
    `<th>增量(${unit})</th><th>增幅</th><th>占比(本周)</th>` +
    (extra ? `<th>${extra.head}</th>` : '') + `</tr></thead><tbody>${trs}</tbody></table>`;
}

const html = `<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<title>广电分地市周报对比（${meta.range1} vs ${meta.range2}）</title>
<style>
  body{font-family:"Microsoft YaHei",sans-serif;margin:24px;color:#333}
  h1{font-size:20px} h2{font-size:16px;margin-top:32px}
  .chart{width:100%;height:520px;border:1px solid #eee;margin-top:8px}
  table{border-collapse:collapse;font-size:13px;margin-top:8px}
  th,td{border:1px solid #ccc;padding:5px 12px;text-align:right}
  th:first-child,td:first-child{text-align:left}
  thead{background:#f2f5f9}
  .up{color:#c0392b} .down{color:#27ae60}
  .note{color:#888;font-size:12px}
</style></head><body>
<h1>广电分地市周报数据对比（${meta.range1} → ${meta.range2}）</h1>
<p class="note">数据来源：${meta.file1}（${meta.range1}）与 ${meta.file2}（${meta.range2}）两份周报；用户数（单位：万）与流量（单位：T）均取各周7天日均值，保证口径一致。增量/增幅为本周相对上周，占比为该地市占全省比例（本周）。图表按增幅升序排列（最低在左、最高在右），柱状图浅色为上周、深色为本周。附表三为两周最后一天（期末）全省总用户数及增减额。月均DOU(GB) = 本周日均流量 ÷ 本周日均用户数 × 当月天数（1T按1024GB折算，月天数取本周多数天所在月份）。</p>
<h2>图表一：分地市周日均用户数（柱状）与增幅（折线）</h2>
<div id="chartUsers" class="chart"></div>
<h2>图表二：分地市日均流量（柱状）与增幅（折线）</h2>
<div id="chartTraf" class="chart"></div>
${tableHtml('附表一：分地市周日均用户数明细（万）', userRows, '万')}
${tableHtml('附表二：分地市日均流量明细（T）', trafRows, 'T', { head: '月均DOU(GB)', key: 'dou' })}
${lastDayHtml}
<script>${echartsJs}</script>
<script>
const optUsers = ${JSON.stringify(optUsers)};
const optTraf = ${JSON.stringify(optTraf)};
const c1 = echarts.init(document.getElementById('chartUsers'));
c1.setOption(optUsers);
const c2 = echarts.init(document.getElementById('chartTraf'));
c2.setOption(optTraf);
window.addEventListener('resize', () => { c1.resize(); c2.resize(); });
</script>
</body></html>`;

fs.writeFileSync(`${meta.outBase}.html`, html);
console.log(`HTML written: ${meta.outBase}.html`);

// 共享数据层：自动识别目录中最新两份周报，计算分地市用户数/流量的周日均、增量、增幅、占比，并生成图表配置
const XLSX = require('xlsx');
const fs = require('fs');

// ---------- 自动识别最新两份周报（文件名形如 gd_2026_07_31_*.xlsx，按日期排序取最新两份） ----------
const files = fs.readdirSync('.')
  .filter(f => /^gd_\d{4}_\d{2}_\d{2}_.*\.xlsx$/.test(f) && !f.startsWith('~$'))
  .sort();
if (files.length < 2) {
  console.error('错误：目录中周报文件不足两份（需要形如 gd_YYYY_MM_DD_*.xlsx 的文件）');
  process.exit(1);
}
const [F1, F2] = files.slice(-2); // F1=上周, F2=本周
console.log(`上周周报: ${F1}\n本周周报: ${F2}`);

// ---------- 读取右表（地市 + 每日数据），并取表头日期 ----------
function readSheetRight(file, sheet, colCount) {
  const wb = XLSX.readFile(file);
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheet], { header: 1, defval: null });
  const header = rows[1];
  const dataRows = rows.slice(2).filter(r => r[0]).map(r => {
    const startIdx = r.findIndex((v, i) => i > 0 && v === r[0]); // 右表地市列
    return { city: r[0], days: r.slice(startIdx + 1, startIdx + 1 + colCount) };
  });
  const cityColIdx = header.findIndex((v, i) => i > 0 && v === '地市');
  const dates = header.slice(cityColIdx + 1, cityColIdx + 1 + colCount); // 形如 "07月24日"
  return { dataRows, dates };
}

// "07月24日" -> "07-24"
function fmtDate(s) {
  const m = /(\d{2})月(\d{2})日/.exec(String(s));
  return m ? `${m[1]}-${m[2]}` : String(s);
}
function fmtRange(dates) {
  const a = fmtDate(dates[0]), b = fmtDate(dates[dates.length - 1]);
  const [am] = a.split('-');
  const [bm, bd] = b.split('-');
  return am === bm ? `${a}~${bd}` : `${a}~${b}`; // 同月 "07-24~30"，跨月 "07-31~08-06"
}

// 日均：只对非空（数值）单元格求平均，源表某天数据缺失时按实际天数计
const avg = a => {
  const v = a.filter(x => typeof x === 'number');
  return v.length ? v.reduce((s, x) => s + x, 0) / v.length : 0;
};

// ---------- 用户数与流量（两周各7天日均值对比，口径一致；用户数单位为万） ----------
const u1 = readSheetRight(F1, '地市用户数', 7);
const u2 = readSheetRight(F2, '地市用户数', 7);
const t1 = readSheetRight(F1, '流量', 7);
const t2 = readSheetRight(F2, '流量', 7);

const prevUsers = {}, currUsers = {}, prevTraf = {}, currTraf = {};
u1.dataRows.forEach(r => prevUsers[r.city] = avg(r.days) / 10000);
u2.dataRows.forEach(r => currUsers[r.city] = avg(r.days) / 10000);
t1.dataRows.forEach(r => prevTraf[r.city] = avg(r.days));
t2.dataRows.forEach(r => currTraf[r.city] = avg(r.days));

// ---------- 汇总计算 ----------
const cities = Object.keys(currUsers).filter(c => c !== '全省');
const totalU = currUsers['全省'];
const totalT = currTraf['全省'];

function buildRows(curr, prev, total, digits, deltaDigits = digits) {
  return cities.map(c => {
    const cur = curr[c], prv = prev[c];
    const delta = cur - prv;
    return {
      city: c,
      prev: +prv.toFixed(digits),
      curr: +cur.toFixed(digits),
      delta: +delta.toFixed(deltaDigits),
      pct: +(delta / prv * 100).toFixed(2),
      share: +(cur / total * 100).toFixed(2),
    };
  });
}
// 按增幅升序：增幅最低在最左、最高在最右
const userRows = buildRows(currUsers, prevUsers, totalU, 2).sort((a, b) => a.pct - b.pct);
// 流量：分地市流量取整数，增量保留1位小数
const trafRows = buildRows(currTraf, prevTraf, totalT, 0, 1).sort((a, b) => a.pct - b.pct);

const totals = {
  prevU: +prevUsers['全省'].toFixed(2), currU: +currUsers['全省'].toFixed(2),
  prevT: +prevTraf['全省'].toFixed(2), currT: +currTraf['全省'].toFixed(2),
};

// ---------- 期末总用户数：前后两周最后一天全省总用户数（万）及增减额 ----------
const lastDayUsers = {
  prevDate: fmtDate(u1.dates[u1.dates.length - 1]),
  currDate: fmtDate(u2.dates[u2.dates.length - 1]),
  prev: +(u1.dataRows.find(r => r.city === '全省').days[6] / 10000).toFixed(2),
  curr: +(u2.dataRows.find(r => r.city === '全省').days[6] / 10000).toFixed(2),
};
lastDayUsers.delta = +(lastDayUsers.curr - lastDayUsers.prev).toFixed(2);

// ---------- 周期标签（从两份周报表头日期自动生成） ----------
const range1 = fmtRange(u1.dates); // 上周，如 "07-24~30"
const range2 = fmtRange(u2.dates); // 本周，如 "07-31~08-06"

// 最新统计日期 = 本周周报最后一个统计日（年份取自文件名 gd_YYYY_...）
const year = Number(/gd_(\d{4})_/.exec(F2)[1]);
const lastDay = fmtDate(u2.dates[u2.dates.length - 1]); // "08-06"
const [mm, dd] = lastDay.split('-').map(Number);
const statDate = `${year}-${String(mm).padStart(2, '0')}-${String(dd).padStart(2, '0')}`;

// ISO 8601 年内第几周
function isoWeek(y, m, d) {
  const date = new Date(Date.UTC(y, m - 1, d));
  const dayNum = (date.getUTCDay() + 6) % 7;
  date.setUTCDate(date.getUTCDate() - dayNum + 3);
  const firstThursday = new Date(Date.UTC(date.getUTCFullYear(), 0, 4));
  const fDay = (firstThursday.getUTCDay() + 6) % 7;
  firstThursday.setUTCDate(firstThursday.getUTCDate() - fDay + 3);
  return 1 + Math.round((date - firstThursday) / (7 * 864e5));
}
const weekNo = isoWeek(year, mm, dd);

const meta = {
  file1: F1, file2: F2,
  range1, range2,
  prevLabel: `上周(${range1}日均)`,
  currLabel: `本周(${range2}日均)`,
  statDate, weekNo,
  outBase: `分地市周报对比_用户数与流量_${statDate}_第${weekNo}周`,
};

// ---------- 月均DOU（GB）：本周日均流量 ÷ 本周日均用户数 × 当月天数（1T=1024GB；月天数取本周多数天所在月份，即7天中间日所在月份，避免跨月周口径跳变） ----------
const midMonth = Number(/(\d{2})月\d{2}日/.exec(u2.dates[3])[1]); // 7天中第4天必属占多数的月份
const monthDays = new Date(year, midMonth, 0).getDate();
trafRows.forEach(r => {
  r.dou = +(currTraf[r.city] / currUsers[r.city] * monthDays * 1024 / 10000).toFixed(2);
});
totals.douT = +(totalT / totalU * monthDays * 1024 / 10000).toFixed(2);
meta.monthDays = monthDays;

// ---------- 图表配置 ----------
function chartOption(title, barName, prevLabel, currLabel, rows, barUnit, barDigits) {
  // 两轴留白：柱形压低至约 70% 高度，增幅折线整体抬到上部空白区，
  // 使折线数据标签位于折线上方且不遮挡柱形与柱标签；副轴刻度无参考意义，隐藏
  const maxBar = Math.max(...rows.map(r => Math.max(r.prev, r.curr)));
  const y1Max = Math.ceil(maxBar * 1.4);
  const pcts = rows.map(r => r.pct);
  const pMin = Math.min(...pcts), pMax = Math.max(...pcts);
  const span = Math.max(pMax - pMin, 1);
  const y2Min = +(pMin - span * 2.7).toFixed(1);
  const y2Max = +(pMax + span * 0.65).toFixed(1);
  return {
    title: { text: title, left: 'center' },
    tooltip: {
      trigger: 'axis',
      axisPointer: { type: 'shadow' },
      formatter: ps => {
        const i = ps[0].dataIndex, r = rows[i];
        const sign = r.delta >= 0 ? '+' : '';
        return `${r.city}<br/>${prevLabel}：${r.prev.toLocaleString()} ${barUnit}` +
          `<br/>${currLabel}：${r.curr.toLocaleString()} ${barUnit}` +
          `<br/>增量：${sign}${r.delta.toLocaleString()} ${barUnit}` +
          `<br/>增幅：${sign}${r.pct}%<br/>占比：${r.share}%`;
      }
    },
    legend: { bottom: 0, data: [prevLabel, currLabel, '增幅'] },
    grid: { left: 70, right: 70, top: 60, bottom: 60 },
    xAxis: { type: 'category', data: rows.map(r => r.city), axisLabel: { interval: 0, rotate: 30 } },
    yAxis: [
      { type: 'value', name: barUnit, max: y1Max, axisLabel: { formatter: v => v.toLocaleString() }, splitLine: { show: false } },
      { type: 'value', name: '增幅(%)', nameTextStyle: { fontSize: 10 }, min: y2Min, max: y2Max, axisLabel: { show: false }, axisLine: { show: false }, axisTick: { show: false }, splitLine: { show: false } }
    ],
    series: [
      {
        name: prevLabel, type: 'bar', data: rows.map(r => +r.prev.toFixed(barDigits)), barMaxWidth: 28,
        itemStyle: { color: '#9db8d2' },
        label: { show: true, position: 'top', formatter: p => rows[p.dataIndex].prev.toLocaleString(), fontSize: 9, color: '#6b8cae' }
      },
      {
        name: currLabel, type: 'bar', data: rows.map(r => +r.curr.toFixed(barDigits)), barMaxWidth: 28,
        itemStyle: { color: '#4e79a7' },
        label: { show: true, position: 'top', formatter: p => rows[p.dataIndex].curr.toLocaleString(), fontSize: 9 }
      },
      {
        name: '增幅', type: 'line', yAxisIndex: 1,
        // 逐点着色：与表格一致，正增幅红色、负增幅绿色
        data: rows.map(r => ({ value: +r.pct.toFixed(1), label: { color: r.pct >= 0 ? '#c0392b' : '#27ae60' } })),
        itemStyle: { color: '#e15759' }, symbolSize: 8,
        label: { show: true, position: 'top', distance: 8, fontSize: 10 },
        markLine: { silent: true, symbol: 'none', lineStyle: { type: 'dashed', color: '#999' }, data: [{ yAxis: 0 }] }
      }
    ]
  };
}

const optUsers = chartOption('广电分地市用户数（周日均）及周增幅（按增幅升序）', '周日均用户数', meta.prevLabel, meta.currLabel, userRows, '万', 1);
const optTraf = chartOption('广电分地市流量（日均）及周增幅（按增幅升序）', '日均流量', meta.prevLabel, meta.currLabel, trafRows, 'T', 2);

module.exports = { userRows, trafRows, totals, lastDayUsers, meta, optUsers, optTraf };

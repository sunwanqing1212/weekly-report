// 将计算结果导出为 JSON，供 gen_xlsx.py 生成原生图表 xlsx
const fs = require('fs');
const { userRows, trafRows, totals, lastDayUsers, meta } = require('./report_data');
fs.writeFileSync('report_data.json', JSON.stringify({ userRows, trafRows, totals, lastDayUsers, meta }, null, 2));
console.log('JSON written: report_data.json');

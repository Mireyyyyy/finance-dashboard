/**
 * 财务 Excel 自动同步模块
 * 每次网页数据变更时，同步到财务部门的 Excel 文件
 * 使用 xlsx (SheetJS) 保持公式和格式
 */
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

// Excel 日期纪元 (1899-12-30 UTC, 考虑 Lotus 123 闰年bug)
// 注意: 必须用 Date.UTC 避免本地时区历史偏差 (中国1901年前为GMT+0805)
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);

function serialToDate(serial) {
  const d = new Date(EXCEL_EPOCH_MS + serial * 86400000);
  return d.toISOString().slice(0, 10);
}

function dateToSerial(dateStr) {
  const parts = dateStr.split('-').map(Number);
  return Math.round((Date.UTC(parts[0], parts[1] - 1, parts[2]) - EXCEL_EPOCH_MS) / 86400000);
}

/** 在工作表中查找包含关键字的表头行 */
function findHeaderRow(ws, range, keyword) {
  for (let r = range.s.r; r <= range.e.r; r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (cell && cell.v !== undefined && String(cell.v).includes(keyword)) {
        return r;
      }
    }
  }
  return -1;
}

/**
 * 同步财务部 Excel 文件（带公式的模板）
 * - 只更新输入列（饮食-W/H、交通、其他、备注），保持公式列不变
 * - 更新汇总页的初始金额
 */
function syncFinanceExcel(filepath, dailyData, monthlyBudgets) {
  if (!fs.existsSync(filepath)) {
    console.log(`[Excel同步] 文件不存在，跳过: ${path.basename(filepath)}`);
    return false;
  }

  const wb = XLSX.readFile(filepath, { cellFormula: true, cellStyles: true });
  const ws = wb.Sheets['明细'];
  if (!ws || !ws['!ref']) {
    console.log(`[Excel同步] ${path.basename(filepath)} 缺少明细表`);
    return false;
  }

  const range = XLSX.utils.decode_range(ws['!ref']);
  const hRow = findHeaderRow(ws, range, '饮食-W');
  if (hRow < 0) {
    console.log(`[Excel同步] ${path.basename(filepath)} 未找到表头`);
    return false;
  }

  // 映射列索引
  let colDate = -1, colFW = -1, colFH = -1, colT = -1, colO = -1, colNote = -1;
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: hRow, c })];
    if (!cell || cell.v === undefined) continue;
    const v = String(cell.v).trim();
    if (v === '日期') colDate = c;
    else if (v.includes('饮食-W')) colFW = c;
    else if (v.includes('饮食-H')) colFH = c;
    else if (v === '交通') colT = c;
    else if (v === '其他') colO = c;
    else if (v === '备注') colNote = c;
  }
  if (colDate < 0) {
    console.log(`[Excel同步] ${path.basename(filepath)} 未找到日期列`);
    return false;
  }

  // 日期 -> 数据 映射
  const dataMap = {};
  (dailyData || []).forEach(d => { dataMap[d.d] = d; });

  // 更新数据行（只更新输入列，跳过公式的列）
  for (let r = hRow + 1; r <= range.e.r; r++) {
    const dateAddr = XLSX.utils.encode_cell({ r, c: colDate });
    const dateCell = ws[dateAddr];
    if (!dateCell || dateCell.v === undefined || dateCell.v === '') continue;

    let dateStr;
    if (typeof dateCell.v === 'number') {
      dateStr = serialToDate(dateCell.v);
    } else if (dateCell.v instanceof Date) {
      dateStr = dateCell.v.toISOString().slice(0, 10);
    } else {
      dateStr = String(dateCell.v).slice(0, 10);
    }

    const entry = dataMap[dateStr];
    if (!entry) continue;

    // 设置单元格值（保留存在的 f 公式属性）
    const setCell = (col, value, type) => {
      if (col < 0) return;
      const addr = XLSX.utils.encode_cell({ r, c: col });
      const existing = ws[addr];
      if (existing && existing.f) {
        // 公式单元格：只更新缓存值
        existing.v = value;
      } else {
        // 普通单元格：直接覆写
        ws[addr] = { v: value, t: type };
      }
    };

    if (colFW >= 0) setCell(colFW, Number(entry.fw) || 0, 'n');
    if (colFH >= 0) setCell(colFH, Number(entry.fh) || 0, 'n');
    if (colT >= 0) setCell(colT, Number(entry.t) || 0, 'n');
    if (colO >= 0) setCell(colO, Number(entry.o) || 0, 'n');
    if (colNote >= 0) setCell(colNote, String(entry.note || ''), 's');
  }

  // 更新汇总页的初始金额
  const wsSum = wb.Sheets['汇总'];
  if (wsSum && wsSum['!ref']) {
    const sRange = XLSX.utils.decode_range(wsSum['!ref']);
    for (let r = sRange.s.r; r <= sRange.e.r; r++) {
      for (let c = sRange.s.c; c <= sRange.e.c; c++) {
        const cell = wsSum[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v !== undefined && String(cell.v).includes('初始金额')) {
          const valAddr = XLSX.utils.encode_cell({ r, c: c + 1 });
          // 用最新的月份预算
          const months = Object.keys(monthlyBudgets || {}).sort();
          const budget = months.length > 0 ? monthlyBudgets[months[months.length - 1]] : 2370;
          const existingVal = wsSum[valAddr];
          if (existingVal && existingVal.f) {
            existingVal.v = budget;
          } else {
            wsSum[valAddr] = { v: budget, t: 'n' };
          }
          break;
        }
      }
    }
  }

  XLSX.writeFile(wb, filepath);
  console.log(`[Excel同步] ${path.basename(filepath)} 已更新`);
  return true;
}

/**
 * 同步导出格式 Excel（无公式，完全覆写）
 * 格式：明细表 + 汇总表(月水电)
 */
function syncExportExcel(filepath, dailyData, monthlyWater) {
  const wb = XLSX.utils.book_new();

  // 明细表
  const detailRows = [
    ['日期', '星期', '饮食-W', '饮食-H', '饮食小计', '交通', '其他', '日合计', '余额', '备注']
  ];
  (dailyData || []).forEach(d => {
    detailRows.push([
      d.d, '周' + (d.w || ''),
      d.fw || 0, d.fh || 0, d.fs || 0,
      d.t || 0, d.o || 0, d.total || 0, d.rem || 0,
      d.note || ''
    ]);
  });
  const ws1 = XLSX.utils.aoa_to_sheet(detailRows);
  ws1['!cols'] = [
    { wch: 12 }, { wch: 5 }, { wch: 8 }, { wch: 8 },
    { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 },
    { wch: 8 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, ws1, '明细');

  // 汇总表（水电按月）
  const sumRows = [['月份', '水电金额']];
  if (monthlyWater) {
    Object.keys(monthlyWater)
      .sort()
      .forEach(m => {
        const monthNum = parseInt(m.split('-')[1]);
        sumRows.push([monthNum + '月', monthlyWater[m] || 0]);
      });
  }
  const ws2 = XLSX.utils.aoa_to_sheet(sumRows);
  ws2['!cols'] = [{ wch: 6 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws2, '汇总');

  XLSX.writeFile(wb, filepath);
  console.log(`[Excel同步] ${path.basename(filepath)} 已更新`);
  return true;
}

/**
 * 同步所有 Excel 文件
 * @param {Object} data 来自网页的数据 { dailyData, monthlyWater, monthlyBudgets, currentMonth }
 * @param {string} baseDir 项目根目录
 * @returns {Object} 各文件同步结果
 */
function syncAll(data, baseDir) {
  const dailyData = data.dailyData || [];
  const monthlyBudgets = data.monthlyBudgets || {};
  const monthlyWater = data.monthlyWater || {};
  const results = {};

  // 1. 财务部 Excel（带公式）
  const financeFiles = [
    '财务部门_new.xlsx',
    fs.existsSync(path.join(baseDir, '财务部门_月结算.xlsx')) ? '财务部门_月结算.xlsx' : '财务部门_周结算.xlsx'
  ];
  for (const f of financeFiles) {
    const fp = path.join(baseDir, f);
    results[f] = syncFinanceExcel(fp, dailyData, monthlyBudgets);
  }

  // 2. 导出格式 Excel
  const exportFile = '财务支出明细_20260706.xlsx';
  const exportFp = path.join(baseDir, exportFile);
  results[exportFile] = syncExportExcel(exportFp, dailyData, monthlyWater);

  return results;
}

/**
 * 生成财务部格式的 Excel 文件（返回 Buffer，供下载）
 * 格式与 财务部门_月结算.xlsx 一致
 */
function generateFinanceExcel(data) {
  const dailyData = data.dailyData || [];
  const monthlyBudgets = data.monthlyBudgets || {};
  const monthlyWater = data.monthlyWater || {};

  const wb = XLSX.utils.book_new();

  // ===== 明细表 =====
  const detailHeader = ['日期', '周起始日', '周', '饮食-W', '饮食-H', '饮食小计', '交通', '水电(月结算)', '其他', '天总计', '支出后余额', '备注'];
  const detailRows = [detailHeader];

  (dailyData || []).forEach(d => {
    const dt = new Date(d.d + 'T00:00:00');
    const weekStart = new Date(dt);
    weekStart.setDate(dt.getDate() - dt.getDay() + 1); // 周一
    const wsStr = weekStart.getMonth() + 1 + '/' + weekStart.getDate();

    detailRows.push([
      d.d,                          // 日期(文本)
      wsStr,                        // 周起始日
      d.w ? '周' + d.w : '',        // 周
      d.fw || 0,                    // 饮食-W
      d.fh || 0,                    // 饮食-H
      d.fs || 0,                    // 饮食小计
      d.t || 0,                     // 交通
      0,                            // 水电(月结算) — 留空手动填
      d.o || 0,                     // 其他
      d.total || 0,                 // 天总计
      d.rem || 0,                   // 支出后余额
      d.note || ''                  // 备注
    ]);
  });

  const ws1 = XLSX.utils.aoa_to_sheet(detailRows);
  ws1['!cols'] = [
    { wch: 12 }, { wch: 8 }, { wch: 5 }, { wch: 8 }, { wch: 8 },
    { wch: 8 }, { wch: 8 }, { wch: 10 }, { wch: 8 }, { wch: 8 },
    { wch: 10 }, { wch: 16 }
  ];
  XLSX.utils.book_append_sheet(wb, ws1, '明细');

  // ===== 汇总表 =====
  // 计算月度汇总
  const monthMap = {};
  dailyData.forEach(d => {
    const m = d.d.slice(0, 7);
    if (!monthMap[m]) monthMap[m] = { food: 0, transport: 0, other: 0, total: 0, count: 0 };
    monthMap[m].food += (d.fs || 0);
    monthMap[m].transport += (d.t || 0);
    monthMap[m].other += (d.o || 0);
    monthMap[m].total += (d.total || 0);
    if (d.total > 0) monthMap[m].count++;
  });

  // 获取最新月份预算
  const months = Object.keys(monthlyBudgets).sort();
  const latestBudget = months.length > 0 ? monthlyBudgets[months[months.length - 1]] : 0;

  const sumRows = [
    ['财务部门支出汇总'],
    ['数据来源：网页自动同步'],
    [],
    ['月份', '天数', '预算', '饮食', '交通', '水电', '其他', '月合计', '日均'],
  ];

  Object.keys(monthMap).sort().forEach(m => {
    const label = parseInt(m.split('-')[1]) + '月';
    const mm = monthMap[m];
    const budget = monthlyBudgets[m] || 0;
    const water = monthlyWater[m] || 0;
    const monthTotal = mm.total + water;
    sumRows.push([
      label,
      mm.count + '天',
      budget,
      Math.round(mm.food * 100) / 100,
      Math.round(mm.transport * 100) / 100,
      water,
      Math.round(mm.other * 100) / 100,
      Math.round(monthTotal * 100) / 100,
      mm.count > 0 ? Math.round(monthTotal / mm.count * 100) / 100 : 0
    ]);
  });

  sumRows.push([]);
  sumRows.push(['初始金额', latestBudget]);

  const ws2 = XLSX.utils.aoa_to_sheet(sumRows);
  ws2['!cols'] = [{ wch: 6 }, { wch: 6 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ws2, '汇总');

  // 导出为 Buffer
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

module.exports = { syncAll, generateFinanceExcel };

const express = require('express');
const fs = require('fs');
const path = require('path');
const { syncAll, generateFinanceExcel } = require('./syncFinance.js');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data.json');

const WEEK_MONTHS = ['2026-06', '2026-06', '2026-07', '2026-07'];
const WEEK_KEYS = ['w1', 'w2', 'w3', 'w4'];

function normalizeFinanceData(data) {
  const normalized = { ...data };

  if (!normalized.monthlyWater) {
    normalized.monthlyWater = {};
    if (normalized.weeklyUtils) {
      WEEK_KEYS.forEach((key, index) => {
        const month = WEEK_MONTHS[index];
        normalized.monthlyWater[month] =
          (normalized.monthlyWater[month] || 0) + (parseFloat(normalized.weeklyUtils[key]) || 0);
      });
    }
  }

  if (!normalized.monthlyBudgets && normalized.initBalance) {
    normalized.monthlyBudgets = {
      '2026-06': normalized.initBalance,
      '2026-07': normalized.initBalance
    };
  }

  delete normalized.weeklyUtils;
  delete normalized.initBalance;
  return normalized;
}

app.use(express.json({ limit: '5mb' }));

// 设备检测路由：手机 → mobile.html，电脑 → index.html
app.get('/', (req, res) => {
  const ua = (req.headers['user-agent'] || '').toLowerCase();
  const isMobile = /mobile|android|iphone|ipad|phone|webos|blackberry|iemobile|opera mini/i.test(ua);
  res.sendFile(path.join(__dirname, 'public', isMobile ? 'mobile.html' : 'index.html'));
});

app.use(express.static(path.join(__dirname, 'public')));

// CORS
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

// GET /api/data
app.get('/api/data', (req, res) => {
  try {
    if (fs.existsSync(DATA_FILE)) {
      const raw = fs.readFileSync(DATA_FILE, 'utf-8');
      return res.json({ success: true, data: JSON.parse(raw), timestamp: Date.now() });
    }
    res.json({ success: true, data: null, timestamp: Date.now() });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// POST /api/data
app.post('/api/data', (req, res) => {
  try {
    const body = req.body;
    let saveData;
    if (body.data) {
      saveData = normalizeFinanceData(body.data);
    } else if (body.dailyData) {
      saveData = normalizeFinanceData({
        dailyData: body.dailyData,
        monthlyWater: body.monthlyWater,
        weeklyUtils: body.weeklyUtils,
        monthlyBudgets: body.monthlyBudgets,
        initBalance: body.initBalance,
        currentMonth: body.currentMonth
      });
    } else {
      return res.status(400).json({ success: false, error: '无效数据' });
    }
    const payload = { savedAt: new Date().toISOString(), data: saveData };
    fs.writeFileSync(DATA_FILE, JSON.stringify(payload, null, 2), 'utf-8');

    // 如果服务器本地有 Excel 文件，也同步（本地部署时有效）
    let syncResults = {};
    try {
      syncResults = syncAll(saveData, __dirname);
    } catch (syncErr) {
      // 忽略（云端部署没有本地 Excel 文件）
    }

    res.json({ success: true, timestamp: Date.now(), excelSync: syncResults });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/export-excel — 下载财务部格式 Excel
app.get('/api/export-excel', (req, res) => {
  try {
    if (!fs.existsSync(DATA_FILE)) {
      return res.status(404).json({ success: false, error: '无数据' });
    }
    const raw = fs.readFileSync(DATA_FILE, 'utf-8');
    const parsed = JSON.parse(raw);
    const data = parsed.data || {};

    const buffer = generateFinanceExcel(data);
    const filename = '财务部门_月结算.xlsx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename*=UTF-8\'\'' + encodeURIComponent(filename));
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  }
});

// GET /api/status
app.get('/api/status', (req, res) => {
  res.json({ success: true, status: 'running', uptime: process.uptime() });
});

app.listen(PORT, '0.0.0.0', () => {
  const url = process.env.RENDER_EXTERNAL_URL || `http://localhost:${PORT}`;
  console.log('========================================');
  console.log('  财务支出看板 - 服务已启动');
  console.log('========================================');
  console.log(`  访问地址: ${url}`);
  console.log('  Ctrl+C 停止服务');
  console.log('========================================');
});

const express = require('express');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const DATA_FILE = path.join(__dirname, 'data.json');

app.use(express.json({ limit: '5mb' }));
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

// POST /api/data (支持两种格式: {data:{...}} 或 {dailyData,weeklyUtils,initBalance})
app.post('/api/data', (req, res) => {
  try {
    const body = req.body;
    let saveData;
    if (body.data) {
      saveData = body.data;
    } else if (body.dailyData) {
      saveData = { dailyData: body.dailyData, weeklyUtils: body.weeklyUtils, initBalance: body.initBalance };
    } else {
      return res.status(400).json({ success: false, error: '无效数据' });
    }
    const payload = { savedAt: new Date().toISOString(), data: saveData };
    fs.writeFileSync(DATA_FILE, JSON.stringify(payload, null, 2), 'utf-8');
    res.json({ success: true, timestamp: Date.now() });
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

# 财务支出看板

个人/家庭财务支出管理工具，支持多设备实时同步。

## 部署到公网（Render.com）

### 第一步：创建 GitHub 仓库

1. 打开 https://github.com/new
2. 仓库名填 `finance-dashboard`
3. 选择 **Private**（私有仓库）
4. 点击 **Create repository**

### 第二步：上传代码

在终端中执行：

```bash
cd C:\Users\51299\Desktop\财务看板
git init
git add .
git commit -m "first commit"
git remote add origin https://github.com/<你的用户名>/finance-dashboard.git
git push -u origin main
```

### 第三步：部署到 Render

1. 打开 https://dashboard.render.com/
2. 用 GitHub 账号登录（点 "Sign in with GitHub"）
3. 点 **New +** → **Blueprint**
4. 选择 `finance-dashboard` 仓库
5. 点 **Apply**，等待 2-3 分钟
6. 部署完成后会显示 `https://finance-dashboard.onrender.com`

### 第四步：访问

打开 `https://finance-dashboard.onrender.com` 即可使用。

> **注意**：Render 免费实例如果 15 分钟无人访问会休眠，再次访问时会自动唤醒（等待 10-30 秒）。

## 本地运行

```bash
npm install
node server.js
# 访问 http://localhost:3000
```

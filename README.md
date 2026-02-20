# 市人大常委会监督协调处管理系统

市人大常委会监督协调处业务管理平台：Node.js + Express 后端、JSON 文件存储、Bootstrap 5 前端，支持统一管理员登录与多业务模块管理。

---

## 技术栈

| 层级 | 技术 |
|------|------|
| 后端 | Node.js、Express 4.x |
| 存储 | JSON 文件（无需独立数据库） |
| 前端 | Bootstrap 5、Font Awesome、Chart.js、PDF.js（CDN） |
| 其他 | compression（gzip）、multer（上传）、xlsx（Excel）、mammoth（Word） |

---

## 配置

| 配置项 | 默认值 | 说明 |
|--------|--------|------|
| 数据库名 | `jiandu` | 数据目录 `data/jiandu/` |
| 端口 | `3000` | 服务监听端口 |

修改方式：编辑根目录 `config.js` 中的 `DB_NAME`、`PORT`。

---

## 本地运行

### 环境要求

- Node.js 14+
- npm

### 安装与启动

```bash
npm install
npm start
```

### 访问

浏览器打开：**http://localhost:3000**

### 默认账号

| 角色 | 用户名 | 密码 |
|------|--------|------|
| 超级管理员 | `1312` | `1312` |
| 普通管理员 | `1645` | `4688633` |

---

## 业务模块

| 模块 | 入口页面 | 功能要点 |
|------|----------|----------|
| 监督议题 | `index.html` | 年度/月份议题、监督形式、部门、数据概览、文件资料上传与预览 |
| 批示督办 | `pishi.html` | 期数、来文单位、批示内容、完成情况、办理情况报告 |
| 民生实事 | `minsheng.html` | 年度项目、监督部门、小组成员、完成情况 |
| 每月工作 | `meiyue.html` | 按月份 PDF 上传与预览 |
| 每周工作 | `meizhou.html` | 按周 PDF 上传与预览 |
| 法律法规 | `falv.html` | 法规分类、制定机关、Word/PDF 上传与预览 |
| 计划月历 | `jihua.html` | 常委会每月/机关每周重点工作、弹窗预览 |

详情页：`jiandu-file-detail.html`、`pishi-report-detail.html`、`meiyue-detail.html`、`meizhou-detail.html`、`minsheng-progress-detail.html`、`falv-detail.html` 等，用于文件预览或嵌入 iframe。

---

## 功能概览

- **通用**：列表筛选、模糊搜索、排序；管理员登录；新增/编辑/删除、批量删除；JSON 备份与恢复。
- **监督议题**：数据概览、点击筛选、部门管理、议题下 Word/PDF 资料上传与预览。
- **批示督办 / 民生实事**：Excel 导入、统计图表。
- **法律法规**：分类与制定机关筛选、批量下载/导出目录。
- **计划月历**：按年度/月份/周展示，仅显示有内容的项，点击弹窗预览（embed 模式无导航）。

---

## 项目结构

```
jiandu/
├── config.js                 # 配置（DB_NAME, PORT）
├── server.js                 # Express 后端与 API
├── package.json
├── index.html                # 监督议题
├── pishi.html                # 批示督办
├── minsheng.html             # 民生实事
├── meiyue.html               # 每月工作
├── meizhou.html              # 每周工作
├── falv.html                 # 法律法规
├── jihua.html                # 计划月历
├── *-detail.html             # 各模块详情/预览页
├── css/
│   └── common.css            # 公共样式（导航、按钮、布局）
├── data/
│   └── jiandu/               # 数据目录（首次运行自动创建）
│       ├── admins.json
│       ├── jiandu_topics.json
│       ├── jiandu_topic_files.json
│       ├── departments.json
│       ├── form_sort.json
│       ├── pishi.json
│       ├── pishi_report.json
│       ├── projects.json
│       ├── minsheng_progress.json
│       ├── meiyue.json
│       ├── meizhou.json
│       ├── falv.json
│       └── uploads/           # 各模块上传文件
└── README.md
```

前端无构建步骤：页面直接引用 CDN（Bootstrap、Chart.js、PDF.js 等）和本地 `css/common.css`，脚本内联在各 HTML 中。

---

## 数据与迁移

- 所有业务数据存于 `data/jiandu/` 下的 JSON 文件。
- 首次启动会自动创建缺失的 JSON 及 `uploads` 子目录；若存在旧版 `data/minsheng/` 或根目录 `data/projects.json`、`data/admins.json`，会迁移到 `data/jiandu/`。

---

## 性能与缓存

### 后端

- **JSON 内存缓存**：读操作使用内存缓存，**写入时自动失效**（`invalidateCache`），进程重启后缓存清空，避免长期占用或脏数据。
- **gzip**：启用 `compression` 中间件，响应体压缩。
- **API 不缓存**：所有 `/api` 响应带 `Cache-Control: no-store`，防止浏览器缓存接口数据，保证登录/数据变更后看到最新结果。

### 静态资源

- **CSS/JS**：`Cache-Control: public, max-age=3600`（1 小时），减少重复请求。
- **HTML**：`Cache-Control: no-cache`，便于发布后及时更新页面。

### 前端

- 列表模糊搜索使用防抖，降低输入时的渲染与请求频率。

---

## 脚本说明

| 命令 | 说明 |
|------|------|
| `npm start` | 启动服务（同 `node server.js`） |
| `npm run dev` | 同上，无热重载 |

---

## 许可证

MIT

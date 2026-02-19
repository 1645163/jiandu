# 市人大常委会监督协调处管理系统

市人大常委会监督协调处业务管理平台，采用 Node.js + Express 后端、JSON 文件存储、Bootstrap 5 前端，支持统一管理员登录与多业务模块管理。

## 系统配置

| 配置项 | 值 | 说明 |
|--------|-----|------|
| 数据库名 | jiandu | 数据目录 `data/jiandu/` |
| 端口 | 3000 | 服务监听端口 |

## 业务模块

| 模块 | 页面 | 功能要点 |
|------|------|----------|
| 监督议题 | index.html | 年度/月份监督议题、监督形式、部门、数据概览、文件资料上传与管理 |
| 批示督办 | pishi.html | 期数、来文单位、批示内容、批示日期、领导与责任部门、类别、落实举措、完成情况 |
| 民生实事 | minsheng.html | 年度、项目名称、监督部门、小组成员、重点监督内容、完成情况 |
| 每月工作 | meiyue.html | 年度、月份、时间、地点、内容、出席领导、备注 |
| 每周工作 | meizhou.html | 年度、月份、周数、部门、内容、备注 |
| 法律法规 | falv.html | 法规名称、分类、制定机关、公布/施行日期、Word/PDF 文档 |
| 工作计划 | jihua.html | 工作计划管理 |

## 功能说明

### 通用功能

- **列表展示**：年度/期数/月份筛选、模糊搜索、排序
- **权限管理**：管理员登录、超级管理员权限管理、强制下线
- **数据操作**：新增 / 编辑 / 删除、批量删除
- **数据备份**：JSON 备份与恢复
- **数据导入**：批示督办、每月/每周工作支持 Excel 导入
- **统计图表**：民生实事监督部门统计、批示督办概览、监督议题数据概览

### 监督议题特色

- **数据概览**：全年监督议题数量、按监督形式分布、按部门分布、按工作时间分布
- **点击筛选**：点击概览中对应类别可筛选议题列表
- **部门管理**：按民生实事模式管理监督部门
- **文件资料**：每个议题可上传 Word/PDF 资料，支持预览、下载、删除
- **图标区分**：操作列图钉图标有/无资料时不同样式（蓝色填充 / 灰色描边）

### 法律法规

- Word/PDF 上传、在线预览、下载、分类与制定机关筛选

## 本地运行

### 环境要求

- Node.js 14+
- npm

### 安装与启动

```bash
npm install
npm start
```

### 访问系统

浏览器打开：**http://localhost:3000**

### 默认账号

- 超级管理员：用户名 `1312`，密码 `1312`
- 普通管理员：用户名 `1645`，密码 `4688633`

## 项目结构

```
jiandu/
├── config.js              # 配置（数据库名、端口）
├── server.js              # 后端服务
├── index.html             # 监督议题
├── pishi.html             # 批示督办
├── minsheng.html          # 民生实事
├── meiyue.html            # 每月工作
├── meizhou.html           # 每周工作
├── falv.html              # 法律法规
├── jihua.html             # 工作计划
├── jiandu-file-detail.html    # 监督议题文件预览
├── minsheng-progress-detail.html
├── pishi-report-detail.html
├── meiyue-detail.html
├── meizhou-detail.html
├── falv-detail.html
├── css/
│   └── common.css
├── data/
│   └── jiandu/
│       ├── projects.json      # 民生实事
│       ├── admins.json        # 管理员
│       ├── jiandu_topics.json # 监督议题
│       ├── jiandu_topic_files.json
│       ├── departments.json   # 监督部门
│       ├── form_sort.json     # 监督形式排序
│       ├── pishi.json
│       ├── pishi_report.json
│       ├── minsheng_progress.json
│       ├── meiyue.json
│       ├── meizhou.json
│       ├── falv.json
│       └── uploads/
│           ├── jiandu_topics/ # 监督议题资料
│           ├── falv/          # 法律法规文档
│           ├── pishi_report/
│           ├── meiyue/
│           ├── meizhou/
│           └── minsheng/
├── package.json
└── README.md
```

## 数据存储

- 使用 JSON 文件存储，无需单独安装数据库
- 数据目录：`data/jiandu/`
- 首次运行自动创建数据目录及示例数据
- 若存在旧版 `data/minsheng/` 或根目录 `data/projects.json`、`data/admins.json`，将自动迁移到 `data/jiandu/`

## 性能与优化

- 后端：JSON 文件内存缓存（写入时失效）、gzip 压缩响应
- 前端：模糊搜索防抖，减少无效渲染

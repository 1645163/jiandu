/**
 * 市人大常委会监督协调处管理系统 - 后端服务
 * 数据库：jiandu | 端口：3000
 * Node.js + Express + JSON 文件存储
 */

const express = require('express');
const path = require('path');
const fs = require('fs');
const XLSX = require('xlsx');
const multer = require('multer');
const mammoth = require('mammoth');
const config = require('./config.js');
let compression;
try { compression = require('compression'); } catch (e) { compression = null; }

const app = express();
const PORT = config.PORT;
const DB_DIR = path.join(__dirname, 'data', config.DB_NAME);
const DATA_FILE = path.join(DB_DIR, 'projects.json');
const ADMINS_FILE = path.join(DB_DIR, 'admins.json');
const PISHI_FILE = path.join(DB_DIR, 'pishi.json');
const MEIYUE_FILE = path.join(DB_DIR, 'meiyue.json');
const MEIZHOU_FILE = path.join(DB_DIR, 'meizhou.json');
const FALV_FILE = path.join(DB_DIR, 'falv.json');
const FALV_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'falv');
const MEIZHOU_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'meizhou');
const MEIYUE_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'meiyue');
const MINSHENG_PROGRESS_FILE = path.join(DB_DIR, 'minsheng_progress.json');
const MINSHENG_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'minsheng');
const DEPARTMENTS_FILE = path.join(DB_DIR, 'departments.json');
const PISHI_REPORT_FILE = path.join(DB_DIR, 'pishi_report.json');
const PISHI_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'pishi_report');
const JIANDU_TOPICS_FILE = path.join(DB_DIR, 'jiandu_topics.json');
const JIANDU_TOPIC_FILES_FILE = path.join(DB_DIR, 'jiandu_topic_files.json');
const FORM_SORT_FILE = path.join(DB_DIR, 'form_sort.json');
const JIANDU_UPLOAD_DIR = path.join(DB_DIR, 'uploads', 'jiandu_topics');

// 批示督办表头（与 Excel 模板一致）
const PISHI_HEADERS = ['期数', '来文单位', '文号', '批示内容', '批示日期', '领导和责任部门', '类别', '落实举措', '完成情况'];
const PISHI_KEYS = ['qishu', 'laiwenUnit', 'wenhao', 'pishiContent', 'pishiDate', 'leaderDept', 'category', 'luoshiCuoshi', 'completeStatus'];

var meizhouStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(MEIZHOU_UPLOAD_DIR)) fs.mkdirSync(MEIZHOU_UPLOAD_DIR, { recursive: true });
    cb(null, MEIZHOU_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'meizhou_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var meizhouUpload = multer({
  storage: meizhouStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF 格式'));
  }
});

var meiyueStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(MEIYUE_UPLOAD_DIR)) fs.mkdirSync(MEIYUE_UPLOAD_DIR, { recursive: true });
    cb(null, MEIYUE_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'meiyue_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var meiyueUpload = multer({
  storage: meiyueStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF 格式'));
  }
});

const OLD_DATA_FILE = path.join(__dirname, 'data', 'projects.json');
const OLD_ADMINS_FILE = path.join(__dirname, 'data', 'admins.json');
const LEGACY_DIR = path.join(__dirname, 'data', 'minsheng');

// 首次启动迁移：data/projects|admins.json 或 data/minsheng/* -> data/jiandu/
function migrateToJiandu() {
  if (!fs.existsSync(DB_DIR)) fs.mkdirSync(DB_DIR, { recursive: true });
  const files = [
    ['projects.json', DATA_FILE, OLD_DATA_FILE],
    ['admins.json', ADMINS_FILE, OLD_ADMINS_FILE],
    ['pishi.json', PISHI_FILE],
    ['meiyue.json', MEIYUE_FILE],
    ['meizhou.json', MEIZHOU_FILE],
    ['falv.json', FALV_FILE]
  ];
  for (const arr of files) {
    const name = arr[0], dest = arr[1], legacy1 = arr[2], legacy2 = path.join(LEGACY_DIR, name);
    if (!fs.existsSync(dest)) {
      if (legacy1 && fs.existsSync(legacy1)) {
        fs.copyFileSync(legacy1, dest);
        console.log('  已迁移 ' + name + ' 到 jiandu');
      } else if (fs.existsSync(legacy2)) {
        fs.copyFileSync(legacy2, dest);
        console.log('  已迁移 ' + name + ' 从 minsheng 到 jiandu');
      }
    }
  }
}

var tokenStore = {};

// 文件读取缓存（写入时失效，减少磁盘 I/O）
var fileCache = {};
function readCache(filePath, readFn) {
  var mtime = 0;
  try { if (fs.existsSync(filePath)) mtime = fs.statSync(filePath).mtimeMs; } catch (e) {}
  if (fileCache[filePath] && fileCache[filePath].m === mtime) return fileCache[filePath].d;
  var d = readFn();
  fileCache[filePath] = { m: mtime, d: d };
  return d;
}
function invalidateCache(filePath) { delete fileCache[filePath]; }

// 中间件
if (compression) app.use(compression());
app.use(express.json({ limit: '10mb' }));

// 模板下载路由（需在 static 之前，避免被静态文件拦截）
app.get('/api/projects/template', function (req, res) {
  try {
    var XLSX = require('xlsx');
    var PROJECT_HEADERS = ['年度', '项目名称', '监督部门', '小组成员名单', '重点监督内容', '完成情况'];
    var ws = XLSX.utils.aoa_to_sheet([PROJECT_HEADERS]);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
    if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
    res.setHeader('Content-Disposition', 'attachment; filename="minsheng_template.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Cache-Control', 'no-cache');
    res.send(buf);
  } catch (e) {
    console.error('projects template error:', e);
    res.status(500).json({ code: 1, msg: e.message || '生成模板失败' });
  }
});
app.get('/api/pishi/template', function (req, res) {
  try {
    var XLSX = require('xlsx');
    var ws = XLSX.utils.aoa_to_sheet([PISHI_HEADERS]);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
    if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
    res.setHeader('Content-Disposition', 'attachment; filename="pishi_template.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Cache-Control', 'no-cache');
    res.send(buf);
  } catch (e) {
    console.error('pishi template error:', e);
    res.status(500).json({ code: 1, msg: e.message || '生成模板失败' });
  }
});
app.get('/api/jiandu/template', function (req, res) {
  try {
    var XLSX = require('xlsx');
    var KEYS = ['年度', '月份', '监督内容', '监督形式', '部门/处室'];
    var ws = XLSX.utils.aoa_to_sheet([KEYS]);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    var buf = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
    if (!Buffer.isBuffer(buf)) buf = Buffer.from(buf);
    res.setHeader('Content-Disposition', 'attachment; filename="jiandu_topics_template.xlsx"');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Cache-Control', 'no-cache');
    res.send(buf);
  } catch (e) {
    res.status(500).json({ code: 1, msg: e.message || '生成模板失败' });
  }
});

app.use(express.static(path.join(__dirname)));

// 默认管理员（超级管理员 1312 + 普通管理员 1645）
function getDefaultAdmins() {
  return [
    { id: 1, username: '1312', password: '1312', role: 'super_admin' },
    { id: 2, username: '1645', password: '4688633', role: 'admin' }
  ];
}

function loadAdmins() {
  return readCache(ADMINS_FILE, function() {
    var dir = path.dirname(ADMINS_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(ADMINS_FILE)) {
      var admins = getDefaultAdmins();
      fs.writeFileSync(ADMINS_FILE, JSON.stringify({ admins: admins, nextId: 3 }, null, 2));
      return { admins: admins.slice(), nextId: 3 };
    }
    var raw = fs.readFileSync(ADMINS_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { admins: data.admins || getDefaultAdmins(), nextId: data.nextId || 3 };
  });
}

function saveAdmins(data) {
  var dir = path.dirname(ADMINS_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(ADMINS_FILE, JSON.stringify(data, null, 2));
  invalidateCache(ADMINS_FILE);
}

// 根据 token 获取当前用户（用于权限校验）
function getCurrentUser(req) {
  var auth = req.headers.authorization;
  if (!auth || auth.indexOf('Bearer ') !== 0) return null;
  var token = auth.slice(7);
  return tokenStore[token] || null;
}

// 需要登录的操作：无有效 token 则返回 401
function requireAuth(req, res, next) {
  var user = getCurrentUser(req);
  if (!user) {
    res.status(401).json(err('未登录或已被强制下线，请重新登录'));
    return;
  }
  req.currentUser = user;
  next();
}

// 默认数据
const DEFAULT_DATA = [
  { id: 1, year: 2025, name: '老旧小区改造', department: '住建局', members: '张军、李红、王强', supervise: '改造进度、工程质量、居民满意度', status: '进行中' },
  { id: 2, year: 2025, name: '社区养老服务中心建设', department: '民政局', members: '刘芳、赵伟、孙丽', supervise: '场地建设、人员配置、服务落地', status: '已完成' },
  { id: 3, year: 2024, name: '农村饮水安全工程', department: '水利局', members: '陈明、周杰、吴丹', supervise: '水质检测、管网铺设、供水稳定性', status: '已完成' },
  { id: 4, year: 2023, name: '义务教育学校扩建', department: '教育局', members: '郑华、马涛、钱静', supervise: '施工进度、师资配套、招生计划', status: '已完成' }
];

function loadData() {
  return readCache(DATA_FILE, function() {
    var dir = path.dirname(DATA_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(DATA_FILE)) {
      fs.writeFileSync(DATA_FILE, JSON.stringify({ projects: DEFAULT_DATA, nextId: 5 }, null, 2));
      return { projects: DEFAULT_DATA.slice(), nextId: 5 };
    }
    var raw = fs.readFileSync(DATA_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { projects: data.projects || [], nextId: data.nextId || 5 };
  });
}

function saveData(data) {
  var dir = path.dirname(DATA_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2));
  invalidateCache(DATA_FILE);
}

// 统一响应格式
function ok(data) {
  return { code: 0, msg: '成功', data: data };
}
function err(msg) {
  return { code: 1, msg: msg || '操作失败' };
}

// ========== API 路由 ==========

// 登录（从 admins 表校验）
app.post('/api/login', function (req, res) {
  var body = req.body || {};
  var username = (body.username || '').trim();
  var password = (body.password || '').trim();
  if (!username || !password) return res.json(err('请输入用户名和密码'));
  var data = loadAdmins();
  var admin = data.admins.find(function (a) { return a.username === username && a.password === password; });
  if (!admin) return res.json(err('用户名或密码错误'));
  var token = Buffer.from(username + ':' + Date.now() + ':' + Math.random()).toString('base64');
  tokenStore[token] = { username: admin.username, userId: admin.id, role: admin.role || 'admin' };
  res.json(ok({ token: token, username: admin.username, role: admin.role || 'admin' }));
});

// 登出（清除服务端 token）
app.post('/api/logout', function (req, res) {
  var auth = req.headers.authorization;
  if (auth && auth.indexOf('Bearer ') === 0) {
    var token = auth.slice(7);
    delete tokenStore[token];
  }
  res.json(ok({}));
});

// 校验 token 是否有效
app.get('/api/me', function (req, res) {
  var user = getCurrentUser(req);
  if (!user) return res.status(401).json(err('未登录或已下线'));
  res.json(ok({ username: user.username, role: user.role }));
});

// ---------- 超级管理员专用 ----------

// 强制指定管理员下线
app.post('/api/admin/force-logout', function (req, res) {
  var cur = getCurrentUser(req);
  if (!cur || cur.role !== 'super_admin') return res.status(403).json(err('仅超级管理员可操作'));
  var body = req.body || {};
  var targetUsername = (body.username || '').trim();
  if (!targetUsername) return res.json(err('请指定要下线的用户名'));
  if (targetUsername === cur.username) return res.json(err('不能强制自己下线'));
  var count = 0;
  for (var t in tokenStore) {
    if (tokenStore[t].username === targetUsername) {
      delete tokenStore[t];
      count++;
    }
  }
  res.json(ok({ message: '已强制下线', count: count }));
});

// 获取管理员列表（超级管理员可见，不返回密码）
app.get('/api/admin/list', function (req, res) {
  var cur = getCurrentUser(req);
  if (!cur || cur.role !== 'super_admin') return res.status(403).json(err('仅超级管理员可操作'));
  var data = loadAdmins();
  var list = data.admins.map(function (a) {
    return { id: a.id, username: a.username, role: a.role || 'admin' };
  });
  res.json(ok(list));
});

// 新增管理员
app.post('/api/admin/add', function (req, res) {
  var cur = getCurrentUser(req);
  if (!cur || cur.role !== 'super_admin') return res.status(403).json(err('仅超级管理员可操作'));
  var body = req.body || {};
  var username = (body.username || '').trim();
  var password = (body.password || '').trim();
  var role = (body.role || 'admin').trim();
  if (!username || !password) return res.json(err('请输入用户名和密码'));
  if (role !== 'admin' && role !== 'super_admin') role = 'admin';
  var data = loadAdmins();
  if (data.admins.some(function (a) { return a.username === username; })) return res.json(err('用户名已存在'));
  var id = data.nextId++;
  data.admins.push({ id: id, username: username, password: password, role: role });
  saveAdmins(data);
  res.json(ok({ id: id, username: username, role: role }));
});

// 修改管理员（密码、权限）
app.put('/api/admin/:id', function (req, res) {
  var cur = getCurrentUser(req);
  if (!cur || cur.role !== 'super_admin') return res.status(403).json(err('仅超级管理员可操作'));
  var id = parseInt(req.params.id, 10);
  if (isNaN(id)) return res.json(err('ID无效'));
  var body = req.body || {};
  var data = loadAdmins();
  var idx = data.admins.findIndex(function (a) { return a.id === id; });
  if (idx === -1) return res.json(err('管理员不存在'));
  var a = data.admins[idx];
  if (a.role === 'super_admin' && a.username !== cur.username) return res.json(err('不可修改其他超级管理员'));
  if (body.password !== undefined && body.password !== '') a.password = String(body.password).trim();
  if (body.role !== undefined) {
    var newRole = String(body.role).trim();
    if (newRole === 'admin' || newRole === 'super_admin') a.role = newRole;
  }
  saveAdmins(data);
  res.json(ok({ id: a.id, username: a.username, role: a.role }));
});

// 删除管理员
app.delete('/api/admin/:id', function (req, res) {
  var cur = getCurrentUser(req);
  if (!cur || cur.role !== 'super_admin') return res.status(403).json(err('仅超级管理员可操作'));
  var id = parseInt(req.params.id, 10);
  if (isNaN(id)) return res.json(err('ID无效'));
  var data = loadAdmins();
  var idx = data.admins.findIndex(function (a) { return a.id === id; });
  if (idx === -1) return res.json(err('管理员不存在'));
  var a = data.admins[idx];
  if (a.role === 'super_admin') return res.json(err('不可删除超级管理员'));
  data.admins.splice(idx, 1);
  saveAdmins(data);
  for (var t in tokenStore) {
    if (tokenStore[t].username === a.username) delete tokenStore[t];
  }
  res.json(ok({ deleted: id }));
});

const PROJECT_KEYS = ['year', 'name', 'department', 'members', 'supervise', 'status'];

// 获取项目列表
app.get('/api/projects', (req, res) => {
  try {
    const data = loadData();
    let list = data.projects;
    const year = req.query.year;
    if (year && year !== 'all') {
      const y = parseInt(year, 10);
      if (!isNaN(y)) list = list.filter(function (p) { return p.year === y; });
    }
    list = list.sort(function (a, b) {
      if (a.year !== b.year) return b.year - a.year;
      return (a.id || 0) - (b.id || 0);
    });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 导入 Excel 项目数据（需登录），body: { rows: [ { year, name, department, ... } ] }
function normalizeProjectRow(row) {
  var map = { '年度': 'year', '项目名称': 'name', '监督部门': 'department', '小组成员名单': 'members', '重点监督内容': 'supervise', '完成情况': 'status' };
  var out = { year: '', name: '', department: '', members: '', supervise: '', status: '' };
  for (var key in row) {
    if (row.hasOwnProperty(key)) {
      var k = map[key] || key;
      if (['year', 'name', 'department', 'members', 'supervise', 'status'].indexOf(k) !== -1) {
        var v = row[key] != null ? String(row[key]).trim() : '';
        if (k === 'year' && v) out[k] = parseInt(v, 10) || v;
        else out[k] = v;
      }
    }
  }
  if (!out.status || !['未开始', '进行中', '已完成'].includes(out.status)) out.status = '未开始';
  return out;
}
app.post('/api/projects/import', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var rows = body.rows;
    if (!Array.isArray(rows) || rows.length === 0) return res.json(err('请上传有效数据'));
    var data = loadData();
    var added = 0;
    rows.forEach(function (row) {
      var r = normalizeProjectRow(row);
      if (!r.year || !r.name || !r.department) return;
      var id = data.nextId++;
      data.projects.push({ id: id, year: parseInt(r.year, 10) || new Date().getFullYear(), name: r.name, department: r.department, members: r.members || '', supervise: r.supervise || '', status: r.status || '未开始' });
      added++;
    });
    saveData(data);
    res.json(ok({ imported: added }));
  } catch (e) {
    res.status(500).json(err(e.message || '导入失败'));
  }
});

// 新增项目（需登录）
app.post('/api/projects', requireAuth, function (req, res) {
  try {
    const body = req.body || {};
    const year = parseInt(body.year, 10);
    const name = (body.name || '').trim();
    const department = (body.department || '').trim();
    const members = (body.members || '').trim();
    const supervise = (body.supervise || '').trim();
    const status = (body.status || '').trim();
    if (!year || !name || !department || !members || !supervise || !status) {
      return res.json(err('缺少必填字段'));
    }
    const data = loadData();
    const id = data.nextId++;
    const row = { id: id, year: year, name: name, department: department, members: members, supervise: supervise, status: status };
    data.projects.push(row);
    saveData(data);
    res.json(ok(row));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 修改项目（需登录）
app.put('/api/projects/:id', requireAuth, function (req, res) {
  try {
    const id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    const body = req.body || {};
    const data = loadData();
    const idx = data.projects.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('项目不存在'));
    const p = data.projects[idx];
    p.year = parseInt(body.year, 10) || p.year;
    p.name = (body.name || '').trim() || p.name;
    p.department = (body.department || '').trim() || p.department;
    p.members = (body.members || '').trim() || p.members;
    p.supervise = (body.supervise || '').trim() || p.supervise;
    p.status = (body.status || '').trim() || p.status;
    saveData(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 删除单个项目（需登录）
app.delete('/api/projects/:id', requireAuth, function (req, res) {
  try {
    const id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    const data = loadData();
    const idx = data.projects.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('项目不存在'));
    data.projects.splice(idx, 1);
    saveData(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 调整排序（需登录）
app.post('/api/projects/reorder', requireAuth, function (req, res) {
  try {
    const body = req.body || {};
    const orderedIds = body.orderedIds;
    if (!Array.isArray(orderedIds) || orderedIds.length === 0) {
      return res.json(err('排序数据无效'));
    }
    const data = loadData();
    const map = {};
    data.projects.forEach(function (p) { map[p.id] = p; });
    const ordered = [];
    for (var i = 0; i < orderedIds.length; i++) {
      var pid = parseInt(orderedIds[i], 10);
      if (map[pid]) ordered.push(map[pid]);
    }
    data.projects = ordered;
    saveData(data);
    res.json(ok({}));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 批量删除（需登录）
app.post('/api/projects/batch-delete', requireAuth, function (req, res) {
  try {
    const body = req.body || {};
    const ids = body.ids;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.json(err('请选择要删除的项目'));
    }
    const idSet = {};
    ids.forEach(function (id) { idSet[parseInt(id, 10)] = true; });
    const data = loadData();
    data.projects = data.projects.filter(function (p) { return !idSet[p.id]; });
    saveData(data);
    res.json(ok({ deleted: ids.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 备份（需登录）
app.get('/api/backup', requireAuth, function (req, res) {
  try {
    const data = loadData();
    res.json(ok(data.projects));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 恢复（需登录）
app.post('/api/restore', requireAuth, function (req, res) {
  try {
    const body = req.body || {};
    const items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var projects = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      return {
        id: id,
        year: parseInt(item.year, 10) || 2025,
        name: (item.name || '').trim(),
        department: (item.department || '').trim(),
        members: (item.members || '').trim(),
        supervise: (item.supervise || '').trim(),
        status: (item.status || '未开始').trim()
      };
    });
    if (projects.length > 0) {
      var maxId = Math.max.apply(null, projects.map(function (p) { return p.id; }));
      nextId = maxId + 1;
    }
    saveData({ projects: projects, nextId: nextId });
    res.json(ok({ restored: items.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 监督方式排序 API（全局） ==========
function loadFormSort() {
  return readCache(FORM_SORT_FILE, function () {
    var dir = path.dirname(FORM_SORT_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(FORM_SORT_FILE)) {
      var defaultOrder = ['听取审议报告', '财经工作监督', '执法检查', '专题询问', '视察'];
      fs.writeFileSync(FORM_SORT_FILE, JSON.stringify({ order: defaultOrder }, null, 2));
      return { order: defaultOrder.slice() };
    }
    var raw = fs.readFileSync(FORM_SORT_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { order: Array.isArray(data.order) ? data.order : ['听取审议报告', '财经工作监督', '执法检查', '专题询问', '视察'] };
  });
}
function saveFormSort(data) {
  var dir = path.dirname(FORM_SORT_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FORM_SORT_FILE, JSON.stringify(data, null, 2));
  invalidateCache(FORM_SORT_FILE);
}
app.get('/api/form-sort', function (req, res) {
  try {
    var data = loadFormSort();
    res.json(ok(data.order));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});
app.put('/api/form-sort', requireAuth, function (req, res) {
  try {
    var order = req.body && req.body.order;
    if (!Array.isArray(order) || order.length !== 5) return res.json(err('排序数组必须包含5项'));
    var valid = ['听取审议报告', '财经工作监督', '执法检查', '专题询问', '视察'];
    var set = {};
    valid.forEach(function (v) { set[v] = true; });
    for (var i = 0; i < order.length; i++) {
      if (!set[order[i]]) return res.json(err('包含无效的监督形式'));
    }
    saveFormSort({ order: order });
    res.json(ok({ order: order }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 监督部门管理 API ==========
app.get('/api/departments', function (req, res) {
  try {
    var data = loadDepartments();
    var list = (data.items || []).slice().sort(function (a, b) { return (a.id || 0) - (b.id || 0); });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/departments', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var id = parseInt(body.id, 10);
    var name = (body.name || '').trim();
    if (!name) return res.json(err('部门名称不能为空'));
    var data = loadDepartments();
    if (data.items.some(function (d) { return d.name === name; })) return res.json(err('部门名称已存在'));
    if (isNaN(id) || id <= 0 || data.items.some(function (d) { return d.id === id; })) id = data.nextId++;
    data.items.push({ id: id, name: name });
    data.items.sort(function (a, b) { return (a.id || 0) - (b.id || 0); });
    if (data.nextId <= id) data.nextId = id + 1;
    saveDepartments(data);
    res.json(ok({ id: id, name: name }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.put('/api/departments/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var newId = parseInt(body.id, 10);
    var newName = (body.name || '').trim();
    var data = loadDepartments();
    var idx = data.items.findIndex(function (d) { return d.id === id; });
    if (idx === -1) return res.json(err('部门不存在'));
    var oldName = data.items[idx].name;
    if (newName) {
      data.items[idx].name = newName;
      var projData = loadData();
      projData.projects.forEach(function (p) {
        if (p.department === oldName) p.department = newName;
      });
      saveData(projData);
      var progData = loadMinshengProgress();
      progData.items.forEach(function (p) {
        if (p.department === oldName) p.department = newName;
      });
      saveMinshengProgress(progData);
      var jianduData = loadJianduTopics();
      jianduData.items.forEach(function (p) {
        if (p.department === oldName) p.department = newName;
      });
      saveJianduTopics(jianduData);
    }
    if (!isNaN(newId) && newId > 0 && newId !== id) {
      data.items[idx].id = newId;
    }
    data.items.sort(function (a, b) { return (a.id || 0) - (b.id || 0); });
    saveDepartments(data);
    res.json(ok(data.items[idx]));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.delete('/api/departments/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadDepartments();
    var idx = data.items.findIndex(function (d) { return d.id === id; });
    if (idx === -1) return res.json(err('部门不存在'));
    var name = data.items[idx].name;
    var projData = loadData();
    var hasProj = projData.projects.some(function (p) { return p.department === name; });
    var progData = loadMinshengProgress();
    var hasProg = progData.items.some(function (p) { return p.department === name; });
    var jianduData = loadJianduTopics();
    var hasJiandu = (jianduData.items || []).some(function (p) { return p.department === name; });
    if (hasProj || hasProg || hasJiandu) return res.json(err('该部门下存在项目、进展资料或监督议题，无法删除'));
    data.items.splice(idx, 1);
    saveDepartments(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 领导批示督办 API ==========

function loadPishi() {
  return readCache(PISHI_FILE, function() {
    var dir = path.dirname(PISHI_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(PISHI_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(PISHI_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(PISHI_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function savePishi(data) {
  var dir = path.dirname(PISHI_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(PISHI_FILE, JSON.stringify(data, null, 2));
  invalidateCache(PISHI_FILE);
}

// Excel 日期序列号转 YYYY-MM-DD（Excel 1900 日期系统）
function excelDateToYMD(val) {
  if (val == null || val === '') return val;
  var n = typeof val === 'number' ? val : parseFloat(String(val).trim());
  if (isNaN(n) || n < 1 || n > 2958465) return val;
  if (n >= 60) n -= 1;
  var d = new Date((n - 25569) * 86400 * 1000);
  if (isNaN(d.getTime())) return val;
  var y = d.getUTCFullYear(), m = d.getUTCMonth() + 1, day = d.getUTCDate();
  var pad = function (x) { return (x < 10 ? '0' : '') + x; };
  return y + '-' + pad(m) + '-' + pad(day);
}

function formatPishiDate(val) {
  if (!val) return val;
  var s = String(val).trim();
  var n = parseFloat(s);
  if (!isNaN(n) && n >= 1 && n <= 2958465 && s.match(/^\d+$/)) return excelDateToYMD(n);
  return val;
}

// 将 Excel 行对象（中文表头或英文字段）转为统一格式
function normalizePishiRow(row) {
  var map = {
    '期数': 'qishu', '来文单位': 'laiwenUnit', '文号': 'wenhao', '批示内容': 'pishiContent',
    '批示日期': 'pishiDate', '领导和责任部门': 'leaderDept', '类别': 'category',
    '落实举措': 'luoshiCuoshi', '完成情况': 'completeStatus'
  };
  var out = {};
  PISHI_KEYS.forEach(function (k) { out[k] = ''; });
  for (var key in row) {
    if (row.hasOwnProperty(key)) {
      var k = map[key] || key;
      if (PISHI_KEYS.indexOf(k) !== -1) {
        var v = row[key] != null ? String(row[key]).trim() : '';
        if (k === 'pishiDate' && v) v = formatPishiDate(row[key]);
        out[k] = v;
      }
    }
  }
  if (!out.completeStatus || (out.completeStatus !== '已完成' && out.completeStatus !== '推进中')) {
    out.completeStatus = '推进中';
  }
  return out;
}

// 获取批示列表（支持筛选：期数、批示日期、领导和责任部门、完成情况）
app.get('/api/pishi', function (req, res) {
  try {
    var data = loadPishi();
    var list = data.items;
    var qishu = (req.query.qishu || '').trim();
    var pishiDate = (req.query.pishiDate || '').trim();
    var leaderDept = (req.query.leaderDept || '').trim();
    var completeStatus = (req.query.completeStatus || '').trim();
    if (qishu) list = list.filter(function (p) { return String(p.qishu) === qishu; });
    if (pishiDate) list = list.filter(function (p) { return String(formatPishiDate(p.pishiDate) || '').indexOf(pishiDate) !== -1; });
    if (leaderDept) list = list.filter(function (p) { return String(p.leaderDept).indexOf(leaderDept) !== -1; });
    if (completeStatus) list = list.filter(function (p) { return p.completeStatus === completeStatus; });
    list = list.map(function (p) {
      var v = { ...p };
      v.pishiDate = formatPishiDate(p.pishiDate) || p.pishiDate;
      return v;
    }).sort(function (a, b) {
      if (a.pishiDate !== b.pishiDate) return (b.pishiDate || '').localeCompare(a.pishiDate || '');
      return (b.id || 0) - (a.id || 0);
    });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 按“领导和责任部门”统计数量（用于数据概览图表）
app.get('/api/pishi/stats', function (req, res) {
  try {
    var data = loadPishi();
    var stats = {};
    data.items.forEach(function (p) {
      var dept = (p.leaderDept || '未分类').trim() || '未分类';
      stats[dept] = (stats[dept] || 0) + 1;
    });
    res.json(ok(stats));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 新增批示（需登录）
app.post('/api/pishi', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var row = normalizePishiRow(body);
    var data = loadPishi();
    var id = data.nextId++;
    var item = { id: id, qishu: row.qishu, laiwenUnit: row.laiwenUnit, wenhao: row.wenhao, pishiContent: row.pishiContent, pishiDate: row.pishiDate, leaderDept: row.leaderDept, category: row.category, luoshiCuoshi: row.luoshiCuoshi, completeStatus: row.completeStatus };
    data.items.push(item);
    savePishi(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 修改批示（需登录）
app.put('/api/pishi/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadPishi();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var row = normalizePishiRow(body);
    var p = data.items[idx];
    p.qishu = row.qishu; p.laiwenUnit = row.laiwenUnit; p.wenhao = row.wenhao; p.pishiContent = row.pishiContent;
    p.pishiDate = row.pishiDate; p.leaderDept = row.leaderDept; p.category = row.category; p.luoshiCuoshi = row.luoshiCuoshi;
    p.completeStatus = row.completeStatus;
    savePishi(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 删除批示（需登录）
app.delete('/api/pishi/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadPishi();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    data.items.splice(idx, 1);
    savePishi(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 批量删除批示（需登录）
app.post('/api/pishi/batch-delete', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var ids = body.ids;
    if (!Array.isArray(ids) || ids.length === 0) return res.json(err('请选择要删除的记录'));
    var idSet = {};
    ids.forEach(function (id) { idSet[parseInt(id, 10)] = true; });
    var data = loadPishi();
    data.items = data.items.filter(function (p) { return !idSet[p.id]; });
    savePishi(data);
    res.json(ok({ deleted: ids.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 批示督办备份（需登录），返回全部批示数据
app.get('/api/pishi/backup', requireAuth, function (req, res) {
  try {
    var data = loadPishi();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 批示督办恢复（需登录），body: { data: [ items ] }
app.post('/api/pishi/restore', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var list = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      return {
        id: id,
        qishu: (item.qishu || '').trim(),
        laiwenUnit: (item.laiwenUnit || '').trim(),
        wenhao: (item.wenhao || '').trim(),
        pishiContent: (item.pishiContent || '').trim(),
        pishiDate: (item.pishiDate || '').trim(),
        leaderDept: (item.leaderDept || '').trim(),
        category: (item.category || '').trim(),
        luoshiCuoshi: (item.luoshiCuoshi || '').trim(),
        completeStatus: (item.completeStatus === '已完成' ? '已完成' : '推进中')
      };
    });
    if (list.length > 0) {
      var maxId = Math.max.apply(null, list.map(function (p) { return p.id; }));
      nextId = maxId + 1;
    }
    savePishi({ items: list, nextId: nextId });
    res.json(ok({ restored: list.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 导入 Excel 数据（需登录），body: { rows: [ { 期数, 来文单位, ... } ] }
app.post('/api/pishi/import', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var rows = body.rows;
    if (!Array.isArray(rows) || rows.length === 0) return res.json(err('请上传有效数据'));
    var data = loadPishi();
    var added = 0;
    rows.forEach(function (row) {
      var r = normalizePishiRow(row);
      if (!r.qishu && !r.pishiContent && !r.leaderDept) return;
      var id = data.nextId++;
      data.items.push({ id: id, qishu: r.qishu, laiwenUnit: r.laiwenUnit, wenhao: r.wenhao, pishiContent: r.pishiContent, pishiDate: r.pishiDate, leaderDept: r.leaderDept, category: r.category, luoshiCuoshi: r.luoshiCuoshi, completeStatus: r.completeStatus });
      added++;
    });
    savePishi(data);
    res.json(ok({ imported: added }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 常委会每月重点工作 API ==========

function loadMeiyue() {
  return readCache(MEIYUE_FILE, function() {
    var dir = path.dirname(MEIYUE_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(MEIYUE_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(MEIYUE_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(MEIYUE_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function saveMeiyue(data) {
  var dir = path.dirname(MEIYUE_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(MEIYUE_FILE, JSON.stringify(data, null, 2));
  invalidateCache(MEIYUE_FILE);
}

function parseYearFromMonth(monthStr) {
  if (!monthStr) return '';
  var s = String(monthStr).trim();
  var m = s.match(/^(\d{4})/);
  return m ? m[1] : s;
}

app.get('/api/meiyue/file/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMeiyue();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p || !p.filePath) return res.status(404).send('文件不存在');
    var fp = path.join(MEIYUE_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.setHeader('Content-Type', 'application/pdf');
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/meiyue/download/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMeiyue();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p || !p.filePath) return res.status(404).send('文件不存在');
    var fp = path.join(MEIYUE_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var baseName = (p.fileName || p.content || 'document').replace(/[/\\:*?"<>|]/g, '_').trim() || 'document';
    var ext = path.extname(p.filePath) || '.pdf';
    var fname = baseName + (ext.charAt(0) === '.' ? ext : '.' + ext);
    var fnameEnc = encodeURIComponent(fname);
    res.setHeader('Content-Disposition', 'attachment; filename="download' + ext + '"; filename*=UTF-8\'\'' + fnameEnc);
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.post('/api/meiyue/upload', requireAuth, function (req, res, next) {
  meiyueUpload.single('file')(req, res, function (multerErr) {
    if (multerErr) return res.json(err(multerErr.message || '文件上传失败'));
    next();
  });
}, function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择文件'));
    var rawName = req.file.originalname || '';
    try { rawName = Buffer.from(rawName, 'latin1').toString('utf8'); } catch (e) {}
    var body = req.body || {};
    var year = (body.year || '').trim();
    var month = (body.month || '').trim();
    var data = loadMeiyue();
    var id = data.nextId++;
    var ext = (path.extname(rawName) || '').toLowerCase() || '.pdf';
    var item = {
      id: id,
      year: year,
      month: month,
      filePath: path.basename(req.file.path),
      fileName: rawName.replace(/\.pdf$/i, '') || rawName,
      fileType: ext.replace(/^\./, '') || 'pdf'
    };
    data.items.push(item);
    saveMeiyue(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meiyue/:id/replace-file', requireAuth, meiyueUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var rawName = req.file.originalname || '';
    try { rawName = Buffer.from(rawName, 'latin1').toString('utf8'); } catch (e) {}
    var body = req.body || {};
    var data = loadMeiyue();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (p.filePath) {
      var oldFp = path.join(MEIYUE_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(oldFp)) try { fs.unlinkSync(oldFp); } catch (e) {}
    }
    p.filePath = path.basename(req.file.path);
    var fn = (body.fileName !== undefined && body.fileName) ? body.fileName : (rawName ? rawName.replace(/\.pdf$/i, '') : '') || p.fileName || '';
    p.fileName = String(fn).trim();
    p.fileType = (path.extname(rawName) || '.pdf').toLowerCase().replace(/^\./, '') || 'pdf';
    if (body.year !== undefined) p.year = String(body.year || '').trim();
    if (body.month !== undefined) p.month = String(body.month || '').trim();
    saveMeiyue(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/meiyue', function (req, res) {
  try {
    var data = loadMeiyue();
    var list = data.items;
    var year = (req.query.year || '').trim();
    var month = (req.query.month || '').trim();
    if (year) list = list.filter(function (p) { return String(p.year || parseYearFromMonth(p.month)) === year; });
    if (month) list = list.filter(function (p) { return String(p.month) === month; });
    list = list.slice().sort(function (a, b) {
      var ya = String(a.year || parseYearFromMonth(a.month) || '');
      var yb = String(b.year || parseYearFromMonth(b.month) || '');
      if (ya !== yb) return yb.localeCompare(ya);
      var c = (b.month || '').localeCompare(a.month || '');
      if (c !== 0) return c;
      return (b.id || 0) - (a.id || 0);
    });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meiyue', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var month = (body.month || '').trim();
    var year = (body.year || '').trim() || parseYearFromMonth(month);
    var data = loadMeiyue();
    var id = data.nextId++;
    var item = {
      id: id,
      year: year,
      month: month,
      time1: (body.time1 || '').trim(),
      time2: (body.time2 || '').trim(),
      location: (body.location || '').trim(),
      content: (body.content || '').trim(),
      leaders: (body.leaders || '').trim(),
      remark: (body.remark || '').trim()
    };
    data.items.push(item);
    saveMeiyue(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.put('/api/meiyue/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadMeiyue();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    p.month = String(body.month !== undefined ? body.month : p.month || '').trim();
    p.year = String(body.year !== undefined ? body.year : p.year || parseYearFromMonth(p.month) || '').trim() || parseYearFromMonth(p.month);
    if (body.fileName !== undefined) p.fileName = String(body.fileName || '').trim();
    p.time1 = String(body.time1 !== undefined ? body.time1 : p.time1 || '').trim();
    p.time2 = String(body.time2 !== undefined ? body.time2 : p.time2 || '').trim();
    p.location = String(body.location !== undefined ? body.location : p.location || '').trim();
    p.content = String(body.content !== undefined ? body.content : p.content || '').trim();
    p.leaders = String(body.leaders !== undefined ? body.leaders : p.leaders || '').trim();
    p.remark = String(body.remark !== undefined ? body.remark : p.remark || '').trim();
    saveMeiyue(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.delete('/api/meiyue/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadMeiyue();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (p.filePath) {
      var fp = path.join(MEIYUE_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
    }
    data.items.splice(idx, 1);
    saveMeiyue(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meiyue/batch-delete', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var ids = body.ids;
    if (!Array.isArray(ids) || ids.length === 0) return res.json(err('请选择要删除的记录'));
    var idSet = {};
    ids.forEach(function (id) { idSet[parseInt(id, 10)] = true; });
    var data = loadMeiyue();
    data.items = data.items.filter(function (p) {
      if (idSet[p.id]) {
        if (p.filePath) {
          var fp = path.join(MEIYUE_UPLOAD_DIR, p.filePath);
          if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
        }
        return false;
      }
      return true;
    });
    saveMeiyue(data);
    res.json(ok({ deleted: ids.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/meiyue/backup', requireAuth, function (req, res) {
  try {
    var data = loadMeiyue();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meiyue/restore', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var list = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      var month = (item.month || '').trim();
      var year = (item.year || '').trim() || parseYearFromMonth(month);
      var o = {
        id: id,
        year: year,
        month: month,
        time1: (item.time1 || '').trim(),
        time2: (item.time2 || '').trim(),
        location: (item.location || '').trim(),
        content: (item.content || '').trim(),
        leaders: (item.leaders || '').trim(),
        remark: (item.remark || '').trim()
      };
      if (item.filePath) { o.filePath = item.filePath; o.fileName = (item.fileName || '').trim(); o.fileType = (item.fileType || 'pdf').toLowerCase(); }
      return o;
    });
    if (list.length > 0) {
      var maxId = Math.max.apply(null, list.map(function (p) { return p.id; }));
      nextId = maxId + 1;
    }
    saveMeiyue({ items: list, nextId: nextId });
    res.json(ok({ restored: list.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meiyue/import', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var rows = body.rows;
    if (!Array.isArray(rows) || rows.length === 0) return res.json(err('请上传有效数据'));
    var data = loadMeiyue();
    var map = { '年度': 'year', '月份': 'month', '时间1': 'time1', '时间2': 'time2', '地点': 'location', '内容': 'content', '出席领导': 'leaders', '备注': 'remark' };
    var added = 0;
    rows.forEach(function (row) {
      var year = (row.year || row.年度 || '').toString().trim() || parseYearFromMonth(row.month || row.月份);
      var month = (row.month || row.月份 || '').toString().trim();
      if (!month && !year) return;
      var id = data.nextId++;
      data.items.push({
        id: id,
        year: year,
        month: month,
        time1: (row.time1 || row.时间1 || '').toString().trim(),
        time2: (row.time2 || row.时间2 || '').toString().trim(),
        location: (row.location || row.地点 || '').toString().trim(),
        content: (row.content || row.内容 || '').toString().trim(),
        leaders: (row.leaders || row.出席领导 || '').toString().trim(),
        remark: (row.remark || row.备注 || '').toString().trim()
      });
      added++;
    });
    saveMeiyue(data);
    res.json(ok({ imported: added }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 每周工作 API ==========

function loadMeizhou() {
  return readCache(MEIZHOU_FILE, function() {
    var dir = path.dirname(MEIZHOU_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(MEIZHOU_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(MEIZHOU_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(MEIZHOU_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function saveMeizhou(data) {
  var dir = path.dirname(MEIZHOU_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(MEIZHOU_FILE, JSON.stringify(data, null, 2));
  invalidateCache(MEIZHOU_FILE);
}

function fixMeizhouFileName(name) {
  if (!name || typeof name !== 'string') return name;
  if (/[\u4e00-\u9fff]/.test(name)) return name;
  try {
    var decoded = Buffer.from(name, 'latin1').toString('utf8');
    if (/[\u4e00-\u9fff]/.test(decoded)) return decoded;
  } catch (e) {}
  return name;
}

app.get('/api/meizhou/file/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMeizhou();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p || !p.filePath) return res.status(404).send('文件不存在');
    var fp = path.join(MEIZHOU_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.setHeader('Content-Type', 'application/pdf');
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/meizhou/download/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMeizhou();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p || !p.filePath) return res.status(404).send('文件不存在');
    var fp = path.join(MEIZHOU_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var baseName = (p.fileName || p.content || 'document').replace(/[/\\:*?"<>|]/g, '_').trim() || 'document';
    var ext = path.extname(p.filePath) || '.pdf';
    var fname = baseName + (ext.charAt(0) === '.' ? ext : '.' + ext);
    var fnameEnc = encodeURIComponent(fname);
    res.setHeader('Content-Disposition', 'attachment; filename="download' + ext + '"; filename*=UTF-8\'\'' + fnameEnc);
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/meizhou', function (req, res) {
  try {
    var data = loadMeizhou();
    var list = data.items;
    var yearQ = req.query.year;
    var monthQ = req.query.month;
    var weekQ = req.query.week;
    var years = Array.isArray(yearQ) ? yearQ : (yearQ ? String(yearQ).split(',').map(function (s) { return s.trim(); }).filter(Boolean) : []);
    var months = Array.isArray(monthQ) ? monthQ : (monthQ ? String(monthQ).split(',').map(function (s) { return s.trim(); }).filter(Boolean) : []);
    var weeks = Array.isArray(weekQ) ? weekQ : (weekQ ? String(weekQ).split(',').map(function (s) { return s.trim(); }).filter(Boolean) : []);
    if (years.length) list = list.filter(function (p) { return years.indexOf(String(p.year || '')) !== -1; });
    if (months.length) list = list.filter(function (p) { return months.indexOf(String(p.month || '')) !== -1; });
    if (weeks.length) list = list.filter(function (p) { return weeks.indexOf(String(p.week || '')) !== -1; });
    list = list.slice().sort(function (a, b) {
      var ya = String(a.year || '');
      var yb = String(b.year || '');
      if (ya !== yb) return yb.localeCompare(ya);
      var ma = String(a.month || '');
      var mb = String(b.month || '');
      if (ma !== mb) return mb.localeCompare(ma);
      var wa = String(a.week || '');
      var wb = String(b.week || '');
      if (wa !== wb) return wb.localeCompare(wa);
      return (b.id || 0) - (a.id || 0);
    });
    list = list.map(function (p) {
      var q = Object.assign({}, p);
      if (q.fileName) q.fileName = fixMeizhouFileName(q.fileName);
      return q;
    });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var data = loadMeizhou();
    var id = data.nextId++;
    var item = {
      id: id,
      year: (body.year || '').trim(),
      month: (body.month || '').trim(),
      week: (body.week || '').trim(),
      department: (body.department || '').trim(),
      content: (body.content || '').trim(),
      remark: (body.remark || '').trim()
    };
    data.items.push(item);
    saveMeizhou(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou/upload', requireAuth, function (req, res, next) {
  meizhouUpload.single('file')(req, res, function (multerErr) {
    if (multerErr) return res.json(err(multerErr.message || '文件上传失败'));
    next();
  });
}, function (req, res) {
  try {
    var year = (req.body.year || '').toString().trim();
    var month = (req.body.month || '').toString().trim();
    var week = (req.body.week || '').toString().trim();
    if (!req.file) return res.json(err('请选择 PDF 文件'));
    var rawName = req.file.originalname || '';
    try {
      rawName = Buffer.from(rawName, 'latin1').toString('utf8');
    } catch (e) {}
    var data = loadMeizhou();
    var id = data.nextId++;
    var ext = (path.extname(rawName) || '').toLowerCase() || '.pdf';
    var item = {
      id: id,
      year: year,
      month: month,
      week: week,
      filePath: path.basename(req.file.path),
      fileName: rawName,
      fileType: ext.replace(/^\./, '') || 'pdf'
    };
    data.items.push(item);
    saveMeizhou(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou/:id/replace-file', requireAuth, meizhouUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var rawName = req.file.originalname || '';
    try { rawName = Buffer.from(rawName, 'latin1').toString('utf8'); } catch (e) {}
    var body = req.body || {};
    var data = loadMeizhou();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (p.filePath) {
      var oldFp = path.join(MEIZHOU_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(oldFp)) try { fs.unlinkSync(oldFp); } catch (e) {}
    }
    var ext = (path.extname(rawName) || '').toLowerCase() || '.pdf';
    p.filePath = path.basename(req.file.path);
    p.fileName = (body.fileName !== undefined && body.fileName ? body.fileName : rawName.replace(/\.(pdf|doc|docx)$/i, '') || p.fileName).toString().trim();
    p.fileType = ext.replace(/^\./, '') || 'pdf';
    if (body.year !== undefined) p.year = body.year.toString().trim();
    if (body.month !== undefined) p.month = body.month.toString().trim();
    if (body.week !== undefined) p.week = body.week.toString().trim();
    saveMeizhou(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.put('/api/meizhou/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadMeizhou();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    p.year = (body.year !== undefined ? body.year : p.year).toString().trim();
    p.month = (body.month !== undefined ? body.month : p.month).toString().trim();
    p.week = (body.week !== undefined ? body.week : p.week).toString().trim();
    if (body.fileName !== undefined) p.fileName = body.fileName.toString().trim();
    p.department = (body.department !== undefined ? body.department : p.department).toString().trim();
    p.content = (body.content !== undefined ? body.content : p.content).toString().trim();
    p.remark = (body.remark !== undefined ? body.remark : p.remark).toString().trim();
    saveMeizhou(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.delete('/api/meizhou/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadMeizhou();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (p.filePath) {
      var fp = path.join(MEIZHOU_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
    }
    data.items.splice(idx, 1);
    saveMeizhou(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou/batch-delete', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var ids = body.ids;
    if (!Array.isArray(ids) || ids.length === 0) return res.json(err('请选择要删除的记录'));
    var idSet = {};
    ids.forEach(function (id) { idSet[parseInt(id, 10)] = true; });
    var data = loadMeizhou();
    data.items = data.items.filter(function (p) {
      if (idSet[p.id]) {
        if (p.filePath) {
          var fp = path.join(MEIZHOU_UPLOAD_DIR, p.filePath);
          if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
        }
        return false;
      }
      return true;
    });
    saveMeizhou(data);
    res.json(ok({ deleted: ids.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/meizhou/backup', requireAuth, function (req, res) {
  try {
    var data = loadMeizhou();
    var items = (data.items || []).map(function (p) {
      var q = Object.assign({}, p);
      if (q.fileName) q.fileName = fixMeizhouFileName(q.fileName);
      return q;
    });
    res.json(ok(items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/meizhou/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadMeizhou();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).json(err('记录不存在'));
    var out = Object.assign({}, p);
    if (out.fileName) out.fileName = fixMeizhouFileName(out.fileName);
    res.json(ok(out));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou/restore', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var list = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      var obj = {
        id: id,
        year: (item.year || '').trim(),
        month: (item.month || '').trim(),
        week: (item.week || '').trim(),
        department: (item.department || '').trim(),
        content: (item.content || '').trim(),
        remark: (item.remark || '').trim()
      };
      if (item.filePath) obj.filePath = item.filePath;
      if (item.fileName) obj.fileName = fixMeizhouFileName(item.fileName);
      if (item.fileType) obj.fileType = item.fileType;
      return obj;
    });
    if (list.length > 0) {
      var maxId = Math.max.apply(null, list.map(function (p) { return p.id; }));
      nextId = maxId + 1;
    }
    saveMeizhou({ items: list, nextId: nextId });
    res.json(ok({ restored: list.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/meizhou/import', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var rows = body.rows;
    if (!Array.isArray(rows) || rows.length === 0) return res.json(err('请上传有效数据'));
    var data = loadMeizhou();
    var map = { '年度': 'year', '月份': 'month', '周数': 'week', '部门': 'department', '备注': 'remark' };
    var added = 0;
    rows.forEach(function (row) {
      var year = (row.year || row.年度 || '').toString().trim();
      var month = (row.month || row.月份 || '').toString().trim();
      var week = (row.week || row.周数 || '').toString().trim();
      if (!year && !month && !week) return;
      var id = data.nextId++;
      data.items.push({
        id: id,
        year: year,
        month: month,
        week: week,
        department: (row.department || row.部门 || '').toString().trim(),
        content: (row.content || row.内容 || '').toString().trim(),
        remark: (row.remark || row.备注 || '').toString().trim()
      });
      added++;
    });
    saveMeizhou(data);
    res.json(ok({ imported: added }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// ========== 法律法规 API ==========

var falvStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(FALV_UPLOAD_DIR)) fs.mkdirSync(FALV_UPLOAD_DIR, { recursive: true });
    cb(null, FALV_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'falv_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var falvUpload = multer({
  storage: falvStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf', '.doc', '.docx'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF、DOC、DOCX 格式'));
  }
});

var minshengStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(MINSHENG_UPLOAD_DIR)) fs.mkdirSync(MINSHENG_UPLOAD_DIR, { recursive: true });
    cb(null, MINSHENG_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'minsheng_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var minshengUpload = multer({
  storage: minshengStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf', '.doc', '.docx'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF、DOC、DOCX 格式'));
  }
});

var pishiReportStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(PISHI_UPLOAD_DIR)) fs.mkdirSync(PISHI_UPLOAD_DIR, { recursive: true });
    cb(null, PISHI_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'pishi_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var pishiReportUpload = multer({
  storage: pishiReportStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf', '.doc', '.docx'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF、DOC、DOCX 格式'));
  }
});

// ========== 民生实事进展资料 API ==========
app.get('/api/minsheng/progress', function (req, res) {
  try {
    var data = loadMinshengProgress();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/minsheng/progress/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).json(err('Not found'));
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).json(err('Not found'));
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

function decodeFilename(str) {
  if (!str || typeof str !== 'string') return str || '';
  try { return Buffer.from(str, 'latin1').toString('utf8'); } catch (e) { return str; }
}

app.post('/api/minsheng/progress/upload', requireAuth, minshengUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var rawName = decodeFilename(req.file.originalname || '');
    var body = req.body || {};
    var department = (body.department || '').trim();
    var title = (body.title || '').trim() || rawName;
    var uploadDate = (body.uploadDate || '').trim();
    if (!uploadDate) {
      var d = new Date();
      uploadDate = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    if (!department) return res.json(err('请填写监督部门'));
    var ext = (path.extname(rawName) || '').toLowerCase();
    var fileType = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : (ext === '.doc' ? 'doc' : 'pdf'));
    var data = loadMinshengProgress();
    var id = data.nextId++;
    var item = {
      id: id,
      department: department,
      title: title,
      filePath: req.file.filename,
      originalName: rawName,
      fileType: fileType,
      uploadDate: uploadDate,
      createdAt: new Date().toISOString()
    };
    data.items.push(item);
    saveMinshengProgress(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message || '上传失败'));
  }
});

app.put('/api/minsheng/progress/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.json(err('进展资料不存在'));
    if (body.title !== undefined) p.title = String(body.title || '').trim() || p.originalName;
    if (body.department !== undefined) p.department = String(body.department || '').trim();
    if (body.uploadDate !== undefined) p.uploadDate = String(body.uploadDate || '').trim();
    if (body.pinned !== undefined) p.pinned = !!body.pinned;
    saveMinshengProgress(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/minsheng/progress/:id/replace-file', requireAuth, minshengUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var rawName = decodeFilename(req.file.originalname || '');
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.json(err('进展资料不存在'));
    if (p.filePath) {
      var oldFp = path.join(MINSHENG_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(oldFp)) try { fs.unlinkSync(oldFp); } catch (e) {}
    }
    var ext = (path.extname(rawName) || '').toLowerCase();
    var fileType = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : (ext === '.doc' ? 'doc' : 'pdf'));
    p.filePath = req.file.filename;
    p.originalName = rawName;
    p.fileType = fileType;
    if ((body.title || '').trim()) p.title = String(body.title).trim();
    else p.title = rawName;
    if ((body.uploadDate || '').trim()) p.uploadDate = String(body.uploadDate).trim();
    if ((body.department || '').trim()) p.department = String(body.department).trim();
    saveMinshengProgress(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message || '重新上传失败'));
  }
});

app.delete('/api/minsheng/progress/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadMinshengProgress();
    var idx = data.items.findIndex(function (x) { return x.id === id; });
    if (idx === -1) return res.json(err('进展资料不存在'));
    var p = data.items[idx];
    data.items.splice(idx, 1);
    saveMinshengProgress(data);
    var fp = path.join(MINSHENG_UPLOAD_DIR, p.filePath);
    if (fs.existsSync(fp)) {
      try { fs.unlinkSync(fp); } catch (e) {}
    }
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/minsheng/progress/file/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(MINSHENG_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.sendFile(path.resolve(fp), { headers: { 'Content-Disposition': 'inline' } });
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/minsheng/progress/download/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(MINSHENG_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var name = (p.originalName || p.filePath || 'download');
    var encoded = encodeURIComponent(name);
    res.setHeader('Content-Disposition', 'attachment; filename="download"; filename*=UTF-8\'\'' + encoded);
    res.sendFile(path.resolve(fp));
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/minsheng/progress/preview/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadMinshengProgress();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(MINSHENG_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    if (p.fileType === 'pdf') {
      res.redirect('/api/minsheng/progress/file/' + id);
      return;
    }
    if (['doc', 'docx'].indexOf(p.fileType) !== -1) {
      var styleMap = [
        'p[style-name="标题 1"] => h1:fresh',
        'p[style-name="标题 2"] => h2:fresh',
        'p[style-name="标题 3"] => h3:fresh',
        'p[style-name="Heading 1"] => h1:fresh',
        'p[style-name="Heading 2"] => h2:fresh',
        'p[style-name="Heading 3"] => h3:fresh',
        'p[style-name="正文"] => p:fresh',
        'r[style-name="强调"] => strong',
        'r[style-name="标题 1 Char"] => strong'
      ].join('\n');
      mammoth.convertToHtml({
        path: fp,
        styleMap: styleMap,
        includeDefaultStyleMap: true
      }).then(function (result) {
        var css = '<style>body{font-family:SimSun,serif;font-size:16px;line-height:1.8;margin:0;padding:24px 48px}'
          + 'p{margin:0.5em 0;text-indent:2em}p:first-child{text-indent:0}'
          + 'h1,h2,h3{margin:1em 0 0.5em;font-weight:bold}table{border-collapse:collapse;width:100%;margin:1em 0}'
          + 'td,th{border:1px solid #333;padding:6px 10px;text-align:left}</style>';
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send('<!DOCTYPE html><html><head><meta charset="utf-8">' + css + '</head><body>' + result.value + '</body></html>');
      }).catch(function (err) {
        res.status(500).send('Word 文档预览失败：' + (err.message || '未知错误'));
      });
      return;
    }
    res.status(400).send('不支持预览');
  } catch (e) {
    res.status(500).send('Error');
  }
});

// ========== 批示办理情况报告 API ==========
app.get('/api/pishi/report', function (req, res) {
  try {
    var data = loadPishiReport();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/pishi/report/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).json(err('Not found'));
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).json(err('Not found'));
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/pishi/report/upload', requireAuth, pishiReportUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var rawName = decodeFilename(req.file.originalname || '');
    var body = req.body || {};
    var title = (body.title || '').trim() || rawName;
    var uploadDate = (body.uploadDate || '').trim();
    if (!uploadDate) {
      var d = new Date();
      uploadDate = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    var ext = (path.extname(rawName) || '').toLowerCase();
    var fileType = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : (ext === '.doc' ? 'doc' : 'pdf'));
    var data = loadPishiReport();
    var id = data.nextId++;
    var item = { id: id, title: title, filePath: req.file.filename, originalName: rawName, fileType: fileType, uploadDate: uploadDate, createdAt: new Date().toISOString() };
    data.items.push(item);
    savePishiReport(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message || '上传失败'));
  }
});

app.put('/api/pishi/report/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.json(err('报告不存在'));
    if (body.title !== undefined) p.title = String(body.title || '').trim() || p.originalName;
    if (body.uploadDate !== undefined) p.uploadDate = String(body.uploadDate || '').trim();
    savePishiReport(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/pishi/report/:id/replace-file', requireAuth, pishiReportUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var rawName = decodeFilename(req.file.originalname || '');
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.json(err('报告不存在'));
    if (p.filePath) {
      var oldFp = path.join(PISHI_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(oldFp)) try { fs.unlinkSync(oldFp); } catch (e) {}
    }
    var ext = (path.extname(rawName) || '').toLowerCase();
    var fileType = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : (ext === '.doc' ? 'doc' : 'pdf'));
    p.filePath = req.file.filename;
    p.originalName = rawName;
    p.fileType = fileType;
    if ((body.title || '').trim()) p.title = String(body.title).trim();
    else p.title = rawName;
    if ((body.uploadDate || '').trim()) p.uploadDate = String(body.uploadDate).trim();
    savePishiReport(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message || '重新上传失败'));
  }
});

app.delete('/api/pishi/report/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadPishiReport();
    var idx = data.items.findIndex(function (x) { return x.id === id; });
    if (idx === -1) return res.json(err('报告不存在'));
    var p = data.items[idx];
    data.items.splice(idx, 1);
    savePishiReport(data);
    var fp = path.join(PISHI_UPLOAD_DIR, p.filePath);
    if (fs.existsSync(fp)) { try { fs.unlinkSync(fp); } catch (e) {} }
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/pishi/report/file/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(PISHI_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.sendFile(path.resolve(fp), { headers: { 'Content-Disposition': 'inline' } });
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/pishi/report/download/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(PISHI_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var name = (p.originalName || p.filePath || 'download');
    var encoded = encodeURIComponent(name);
    res.setHeader('Content-Disposition', 'attachment; filename="download"; filename*=UTF-8\'\'' + encoded);
    res.sendFile(path.resolve(fp));
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/pishi/report/preview/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadPishiReport();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(PISHI_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    if (p.fileType === 'pdf') {
      res.redirect('/api/pishi/report/file/' + id);
      return;
    }
    if (['doc', 'docx'].indexOf(p.fileType) !== -1) {
      var styleMap = 'p[style-name="标题 1"] => h1:fresh\np[style-name="标题 2"] => h2:fresh\np[style-name="标题 3"] => h3:fresh\np[style-name="Heading 1"] => h1:fresh\np[style-name="Heading 2"] => h2:fresh\np[style-name="Heading 3"] => h3:fresh\np[style-name="正文"] => p:fresh\nr[style-name="强调"] => strong\nr[style-name="标题 1 Char"] => strong';
      mammoth.convertToHtml({ path: fp, styleMap: styleMap, includeDefaultStyleMap: true }).then(function (result) {
        var css = '<style>body{font-family:SimSun,serif;font-size:16px;line-height:1.8;margin:0;padding:24px 48px}p{margin:0.5em 0;text-indent:2em}p:first-child{text-indent:0}h1,h2,h3{margin:1em 0 0.5em;font-weight:bold}table{border-collapse:collapse;width:100%;margin:1em 0}td,th{border:1px solid #333;padding:6px 10px;text-align:left}</style>';
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send('<!DOCTYPE html><html><head><meta charset="utf-8">' + css + '</head><body>' + result.value + '</body></html>');
      }).catch(function (err) {
        res.status(500).send('Word 文档预览失败：' + (err.message || '未知错误'));
      });
      return;
    }
    res.status(400).send('不支持预览');
  } catch (e) {
    res.status(500).send('Error');
  }
});

// ========== 监督议题 API ==========
var JIANDU_FORM_OPTIONS = ['听取审议报告', '财经工作监督', '执法检查', '专题询问', '视察'];
var jianduTopicStorage = multer.diskStorage({
  destination: function (req, file, cb) {
    if (!fs.existsSync(JIANDU_UPLOAD_DIR)) fs.mkdirSync(JIANDU_UPLOAD_DIR, { recursive: true });
    cb(null, JIANDU_UPLOAD_DIR);
  },
  filename: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase() || '.pdf';
    cb(null, 'jiandu_' + Date.now() + '_' + Math.random().toString(36).slice(2) + ext);
  }
});
var jianduTopicUpload = multer({
  storage: jianduTopicStorage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: function (req, file, cb) {
    var ext = (path.extname(file.originalname) || '').toLowerCase();
    if (['.pdf', '.doc', '.docx'].indexOf(ext) !== -1) cb(null, true);
    else cb(new Error('仅支持 PDF、DOC、DOCX 格式'));
  }
});

function loadJianduTopics() {
  return readCache(JIANDU_TOPICS_FILE, function () {
    var dir = path.dirname(JIANDU_TOPICS_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(JIANDU_TOPICS_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(JIANDU_TOPICS_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(JIANDU_TOPICS_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}
function saveJianduTopics(data) {
  var dir = path.dirname(JIANDU_TOPICS_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(JIANDU_TOPICS_FILE, JSON.stringify(data, null, 2));
  invalidateCache(JIANDU_TOPICS_FILE);
}

function loadJianduTopicFiles() {
  return readCache(JIANDU_TOPIC_FILES_FILE, function () {
    var dir = path.dirname(JIANDU_TOPIC_FILES_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(JIANDU_TOPIC_FILES_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(JIANDU_TOPIC_FILES_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(JIANDU_TOPIC_FILES_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}
function saveJianduTopicFiles(data) {
  var dir = path.dirname(JIANDU_TOPIC_FILES_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(JIANDU_TOPIC_FILES_FILE, JSON.stringify(data, null, 2));
  invalidateCache(JIANDU_TOPIC_FILES_FILE);
}

app.get('/api/jiandu/topics', function (req, res) {
  try {
    var year = (req.query.year || '').trim();
    var data = loadJianduTopics();
    var list = data.items || [];
    if (year) {
      var y = parseInt(year, 10);
      if (!isNaN(y)) list = list.filter(function (t) { return t.year === y; });
    }
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/jiandu/topics', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var year = parseInt(body.year, 10);
    var month = parseInt(body.month, 10);
    if (isNaN(month) || month < 1 || month > 12) month = 0;
    var content = (body.content || '').trim();
    var form = (body.form || '').trim();
    var department = (body.department || '').trim();
    if (!year || !content || !form || !department) return res.json(err('年度、监督内容、监督形式、部门/处室不能为空'));
    if (JIANDU_FORM_OPTIONS.indexOf(form) === -1) return res.json(err('监督形式无效'));
    var data = loadJianduTopics();
    var id = data.nextId++;
    var item = { id: id, year: year, month: month, content: content, form: form, department: department };
    data.items.push(item);
    saveJianduTopics(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.put('/api/jiandu/topics/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadJianduTopics();
    var t = data.items.find(function (x) { return x.id === id; });
    if (!t) return res.json(err('议题不存在'));
    if (body.year !== undefined) t.year = parseInt(body.year, 10) || t.year;
    if (body.month !== undefined) { var m = parseInt(body.month, 10); t.month = (m >= 1 && m <= 12) ? m : 0; }
    if (body.content !== undefined) t.content = String(body.content || '').trim();
    if (body.form !== undefined && JIANDU_FORM_OPTIONS.indexOf(body.form) !== -1) t.form = body.form;
    if (body.department !== undefined) t.department = String(body.department || '').trim();
    saveJianduTopics(data);
    res.json(ok(t));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.delete('/api/jiandu/topics/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadJianduTopics();
    var idx = data.items.findIndex(function (x) { return x.id === id; });
    if (idx === -1) return res.json(err('议题不存在'));
    data.items.splice(idx, 1);
    saveJianduTopics(data);
    var filesData = loadJianduTopicFiles();
    var files = filesData.items.filter(function (f) { return f.topicId === id; });
    files.forEach(function (f) {
      var fp = path.join(JIANDU_UPLOAD_DIR, f.filePath);
      if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
    });
    filesData.items = filesData.items.filter(function (f) { return f.topicId !== id; });
    saveJianduTopicFiles(filesData);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/jiandu/topics/:topicId/files', function (req, res) {
  try {
    var topicId = parseInt(req.params.topicId, 10);
    if (isNaN(topicId)) return res.json(ok([]));
    var data = loadJianduTopicFiles();
    var list = (data.items || []).filter(function (f) { return f.topicId === topicId; });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/jiandu/topics/:topicId/files/upload', requireAuth, jianduTopicUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var topicId = parseInt(req.params.topicId, 10);
    if (isNaN(topicId)) return res.json(err('议题ID无效'));
    var topicsData = loadJianduTopics();
    if (!topicsData.items.find(function (t) { return t.id === topicId; })) return res.json(err('议题不存在'));
    var rawName = decodeFilename(req.file.originalname || '');
    var body = req.body || {};
    var title = (body.title || '').trim() || rawName;
    var uploadDate = (body.uploadDate || '').trim();
    if (!uploadDate) {
      var d = new Date();
      uploadDate = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    var ext = (path.extname(rawName) || '').toLowerCase();
    var fileType = ext === '.pdf' ? 'pdf' : (ext === '.docx' ? 'docx' : (ext === '.doc' ? 'doc' : 'pdf'));
    var filesData = loadJianduTopicFiles();
    var id = filesData.nextId++;
    var item = { id: id, topicId: topicId, title: title, filePath: req.file.filename, originalName: rawName, fileType: fileType, uploadDate: uploadDate, createdAt: new Date().toISOString() };
    filesData.items.push(item);
    saveJianduTopicFiles(filesData);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message || '上传失败'));
  }
});

app.delete('/api/jiandu/topics/files/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadJianduTopicFiles();
    var idx = data.items.findIndex(function (x) { return x.id === id; });
    if (idx === -1) return res.json(err('文件不存在'));
    var f = data.items[idx];
    data.items.splice(idx, 1);
    saveJianduTopicFiles(data);
    var fp = path.join(JIANDU_UPLOAD_DIR, f.filePath);
    if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/jiandu/topics/files/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).json(err('Not found'));
    var data = loadJianduTopicFiles();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).json(err('Not found'));
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/jiandu/topics/files/:id/file', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadJianduTopicFiles();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(JIANDU_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.sendFile(path.resolve(fp), { headers: { 'Content-Disposition': 'inline' } });
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/jiandu/topics/files/:id/preview', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadJianduTopicFiles();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(JIANDU_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    if (p.fileType === 'pdf') {
      res.redirect('/api/jiandu/topics/files/' + id + '/file');
      return;
    }
    if (['doc', 'docx'].indexOf(p.fileType) !== -1) {
      var styleMap = 'p[style-name="标题 1"] => h1:fresh\np[style-name="标题 2"] => h2:fresh\np[style-name="正文"] => p:fresh';
      mammoth.convertToHtml({ path: fp, styleMap: styleMap }).then(function (result) {
        var css = '<style>body{font-family:SimSun;font-size:16px;line-height:1.8;margin:0;padding:24px}p{margin:0.5em 0;text-indent:2em}</style>';
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send('<!DOCTYPE html><html><head><meta charset="utf-8">' + css + '</head><body>' + result.value + '</body></html>');
      }).catch(function () { res.status(500).send('Word 预览失败'); });
      return;
    }
    res.redirect('/api/jiandu/topics/files/' + id + '/file');
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/jiandu/topics/files/:id/download', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadJianduTopicFiles();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(JIANDU_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var name = (p.originalName || p.filePath || 'download');
    var encoded = encodeURIComponent(name);
    res.setHeader('Content-Disposition', 'attachment; filename="download"; filename*=UTF-8\'\'' + encoded);
    res.sendFile(path.resolve(fp));
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/jiandu/form-options', function (req, res) {
  try {
    var data = loadFormSort();
    res.json(ok(data.order));
  } catch (e) {
    res.json(ok(JIANDU_FORM_OPTIONS));
  }
});

// 监督议题备份
app.get('/api/jiandu/backup', requireAuth, function (req, res) {
  try {
    var data = loadJianduTopics();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 监督议题恢复
app.post('/api/jiandu/restore', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var list = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      var form = (item.form || '').trim();
      if (JIANDU_FORM_OPTIONS.indexOf(form) === -1) form = JIANDU_FORM_OPTIONS[0] || '听取审议报告';
      var m = parseInt(item.month, 10);
      return {
        id: id,
        year: parseInt(item.year, 10) || new Date().getFullYear(),
        month: (m >= 1 && m <= 12) ? m : 0,
        content: (item.content || '').trim(),
        form: form || '听取审议报告',
        department: (item.department || '').trim()
      };
    });
    if (list.length > 0) {
      var maxId = Math.max.apply(null, list.map(function (t) { return t.id; }));
      nextId = maxId + 1;
    }
    saveJianduTopics({ items: list, nextId: nextId });
    res.json(ok({ restored: list.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

// 监督议题 Excel 导入
app.post('/api/jiandu/topics/import', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var rows = body.rows;
    if (!Array.isArray(rows) || rows.length === 0) return res.json(err('请上传有效数据'));
    var data = loadJianduTopics();
    var map = { 'year': 'year', 'month': 'month', 'content': 'content', 'form': 'form', 'department': 'department', '年度': 'year', '月份': 'month', '监督内容': 'content', '监督形式': 'form', '部门/处室': 'department' };
    var added = 0;
    rows.forEach(function (row) {
      var year = parseInt((row.year || row.年度 || '').toString().trim(), 10);
      var month = parseInt((row.month || row.月份 || '').toString().trim(), 10);
      if (isNaN(month) || month < 1 || month > 12) month = 0;
      var content = (row.content || row.监督内容 || '').toString().trim();
      var form = (row.form || row.监督形式 || '').toString().trim();
      var department = (row.department || row['部门/处室'] || '').toString().trim();
      if (!year || !content || !form || !department) return;
      if (JIANDU_FORM_OPTIONS.indexOf(form) === -1) form = JIANDU_FORM_OPTIONS[0];
      var id = data.nextId++;
      data.items.push({ id: id, year: year, month: month, content: content, form: form, department: department });
      added++;
    });
    saveJianduTopics(data);
    res.json(ok({ imported: added }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

var FALV_CATEGORIES = ['宪法', '法律', '行政法规', '监察法规', '地方法规', '司法解释'];
var FALV_AUTHORITIES = ['全国人大及其常委会', '国务院', '国家监察委员会', '最高人民法院', '最高人民检察院', '大连市人大及其常委会'];

function loadFalv() {
  return readCache(FALV_FILE, function () {
    var dir = path.dirname(FALV_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(FALV_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(FALV_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(FALV_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function saveFalv(data) {
  var dir = path.dirname(FALV_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(FALV_FILE, JSON.stringify(data, null, 2));
  invalidateCache(FALV_FILE);
}

function loadMinshengProgress() {
  return readCache(MINSHENG_PROGRESS_FILE, function () {
    var dir = path.dirname(MINSHENG_PROGRESS_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(MINSHENG_PROGRESS_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(MINSHENG_PROGRESS_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(MINSHENG_PROGRESS_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function saveMinshengProgress(data) {
  var dir = path.dirname(MINSHENG_PROGRESS_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(MINSHENG_PROGRESS_FILE, JSON.stringify(data, null, 2));
  invalidateCache(MINSHENG_PROGRESS_FILE);
}

function loadDepartments() {
  return readCache(DEPARTMENTS_FILE, function () {
    var dir = path.dirname(DEPARTMENTS_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(DEPARTMENTS_FILE)) {
      var proj = loadData();
      var names = {};
      (proj.projects || []).forEach(function (p) {
        var n = (p.department || '').trim();
        if (n) names[n] = true;
      });
      var items = Object.keys(names).map(function (name, i) { return { id: i + 1, name: name }; });
      var init = { items: items, nextId: items.length + 1 };
      fs.writeFileSync(DEPARTMENTS_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(DEPARTMENTS_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function saveDepartments(data) {
  var dir = path.dirname(DEPARTMENTS_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(DEPARTMENTS_FILE, JSON.stringify(data, null, 2));
  invalidateCache(DEPARTMENTS_FILE);
}

function loadPishiReport() {
  return readCache(PISHI_REPORT_FILE, function () {
    var dir = path.dirname(PISHI_REPORT_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    if (!fs.existsSync(PISHI_REPORT_FILE)) {
      var init = { items: [], nextId: 1 };
      fs.writeFileSync(PISHI_REPORT_FILE, JSON.stringify(init, null, 2));
      return init;
    }
    var raw = fs.readFileSync(PISHI_REPORT_FILE, 'utf8');
    var data = JSON.parse(raw);
    return { items: data.items || [], nextId: data.nextId || 1 };
  });
}

function savePishiReport(data) {
  var dir = path.dirname(PISHI_REPORT_FILE);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  fs.writeFileSync(PISHI_REPORT_FILE, JSON.stringify(data, null, 2));
  invalidateCache(PISHI_REPORT_FILE);
}

function computeValidity(effectiveDate) {
  if (!effectiveDate) return '有效';
  var d = String(effectiveDate).trim();
  var m = d.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '有效';
  var t = new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10)).getTime();
  return t > Date.now() ? '尚未生效' : '有效';
}

app.get('/api/falv/categories', function (req, res) { res.json(ok(FALV_CATEGORIES)); });
app.get('/api/falv/authorities', function (req, res) { res.json(ok(FALV_AUTHORITIES)); });

app.get('/api/falv', function (req, res) {
  try {
    var data = loadFalv();
    var list = data.items.map(function (p) {
      var v = { ...p };
      v.validity = v.validity || computeValidity(p.effectiveDate);
      return v;
    });
    var category = (req.query.category || '').trim();
    var authority = (req.query.authority || '').trim();
    if (category) list = list.filter(function (p) { return p.category === category; });
    if (authority) list = list.filter(function (p) { return p.issuingAuthority === authority; });
    list = list.slice().sort(function (a, b) {
      var c = (b.publicationDate || '').localeCompare(a.publicationDate || '');
      if (c !== 0) return c;
      return (b.id || 0) - (a.id || 0);
    });
    res.json(ok(list));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/falv/backup', requireAuth, function (req, res) {
  try {
    var data = loadFalv();
    res.json(ok(data.items));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/falv/restore', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var items = body.data;
    if (!Array.isArray(items)) return res.json(err('备份数据格式错误'));
    var nextId = 1;
    var list = items.map(function (item) {
      var id = item.id || nextId++;
      if (id >= nextId) nextId = id + 1;
      return {
        id: id,
        name: (item.name || '').trim(),
        category: (item.category || '法律').trim(),
        issuingAuthority: (item.issuingAuthority || '').trim(),
        publicationDate: (item.publicationDate || '').trim(),
        effectiveDate: (item.effectiveDate || '').trim(),
        validity: item.validity || computeValidity(item.effectiveDate),
        fileName: (item.fileName || '').trim(),
        filePath: item.filePath || '',
        fileType: (item.fileType || 'pdf').trim(),
        history: item.history || [],
        relatedDocs: item.relatedDocs || []
      };
    });
    if (list.length > 0) {
      var maxId = Math.max.apply(null, list.map(function (p) { return p.id; }));
      nextId = maxId + 1;
    }
    saveFalv({ items: list, nextId: nextId });
    res.json(ok({ restored: list.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/falv/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadFalv();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).json(err('记录不存在'));
    var v = { ...p };
    v.validity = v.validity || computeValidity(p.effectiveDate);
    res.json(ok(v));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/falv/upload', requireAuth, function (req, res, next) {
  falvUpload.single('file')(req, res, function (multerErr) {
    if (multerErr) return res.json(err(multerErr.message || '文件上传失败'));
    next();
  });
}, function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择文件'));
    var rawName = req.file.originalname || '';
    try { rawName = Buffer.from(rawName, 'latin1').toString('utf8'); } catch (e) {}
    var body = req.body || {};
    var name = (body.name || rawName.replace(/\.(pdf|doc|docx)$/i, '') || '未命名').trim();
    var category = (body.category || '法律').trim();
    var issuingAuthority = (body.issuingAuthority || '').trim();
    var publicationDate = (body.publicationDate || '').trim();
    var effectiveDate = (body.effectiveDate || '').trim();
    var data = loadFalv();
    var id = data.nextId++;
    var validity = computeValidity(effectiveDate);
    var item = {
      id: id,
      name: name,
      category: category,
      issuingAuthority: issuingAuthority,
      publicationDate: publicationDate,
      effectiveDate: effectiveDate,
      validity: validity,
      fileName: rawName,
      filePath: path.basename(req.file.path),
      fileType: (path.extname(rawName) || '').toLowerCase().replace(/^\./, '') || 'pdf',
      history: [{ date: publicationDate, name: name }],
      relatedDocs: []
    };
    data.items.push(item);
    saveFalv(data);
    res.json(ok(item));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/falv/:id/replace-file', requireAuth, falvUpload.single('file'), function (req, res) {
  try {
    if (!req.file) return res.json(err('请选择要上传的文件'));
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var rawName = req.file.originalname || '';
    try { rawName = Buffer.from(rawName, 'latin1').toString('utf8'); } catch (e) {}
    var body = req.body || {};
    var data = loadFalv();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (p.filePath) {
      var oldFp = path.join(FALV_UPLOAD_DIR, p.filePath);
      if (fs.existsSync(oldFp)) try { fs.unlinkSync(oldFp); } catch (e) {}
    }
    p.filePath = path.basename(req.file.path);
    p.fileName = rawName;
    p.fileType = (path.extname(rawName) || '').toLowerCase().replace(/^\./, '') || 'pdf';
    if (body.name !== undefined) p.name = String(body.name || '').trim() || p.name;
    if (body.category !== undefined) p.category = String(body.category || '法律').trim();
    if (body.issuingAuthority !== undefined) p.issuingAuthority = String(body.issuingAuthority || '').trim();
    if (body.publicationDate !== undefined) p.publicationDate = String(body.publicationDate || '').trim();
    if (body.effectiveDate !== undefined) {
      p.effectiveDate = String(body.effectiveDate || '').trim();
      p.validity = computeValidity(p.effectiveDate);
    }
    saveFalv(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.put('/api/falv/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var body = req.body || {};
    var data = loadFalv();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    if (body.name !== undefined) p.name = String(body.name || '').trim() || p.name;
    if (body.category !== undefined) p.category = String(body.category || '法律').trim();
    if (body.issuingAuthority !== undefined) p.issuingAuthority = String(body.issuingAuthority || '').trim();
    if (body.publicationDate !== undefined) p.publicationDate = String(body.publicationDate || '').trim();
    if (body.effectiveDate !== undefined) {
      p.effectiveDate = String(body.effectiveDate || '').trim();
      p.validity = computeValidity(p.effectiveDate);
    }
    saveFalv(data);
    res.json(ok(p));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.delete('/api/falv/:id', requireAuth, function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.json(err('ID无效'));
    var data = loadFalv();
    var idx = data.items.findIndex(function (p) { return p.id === id; });
    if (idx === -1) return res.json(err('记录不存在'));
    var p = data.items[idx];
    var fp = path.join(FALV_UPLOAD_DIR, p.filePath);
    if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
    data.items.splice(idx, 1);
    saveFalv(data);
    res.json(ok({ deleted: id }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.post('/api/falv/batch-delete', requireAuth, function (req, res) {
  try {
    var body = req.body || {};
    var ids = body.ids;
    if (!Array.isArray(ids) || ids.length === 0) return res.json(err('请选择要删除的记录'));
    var idSet = {};
    ids.forEach(function (id) { idSet[parseInt(id, 10)] = true; });
    var data = loadFalv();
    data.items = data.items.filter(function (p) {
      if (idSet[p.id]) {
        var fp = path.join(FALV_UPLOAD_DIR, p.filePath);
        if (fs.existsSync(fp)) try { fs.unlinkSync(fp); } catch (e) {}
        return false;
      }
      return true;
    });
    saveFalv(data);
    res.json(ok({ deleted: ids.length }));
  } catch (e) {
    res.status(500).json(err(e.message));
  }
});

app.get('/api/falv/file/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadFalv();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(FALV_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    res.setHeader('Content-Type', p.fileType === 'pdf' ? 'application/pdf' : 'application/octet-stream');
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/falv/download/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadFalv();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(FALV_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    var baseName = (p.name || p.fileName || 'document').replace(/[/\\:*?"<>|]/g, '_').trim() || 'document';
    var ext = path.extname(p.filePath) || (p.fileType === 'pdf' ? '.pdf' : '.docx');
    var fname = baseName + (ext.charAt(0) === '.' ? ext : '.' + ext);
    var fnameEnc = encodeURIComponent(fname);
    res.setHeader('Content-Disposition', 'attachment; filename="download' + ext + '"; filename*=UTF-8\'\'' + fnameEnc);
    res.sendFile(fp);
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/api/falv/preview/:id', function (req, res) {
  try {
    var id = parseInt(req.params.id, 10);
    if (isNaN(id)) return res.status(404).send('Not found');
    var data = loadFalv();
    var p = data.items.find(function (x) { return x.id === id; });
    if (!p) return res.status(404).send('Not found');
    var fp = path.join(FALV_UPLOAD_DIR, p.filePath);
    if (!fs.existsSync(fp)) return res.status(404).send('文件不存在');
    if (p.fileType === 'pdf') {
      res.redirect('/api/falv/file/' + id);
      return;
    }
    if (['doc', 'docx'].indexOf(p.fileType) !== -1) {
      var styleMap = [
        'p[style-name="标题 1"] => h1:fresh',
        'p[style-name="标题 2"] => h2:fresh',
        'p[style-name="标题 3"] => h3:fresh',
        'p[style-name="Heading 1"] => h1:fresh',
        'p[style-name="Heading 2"] => h2:fresh',
        'p[style-name="Heading 3"] => h3:fresh',
        'p[style-name="正文"] => p:fresh',
        'r[style-name="强调"] => strong',
        'r[style-name="标题 1 Char"] => strong'
      ].join('\n');
      mammoth.convertToHtml({
        path: fp,
        styleMap: styleMap,
        includeDefaultStyleMap: true
      }).then(function (result) {
        var css = '<style>body{font-family:SimSun,serif;font-size:16px;line-height:1.8;margin:0;padding:24px 48px}'
          + 'p{margin:0.5em 0;text-indent:2em}p:first-child{text-indent:0}'
          + 'h1,h2,h3{margin:1em 0 0.5em;font-weight:bold}table{border-collapse:collapse;width:100%;margin:1em 0}'
          + 'td,th{border:1px solid #333;padding:6px 10px;text-align:left}</style>';
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send('<!DOCTYPE html><html><head><meta charset="utf-8">' + css + '</head><body>' + result.value + '</body></html>');
      }).catch(function (err) {
        res.status(500).send('Word 文档预览失败：' + (err.message || '未知错误'));
      });
      return;
    }
    res.status(400).send('不支持预览');
  } catch (e) {
    res.status(500).send('Error');
  }
});

app.get('/', function (req, res) {
  res.sendFile(path.join(__dirname, 'index.html'));
});

// 启动
migrateToJiandu();
loadData();
loadAdmins();
app.listen(PORT, function () {
  console.log('========================================');
  console.log('  市人大常委会监督协调处管理系统 [jiandu]');
  console.log('  数据库: ' + config.DB_NAME + ' | 端口: ' + PORT);
  console.log('  访问地址: http://localhost:' + PORT);
  console.log('  超级管理员: 用户名 1312  密码 1312');
  console.log('  普通管理员: 用户名 1645  密码 4688633');
  console.log('========================================');
});

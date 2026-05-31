// CDM Toolkit v3.0 — SPA Router
// If you see "SPA: ACTIVE" in the footer, the router is working.

// Register custom Web Components (side-effect imports)
import './components/file-upload.js';
import './components/progress-bar.js';
import './components/data-table.js';
import './components/folder-tree.js';
import './components/workspace.js';
import './components/pipeline-nav.js';

const ROUTES = {
  '/':       { title: '仪表盘', loader: null },
  '/merge':  { title: '工作簿合并', page: 'merge.js' },
  '/split':  { title: '工作表拆分', page: 'split.js' },
  '/scan':   { title: '特殊字符扫描', page: 'scan.js' },
  '/files':  { title: '文件目录浏览', page: 'files.js' },
  '/dates':  { title: '日期计算', page: 'dates.js' },
  '/units':  { title: '单位转换', page: 'units.js' },
  '/clean':     { title: '数据清理', page: 'clean.js' },
  '/sheets':    { title: '工作表管理', page: 'sheets.js' },
  '/toc':       { title: 'TOC 目录', page: 'toc.js' },
  '/changes':   { title: '修改痕迹', page: 'changes.js' },
  '/creatinine':{ title: '医学计算器', page: 'creatinine.js' },
  '/saschart': { title: 'SAS 图表', page: 'saschart.js' },
  '/docx2pdf': { title: 'Word 转 PDF', page: 'docx2pdf.js' },
  '/explorer': { title: 'Data Explorer', page: 'explorer.js' },
};

let dashboardHTML = null;

// ---- UI helpers ----
function $(id) { return document.getElementById(id); }

function status(msg) {
  const el = $('spa-status');
  if (el) el.textContent = msg;
}

function saveDashboard() {
  if (!dashboardHTML) {
    const main = $('app-main');
    if (main) dashboardHTML = main.innerHTML;
  }
}

function restoreDashboard() {
  const main = $('app-main');
  if (main && dashboardHTML) {
    main.innerHTML = dashboardHTML;
  }
}

function updateNav(route) {
  document.querySelectorAll('[data-route]').forEach(a => {
    a.classList.toggle('active', a.dataset.route === route);
  });
}

// ---- Toast ----
// ---- Workspace helper ----
window._cdmUploadToWorkspace = async function(blob, filename) {
  const sid = localStorage.getItem('cdm_session') || '';
  const fd = new FormData();
  fd.append('file', new File([blob], filename || 'result.xlsx'));
  try {
    const res = await fetch('/api/workspace/upload', {
      method: 'POST',
      headers: { 'X-CDM-Session': sid },
      body: fd,
    });
    const data = await res.json();
    if (data.session_id) localStorage.setItem('cdm_session', data.session_id);
    // Refresh workspace panel if visible
    document.querySelectorAll('workspace-panel').forEach(p => p.refresh());
  } catch (_) {}
};

export function showToast(message, type) {
  type = type || 'success';
  var t = document.createElement('div');
  t.className = 'toast toast-' + type;
  t.textContent = message;
  document.body.appendChild(t);
  setTimeout(function () { t.remove(); }, 3000);
}

// ---- API helper ----
export async function api(endpoint, formData) {
  var res = await fetch(endpoint, { method: 'POST', body: formData });
  if (!res.ok) {
    var msg;
    try { var err = await res.json(); msg = err.detail || JSON.stringify(err); } catch (_) { msg = res.statusText; }
    throw new Error(msg || 'HTTP ' + res.status);
  }
  return res;
}

// ---- Page loader ----
async function loadPage(route) {
  var cfg = ROUTES[route];
  if (!cfg) { status('Unknown route: ' + route); return; }

  updateNav(route);
  document.title = cfg.title + ' — CDM Toolkit';
  status('Loading ' + cfg.title + '...');

  if (route === '/') {
    restoreDashboard();
    status('SPA: ACTIVE');
    return;
  }

  saveDashboard();
  var main = $('app-main');
  if (!main) { status('ERROR: #app-main not found'); return; }

  main.innerHTML = '<div style="text-align:center;padding:48px;"><p>加载中...</p></div>';

  try {
    var mod = await import('/static/js/pages/' + cfg.page);
    if (mod.default) {
      main.innerHTML = mod.default();
    }
    if (mod.init) await mod.init();
    status('SPA: ACTIVE | ' + cfg.title);
  } catch (err) {
    console.error('Page load failed:', route, err);
    main.innerHTML = '<div style="text-align:center;padding:48px;color:#C62828;">' +
      '<h3>页面加载失败</h3><p>' + err.message + '</p>' +
      '<a href="/" data-route="/" class="btn btn-outline mt-16">返回仪表盘</a></div>';
    status('ERROR: ' + err.message);
  }
}

// ---- Init ----
status('SPA: booting...');
console.log('CDM Toolkit SPA booting, routes:', Object.keys(ROUTES));

// ---- SPA click handler ----
document.addEventListener('click', function (e) {
  var link = e.target.closest('[data-route]');
  if (!link) return;
  e.preventDefault();
  e.stopPropagation();
  var route = link.dataset.route;
  console.log('SPA navigate to:', route);
  if (location.pathname === route) return;
  history.pushState(null, '', route);
  loadPage(route);
});

window.addEventListener('popstate', function () {
  loadPage(location.pathname);
});

// ---- Boot ----
function boot() {
  var path = location.pathname;
  console.log('SPA boot, path:', path);
  if (path !== '/' && path !== '') {
    loadPage(path);
  } else {
    updateNav('/');
    saveDashboard();
    status('SPA: ACTIVE');
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot);
} else {
  boot();
}

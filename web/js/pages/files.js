import { showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📁 文件目录浏览</h2><p class="section-subtitle">点击文件夹展开，浏览目录结构</p></div>
    <div id="files-breadcrumb" style="font-size:13px;margin-bottom:8px;color:#666;"></div>
    <div id="files-tree" style="background:#1E1E1E;color:#D4D4D4;padding:12px 16px;border-radius:8px;font-family:'Consolas',monospace;font-size:12px;line-height:1.8;max-height:65vh;overflow:auto;"></div>
    <div style="margin-top:8px;display:flex;gap:8px;">
      <button id="files-download" class="btn btn-outline" style="font-size:11px;">📥 导出 TXT</button>
      <span id="files-info" style="font-size:11px;color:#999;align-self:center;"></span>
    </div>
    <style>.ft-dir { cursor:pointer; } .ft-dir:hover span { color:#4472C4; text-decoration:underline; } .ft-file { color:#999; }</style>`;
}

export async function init() {
  const treeEl = document.getElementById('files-tree');
  const breadEl = document.getElementById('files-breadcrumb');
  const infoEl = document.getElementById('files-info');
  let allText = '';

  async function loadPath(path) {
    treeEl.innerHTML = '<span style="color:#888;">加载中...</span>';
    const url = '/api/file-tree' + (path ? '?path=' + encodeURIComponent(path) : '');
    const res = await fetch(url);
    const data = await res.json();
    if (data.error) { treeEl.innerHTML = `<span style="color:#C62828;">${data.error}</span>`; return; }

    // Breadcrumb
    const parts = path ? path.split('\\').filter(Boolean) : [];
    breadEl.innerHTML = parts.map((p, i) => {
      const pp = parts.slice(0, i + 1).join('\\') + '\\';
      return `<a href="#" data-path="${pp}" style="color:var(--color-primary);text-decoration:none;">${p}</a>`;
    }).join(' <span style="color:#999;">›</span> ') || '<span style="color:#999;">我的电脑</span>';

    // Tree
    const dirs = data.items.filter(i => i.is_dir);
    const files = data.items.filter(i => !i.is_dir);
    let html = '';

    if (!path) {
      html += '<div style="color:#888;margin-bottom:4px;">📁 我的电脑</div>';
    }

    dirs.forEach(d => {
      html += `<div class="ft-item ft-dir" data-path="${d.path.replace(/\\/g,'\\\\')}">📁 <span>${d.name}</span></div>`;
    });
    files.forEach(f => {
      const s = f.size ? formatSize(f.size) : '';
      html += `<div class="ft-file">📄 ${f.name} <span style="color:#888;">${s}</span></div>`;
    });
    if (!dirs.length && !files.length) {
      html += '<div style="color:#888;">空文件夹</div>';
    }

    treeEl.innerHTML = html;

    // Click handlers
    treeEl.querySelectorAll('.ft-dir').forEach(el => {
      el.addEventListener('click', () => loadPath(el.dataset.path.replace(/\\\\/g, '\\')));
    });

    // Breadcrumb clicks
    breadEl.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', e => { e.preventDefault(); loadPath(a.dataset.path); });
    });

    // Collect for TXT export
    infoEl.textContent = `${dirs.length} 文件夹, ${files.length} 文件`;
    allText = buildTextTree(data.items);
  }

  function buildTextTree(items, indent = '') {
    let lines = [];
    items.filter(i => i.is_dir).forEach(d => {
      lines.push(`${indent}📁 ${d.name}`);
    });
    items.filter(i => !i.is_dir).forEach(f => {
      lines.push(`${indent}📄 ${f.name}  ${f.size ? '(' + formatSize(f.size) + ')' : ''}`);
    });
    return lines.join('\n');
  }

  document.getElementById('files-download').addEventListener('click', () => {
    if (!allText) { showToast('请先浏览目录', 'error'); return; }
    const blob = new Blob([allText], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `file_list.txt`; a.click();
    showToast('已下载');
  });

  // Start at root
  loadPath('');
}

function formatSize(bytes) {
  if (!bytes) return '';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

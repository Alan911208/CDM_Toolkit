import { showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📁 文件目录浏览</h2><p class="section-subtitle">点击文件夹展开浏览，支持导出 TXT 格式</p></div>
    <div id="files-bread" style="font-size:12px;margin-bottom:8px;color:#666;"></div>
    <div id="files-tree" style="background:#1E1E1E;color:#D4D4D4;padding:12px 16px;border-radius:8px;font-family:'Consolas',monospace;font-size:12px;line-height:1.7;max-height:65vh;overflow:auto;"></div>
    <div style="margin-top:8px;display:flex;gap:8px;align-items:center;">
      <button id="files-txt" class="btn btn-outline" style="font-size:11px;">📥 导出 TXT</button>
      <span id="files-info" style="font-size:11px;color:#999;"></span>
    </div>
    <style>.ft-dir{cursor:pointer;}.ft-dir:hover span{color:#4472C4;text-decoration:underline;}.ft-file{color:#999;}</style>`;
}

export async function init() {
  const treeEl = document.getElementById('files-tree');
  const breadEl = document.getElementById('files-bread');
  const infoEl = document.getElementById('files-info');
  let currentPath = '';

  async function load(path) {
    currentPath = path;
    treeEl.innerHTML = '<span style="color:#888;">加载中...</span>';
    const url = '/api/file-tree' + (path ? '?path=' + encodeURIComponent(path) : '');
    const res = await fetch(url);
    const data = await res.json();
    if (data.error) { treeEl.innerHTML = `<span style="color:#C62828;">${data.error}</span>`; return; }

    // Breadcrumb
    const parts = path ? path.split('\\').filter(Boolean) : [];
    breadEl.innerHTML = parts.map((p, i) => {
      const pp = parts.slice(0, i + 1).join('\\') + '\\';
      return `<a href="#" data-path="${pp}" style="color:var(--color-primary);">${p}</a>`;
    }).join(' › ') || '我的电脑';

    // Render
    const dirs = data.items.filter(i => i.is_dir);
    const files = data.items.filter(i => !i.is_dir);
    let html = !path ? '<div style="color:#888;">📁 我的电脑</div>' : '';
    dirs.forEach(d => html += `<div class="ft-item ft-dir" data-path="${d.path.replace(/\\/g,'\\\\')}">📁 <span>${d.name}</span></div>`);
    files.forEach(f => html += `<div class="ft-file">📄 ${f.name} ${f.size ? '<span style="color:#666;">('+fmt(f.size)+')</span>' : ''}</div>`);
    if (!dirs.length && !files.length) html += '<div style="color:#888;">空文件夹</div>';
    treeEl.innerHTML = html;

    treeEl.querySelectorAll('.ft-dir').forEach(el => {
      el.addEventListener('click', () => load(el.dataset.path.replace(/\\\\/g, '\\')));
    });
    breadEl.querySelectorAll('a').forEach(a => {
      a.addEventListener('click', e => { e.preventDefault(); load(a.dataset.path); });
    });

    infoEl.textContent = `${dirs.length} 文件夹, ${files.length} 文件`;
  }

  document.getElementById('files-txt').addEventListener('click', async () => {
    if (!currentPath) { showToast('请先选择一个文件夹', 'error'); return; }
    const fd = new FormData(); fd.append('path', currentPath);
    const res = await fetch('/api/file-tree-txt', { method: 'POST', body: fd });
    const text = await res.text();
    const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `file_tree_${new Date().toISOString().slice(0,10)}.txt`; a.click();
    showToast('已下载');
  });

  load('');
}

function fmt(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024*1024) return (b/1024).toFixed(1) + ' KB';
  return (b/(1024*1024)).toFixed(1) + ' MB';
}

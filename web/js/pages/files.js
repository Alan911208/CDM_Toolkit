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
    try {
      const url = '/api/file-tree' + (path ? '?path=' + encodeURIComponent(path) : '');
      const res = await fetch(url);
      if (!res.ok) throw new Error('HTTP ' + res.status);
      const data = await res.json();
      if (data.error) { treeEl.innerHTML = '<span style="color:#C62828;">' + data.error + '</span>'; return; }

      // Breadcrumb
      const parts = path ? path.split('\\').filter(Boolean) : [];
      breadEl.innerHTML = '';
      if (parts.length === 0) {
        breadEl.textContent = '我的电脑';
      } else {
        parts.forEach((p, i) => {
          const pp = parts.slice(0, i + 1).join('\\') + '\\';
          const a = document.createElement('a');
          a.href = '#';
          a.textContent = p;
          a.style.cssText = 'color:var(--color-primary);text-decoration:none;';
          a.addEventListener('click', e => { e.preventDefault(); load(pp); });
          breadEl.appendChild(a);
          if (i < parts.length - 1) {
            breadEl.appendChild(document.createTextNode(' › '));
          }
        });
      }

      // Tree
      const dirs = data.items.filter(i => i.is_dir);
      const files = data.items.filter(i => !i.is_dir);
      treeEl.innerHTML = '';
      if (!path) {
        const hdr = document.createElement('div');
        hdr.style.color = '#888';
        hdr.textContent = '📁 我的电脑';
        treeEl.appendChild(hdr);
      }
      dirs.forEach(d => {
        const div = document.createElement('div');
        div.className = 'ft-item ft-dir';
        div.innerHTML = '📁 <span>' + esc(d.name) + '</span>';
        div.addEventListener('click', () => load(d.path));
        treeEl.appendChild(div);
      });
      files.forEach(f => {
        const div = document.createElement('div');
        div.className = 'ft-file';
        const s = f.size ? '(' + fmt(f.size) + ')' : '';
        div.textContent = '📄 ' + f.name + ' ' + s;
        treeEl.appendChild(div);
      });
      if (!dirs.length && !files.length) {
        const empty = document.createElement('div');
        empty.style.color = '#888';
        empty.textContent = '空文件夹';
        treeEl.appendChild(empty);
      }

      infoEl.textContent = dirs.length + ' 文件夹, ' + files.length + ' 文件';
    } catch (err) {
      treeEl.innerHTML = '<span style="color:#C62828;">加载失败: ' + esc(err.message) + '</span>';
    }
  }

  document.getElementById('files-txt').addEventListener('click', async () => {
    if (!currentPath) { showToast('请先选择一个文件夹', 'error'); return; }
    try {
      const fd = new FormData(); fd.append('path', currentPath);
      const res = await fetch('/api/file-tree-txt', { method: 'POST', body: fd });
      const text = await res.text();
      const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'file_tree_' + new Date().toISOString().slice(0, 10) + '.txt';
      a.click();
      showToast('已下载');
    } catch (err) {
      showToast('下载失败: ' + err.message, 'error');
    }
  });

  load('');
}

function esc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function fmt(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024 * 1024) return (b / 1024).toFixed(1) + ' KB';
  return (b / (1024 * 1024)).toFixed(1) + ' MB';
}

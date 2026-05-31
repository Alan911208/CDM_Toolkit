import { showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📁 文件目录浏览</h2><p class="section-subtitle">浏览本地文件夹内容，树形展示，一键导出 TXT 文件</p></div>
    <div class="form-group"><label for="files-path">文件夹路径</label><input type="text" id="files-path" placeholder="如 C:\Data\CDM"></div>
    <div class="form-row"><div class="form-group"><label for="files-filter">文件类型</label><select id="files-filter"><option value="*.*">所有文件</option><option value="*.xlsx">Excel (*.xlsx)</option><option value="*.xls">Excel (*.xls)</option><option value="*.sas7bdat">SAS (*.sas7bdat)</option><option value="*.csv">CSV (*.csv)</option><option value="*.pdf">PDF (*.pdf)</option></select></div><div class="form-group"><label><input type="checkbox" id="files-sub" checked> 包含子文件夹</label></div></div>
    <button id="files-btn" class="btn btn-primary">🔍 浏览</button>
    <progress-bar id="files-progress" style="display:none;"></progress-bar><div id="files-results" style="margin-top:16px;"></div>`;
}

export async function init() {
  const btn = document.getElementById('files-btn');
  const progress = document.getElementById('files-progress');
  const resultsDiv = document.getElementById('files-results');
  let currentPath = '';

  btn.addEventListener('click', async () => {
    const path = document.getElementById('files-path').value.trim();
    if (!path) { showToast('请输入文件夹路径', 'error'); return; }
    currentPath = path;

    const fd = new FormData();
    fd.append('path', path);
    fd.append('include_subfolders', document.getElementById('files-sub').checked);
    fd.append('file_pattern', document.getElementById('files-filter').value);
    fd.append('format', 'txt');

    progress.show(); progress.setProgress(-1, '正在浏览...');
    try {
      const res = await fetch('/api/file-listing', { method: 'POST', body: fd });
      const text = await res.text();
      resultsDiv.innerHTML = `<div style="display:flex;gap:8px;margin-bottom:8px;align-items:center;">
        <span style="font-size:12px;color:#999;">📂 ${path}</span>
        <button id="files-download" class="btn btn-outline" style="font-size:11px;padding:4px 10px;">📥 下载 TXT</button>
      </div>
      <pre style="background:#1E1E1E;color:#D4D4D4;padding:16px;border-radius:8px;font-family:'Consolas',monospace;font-size:12px;line-height:1.5;max-height:60vh;overflow:auto;white-space:pre-wrap;">${esc(text)}</pre>`;
      progress.setProgress(100, '浏览完成');

      document.getElementById('files-download').addEventListener('click', () => {
        const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = `file_list_${new Date().toISOString().slice(0,10)}.txt`; a.click();
        URL.revokeObjectURL(url);
        showToast('文件列表已下载');
      });
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
  });
}

function esc(s) { return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

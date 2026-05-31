import { api, showToast } from '../app.js';
export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📁 文件目录浏览</h2><p class="section-subtitle">浏览本地文件夹内容，一键导出 Excel 清单</p></div>
    <div class="form-group"><label for="files-path">文件夹路径</label><input type="text" id="files-path" placeholder="如 C:\Data\CDM"></div>
    <div class="form-row"><div class="form-group"><label for="files-filter">文件类型</label><select id="files-filter"><option value="*.*">所有文件</option><option value="*.xlsx">Excel (*.xlsx)</option><option value="*.xls">Excel (*.xls)</option><option value="*.sas7bdat">SAS (*.sas7bdat)</option><option value="*.csv">CSV (*.csv)</option><option value="*.pdf">PDF (*.pdf)</option></select></div><div class="form-group"><label><input type="checkbox" id="files-sub" checked> 包含子文件夹</label></div></div>
    <button id="files-btn" class="btn btn-primary">🔍 浏览</button> <button id="files-export" class="btn btn-outline" style="display:none;">📥 导出 JSON</button>
    <progress-bar id="files-progress" style="display:none;"></progress-bar><div id="files-results" style="margin-top:16px;"></div>`;
}
export async function init() {
  let exportAdded = false;
  document.getElementById('files-btn').addEventListener('click', async () => {
    const path = document.getElementById('files-path').value.trim();
    if (!path) { showToast('请输入文件夹路径', 'error'); return; }
    const progress = document.getElementById('files-progress'); const resultsDiv = document.getElementById('files-results');
    progress.show(); progress.setProgress(-1, '正在浏览...');
    try {
      const fd = new FormData(); fd.append('path', path); fd.append('include_subfolders', document.getElementById('files-sub').checked); fd.append('file_pattern', document.getElementById('files-filter').value);
      const res = await fetch('/api/file-listing', { method: 'POST', body: fd }); const json = await res.json();
      const tree = document.createElement('folder-tree'); tree.data = json; resultsDiv.innerHTML = ''; resultsDiv.appendChild(tree);
      document.getElementById('files-export').style.display = '';
      if (!exportAdded) { exportAdded = true; document.getElementById('files-export').addEventListener('click', () => { const blob = new Blob([JSON.stringify(json, null, 2)], { type: 'application/json' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'file_listing.json'; a.click(); }); }
      progress.setProgress(100, `共 ${json.total} 个项目`);
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(`浏览失败: ${err.message}`, 'error'); }
  });
}

import { showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📊 Data Explorer</h2><p class="section-subtitle">上传 SAS 数据集 → 浏览 Metadata（变量列表、类型、样本值）</p></div>
    <file-upload accept=".sas7bdat"></file-upload>
    <progress-bar id="exp-progress" style="display:none;"></progress-bar>
    <div id="exp-results" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const progress = document.getElementById('exp-progress');
  const resultsDiv = document.getElementById('exp-results');

  upload.addEventListener('files-changed', async (e) => {
    if (e.detail.files.length === 0) return;
    progress.show(); progress.setProgress(-1, '扫描中...');
    resultsDiv.innerHTML = '';
    const fd = new FormData(); fd.append('file', upload.files[0]);
    try {
      const res = await fetch('/api/explorer/scan', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { resultsDiv.innerHTML = `<p style="color:#C62828;">${data.error}</p>`; progress.hide(); return; }
      renderData(data);
      progress.setProgress(100, '扫描完成');
    } catch (err) { resultsDiv.innerHTML = `<p style="color:#C62828;">${err.message}</p>`; progress.hide(); }
  });

  function renderData(data) {
    const ds = data.dataset || {};
    const vars = data.variables || [];
    const sum = data.summary || {};
    resultsDiv.innerHTML = `
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">
        <div class="ds-stat"><strong>${ds.row_count}</strong> Rows</div>
        <div class="ds-stat"><strong>${ds.col_count}</strong> Variables</div>
        <div class="ds-stat">🔤 ${sum.char_vars || 0} Char</div>
        <div class="ds-stat">🔢 ${sum.num_vars || 0} Num</div>
        <div class="ds-stat">✅ ${sum.complete_vars || 0} Complete</div>
        <div class="ds-stat" style="font-family:monospace;">📄 ${ds.file_name || ''}</div>
      </div>
      <div style="max-height:60vh;overflow:auto;">
        <table class="exp-table">
          <thead><tr><th>#</th><th>Variable</th><th>Type</th><th>Dtype</th><th>Missing%</th><th>Unique</th><th>Sample</th></tr></thead>
          <tbody>${vars.map((v, i) => `<tr>
            <td>${i + 1}</td>
            <td><strong>${v.name}</strong></td>
            <td>${v.type}</td>
            <td style="font-size:11px;color:#999;">${v.dtype}</td>
            <td style="color:${v.missing_pct > 50 ? '#C62828' : v.missing_pct > 20 ? '#E65100' : '#999'};">${v.missing_pct}%</td>
            <td>${v.unique}</td>
            <td style="font-size:11px;max-width:250px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${(v.sample||[]).join(' | ') || '-'}</td>
          </tr>`).join('')}</tbody>
        </table></div>
      <style>
        .ds-stat { padding:10px 16px; background:var(--color-surface); border-radius:8px; font-size:13px; border:1px solid var(--color-border); }
        .exp-table { width:100%; border-collapse:collapse; font-size:13px; }
        .exp-table th { background:#4472C4; color:white; padding:8px 12px; text-align:left; position:sticky; top:0; }
        .exp-table td { padding:6px 12px; border-bottom:1px solid #E0E4E8; }
        .exp-table tr:hover { background:#F5F7FA; }
      </style>`;
  }
}

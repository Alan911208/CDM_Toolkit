import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔍 数据集比较</h2><p class="section-subtitle">上传旧版和新版 SAS 数据集 → 比对异同 → 网页展示 + Excel 下载</p></div>
    <div class="form-row">
      <div class="form-group"><label style="font-weight:600;">📂 旧版本 (Old)</label><file-upload id="cmp-old" multiple accept=".sas7bdat"></file-upload></div>
      <div class="form-group"><label style="font-weight:600;">📂 新版本 (New)</label><file-upload id="cmp-new" multiple accept=".sas7bdat"></file-upload></div>
    </div>
    <div class="form-row mt-16">
      <div class="form-group"><label for="cmp-key">Key 变量（逗号分隔）</label><input type="text" id="cmp-key" placeholder="USUBJID,Visit"></div>
      <div class="form-group" style="align-self:flex-end;padding-bottom:16px;"><button id="cmp-btn" class="btn btn-primary" disabled>▶ 开始比较</button></div>
    </div>
    <progress-bar id="cmp-progress" style="display:none;"></progress-bar>
    <div id="cmp-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const oldUp = document.getElementById('cmp-old'), newUp = document.getElementById('cmp-new');
  const btn = document.getElementById('cmp-btn'), progress = document.getElementById('cmp-progress');
  const resultDiv = document.getElementById('cmp-result');

  function check() { btn.disabled = !(oldUp.files.length && newUp.files.length); }
  oldUp.addEventListener('files-changed', check); newUp.addEventListener('files-changed', check);

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    oldUp.files.forEach(f => fd.append('old_files', f));
    newUp.files.forEach(f => fd.append('new_files', f));
    fd.append('key_vars', document.getElementById('cmp-key').value);
    fd.append('format', 'json');

    progress.show(); progress.setProgress(-1, '正在比较...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await fetch('/api/compare-datasets', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { resultDiv.innerHTML = `<p style="color:#C62828;">${data.error}</p>`; progress.hide(); return; }
      renderResults(data);
      progress.setProgress(100, '比较完成');
    } catch (err) { resultDiv.innerHTML = `<p style="color:#C62828;">${err.message}</p>`; progress.hide(); }
    finally { btn.disabled = false; }
  });

  function renderResults(data) {
    const sum = data.summary || {}, ds = data.datasets || [];
    const color = sum.changed > 0 ? '#E65100' : '#2E7D32', icon = sum.changed > 0 ? '⚠️' : '✅';

    resultDiv.innerHTML = `
      <div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:16px;align-items:center;">
        <div style="padding:10px 20px;background:${color};color:white;border-radius:8px;font-size:14px;font-weight:600;">${icon} ${sum.changed || 0} 个变化 / ${sum.total} 个数据集</div>
        <div class="cmp-stat">✅ ${sum.identical||0} 一致</div>
        <div class="cmp-stat">🟡 ${sum.changed||0} 有差异</div>
        <div class="cmp-stat">🔴 ${sum.only_old||0} 仅旧版</div>
        <div class="cmp-stat">🟢 ${sum.only_new||0} 仅新版</div>
        <button id="cmp-download" class="btn btn-outline" style="font-size:12px;margin-left:auto;">📥 下载 Excel 报告</button>
      </div>
      <div style="max-height:60vh;overflow:auto;">
        <table class="cmp-table">
          <thead><tr>
            <th>Dataset</th><th>Status</th><th>Old Rows</th><th>New Rows</th>
            <th>+Vars</th><th>-Vars</th><th>~Vars</th>
            <th>+Rows</th><th>-Rows</th><th>~Rows</th>
          </tr></thead>
          <tbody>${ds.map(d => {
            const st = d.status;
            const bg = st === 'identical' ? '#E8F5E9' : st === 'changed' ? '#FFF3E0' : st === 'only_new' ? '#E3F2FD' : '#FFEBEE';
            const label = st === 'identical' ? '✅ 一致' : st === 'changed' ? '⚠️ 有差异' : st === 'only_new' ? '🆕 仅新版' : '🗑 仅旧版';
            return `<tr style="background:${bg};">
              <td><strong>${d.dataset}</strong></td>
              <td>${label}</td>
              <td>${d.old_rows||0}</td><td>${d.new_rows||0}</td>
              <td style="color:#2E7D32;">${(d.added_vars||[]).length}</td>
              <td style="color:#C62828;">${(d.removed_vars||[]).length}</td>
              <td>${(d.modified_vars||[]).length}</td>
              <td style="color:#2E7D32;">${d.added_rows||0}</td>
              <td style="color:#C62828;">${d.removed_rows||0}</td>
              <td>${d.modified_rows||0}</td>
            </tr>
            ${st === 'changed' ? `<tr style="font-size:11px;color:#666;"><td colspan="10">
              ${d.added_vars?.length ? '➕新增变量: '+d.added_vars.join(', ')+'<br>' : ''}
              ${d.removed_vars?.length ? '➖删除变量: '+d.removed_vars.join(', ')+'<br>' : ''}
              ${d.modified_vars?.length ? '✏️修改变量: '+d.modified_vars.slice(0,10).join(', ')+'<br>' : ''}
            </td></tr>` : ''}`;
          }).join('')}</tbody>
        </table></div>
      <style>
        .cmp-stat { padding:8px 14px; background:var(--color-surface); border-radius:8px; font-size:13px; border:1px solid var(--color-border); }
        .cmp-table { width:100%; border-collapse:collapse; font-size:13px; }
        .cmp-table th { background:#4472C4; color:white; padding:8px 10px; text-align:left; position:sticky; top:0; }
        .cmp-table td { padding:6px 10px; border-bottom:1px solid #E0E4E8; }
      </style>`;

    // Download button
    document.getElementById('cmp-download').addEventListener('click', async () => {
      const fd2 = new FormData();
      oldUp.files.forEach(f => fd2.append('old_files', f));
      newUp.files.forEach(f => fd2.append('new_files', f));
      fd2.append('key_vars', document.getElementById('cmp-key').value);
      fd2.append('format', 'xlsx');
      const r = await fetch('/api/compare-datasets', { method: 'POST', body: fd2 });
      const blob = await r.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'comparison_report.xlsx'; a.click();
      URL.revokeObjectURL(url);
      showToast('Excel 报告已下载');
    });
  }
}

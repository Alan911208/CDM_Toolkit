import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📋 工作表管理</h2><p class="section-subtitle">批量创建、删除、列出工作表</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="sheets-action">操作</label><select id="sheets-action"><option value="create">批量创建工作表</option><option value="delete">批量删除工作表</option><option value="list">列出所有工作表</option></select></div></div>
    <div class="form-group" id="sheets-names-group"><label for="sheets-names">工作表名称列表（每行一个）</label><textarea id="sheets-names" rows="4" placeholder="Sheet1&#10;Sheet2&#10;Sheet3" style="width:100%;padding:8px;border:1px solid var(--color-border);border-radius:4px;font-size:14px;"></textarea></div>
    <button id="sheets-btn" class="btn btn-primary" disabled>▶ 执行</button>
    <progress-bar id="sheets-progress" style="display:none;"></progress-bar>
    <div id="sheets-result" style="margin-top:16px;"></div>
    <pipeline-nav context="sheets"></pipeline-nav>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('sheets-btn');
  const progress = document.getElementById('sheets-progress');
  const actionSel = document.getElementById('sheets-action');
  const namesGroup = document.getElementById('sheets-names-group');
  const namesInput = document.getElementById('sheets-names');
  const resultDiv = document.getElementById('sheets-result');

  actionSel.addEventListener('change', () => { namesGroup.style.display = actionSel.value === 'list' ? 'none' : ''; });
  actionSel.dispatchEvent(new Event('change'));
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0 && actionSel.value !== 'list'; });

  btn.addEventListener('click', async () => {
    const action = actionSel.value; const fd = new FormData();
    if (upload.files.length > 0) fd.append('file', upload.files[0]);

    if (action === 'list') {
      if (!upload.files.length) { showToast('请上传文件', 'error'); return; }
      progress.show(); progress.setProgress(-1, '读取中...'); btn.disabled = true; resultDiv.innerHTML = '';
      try {
        const res = await fetch('/api/list-sheets', { method: 'POST', body: fd }); const data = await res.json();
        resultDiv.innerHTML = `<h4>工作表列表 (${data.count})</h4><div style="display:flex;flex-wrap:wrap;gap:8px;">${data.sheets.map(s => `<span style="padding:4px 12px;background:#F5F7FA;border-radius:16px;font-size:13px;">${s.name}</span>`).join('')}</div>`;
        progress.setProgress(100, `共 ${data.count} 个工作表`);
      } catch (err) { progress.setProgress(0, err.message); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
      return;
    }

    const names = namesInput.value.trim(); if (!names) { showToast('请输入工作表名称', 'error'); return; }
    fd.append('sheet_names', names);
    progress.show(); progress.setProgress(-1, '处理中...'); btn.disabled = true; resultDiv.innerHTML = '';

    const endpoint = action === 'create' ? '/api/create-sheets' : '/api/delete-sheets';
    const label = action === 'create' ? '创建' : '删除';
    try {
      const res = await api(endpoint, fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'sheets_result.xlsx'; a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '完成！'); resultDiv.innerHTML = `<p style="color:#2E7D32;padding:8px;">✅ 工作表${label}完成 — 已加入工作区预览</p>`;
      showToast(`工作表${label}成功`); window._cdmUploadToWorkspace(blob, 'sheets_result.xlsx');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

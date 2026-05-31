import { api, showToast } from '../app.js';

const SHEET_LIMIT_EXPLAINER = '每行一个名称，自动处理 31 字符限制和非法字符';

export default function render() {
  return `<div class="page-section"><h2>📋 工作表管理</h2><p class="section-subtitle">批量创建、删除、重命名、列出工作表</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="sheets-action">操作</label>
        <select id="sheets-action">
          <option value="create">批量创建工作表</option>
          <option value="delete">批量删除工作表</option>
          <option value="list">列出所有工作表</option>
        </select>
      </div>
    </div>
    <div class="form-group" id="sheets-names-group">
      <label for="sheets-names">工作表名称列表</label>
      <textarea id="sheets-names" rows="5" placeholder="${SHEET_LIMIT_EXPLAINER}" style="width:100%;padding:8px 12px;border:1px solid var(--color-border);border-radius:4px;font-size:14px;font-family:var(--font);"></textarea>
    </div>
    <div class="form-group" id="sheets-clear-group" style="display:none;">
      <label><input type="checkbox" id="sheets-clear-first"> 创建前删除同名工作表</label>
    </div>
    <button id="sheets-btn" class="btn btn-primary" disabled>▶ 执行</button>
    <progress-bar id="sheets-progress" style="display:none;"></progress-bar>
    <div id="sheets-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('sheets-btn');
  const progress = document.getElementById('sheets-progress');
  const actionSel = document.getElementById('sheets-action');
  const namesGroup = document.getElementById('sheets-names-group');
  const namesInput = document.getElementById('sheets-names');
  const clearGroup = document.getElementById('sheets-clear-group');
  const resultDiv = document.getElementById('sheets-result');

  actionSel.addEventListener('change', () => {
    const action = actionSel.value;
    namesGroup.style.display = (action === 'list') ? 'none' : '';
    clearGroup.style.display = (action === 'create') ? '' : 'none';
  });
  actionSel.dispatchEvent(new Event('change'));

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0 && actionSel.value !== 'list';
  });

  btn.addEventListener('click', async () => {
    const action = actionSel.value;
    const fd = new FormData();
    if (upload.files.length > 0) fd.append('file', upload.files[0]);

    if (action === 'list') {
      // List sheet names — use the engine directly via a small local read
      if (!upload.files.length) { showToast('请上传文件', 'error'); return; }
      progress.show(); progress.setProgress(-1, '读取中...'); btn.disabled = true;
      try {
        const res = await fetch('/api/list-sheets', { method: 'POST', body: fd });
        const data = await res.json();
        resultDiv.innerHTML = `<h4 style="margin-bottom:8px;">工作表列表 (${data.count})</h4>
          <div class="cards" style="grid-template-columns:repeat(auto-fill, minmax(180px, 1fr));">${data.sheets.map(s => `<div class="card" style="padding:12px;"><h4>${s.name}</h4><p style="font-size:11px;">${s.visible ? '可见' : '隐藏'}</p></div>`).join('')}</div>`;
        progress.setProgress(100, `共 ${data.count} 个工作表`);
      } catch (err) { progress.setProgress(0, err.message); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
      return;
    }

    const names = namesInput.value.trim();
    if (!names) { showToast('请输入工作表名称', 'error'); return; }

    fd.append('sheet_names', names);

    if (action === 'create') {
      fd.append('clear_first', document.getElementById('sheets-clear-first').checked);
      progress.show(); progress.setProgress(-1, '正在创建工作表...'); btn.disabled = true;
      try {
        const res = await api('/api/create-sheets', fd);
        const blob = await res.blob();
        downloadBlob(blob, 'sheets_result.xlsx');
        progress.setProgress(100, '完成！'); showToast('工作表创建成功');
      } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
    } else if (action === 'delete') {
      progress.show(); progress.setProgress(-1, '正在删除工作表...'); btn.disabled = true;
      try {
        const res = await api('/api/delete-sheets', fd);
        const blob = await res.blob();
        downloadBlob(blob, 'sheets_result.xlsx');
        progress.setProgress(100, '完成！'); showToast('工作表删除成功');
      } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
    }
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

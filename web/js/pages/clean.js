import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>🧹 数据清理</h2><p class="section-subtitle">文本标准化、重复高亮、空白行管理、删除线处理</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="clean-action">清理操作</label><select id="clean-action"><option value="strikethrough-delete">删除含删除线的行</option><option value="strikethrough-clear">清除所有删除线格式</option></select></div></div>
    <button id="clean-btn" class="btn btn-primary" disabled>▶ 执行清理</button>
    <progress-bar id="clean-progress" style="display:none;"></progress-bar>
    <div id="clean-result" style="margin-top:16px;"></div>
    <pipeline-nav context="clean"></pipeline-nav>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('clean-btn');
  const progress = document.getElementById('clean-progress');
  const actionSel = document.getElementById('clean-action');
  const resultDiv = document.getElementById('clean-result');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const action = actionSel.value; const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('action', action === 'strikethrough-delete' ? 'delete' : 'clear');
    const label = action === 'strikethrough-delete' ? '删除含删除线的行' : '清除删除线格式';
    progress.show(); progress.setProgress(-1, `正在${label}...`); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/strikethrough', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'strikethrough_result.xlsx'; a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '完成！'); resultDiv.innerHTML = `<p style="color:#2E7D32;padding:8px;">✅ ${label}完成 — 已加入工作区预览</p>`;
      showToast(`${label}完成`); window._cdmUploadToWorkspace(blob, 'strikethrough_result.xlsx');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

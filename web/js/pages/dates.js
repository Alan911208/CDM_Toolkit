import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📅 日期计算</h2><p class="section-subtitle">对日期列批量加减天数</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="dates-col">目标列</label><input type="text" id="dates-col" value="F" maxlength="3"></div><div class="form-group"><label for="dates-days">加减天数（负数=减）</label><input type="number" id="dates-days" value="30"></div><div class="form-group"><label for="dates-format">日期格式</label><input type="text" id="dates-format" value="YYYY-MM-DD"></div></div>
    <button id="dates-btn" class="btn btn-primary" disabled>▶ 执行</button>
    <progress-bar id="dates-progress" style="display:none;"></progress-bar>
    <div id="dates-result" style="margin-top:16px;"></div>
    <pipeline-nav context="dates"></pipeline-nav>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('dates-btn');
  const progress = document.getElementById('dates-progress');
  const resultDiv = document.getElementById('dates-result');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('col_letter', document.getElementById('dates-col').value); fd.append('days', document.getElementById('dates-days').value); fd.append('date_format', document.getElementById('dates-format').value);
    progress.show(); progress.setProgress(-1, '处理中...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/date-calc', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'dates_result.xlsx'; a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '完成！'); resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ 日期计算完成 — 已加入工作区预览</p>';
      showToast('日期计算完成'); window._cdmUploadToWorkspace(blob, 'dates_result.xlsx');
    } catch (err) { showToast(`失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

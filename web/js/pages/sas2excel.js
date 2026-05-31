import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📊 SAS 批量转 Excel</h2><p class="section-subtitle">拖拽 .sas7bdat 数据集 → 批量转换为 Excel，表头格式：变量名(label)</p></div>
    <file-upload multiple accept=".sas7bdat"></file-upload>
    <button id="s2e-btn" class="btn btn-primary mt-16" disabled>▶ 批量转换</button>
    <progress-bar id="s2e-progress" style="display:none;"></progress-bar>
    <div id="s2e-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('s2e-btn');
  const progress = document.getElementById('s2e-progress');
  const resultDiv = document.getElementById('s2e-result');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
    btn.textContent = e.detail.files.length ? `▶ 转换 ${e.detail.files.length} 个数据集` : '▶ 批量转换';
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    progress.show(); progress.setProgress(-1, '正在转换...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/sas-batch-to-xlsx', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const ct = res.headers.get('content-type') || '';
      a.href = url;
      a.download = ct.includes('zip') ? 'sas_to_excel.zip' : 'output.xlsx';
      a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '转换完成！');
      resultDiv.innerHTML = '<p style="color:#2E7D32;">✅ 转换完成 — 表头格式：变量名(label)</p>';
      showToast('批量转换完成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

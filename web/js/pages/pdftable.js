import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📊 PDF 表格提取</h2><p class="section-subtitle">上传 PDF → 自动检测并提取表格 → 导出为 Excel</p></div>
    <file-upload accept=".pdf"></file-upload>
    <button id="pdftable-btn" class="btn btn-primary mt-16" disabled>▶ 提取表格</button>
    <progress-bar id="pdftable-progress" style="display:none;"></progress-bar>
    <div id="pdftable-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('pdftable-btn');
  const progress = document.getElementById('pdftable-progress');
  const resultDiv = document.getElementById('pdftable-result');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });

  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]);
    progress.show(); progress.setProgress(-1, '正在提取表格...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/pdf-extract-table', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      const ct = res.headers.get('content-type') || '';
      a.href = url;
      a.download = ct.includes('zip') ? 'pdf_tables.zip' : 'extracted.xlsx';
      a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '提取完成！');
      resultDiv.innerHTML = '<p style="color:#2E7D32;">✅ 表格提取完成 — 已下载</p>';
      showToast('表格提取完成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

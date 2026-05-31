import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📑 PDF 合并</h2><p class="section-subtitle">上传多个 PDF → 合并为一个文件（自动添加书签）</p></div>
    <file-upload multiple accept=".pdf"></file-upload>
    <button id="pdfmerge-btn" class="btn btn-primary mt-16" disabled>▶ 合并 PDF</button>
    <progress-bar id="pdfmerge-progress" style="display:none;"></progress-bar>
    <div id="pdfmerge-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('pdfmerge-btn');
  const progress = document.getElementById('pdfmerge-progress');
  const resultDiv = document.getElementById('pdfmerge-result');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length < 2;
    btn.textContent = e.detail.files.length >= 2 ? `▶ 合并 ${e.detail.files.length} 个 PDF` : '▶ 合并 PDF（需2个以上）';
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    progress.show(); progress.setProgress(-1, '正在合并...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/pdf-merge', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'Combined.pdf'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '合并完成！');
      resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ PDF 合并完成（含书签导航）</p>';
      showToast('PDF 合并完成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

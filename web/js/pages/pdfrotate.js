import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔄 PDF 旋转</h2><p class="section-subtitle">上传 PDF → 旋转所有页面 → 下载</p></div>
    <file-upload accept=".pdf"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="pdf-rotate-angle">旋转角度</label><select id="pdf-rotate-angle"><option value="90">90° 顺时针</option><option value="180">180°</option><option value="270">270° 逆时针</option></select></div></div>
    <button id="pdfrotate-btn" class="btn btn-primary" disabled>▶ 旋转</button>
    <progress-bar id="pdfrotate-progress" style="display:none;"></progress-bar>
    <div id="pdfrotate-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('pdfrotate-btn');
  const progress = document.getElementById('pdfrotate-progress');
  const resultDiv = document.getElementById('pdfrotate-result');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('angle', document.getElementById('pdf-rotate-angle').value);

    progress.show(); progress.setProgress(-1, '旋转中...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/pdf-rotate', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'rotated.pdf'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '完成！'); resultDiv.innerHTML = '<p style="color:#2E7D32;">✅ 旋转完成</p>';
      showToast('PDF 旋转完成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

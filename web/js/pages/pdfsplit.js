import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>✂️ PDF 分割</h2><p class="section-subtitle">上传 PDF → 按指定页数自动分割为多个文件</p></div>
    <file-upload accept=".pdf"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="pdf-group-size">每部分页数</label><input type="number" id="pdf-group-size" value="100" min="1"></div></div>
    <button id="pdfsplit-btn" class="btn btn-primary" disabled>▶ 分割 PDF</button>
    <progress-bar id="pdfsplit-progress" style="display:none;"></progress-bar>
    <div id="pdfsplit-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('pdfsplit-btn');
  const progress = document.getElementById('pdfsplit-progress');
  const resultDiv = document.getElementById('pdfsplit-result');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });

  btn.addEventListener('click', async () => {
    const gs = parseInt(document.getElementById('pdf-group-size').value) || 100;
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('group_size', gs);

    progress.show(); progress.setProgress(-1, '正在分割...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/pdf-split', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = res.headers.get('content-type')?.includes('zip') ? 'pdf_split.zip' : 'split.pdf';
      a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '分割完成！');
      resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ PDF 分割完成 — 文件已下载</p>';
      showToast('PDF 分割完成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

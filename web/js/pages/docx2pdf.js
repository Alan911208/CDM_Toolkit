import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📄 Word 转 PDF</h2><p class="section-subtitle">批量上传 .docx/.doc/.rtf → 自动转换为 PDF（需 Windows + Word）</p></div>
    <file-upload multiple accept=".docx,.doc,.rtf"></file-upload>
    <button id="d2p-btn" class="btn btn-primary mt-16" disabled>▶ 批量转换</button>
    <progress-bar id="d2p-progress" style="display:none;"></progress-bar>
    <div id="d2p-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('d2p-btn');
  const progress = document.getElementById('d2p-progress');
  const resultDiv = document.getElementById('d2p-result');

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
    btn.textContent = e.detail.files.length ? `▶ 转换 ${e.detail.files.length} 个文件` : '▶ 批量转换';
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    progress.show(); progress.setProgress(-1, '正在转换（需要几秒到几分钟）...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/docx2pdf', fd);
      const contentType = res.headers.get('content-type') || '';
      if (contentType.includes('application/pdf') || contentType.includes('application/zip')) {
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const isZip = contentType.includes('zip');
        a.href = url;
        a.download = isZip ? 'pdf_output.zip' : 'output.pdf';
        a.click();
        URL.revokeObjectURL(url);
        progress.setProgress(100, '转换完成！');
        resultDiv.innerHTML = `<p style="color:#2E7D32;">✅ 转换成功 — ${isZip ? '已打包为 ZIP 下载' : 'PDF 已下载'}</p>`;
        showToast(`转换完成`);
      } else {
        const data = await res.json();
        progress.setProgress(0, '转换失败');
        resultDiv.innerHTML = `<p style="color:#C62828;">${data.error || '转换失败'}</p>`;
        if (data.results) {
          resultDiv.innerHTML += data.results.map(r =>
            `<p style="font-size:12px;">${r.status === 'ok' ? '✅' : '❌'} ${r.file} ${r.error || ''}</p>`
          ).join('');
        }
      }
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

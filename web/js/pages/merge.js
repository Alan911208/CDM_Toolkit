import { api, showToast } from '../app.js';
export default function render() {
  return `<div class="page-section"><h2>📑 工作簿合并</h2><p class="section-subtitle">将多个 Excel 文件合并为一个工作簿</p></div>
    <file-upload multiple accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label><input type="checkbox" id="merge-toc" checked> 自动生成 TOC 目录</label></div><div class="form-group"><label><input type="checkbox" id="merge-prefix" checked> 文件名作为 sheet 前缀</label></div></div>
    <button id="merge-btn" class="btn btn-primary" disabled>▶ 开始合并</button>
    <progress-bar id="merge-progress" style="display:none;"></progress-bar>`;
}
export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('merge-btn');
  const progress = document.getElementById('merge-progress');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; btn.textContent = e.detail.files.length ? `▶ 合并 ${e.detail.files.length} 个文件` : '▶ 开始合并'; });
  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    fd.append('create_toc', document.getElementById('merge-toc').checked);
    fd.append('add_prefix', document.getElementById('merge-prefix').checked);
    progress.show(); progress.setProgress(-1, '正在合并文件...'); btn.disabled = true;
    try { const res = await api('/api/merge', fd); const blob = await res.blob(); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'merged.xlsx'; a.click(); URL.revokeObjectURL(url); progress.setProgress(100, '合并完成！'); showToast('文件合并成功'); }
    catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(`合并失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

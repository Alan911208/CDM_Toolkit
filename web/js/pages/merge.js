import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📑 工作簿合并</h2><p class="section-subtitle">将多个 Excel 文件合并为一个工作簿</p></div>
    <file-upload multiple accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label><input type="checkbox" id="merge-toc" checked> 自动生成 TOC 目录</label></div><div class="form-group"><label><input type="checkbox" id="merge-prefix" checked> 文件名作为 sheet 前缀</label></div></div>
    <button id="merge-btn" class="btn btn-primary" disabled>▶ 开始合并</button>
    <progress-bar id="merge-progress" style="display:none;"></progress-bar>
    <div id="merge-result" style="margin-top:16px;"></div>
    <pipeline-nav context="merge"></pipeline-nav>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('merge-btn');
  const progress = document.getElementById('merge-progress');
  const resultDiv = document.getElementById('merge-result');

  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; btn.textContent = e.detail.files.length ? `▶ 合并 ${e.detail.files.length} 个文件` : '▶ 开始合并'; });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    upload.files.forEach(f => fd.append('files', f));
    fd.append('create_toc', document.getElementById('merge-toc').checked);
    fd.append('add_prefix', document.getElementById('merge-prefix').checked);
    progress.show(); progress.setProgress(-1, '正在合并文件...'); btn.disabled = true;
    resultDiv.innerHTML = '';
    try {
      const res = await api('/api/merge', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a'); a.href = url; a.download = 'merged.xlsx'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '合并完成！');
      resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ 文件合并成功 — 已加入工作区预览</p>';
      showToast('文件合并成功');
      window._cdmUploadToWorkspace(blob, 'merged.xlsx');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(`合并失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

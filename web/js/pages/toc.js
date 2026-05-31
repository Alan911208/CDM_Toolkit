import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📑 TOC 目录生成</h2><p class="section-subtitle">为工作簿生成带超链接的目录页（标准版 / SAS Derive 版）</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="toc-style">目录样式</label><select id="toc-style"><option value="standard">标准 TOC — 工作表名 + 超链接</option><option value="derive">SAS Derive TOC — 工作表名 + 描述列</option></select></div>
      <div class="form-group"><label for="toc-name">目录页名称</label><input type="text" id="toc-name" value="TOC" maxlength="31"></div>
    </div>
    <button id="toc-btn" class="btn btn-primary" disabled>▶ 生成目录</button>
    <progress-bar id="toc-progress" style="display:none;"></progress-bar>
    <div id="toc-result" style="margin-top:16px;"></div>
    <pipeline-nav context="toc"></pipeline-nav>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('toc-btn');
  const progress = document.getElementById('toc-progress');
  const resultDiv = document.getElementById('toc-result');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('style', document.getElementById('toc-style').value); fd.append('toc_name', document.getElementById('toc-name').value);
    progress.show(); progress.setProgress(-1, '正在生成目录...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/toc', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'toc_result.xlsx'; a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '目录生成完成！'); resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ TOC 目录已生成 — 已加入工作区预览</p>';
      showToast('TOC 目录已生成'); window._cdmUploadToWorkspace(blob, 'toc_result.xlsx');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

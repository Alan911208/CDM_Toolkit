import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📑 TOC 目录生成</h2><p class="section-subtitle">为工作簿生成带超链接的目录页（标准版 / SAS Derive 版）</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="toc-style">目录样式</label>
        <select id="toc-style">
          <option value="standard">标准 TOC — 工作表名 + 超链接</option>
          <option value="derive">SAS Derive TOC — 工作表名 + 描述列</option>
        </select>
      </div>
      <div class="form-group"><label for="toc-name">目录页名称</label>
        <input type="text" id="toc-name" value="TOC" maxlength="31">
      </div>
    </div>
    <div class="form-group" id="toc-desc-col-group" style="display:none;">
      <label for="toc-desc-col">描述来源列 (SAS Derive)</label>
      <input type="text" id="toc-desc-col" value="V2" maxlength="3" placeholder="如 V2, B1">
    </div>
    <button id="toc-btn" class="btn btn-primary" disabled>▶ 生成目录</button>
    <progress-bar id="toc-progress" style="display:none;"></progress-bar>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('toc-btn');
  const progress = document.getElementById('toc-progress');
  const styleSel = document.getElementById('toc-style');
  const descColGroup = document.getElementById('toc-desc-col-group');

  styleSel.addEventListener('change', () => {
    descColGroup.style.display = styleSel.value === 'derive' ? '' : 'none';
  });

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('style', styleSel.value);
    fd.append('toc_name', document.getElementById('toc-name').value);
    if (styleSel.value === 'derive') {
      // Note: current API doesn't pass desc_col separately — it uses V2 default
    }
    progress.show(); progress.setProgress(-1, '正在生成目录...'); btn.disabled = true;
    try {
      const res = await api('/api/toc', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'toc_result.xlsx'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '目录生成完成！'); showToast('TOC 目录已生成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

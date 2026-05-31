import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section"><h2>✂️ 工作表拆分</h2><p class="section-subtitle">按指定列的值将工作表拆分为多个 sheet</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="split-col">拆分列</label><input type="text" id="split-col" value="A" maxlength="3"></div><div class="form-group"><label for="split-header">表头行数</label><input type="number" id="split-header" value="1" min="0" max="10"></div></div>
    <button id="split-btn" class="btn btn-primary" disabled>▶ 开始拆分</button>
    <progress-bar id="split-progress" style="display:none;"></progress-bar>
    <div id="split-result" style="margin-top:16px;"></div>
    `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('split-btn');
  const progress = document.getElementById('split-progress');
  const resultDiv = document.getElementById('split-result');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('split_col', document.getElementById('split-col').value); fd.append('header_rows', document.getElementById('split-header').value);
    progress.show(); progress.setProgress(-1, '正在拆分...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/split', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'split.xlsx'; a.click(); URL.revokeObjectURL(url);
      showToast('工作表拆分成功');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(`拆分失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

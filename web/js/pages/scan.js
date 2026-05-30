import { api, showToast } from '../app.js';
export default function render() {
  return `<div class="page-section"><h2>🔍 特殊字符扫描</h2><p class="section-subtitle">扫描工作表中的非标准字符，生成 EDC 合规报告</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="scan-system">EDC 系统</label><select id="scan-system"><option value="Rave">Rave</option><option value="Clinflash">Clinflash</option></select></div></div>
    <button id="scan-btn" class="btn btn-primary" disabled>▶ 开始扫描</button>
    <progress-bar id="scan-progress" style="display:none;"></progress-bar>
    <div id="scan-results" style="margin-top:16px;"></div>`;
}
export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('scan-btn');
  const progress = document.getElementById('scan-progress');
  const resultsDiv = document.getElementById('scan-results');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('system', document.getElementById('scan-system').value);
    progress.show(); progress.setProgress(-1, '正在扫描...'); btn.disabled = true; resultsDiv.innerHTML = '';
    try {
      const res = await fetch('/api/scan-chars', { method: 'POST', body: fd }); const json = await res.json();
      progress.setProgress(100, `发现 ${json.total} 个含特殊字符的单元格`);
      if (json.total === 0) { resultsDiv.innerHTML = '<p style="color:#2E7D32;padding:16px;">✅ 未发现特殊字符！</p>'; showToast('扫描完成，未发现特殊字符'); }
      else {
        const table = document.createElement('data-table'); table.setAttribute('columns', JSON.stringify([{ key: 'sheet', label: 'Sheet' }, { key: 'cell', label: '单元格' }, { key: 'value', label: '内容' }])); table.data = json.findings; resultsDiv.appendChild(table);
        const exportBtn = document.createElement('button'); exportBtn.className = 'btn btn-outline mt-16'; exportBtn.textContent = '📥 导出 JSON 报告';
        exportBtn.addEventListener('click', () => { const blob = new Blob([JSON.stringify(json, null, 2)], { type: 'application/json' }); const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = `special_chars_${json.system}.json`; a.click(); });
        resultsDiv.appendChild(exportBtn); showToast(`发现 ${json.total} 个特殊字符`);
      }
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(`扫描失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

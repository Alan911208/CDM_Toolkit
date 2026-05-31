import { api, showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section"><h2>⚖️ 实验室单位转换</h2><p class="section-subtitle">将实验室检验结果统一转换为标准单位</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16"><div class="form-group"><label for="units-test-col">检验项目列</label><input type="text" id="units-test-col" value="D" maxlength="3"></div><div class="form-group"><label for="units-result-col">结果值列</label><input type="text" id="units-result-col" value="E" maxlength="3"></div><div class="form-group"><label for="units-unit-col">单位列</label><input type="text" id="units-unit-col" value="F" maxlength="3"></div></div>
    <div class="form-row"><div class="form-group"><label for="units-std">目标标准单位</label><input type="text" id="units-std" value="mg/dL"></div><div class="form-group"><label for="units-filter">检验项目筛选 (可选)</label><input type="text" id="units-filter" placeholder="留空=全部转换"></div></div>
    <button id="units-btn" class="btn btn-primary" disabled>▶ 执行转换</button>
    <progress-bar id="units-progress" style="display:none;"></progress-bar>
    <div id="units-result" style="margin-top:16px;"></div>
    `;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('units-btn');
  const progress = document.getElementById('units-progress');
  const resultDiv = document.getElementById('units-result');
  upload.addEventListener('files-changed', e => { btn.disabled = e.detail.files.length === 0; });
  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', upload.files[0]); fd.append('test_col', document.getElementById('units-test-col').value); fd.append('result_col', document.getElementById('units-result-col').value); fd.append('unit_col', document.getElementById('units-unit-col').value); fd.append('standard_unit', document.getElementById('units-std').value); fd.append('test_filter', document.getElementById('units-filter').value);
    progress.show(); progress.setProgress(-1, '转换中...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/unit-convert', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'units_result.xlsx'; a.click(); URL.revokeObjectURL(url);
    } catch (err) { showToast(`失败: ${err.message}`, 'error'); }
    finally { btn.disabled = false; }
  });
}

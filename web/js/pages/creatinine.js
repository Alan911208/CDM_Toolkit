import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🩺 医学计算器</h2><p class="section-subtitle">批量计算肌酐清除率 (Cockcroft-Gault) 或 eGFR (CKD-EPI 2021)</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="crcl-method">计算公式</label>
        <select id="crcl-method">
          <option value="CG">Cockcroft-Gault (CrCl) — 需要年龄/体重/肌酐/性别</option>
          <option value="CKD">CKD-EPI 2021 (eGFR) — 需要年龄/肌酐/性别</option>
        </select>
      </div>
    </div>
    <div class="form-row">
      <div class="form-group"><label for="crcl-age-col">年龄列</label><input type="text" id="crcl-age-col" value="C" maxlength="3"></div>
      <div class="form-group"><label for="crcl-scr-col">肌酐列 (Scr mg/dL)</label><input type="text" id="crcl-scr-col" value="D" maxlength="3"></div>
      <div class="form-group"><label for="crcl-gender-col">性别列</label><input type="text" id="crcl-gender-col" value="E" maxlength="3"></div>
      <div class="form-group"><label for="crcl-result-col">结果输出列</label><input type="text" id="crcl-result-col" value="G" maxlength="3"></div>
    </div>
    <div class="form-row" id="crcl-weight-group">
      <div class="form-group"><label for="crcl-weight-col">体重列 (kg) — 仅 CG</label><input type="text" id="crcl-weight-col" value="F" maxlength="3"></div>
      <div class="form-group"><label for="crcl-row-start">数据起始行</label><input type="number" id="crcl-row-start" value="2" min="1" max="100"></div>
    </div>
    <button id="crcl-btn" class="btn btn-primary" disabled>▶ 执行计算</button>
    <progress-bar id="crcl-progress" style="display:none;"></progress-bar>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('crcl-btn');
  const progress = document.getElementById('crcl-progress');
  const methodSel = document.getElementById('crcl-method');
  const weightGroup = document.getElementById('crcl-weight-group');

  methodSel.addEventListener('change', () => {
    weightGroup.style.display = methodSel.value === 'CG' ? '' : 'none';
  });
  methodSel.dispatchEvent(new Event('change'));

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('method', methodSel.value);
    fd.append('age_col', document.getElementById('crcl-age-col').value);
    fd.append('scr_col', document.getElementById('crcl-scr-col').value);
    fd.append('gender_col', document.getElementById('crcl-gender-col').value);
    fd.append('result_col', document.getElementById('crcl-result-col').value);
    fd.append('row_start', document.getElementById('crcl-row-start').value);
    if (methodSel.value === 'CG') {
      fd.append('weight_col', document.getElementById('crcl-weight-col').value);
    }

    progress.show(); progress.setProgress(-1, '正在计算...'); btn.disabled = true;
    try {
      const res = await api('/api/creatinine', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'creatinine_result.xlsx'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '计算完成！');
      showToast('医学计算完成');
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(err.message, 'error');
    }
    finally { btn.disabled = false; }
  });
}

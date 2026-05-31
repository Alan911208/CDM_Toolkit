import { api, showToast } from '../app.js';

// Method definitions with column requirements
const METHODS = [
  { id: 'CG',     name: 'Cockcroft-Gault (CrCl)',     desc: '肌酐清除率', cols: ['age','scr','gender','weight','result'], labels: {age:'年龄列', scr:'肌酐 Scr (mg/dL)', gender:'性别列 (F/M)', weight:'体重 (kg)', result:'结果输出列'} },
  { id: 'CKD',    name: 'CKD-EPI 2021 (eGFR)',        desc: '估算肾小球滤过率', cols: ['age','scr','gender','result'], labels: {age:'年龄列', scr:'肌酐 Scr (mg/dL)', gender:'性别列 (F/M)', result:'结果输出列'} },
  { id: 'MDRD',   name: 'MDRD (eGFR)',                desc: 'MDRD 研究方程', cols: ['age','scr','gender','result'], labels: {age:'年龄列', scr:'肌酐 Scr (mg/dL)', gender:'性别列 (F/M)', result:'结果输出列'} },
  { id: 'BMI',    name: 'BMI 体重指数',                desc: '体重(kg) / 身高(m)²', cols: ['weight','height','result'], labels: {weight:'体重 (kg)', height:'身高 (cm)', result:'结果输出列'} },
  { id: 'BSA',    name: 'BSA 体表面积 (Mosteller)',     desc: '√(身高×体重/3600)', cols: ['weight','height','result'], labels: {weight:'体重 (kg)', height:'身高 (cm)', result:'结果输出列'} },
  { id: 'LDL',    name: 'LDL 胆固醇 (Friedewald)',     desc: 'LDL = TC - HDL - TG/5', cols: ['tc','hdl','tg','result'], labels: {tc:'总胆固醇 TC', hdl:'HDL-C', tg:'甘油三酯 TG', result:'结果输出列'} },
  { id: 'CORRCA', name: '校正血钙',                     desc: '校正 Ca = Ca + 0.8×(4-Alb)', cols: ['ca','albumin','result'], labels: {ca:'血钙 Ca', albumin:'白蛋白 Albumin', result:'结果输出列'} },
  { id: 'IBW',    name: '理想体重 (Devine)',            desc: 'IBW = 50/45.5 + 2.3×(英寸-60)', cols: ['height','gender','result'], labels: {height:'身高 (cm)', gender:'性别列 (F/M)', result:'结果输出列'} },
  { id: 'AG',     name: '阴离子间隙',                   desc: 'AG = Na - (Cl + HCO₃)', cols: ['na','cl','hco3','result'], labels: {na:'Na⁺', cl:'Cl⁻', hco3:'HCO₃⁻', result:'结果输出列'} },
];

const DEFAULT_VALUES = { age:'C', scr:'D', gender:'E', weight:'F', height:'G', tc:'C', hdl:'D', tg:'E', ca:'C', albumin:'D', na:'C', cl:'D', hco3:'E', result:'H' };

export default function render() {
  return `<div class="page-section"><h2>🩺 医学计算器</h2><p class="section-subtitle">批量临床计算 — 支持 9 种常用公式</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group" style="flex:2;"><label for="crcl-method">计算公式</label>
        <select id="crcl-method">${METHODS.map(m => `<option value="${m.id}">${m.name} — ${m.desc}</option>`).join('')}</select>
      </div>
      <div class="form-group"><label for="crcl-row-start">数据起始行</label><input type="number" id="crcl-row-start" value="2" min="1" max="100"></div>
    </div>
    <div class="form-row" id="crcl-params"></div>
    <div class="form-group mt-16" id="crcl-formula-info" style="font-size:12px;color:#999;padding:8px 12px;background:#F5F7FA;border-radius:4px;"></div>
    <button id="crcl-btn" class="btn btn-primary" disabled>▶ 执行计算</button>
    <progress-bar id="crcl-progress" style="display:none;"></progress-bar>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('crcl-btn');
  const progress = document.getElementById('crcl-progress');
  const methodSel = document.getElementById('crcl-method');
  const paramsRow = document.getElementById('crcl-params');
  const formulaInfo = document.getElementById('crcl-formula-info');

  const formulaTexts = {
    CG: 'CrCl = ((140 - Age) × Weight<sub>kg</sub>) / (72 × Scr<sub>mg/dL</sub>) × 0.85 (if female)',
    CKD: 'eGFR = 142 × (Scr/A)<sup>B</sup> × 0.9938<sup>Age</sup> × 1.012 (if female)',
    MDRD: 'eGFR = 175 × Scr<sup>-1.154</sup> × Age<sup>-0.203</sup> × 0.742 (if female) × 1.212 (if black)',
    BMI: 'BMI = Weight<sub>kg</sub> / Height<sub>m</sub>²',
    BSA: 'BSA = √(Height<sub>cm</sub> × Weight<sub>kg</sub> / 3600)',
    LDL: 'LDL = TC - HDL - TG/5 (仅当 TG < 400 mg/dL)',
    CORRCA: 'Corrected Ca = Ca + 0.8 × (4.0 - Albumin)',
    IBW: 'Male: 50 + 2.3×(Ht<sub>in</sub>-60) | Female: 45.5 + 2.3×(Ht<sub>in</sub>-60)',
    AG: 'AG = Na⁺ - (Cl⁻ + HCO₃⁻)  (正常: 8-12 mEq/L)',
  };

  function renderParams() {
    const method = METHODS.find(m => m.id === methodSel.value);
    if (!method) return;

    formulaInfo.innerHTML = `<strong>${method.name}</strong>: ${formulaTexts[method.id] || ''}`;
    paramsRow.innerHTML = method.cols.filter(c => c !== 'result').map(c => {
      const label = method.labels[c] || c;
      return `<div class="form-group"><label>${label}</label><input type="text" id="crcl-${c}" value="${DEFAULT_VALUES[c] || ''}" maxlength="3"></div>`;
    }).join('') + `<div class="form-group"><label>${method.labels['result'] || '结果输出列'}</label><input type="text" id="crcl-result" value="${DEFAULT_VALUES['result'] || 'H'}" maxlength="3"></div>`;
  }

  methodSel.addEventListener('change', renderParams);
  renderParams();

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const method = methodSel.value;
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('method', method);
    fd.append('row_start', document.getElementById('crcl-row-start').value);

    const methodDef = METHODS.find(m => m.id === method);
    if (methodDef) {
      methodDef.cols.forEach(c => {
        const el = document.getElementById('crcl-' + c);
        if (el) fd.append(c + '_col', el.value);
      });
    }

    progress.show(); progress.setProgress(-1, '正在计算...'); btn.disabled = true;
    try {
      const res = await api('/api/medical-calc', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = `${method.toLowerCase()}_result.xlsx`; a.click();
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

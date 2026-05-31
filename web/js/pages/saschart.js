import { api, showToast } from '../app.js';

const CHART_TYPES = [
  { id: 'bar', name: '柱状图', icon: '📊' },
  { id: 'line', name: '折线图', icon: '📈' },
  { id: 'pie', name: '饼图', icon: '🥧' },
  { id: 'scatter', name: '散点图', icon: '🔵' },
];

export default function render() {
  return `<div class="page-section"><h2>📊 SAS 数据集图表</h2><p class="section-subtitle">上传 .sas7bdat 文件 → 选择列 → 选择图表类型 → 生成 Excel 图表</p></div>
    <file-upload accept=".sas7bdat"></file-upload>
    <div id="saschart-preview" style="margin-top:16px;"></div>
    <div class="form-row mt-16" id="saschart-params" style="display:none;">
      <div class="form-group"><label for="saschart-x">X 轴 / 分类列</label><select id="saschart-x"></select></div>
      <div class="form-group"><label for="saschart-y">Y 轴 / 数值列</label><select id="saschart-y"></select></div>
      <div class="form-group"><label for="saschart-group">分组列 (可选)</label><select id="saschart-group"><option value="">无分组</option></select></div>
    </div>
    <div class="form-row" id="saschart-type-row" style="display:none;">
      <div class="form-group"><label>图表类型</label>
        <div style="display:flex;gap:8px;margin-top:4px;">${CHART_TYPES.map(t => `<label style="display:flex;align-items:center;gap:4px;padding:6px 12px;border:1px solid var(--color-border);border-radius:8px;cursor:pointer;" id="saschart-label-${t.id}"><input type="radio" name="saschart-type" value="${t.id}" ${t.id==='bar'?'checked':''}> ${t.icon} ${t.name}</label>`).join('')}</div>
      </div>
      <div class="form-group"><label for="saschart-title">图表标题</label><input type="text" id="saschart-title" value="SAS Data Chart"></div>
    </div>
    <button id="saschart-btn" class="btn btn-primary mt-16" disabled>📊 生成图表</button>
    <progress-bar id="saschart-progress" style="display:none;"></progress-bar>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('saschart-btn');
  const progress = document.getElementById('saschart-progress');
  const previewDiv = document.getElementById('saschart-preview');
  const paramsRow = document.getElementById('saschart-params');
  const typeRow = document.getElementById('saschart-type-row');
  const xSel = document.getElementById('saschart-x');
  const ySel = document.getElementById('saschart-y');
  const gSel = document.getElementById('saschart-group');

  let previewData = null;

  upload.addEventListener('files-changed', async (e) => {
    if (e.detail.files.length === 0) { btn.disabled = true; return; }
    previewDiv.innerHTML = '<p style="color:#999;">加载数据预览...</p>';
    btn.disabled = true; paramsRow.style.display = 'none'; typeRow.style.display = 'none';

    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('rows', '50');
    try {
      const res = await fetch('/api/sas-preview', { method: 'POST', body: fd });
      previewData = await res.json();
      if (previewData.error) { previewDiv.innerHTML = `<p style="color:#C62828;">${previewData.error}</p>`; return; }

      // Show preview table
      const h = previewData.headers || [];
      const r = previewData.rows || [];
      previewDiv.innerHTML = `<div style="font-size:12px;color:#999;margin-bottom:4px;">共 ${previewData.total_rows} 行 | ${h.length} 列</div>
        <div style="max-height:200px;overflow:auto;border:1px solid var(--color-border);border-radius:4px;">
          <table style="width:100%;border-collapse:collapse;font-size:11px;">
            <thead><tr>${h.map(c => `<th style="background:#4472C4;color:white;padding:4px 8px;text-align:left;white-space:nowrap;">${c}</th>`).join('')}</tr></thead>
            <tbody>${r.map(row => `<tr>${row.map(v => `<td style="padding:2px 8px;border-bottom:1px solid #eee;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${v}</td>`).join('')}</tr>`).join('')}</tbody>
          </table></div>`;

      // Populate column selectors
      const opts = h.map(c => `<option value="${c}">${c}</option>`).join('');
      xSel.innerHTML = opts;
      ySel.innerHTML = opts;
      gSel.innerHTML = '<option value="">无分组</option>' + opts;

      // Auto-select numeric columns for Y
      if (previewData.numeric_cols && previewData.numeric_cols.length > 0) {
        ySel.value = previewData.numeric_cols[0];
      }

      paramsRow.style.display = '';
      typeRow.style.display = '';
      btn.disabled = false;
    } catch (err) {
      previewDiv.innerHTML = `<p style="color:#C62828;">预览失败: ${err.message}</p>`;
    }
  });

  btn.addEventListener('click', async () => {
    const chartType = document.querySelector('input[name="saschart-type"]:checked')?.value || 'bar';
    const fd = new FormData();
    fd.append('file', upload.files[0]);
    fd.append('chart_type', chartType);
    fd.append('x_col', xSel.value);
    fd.append('y_col', ySel.value);
    fd.append('group_col', gSel.value || '');
    fd.append('title', document.getElementById('saschart-title').value || 'SAS Data Chart');

    progress.show(); progress.setProgress(-1, '正在生成图表...'); btn.disabled = true;
    try {
      const res = await api('/api/sas-chart', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'sas_chart.xlsx'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '图表生成完成！');
      showToast('图表已生成 — Excel 文件包含数据和图表两个工作表');
    } catch (err) {
      progress.setProgress(0, `失败: ${err.message}`);
      showToast(err.message, 'error');
    }
    finally { btn.disabled = false; }
  });
}

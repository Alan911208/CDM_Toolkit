import { showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔬 Clinical Raw Data QC Portal</h2><p class="section-subtitle">上传 SAS 数据集 → 自动扫描 Metadata → 生成 QC 报告</p></div>
    <file-upload accept=".sas7bdat"></file-upload>
    <progress-bar id="qc-progress" style="display:none;"></progress-bar>
    <div id="qc-results" style="margin-top:16px;"></div>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const progress = document.getElementById('qc-progress');
  const resultsDiv = document.getElementById('qc-results');

  upload.addEventListener('files-changed', async (e) => {
    if (e.detail.files.length === 0) return;
    progress.show(); progress.setProgress(-1, '正在分析数据集...');
    resultsDiv.innerHTML = '';

    const fd = new FormData();
    fd.append('file', upload.files[0]);
    try {
      const res = await fetch('/api/qc/analyze', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { resultsDiv.innerHTML = `<p style="color:#C62828;">${data.error}</p>`; progress.hide(); return; }
      renderReport(data);
      progress.setProgress(100, '分析完成');
    } catch (err) {
      resultsDiv.innerHTML = `<p style="color:#C62828;">分析失败: ${err.message}</p>`;
      progress.hide();
    }
  });

  function renderReport(data) {
    const ds = data.dataset || {};
    const vars = data.variables || [];
    const findings = data.findings || [];
    const sum = data.summary || {};

    const statusColor = sum.errors > 0 ? '#C62828' : sum.warnings > 0 ? '#E65100' : '#2E7D32';
    const statusIcon = sum.errors > 0 ? '❌' : sum.warnings > 0 ? '⚠️' : '✅';

    resultsDiv.innerHTML = `
      <!-- Summary Bar -->
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:16px;">
        <div style="padding:12px 20px;background:${statusColor};color:white;border-radius:8px;font-size:14px;font-weight:600;">${statusIcon} QC Pass Rate: ${sum.pass_rate || 0}%</div>
        <div class="qc-stat"><strong>${ds.row_count}</strong> Rows</div>
        <div class="qc-stat"><strong>${ds.col_count}</strong> Variables</div>
        <div class="qc-stat" style="color:#2E7D32;">📋 <strong>${sum.info || 0}</strong> Info</div>
        <div class="qc-stat" style="color:#E65100;">⚠️ <strong>${sum.warnings || 0}</strong> Warnings</div>
        <div class="qc-stat" style="color:#C62828;">❌ <strong>${sum.errors || 0}</strong> Errors</div>
      </div>

      <!-- Tabs -->
      <div class="qc-tabs">
        <button class="qc-tab active" data-tab="findings">🔍 QC Findings (${findings.length})</button>
        <button class="qc-tab" data-tab="variables">📋 Variables (${vars.length})</button>
        <button class="qc-tab" data-tab="dataset">📦 Dataset Info</button>
      </div>

      <!-- QC Findings Tab -->
      <div class="qc-panel" id="tab-findings">
        ${findings.length === 0 ? '<p style="color:#2E7D32;padding:16px;">✅ 未发现问题！</p>' : `
        <div style="max-height:500px;overflow:auto;">
          <table class="qc-table">
            <thead><tr><th style="width:60px;">级别</th><th style="width:100px;">类别</th><th>问题</th><th>详情</th></tr></thead>
            <tbody>${findings.map(f => `<tr>
              <td><span class="qc-badge qc-${f.level}">${f.level === 'error' ? '❌' : f.level === 'warning' ? '⚠️' : 'ℹ️'}</span></td>
              <td>${f.category}</td>
              <td>${f.message}</td>
              <td style="font-size:11px;color:#999;">${f.detail || ''}</td>
            </tr>`).join('')}</tbody>
          </table></div>`}
      </div>

      <!-- Variables Tab -->
      <div class="qc-panel" id="tab-variables" style="display:none;">
        <div style="max-height:500px;overflow:auto;">
          <table class="qc-table">
            <thead><tr><th>#</th><th>Variable</th><th>Type</th><th>Missing</th><th>Unique</th><th>Sample Values</th></tr></thead>
            <tbody>${vars.map((v, i) => `<tr>
              <td>${i + 1}</td>
              <td><strong>${v.name}</strong></td>
              <td>${v.type}</td>
              <td style="color:${v.missing_pct > 20 ? '#C62828' : '#666'};">${v.missing_pct}%</td>
              <td>${v.unique}</td>
              <td style="font-size:11px;max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${(v.sample || []).join(', ') || '-'}</td>
            </tr>`).join('')}</tbody>
          </table></div>
      </div>

      <!-- Dataset Tab -->
      <div class="qc-panel" id="tab-dataset" style="display:none;">
        <div class="cards" style="grid-template-columns:repeat(auto-fill, minmax(200px, 1fr));">
          <div class="card" style="padding:16px;"><h4>📄 文件名</h4><p>${ds.file_name || '-'}</p></div>
          <div class="card" style="padding:16px;"><h4>📝 行数</h4><p>${ds.row_count || 0}</p></div>
          <div class="card" style="padding:16px;"><h4>📐 变量数</h4><p>${ds.col_count || 0}</p></div>
          <div class="card" style="padding:16px;"><h4>💾 内存</h4><p>${ds.memory_mb || 0} MB</p></div>
        </div>
        <h4 style="margin-top:16px;">Variable Distribution</h4>
        <div style="display:flex;gap:8px;flex-wrap:wrap;">
          ${['Char', 'Num'].map(t => {
            const count = vars.filter(v => v.type === t).length;
            return `<span style="padding:4px 12px;background:#F5F7FA;border-radius:16px;font-size:12px;">${t}: ${count}</span>`;
          }).join('')}
        </div>
      </div>
      <style>
        .qc-stat { padding:12px 16px; background:var(--color-surface); border-radius:8px; font-size:13px; border:1px solid var(--color-border); }
        .qc-tabs { display:flex; gap:4px; margin-bottom:0; border-bottom:2px solid var(--color-border); }
        .qc-tab { padding:10px 20px; border:none; background:none; cursor:pointer; font-size:13px; color:#666; border-bottom:2px solid transparent; margin-bottom:-2px; font-family:var(--font); }
        .qc-tab.active { color:var(--color-primary); border-bottom-color:var(--color-primary); font-weight:600; }
        .qc-panel { padding:16px 0; }
        .qc-table { width:100%; border-collapse:collapse; font-size:13px; }
        .qc-table th { background:#4472C4; color:white; padding:8px 12px; text-align:left; font-weight:600; position:sticky; top:0; }
        .qc-table td { padding:8px 12px; border-bottom:1px solid #E0E4E8; }
        .qc-table tr:hover { background:#F5F7FA; }
        .qc-badge { display:inline-block; padding:2px 8px; border-radius:4px; font-size:11px; font-weight:600; }
        .qc-error { background:#FFEBEE; color:#C62828; }
        .qc-warning { background:#FFF3E0; color:#E65100; }
        .qc-info { background:#E3F2FD; color:#1565C0; }
      </style>`;

    // Tab switching
    resultsDiv.querySelectorAll('.qc-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        resultsDiv.querySelectorAll('.qc-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        resultsDiv.querySelectorAll('.qc-panel').forEach(p => p.style.display = 'none');
        const panel = resultsDiv.querySelector('#tab-' + tab.dataset.tab);
        if (panel) panel.style.display = '';
      });
    });
  }
}

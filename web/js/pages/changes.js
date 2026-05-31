import { api, showToast } from '../app.js';

export default function render() {
  return `<workspace-panel></workspace-panel>
    <div class="page-section"><h2>📝 修改痕迹跟踪</h2><p class="section-subtitle">对比原版文件，标记所有单元格修改（绿色填充 + 红色边框 + 批注）</p></div>
    <div class="form-row"><div class="form-group"><label>当前文件</label><file-upload id="changes-current" accept=".xlsx,.xls"></file-upload></div><div class="form-group"><label>原始文件（备份）</label><file-upload id="changes-original" accept=".xlsx,.xls"></file-upload></div></div>
    <div class="form-row mt-16"><div class="form-group"><label><input type="checkbox" id="changes-comments" checked> 添加批注 (旧值 → 新值)</label></div><div class="form-group"><label><input type="checkbox" id="changes-highlight" checked> 高亮修改单元格</label></div></div>
    <button id="changes-btn" class="btn btn-primary" disabled>▶ 跟踪修改</button>
    <progress-bar id="changes-progress" style="display:none;"></progress-bar>
    <div id="changes-result" style="margin-top:16px;"></div>
    <pipeline-nav context="changes"></pipeline-nav>`;
}

export async function init() {
  const curUp = document.getElementById('changes-current'), origUp = document.getElementById('changes-original');
  const btn = document.getElementById('changes-btn');
  const progress = document.getElementById('changes-progress');
  const resultDiv = document.getElementById('changes-result');
  function chk() { btn.disabled = !(curUp.files.length > 0 && origUp.files.length > 0); }
  curUp.addEventListener('files-changed', chk); origUp.addEventListener('files-changed', chk);

  btn.addEventListener('click', async () => {
    const fd = new FormData(); fd.append('file', curUp.files[0]); fd.append('original', origUp.files[0]); fd.append('add_comments', document.getElementById('changes-comments').checked); fd.append('highlight', document.getElementById('changes-highlight').checked);
    progress.show(); progress.setProgress(-1, '正在对比文件...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/track-changes', fd); const blob = await res.blob();
      const url = URL.createObjectURL(blob); const a = document.createElement('a'); a.href = url; a.download = 'tracked_changes.xlsx'; a.click(); URL.revokeObjectURL(url);
      progress.setProgress(100, '修改跟踪完成！'); resultDiv.innerHTML = '<p style="color:#2E7D32;padding:8px;">✅ 修改痕迹已标记 — 已加入工作区预览</p>';
      showToast('修改痕迹已标记'); window._cdmUploadToWorkspace(blob, 'tracked_changes.xlsx');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

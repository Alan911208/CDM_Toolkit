import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔍 数据集比较</h2><p class="section-subtitle">上传旧版和新版 SAS 数据集 → 比对异同 → 生成 Excel 报告</p></div>
    <div class="form-row">
      <div class="form-group">
        <label style="font-weight:600;">📂 旧版本数据集 (Old)</label>
        <file-upload id="cmp-old" multiple accept=".sas7bdat"></file-upload>
      </div>
      <div class="form-group">
        <label style="font-weight:600;">📂 新版本数据集 (New)</label>
        <file-upload id="cmp-new" multiple accept=".sas7bdat"></file-upload>
      </div>
    </div>
    <div class="form-row mt-16">
      <div class="form-group"><label for="cmp-key">Key 变量（逗号分隔，用于行级比较）</label><input type="text" id="cmp-key" placeholder="如 USUBJID 或 USUBJID,Visit" style="width:100%;"></div>
      <div class="form-group" style="align-self:flex-end;padding-bottom:16px;"><button id="cmp-btn" class="btn btn-primary" disabled>▶ 开始比较</button></div>
    </div>
    <progress-bar id="cmp-progress" style="display:none;"></progress-bar>
    <div id="cmp-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  const oldUp = document.getElementById('cmp-old');
  const newUp = document.getElementById('cmp-new');
  const btn = document.getElementById('cmp-btn');
  const progress = document.getElementById('cmp-progress');
  const resultDiv = document.getElementById('cmp-result');

  function checkReady() {
    btn.disabled = !(oldUp.files.length > 0 && newUp.files.length > 0);
    btn.textContent = (oldUp.files.length && newUp.files.length) ? `▶ 比较 ${oldUp.files.length} vs ${newUp.files.length}` : '▶ 开始比较';
  }
  oldUp.addEventListener('files-changed', checkReady);
  newUp.addEventListener('files-changed', checkReady);

  btn.addEventListener('click', async () => {
    const fd = new FormData();
    oldUp.files.forEach(f => fd.append('old_files', f));
    newUp.files.forEach(f => fd.append('new_files', f));
    fd.append('key_vars', document.getElementById('cmp-key').value);

    progress.show(); progress.setProgress(-1, '正在比较数据集...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      const res = await api('/api/compare-datasets', fd);
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url; a.download = 'comparison_report.xlsx'; a.click();
      URL.revokeObjectURL(url);
      progress.setProgress(100, '比较完成！');
      resultDiv.innerHTML = `<p style="color:#2E7D32;padding:8px;">✅ 比较报告已生成</p>
        <div style="font-size:12px;color:#666;line-height:1.6;">
          <strong>报告内容：</strong><br>
          • Overview 工作表：所有数据集的汇总比较（状态/行数/变量/差异）<br>
          • 各数据集工作表：详细的 Schema 变更 + 行级变更<br>
          • 颜色标记：🟢 新增/一致  🔴 删除  🟡 修改
        </div>`;
      showToast('比较报告已生成');
    } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
    finally { btn.disabled = false; }
  });
}

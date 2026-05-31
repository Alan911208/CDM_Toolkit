import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔗 拼接数据</h2><p class="section-subtitle">上传多个数据集（SAS/Excel/CSV），支持 Union 和 Join 操作</p></div>
    <file-upload multiple accept=".sas7bdat,.xlsx,.xls,.csv"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="dm-op">操作类型</label><select id="dm-op"><option value="union">Union — 纵向拼接（行合并）</option><option value="join">Join — 横向关联（按 Key 合并）</option></select></div>
      <div class="form-group"><label for="dm-how">Join 方式</label><select id="dm-how"><option value="inner">Inner Join</option><option value="left">Left Join</option><option value="right">Right Join</option><option value="outer">Outer Join</option></select></div>
    </div>
    <div class="form-row" id="dm-key-row">
      <div class="form-group"><label for="dm-key">Key 列（逗号分隔多个）</label><input type="text" id="dm-key" placeholder="USUBJID"></div>
      <div class="form-group"><label for="dm-lkey">左表 Key（不同名时填写）</label><input type="text" id="dm-lkey" placeholder="留空=与 Key 相同"></div>
      <div class="form-group"><label for="dm-rkey">右表 Key（不同名时填写）</label><input type="text" id="dm-rkey" placeholder="留空=与 Key 相同"></div>
    </div>
    <button id="dm-btn" class="btn btn-primary" disabled>▶ 执行</button>
    <progress-bar id="dm-progress" style="display:none;"></progress-bar>
    <div id="dm-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  var upload = document.querySelector('file-upload');
  var btn = document.getElementById('dm-btn'), progress = document.getElementById('dm-progress');
  var opSel = document.getElementById('dm-op'), howSel = document.getElementById('dm-how');
  var keyRow = document.getElementById('dm-key-row'), resultDiv = document.getElementById('dm-result');

  opSel.addEventListener('change', function() {
    var isJoin = opSel.value === 'join';
    keyRow.style.display = isJoin ? '' : 'none';
    howSel.style.display = isJoin ? '' : 'none';
    btn.disabled = upload.files.length === 0 || (isJoin && upload.files.length < 2);
  });

  upload.addEventListener('files-changed', function(e) {
    var n = e.detail.files.length;
    var isJoin = opSel.value === 'join';
    btn.disabled = n === 0 || (isJoin && n < 2);
    btn.textContent = n ? '▶ ' + (opSel.value === 'union' ? 'Union ' + n + ' 个文件' : 'Join ' + n + ' 个文件') : '▶ 执行';
  });

  btn.addEventListener('click', function() {
    var fd = new FormData();
    fd.append('operation', opSel.value);
    fd.append('how', howSel.value);
    fd.append('key', document.getElementById('dm-key').value);
    fd.append('left_key', document.getElementById('dm-lkey').value);
    fd.append('right_key', document.getElementById('dm-rkey').value);
    upload.files.forEach(function(f) { fd.append('files', f); });

    progress.show(); progress.setProgress(-1, '处理中...'); btn.disabled = true; resultDiv.innerHTML = '';
    try {
      api('/api/data-merge', fd).then(function(res) {
        return res.blob();
      }).then(function(blob) {
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url; a.download = opSel.value + '_result.xlsx'; a.click();
        URL.revokeObjectURL(url);
        progress.setProgress(100, '完成！');
        resultDiv.innerHTML = '<p style="color:#2E7D32;">✅ ' + (opSel.value === 'union' ? 'Union' : 'Join') + ' 完成 — 已下载</p>';
        showToast('拼接完成');
        btn.disabled = false;
      }).catch(function(err) {
        progress.setProgress(0, '失败'); resultDiv.innerHTML = '<p style="color:#C62828;">' + err.message + '</p>';
        btn.disabled = false;
      });
    } catch (err) {
      progress.setProgress(0, '失败'); btn.disabled = false;
    }
  });
}

import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🔗 拼接数据</h2><p class="section-subtitle">上传多个数据集（SAS/Excel/CSV），先预览 → 再 Union/Join</p></div>
    <file-upload multiple accept=".sas7bdat,.xlsx,.xls,.csv"></file-upload>
    <div id="dm-summary" style="margin-top:16px;"></div>
    <div class="form-row mt-16" id="dm-params" style="display:none;">
      <div class="form-group"><label for="dm-op">操作</label><select id="dm-op"><option value="union">Union — 纵向拼接（行合并）</option><option value="join">Join — 横向关联（按 Key）</option></select></div>
      <div class="form-group"><label for="dm-how">Join 方式</label><select id="dm-how"><option value="inner">Inner</option><option value="left">Left</option><option value="right">Right</option><option value="outer">Outer</option></select></div>
    </div>
    <div class="form-row" id="dm-key-row" style="display:none;">
      <div class="form-group"><label for="dm-key">Key 列</label><input type="text" id="dm-key" placeholder="USUBJID"></div>
      <div class="form-group"><label for="dm-lkey">左表 Key（不同名）</label><input type="text" id="dm-lkey" placeholder="留空=Key"></div>
      <div class="form-group"><label for="dm-rkey">右表 Key（不同名）</label><input type="text" id="dm-rkey" placeholder="留空=Key"></div>
    </div>
    <button id="dm-btn" class="btn btn-primary mt-16" style="display:none;" disabled>▶ 执行</button>
    <progress-bar id="dm-progress" style="display:none;"></progress-bar>
    <div id="dm-result" style="margin-top:16px;"></div>`;
}

export async function init() {
  var upload = document.querySelector('file-upload');
  var btn = document.getElementById('dm-btn'), progress = document.getElementById('dm-progress');
  var opSel = document.getElementById('dm-op'), howSel = document.getElementById('dm-how');
  var keyRow = document.getElementById('dm-key-row'), paramsRow = document.getElementById('dm-params');
  var summaryDiv = document.getElementById('dm-summary'), resultDiv = document.getElementById('dm-result');

  upload.addEventListener('files-changed', function(e) {
    if (e.detail.files.length === 0) {
      summaryDiv.innerHTML = ''; paramsRow.style.display = 'none';
      keyRow.style.display = 'none'; btn.style.display = 'none'; return;
    }
    // Preview files
    var fd = new FormData();
    fd.append('operation', 'union'); fd.append('preview_only', 'true');
    e.detail.files.forEach(function(f) { fd.append('files', f); });

    fetch('/api/data-merge', { method: 'POST', body: fd }).then(function(r) { return r.json(); }).then(function(data) {
      if (data.error) { summaryDiv.innerHTML = '<p style="color:#C62828;">' + data.error + '</p>'; return; }
      var files = data.files || [];
      summaryDiv.innerHTML = '<div style="font-size:12px;font-weight:600;margin-bottom:4px;">📋 上传文件概览</div>' +
        '<table class="dm-table"><thead><tr><th>文件</th><th>格式</th><th>行数</th><th>列数</th><th>列名</th><th>大小</th></tr></thead><tbody>' +
        files.map(function(f) {
          var errStyle = f.error ? 'color:#C62828;' : '';
          return '<tr><td style="' + errStyle + '">' + (f.error ? '⚠️ ' : '') + f.file + '</td><td>' + (f.format || '') + '</td><td>' + f.rows + '</td><td>' + f.cols + '</td><td style="font-size:11px;max-width:300px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">' + (f.columns || []).join(', ') + '</td><td>' + (f.memory || '') + '</td></tr>';
        }).join('') +
        (data.total_rows ? '<tr style="font-weight:600;background:#E3F2FD;"><td colspan="2">Union 预计总行数</td><td>' + data.total_rows + '</td><td colspan="3"></td></tr>' : '') +
        '</tbody></table>' +
        '<style>.dm-table{width:100%;border-collapse:collapse;font-size:13px;}.dm-table th{background:#4472C4;color:white;padding:6px 10px;text-align:left;}.dm-table td{padding:6px 10px;border-bottom:1px solid #E0E4E8;}</style>';

      paramsRow.style.display = ''; btn.style.display = '';
      checkReady();
    }).catch(function(err) {
      summaryDiv.innerHTML = '<p style="color:#C62828;">' + err.message + '</p>';
    });
  });

  function checkReady() {
    var n = upload.files.length;
    var isJoin = opSel.value === 'join';
    btn.disabled = n === 0 || (isJoin && n < 2);
    btn.textContent = n ? '▶ ' + (opSel.value === 'union' ? 'Union ' + n + ' 个文件' : 'Join ' + n + ' 个文件') : '▶ 执行';
  }

  opSel.addEventListener('change', function() {
    var isJoin = opSel.value === 'join';
    keyRow.style.display = isJoin ? '' : 'none';
    howSel.style.display = isJoin ? '' : 'none';
    checkReady();
  });

  btn.addEventListener('click', function() {
    var fd = new FormData();
    fd.append('operation', opSel.value); fd.append('how', howSel.value);
    fd.append('key', document.getElementById('dm-key').value);
    fd.append('left_key', document.getElementById('dm-lkey').value);
    fd.append('right_key', document.getElementById('dm-rkey').value);
    upload.files.forEach(function(f) { fd.append('files', f); });

    progress.show(); progress.setProgress(-1, '处理中...'); btn.disabled = true; resultDiv.innerHTML = '';
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
    }).catch(function(err) {
      progress.setProgress(0, '失败'); resultDiv.innerHTML = '<p style="color:#C62828;">' + err.message + '</p>';
    }).finally(function() { btn.disabled = false; });
  });
}

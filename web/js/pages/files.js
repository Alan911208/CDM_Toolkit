import { showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>📁 文件目录浏览</h2><p class="section-subtitle">点击文件夹展开浏览，支持导出 TXT 格式</p></div>
    <div id="files-bread" style="font-size:12px;margin-bottom:8px;color:#666;">我的电脑</div>
    <div id="files-tree" style="background:#1E1E1E;color:#D4D4D4;padding:12px 16px;border-radius:8px;font-family:'Consolas',monospace;font-size:12px;line-height:1.7;max-height:65vh;overflow:auto;">加载中...</div>
    <div style="margin-top:8px;display:flex;gap:8px;align-items:center;">
      <button id="files-txt" class="btn btn-outline" style="font-size:11px;">📥 导出 TXT</button>
      <span id="files-info" style="font-size:11px;color:#999;"></span>
    </div>
    <style>.ft-dir{cursor:pointer;}.ft-dir:hover{color:#4472C4;text-decoration:underline;}.ft-file{color:#999;}</style>`;
}

export async function init() {
  var treeEl = document.getElementById('files-tree');
  var breadEl = document.getElementById('files-bread');
  var infoEl = document.getElementById('files-info');
  var currentPath = '';

  function load(path) {
    currentPath = path;
    treeEl.innerHTML = '<span style="color:#888;">加载中...</span>';
    var url = '/api/file-tree';
    if (path) url += '?path=' + encodeURIComponent(path);

    fetch(url).then(function(res) {
      if (!res.ok) throw new Error('HTTP ' + res.status);
      return res.json();
    }).then(function(data) {
      if (data.error) { treeEl.innerHTML = '<span style="color:#C62828;">' + data.error + '</span>'; return; }

      // Breadcrumb
      var parts = path ? path.split('\\').filter(function(x) { return x; }) : [];
      if (parts.length === 0) {
        breadEl.textContent = '我的电脑';
      } else {
        breadEl.innerHTML = '';
        parts.forEach(function(p, i) {
          var pp = parts.slice(0, i + 1).join('\\') + '\\';
          var link = document.createElement('a');
          link.href = '#';
          link.textContent = p;
          link.style.cssText = 'color:var(--color-primary);text-decoration:none;cursor:pointer;';
          link.addEventListener('click', function(e) { e.preventDefault(); load(pp); });
          breadEl.appendChild(link);
          if (i < parts.length - 1) breadEl.appendChild(document.createTextNode(' › '));
        });
      }

      // Tree
      treeEl.innerHTML = '';
      var dirs = data.items.filter(function(i) { return i.is_dir; });
      var files = data.items.filter(function(i) { return !i.is_dir; });

      if (!path) {
        var hdr = document.createElement('div');
        hdr.style.color = '#888';
        hdr.textContent = '📁 我的电脑';
        treeEl.appendChild(hdr);
      }

      dirs.forEach(function(d) {
        var div = document.createElement('div');
        div.className = 'ft-dir';
        div.textContent = '📁 ' + d.name;
        div.addEventListener('click', function() { load(d.path); });
        treeEl.appendChild(div);
      });

      files.forEach(function(f) {
        var div = document.createElement('div');
        div.className = 'ft-file';
        var s = f.size ? ' (' + fmt(f.size) + ')' : '';
        div.textContent = '📄 ' + f.name + s;
        treeEl.appendChild(div);
      });

      if (!dirs.length && !files.length) {
        var empty = document.createElement('div');
        empty.style.color = '#888';
        empty.textContent = '空文件夹';
        treeEl.appendChild(empty);
      }

      infoEl.textContent = dirs.length + ' 文件夹, ' + files.length + ' 文件';
    }).catch(function(err) {
      treeEl.innerHTML = '<span style="color:#C62828;">加载失败: ' + err.message + '</span>';
    });
  }

  document.getElementById('files-txt').addEventListener('click', function() {
    if (!currentPath) { showToast('请先选择一个文件夹', 'error'); return; }
    var fd = new FormData(); fd.append('path', currentPath);
    fetch('/api/file-tree-txt', { method: 'POST', body: fd }).then(function(res) {
      return res.text();
    }).then(function(text) {
      var blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'file_tree_' + new Date().toISOString().slice(0, 10) + '.txt';
      a.click();
      showToast('已下载');
    }).catch(function(err) {
      showToast('下载失败: ' + err.message, 'error');
    });
  });

  load('');
}

function fmt(b) {
  if (b < 1024) return b + ' B';
  if (b < 1024 * 1024) return (b / 1024).toFixed(1) + ' KB';
  return (b / (1024 * 1024)).toFixed(1) + ' MB';
}

import { api, showToast } from '../app.js';

const TOOLS = [
  { id:'sas2xlsx', name:'SAS → Excel', desc:'.sas7bdat → .xlsx', accept:'.sas7bdat', ext:'xlsx' },
  { id:'sas2csv',  name:'SAS → CSV',   desc:'.sas7bdat → .csv', accept:'.sas7bdat', ext:'csv' },
  { id:'csv2sas',  name:'CSV → SAS',   desc:'.csv → .sas7bdat', accept:'.csv', ext:'sas7bdat' },
  { id:'xlsx2sas', name:'Excel → SAS', desc:'.xlsx → .sas7bdat', accept:'.xlsx,.xls', ext:'sas7bdat' },
];

export default function render() {
  return `<div class="page-section"><h2>🔄 SAS 转换工具</h2><p class="section-subtitle">SAS / CSV / Excel 文件格式互转（Python 实现，无需 SAS 许可）</p></div>
    <div class="cards" style="margin-bottom:16px;">${TOOLS.map(t => `
      <div class="card" style="padding:16px;" id="card-${t.id}">
        <h4>${t.name}</h4><p>${t.desc}</p>
        <file-upload id="upload-${t.id}" accept="${t.accept}"></file-upload>
        <button id="btn-${t.id}" class="btn btn-primary mt-16" disabled>▶ 转换</button>
        <progress-bar id="prog-${t.id}" style="display:none;"></progress-bar>
      </div>`).join('')}</div>`;
}

export async function init() {
  TOOLS.forEach(t => {
    const upload = document.getElementById('upload-' + t.id);
    const btn = document.getElementById('btn-' + t.id);
    const progress = document.getElementById('prog-' + t.id);

    upload.addEventListener('files-changed', e => {
      btn.disabled = e.detail.files.length === 0;
    });

    btn.addEventListener('click', async () => {
      const fd = new FormData();
      fd.append('file', upload.files[0]);
      progress.show(); progress.setProgress(-1, '转换中...'); btn.disabled = true;
      try {
        const res = await api('/api/convert/' + t.id, fd);
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url; a.download = 'output.' + t.ext; a.click();
        URL.revokeObjectURL(url);
        progress.setProgress(100, '完成！'); showToast(`${t.name} 转换成功`);
      } catch (err) {
        progress.setProgress(0, `失败: ${err.message}`);
        showToast(err.message, 'error');
      }
      finally { btn.disabled = false; }
    });
  });
}

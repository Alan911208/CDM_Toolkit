import { api, showToast } from '../app.js';

export default function render() {
  return `<div class="page-section"><h2>🧹 数据清理</h2><p class="section-subtitle">文本标准化、重复高亮、空白行管理、删除线处理</p></div>
    <file-upload accept=".xlsx,.xls"></file-upload>
    <div class="form-row mt-16">
      <div class="form-group"><label for="clean-action">清理操作</label>
        <select id="clean-action">
          <option value="standardise">文本术语标准化</option>
          <option value="duplicates">重复值高亮（相邻行相同值标黄）</option>
          <option value="blank-every-n">每 N 行插入空白行</option>
          <option value="blank-between">组间插入空白行</option>
          <option value="delete-blank">删除空白行</option>
          <option value="strikethrough-delete">删除含删除线的行</option>
          <option value="strikethrough-clear">清除所有删除线格式</option>
        </select>
      </div>
      <div class="form-group" id="clean-param-group"><label for="clean-param">参数</label>
        <input type="text" id="clean-param" value="" placeholder="根据操作类型输入参数">
      </div>
    </div>
    <button id="clean-btn" class="btn btn-primary" disabled>▶ 执行清理</button>
    <progress-bar id="clean-progress" style="display:none;"></progress-bar>`;
}

export async function init() {
  const upload = document.querySelector('file-upload');
  const btn = document.getElementById('clean-btn');
  const progress = document.getElementById('clean-progress');
  const actionSel = document.getElementById('clean-action');
  const paramInput = document.getElementById('clean-param');
  const paramGroup = document.getElementById('clean-param-group');

  const actionParams = {
    'standardise': { param: false, label: '' },
    'duplicates': { param: false, label: '' },
    'blank-every-n': { param: true, label: '每隔几行 (默认5)', value: '5' },
    'blank-between': { param: true, label: '分组列字母 (默认A)', value: 'A' },
    'delete-blank': { param: false, label: '' },
    'strikethrough-delete': { param: false, label: '' },
    'strikethrough-clear': { param: false, label: '' },
  };

  actionSel.addEventListener('change', () => {
    const cfg = actionParams[actionSel.value];
    if (cfg.param) {
      paramGroup.style.display = '';
      paramGroup.querySelector('label').textContent = cfg.label;
      paramInput.value = cfg.value || '';
    } else {
      paramGroup.style.display = 'none';
    }
  });
  actionSel.dispatchEvent(new Event('change'));

  upload.addEventListener('files-changed', e => {
    btn.disabled = e.detail.files.length === 0;
  });

  btn.addEventListener('click', async () => {
    const action = actionSel.value;
    const fd = new FormData();
    fd.append('file', upload.files[0]);

    if (action === 'strikethrough-delete') {
      fd.append('action', 'delete');
      progress.show(); progress.setProgress(-1, '正在删除含删除线的行...'); btn.disabled = true;
      try {
        const res = await api('/api/strikethrough', fd);
        const blob = await res.blob();
        downloadBlob(blob, 'strikethrough_result.xlsx');
        progress.setProgress(100, '完成！'); showToast('删除线行已删除');
      } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
      return;
    }

    if (action === 'strikethrough-clear') {
      fd.append('action', 'clear');
      progress.show(); progress.setProgress(-1, '正在清除删除线...'); btn.disabled = true;
      try {
        const res = await api('/api/strikethrough', fd);
        const blob = await res.blob();
        downloadBlob(blob, 'strikethrough_result.xlsx');
        progress.setProgress(100, '完成！'); showToast('删除线已清除');
      } catch (err) { progress.setProgress(0, `失败: ${err.message}`); showToast(err.message, 'error'); }
      finally { btn.disabled = false; }
      return;
    }

    // For standardise/duplicates/blanks — these are engine functions without file API wrappers yet.
    // They operate on sheet-level. For now, inform the user these are available as Python API.
    progress.hide();
    showToast('此功能请通过 Python engine 直接调用，Web 文件接口即将完善', 'error');
  });
}

function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

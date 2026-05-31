import { showToast } from '../app.js';

export default function render() {
  return `
    <div class="page-section"><h2>📊 SAS 程序运行器</h2><p class="section-subtitle">上传 SAS 数据集 → 选择宏或粘贴代码 → 执行 → 查看日志 → 导出 Excel</p></div>

    <!-- Dataset Upload Section -->
    <div class="sas-dataset-section" id="sas-dataset-section">
      <div class="sas-dataset-header">
        <span style="font-weight:600;">📦 SAS 数据集</span>
        <span id="sas-dataset-count" style="font-size:12px;color:#999;">未上传</span>
      </div>
      <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="sas-dataset-drop" id="sas-dataset-drop">
          <input type="file" hidden multiple accept=".sas7bdat" id="sas-dataset-input">
          <span style="font-size:20px;">📁</span>
          <span style="font-size:12px;">拖拽或点击上传 .sas7bdat</span>
        </div>
        <div id="sas-dataset-list" style="display:flex;gap:8px;flex-wrap:wrap;font-size:12px;"></div>
      </div>
    </div>

    <!-- Editor Section -->
    <div class="sas-layout">
      <aside class="sas-sidebar" id="sas-sidebar">
        <div class="sas-sidebar-header">通用 SAS 宏程序</div>
        <div style="padding:8px 12px;border-bottom:1px solid var(--color-border);">
          <button id="sas-free-mode" class="btn btn-outline" style="width:100%;font-size:12px;padding:6px;">📝 自由代码模式</button>
        </div>
        <div id="sas-macro-list" style="padding:12px;">加载中...</div>
      </aside>
      <div class="sas-main">
        <div class="sas-toolbar">
          <select id="sas-config" class="sas-config-select">
            <option value="zh">中文配置 (zh/SASV9.CFG)</option>
            <option value="u8">UTF-8 配置 (u8/SASV9.CFG)</option>
          </select>
          <button id="sas-run-btn" class="btn btn-primary">▶ 执行 SAS 程序</button>
          <span id="sas-status-badge" style="font-size:12px;color:#999;margin-left:12px;"></span>
        </div>
        <div id="sas-macro-info" style="padding:0 0 12px 0;font-size:13px;color:#666;min-height:20px;"></div>
        <div class="sas-editor-container">
          <textarea id="sas-editor" class="sas-editor" placeholder="粘贴 SAS 代码，或点击左侧宏列表加载模板..." spellcheck="false"></textarea>
        </div>
        <div class="sas-log-header" style="display:flex;justify-content:space-between;align-items:center;margin-top:12px;">
          <h4 style="margin:0;">📋 SAS 日志输出</h4>
          <button id="sas-clear-log" class="btn btn-outline" style="font-size:12px;padding:4px 12px;">清空</button>
        </div>
        <pre id="sas-log" class="sas-log">等待执行...</pre>
        <div id="sas-workspace-hint" style="margin-top:12px;"></div>
      </div>
    </div>
    <pipeline-nav context="sas"></pipeline-nav>
    <style>
      .sas-dataset-section { background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius); padding: 12px 16px; margin-bottom: 16px; }
      .sas-dataset-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }
      .sas-dataset-drop { border: 2px dashed #ccc; border-radius: 8px; padding: 16px 24px; text-align: center; cursor: pointer; transition: all 0.2s; display: flex; align-items: center; gap: 8px; }
      .sas-dataset-drop:hover { border-color: var(--color-primary); background: #F5F7FA; }
      .sas-dataset-item { display: inline-flex; align-items: center; gap: 4px; padding: 4px 10px; background: #E8F0FE; border-radius: 16px; font-size: 12px; }
      .sas-dataset-item button { border: none; background: none; cursor: pointer; color: #C62828; font-size: 14px; padding: 0 2px; }
      .sas-layout { display: flex; gap: 16px; min-height: 60vh; }
      .sas-sidebar { width: 280px; min-width: 280px; background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius); overflow-y: auto; max-height: 70vh; }
      .sas-sidebar-header { padding: 12px 16px; font-weight: 600; font-size: 14px; background: var(--color-primary); color: white; border-radius: var(--radius) var(--radius) 0 0; }
      .sas-main { flex: 1; min-width: 0; }
      .sas-toolbar { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
      .sas-config-select { padding: 8px 12px; border: 1px solid var(--color-border); border-radius: 4px; font-size: 13px; font-family: var(--font); }
      .sas-editor-container { border: 1px solid var(--color-border); border-radius: var(--radius); overflow: hidden; }
      .sas-editor { width: 100%; min-height: 300px; padding: 16px; border: none; font-family: 'Consolas', 'Courier New', monospace; font-size: 13px; line-height: 1.5; resize: vertical; background: #1E1E1E; color: #D4D4D4; tab-size: 2; }
      .sas-editor:focus { outline: none; }
      .sas-log { background: #0C0C0C; color: #CCCCCC; padding: 16px; border-radius: var(--radius); font-family: 'Consolas', 'Courier New', monospace; font-size: 12px; line-height: 1.4; max-height: 300px; overflow-y: auto; white-space: pre-wrap; word-break: break-all; margin: 8px 0 0 0; }
      .sas-cat { padding: 8px 12px 4px; font-size: 11px; font-weight: 700; color: var(--color-primary); text-transform: uppercase; letter-spacing: 0.5px; }
      .sas-macro-item { display: block; width: 100%; padding: 8px 12px; text-align: left; border: none; background: none; cursor: pointer; font-size: 13px; font-family: var(--font); color: var(--color-text); border-left: 3px solid transparent; transition: all 0.15s; }
      .sas-macro-item:hover { background: #E3F2FD; }
      .sas-macro-item.active { background: #E3F2FD; border-left-color: var(--color-primary); font-weight: 600; }
      .sas-macro-item .sas-macro-name { font-weight: 500; }
      .sas-macro-item .sas-macro-desc { font-size: 11px; color: #999; display: block; margin-top: 2px; }
      @media (max-width: 768px) { .sas-layout { flex-direction: column; } .sas-sidebar { width: 100%; max-height: 200px; } }
    </style>`;
}

let currentMacroId = null;
let pollTimer = null;
let libnamePath = '';

// Shared environment template for SAS code
function sasEnvTemplate() {
  return `/* ── 环境设置 ── */
%let _srcPath = %sysfunc(pathname(work));
%let _tgtPath = %sysfunc(pathname(work));
%let _tmpPath = %sysfunc(pathname(work));
${libnamePath ? `libname _cdm "${libnamePath}";\n%let _cdmData = ${libnamePath};` : '/* 未上传 SAS 数据集 — 上传后可在此处自动生成 libname */'}
`;
}

export async function init() {
  const listEl = document.getElementById('sas-macro-list');
  const editor = document.getElementById('sas-editor');
  const runBtn = document.getElementById('sas-run-btn');
  const logEl = document.getElementById('sas-log');
  const infoEl = document.getElementById('sas-macro-info');
  const statusBadge = document.getElementById('sas-status-badge');
  const freeModeBtn = document.getElementById('sas-free-mode');

  // ── Dataset Upload ──
  const datasetInput = document.getElementById('sas-dataset-input');
  const datasetDrop = document.getElementById('sas-dataset-drop');
  const datasetList = document.getElementById('sas-dataset-list');
  const datasetCount = document.getElementById('sas-dataset-count');

  async function loadDatasets() {
    try {
      const res = await fetch('/api/sas/datasets');
      if (!res.ok) return;
      const data = await res.json();
      libnamePath = data.libname_path || '';
      const ds = data.datasets || [];
      datasetCount.textContent = ds.length ? `${ds.length} 个数据集` : '未上传';
      datasetList.innerHTML = ds.map(d =>
        `<span class="sas-dataset-item">📄 ${d.name} <span style="color:#999;">(${d.size_kb} KB)</span></span>`
      ).join('');
    } catch (_) {}
  }

  datasetDrop.addEventListener('click', () => datasetInput.click());
  datasetDrop.addEventListener('dragover', e => { e.preventDefault(); datasetDrop.style.borderColor = '#4472C4'; });
  datasetDrop.addEventListener('dragleave', () => { datasetDrop.style.borderColor = '#ccc'; });
  datasetDrop.addEventListener('drop', async e => {
    e.preventDefault(); datasetDrop.style.borderColor = '#ccc';
    if (e.dataTransfer.files.length) await uploadDatasets(e.dataTransfer.files);
  });
  datasetInput.addEventListener('change', async () => {
    if (datasetInput.files.length) await uploadDatasets(datasetInput.files);
    datasetInput.value = '';
  });

  async function uploadDatasets(files) {
    const fd = new FormData();
    for (const f of files) fd.append('files', f);
    try {
      const res = await fetch('/api/sas/upload-datasets', { method: 'POST', body: fd });
      const data = await res.json();
      libnamePath = data.libname_path || '';
      showToast(`已上传 ${data.uploaded.length} 个数据集`);
      await loadDatasets();
    } catch (err) {
      showToast('上传失败: ' + err.message, 'error');
    }
  }

  loadDatasets();

  // ── Free Code Mode ──
  freeModeBtn.addEventListener('click', () => {
    currentMacroId = null;
    listEl.querySelectorAll('.sas-macro-item').forEach(b => b.classList.remove('active'));
    infoEl.innerHTML = '<strong>📝 自由代码模式</strong> — 粘贴你的 SAS 业务逻辑程序，点击执行';
    editor.value = `/* ── 自由代码模式 ── 粘贴任意 SAS 程序 ── */
${sasEnvTemplate()}
`;
    logEl.textContent = '就绪 — 粘贴代码后点击"执行 SAS 程序"';
    editor.focus();
  });

  // ── Macro List ──
  let macros = [];
  try {
    const res = await fetch('/api/sas-list');
    const data = await res.json();
    macros = data.macros || [];
  } catch (err) {
    listEl.innerHTML = '<p style="padding:12px;color:#C62828;">加载宏列表失败</p>';
    return;
  }

  const cats = {};
  macros.forEach(m => { if (!cats[m.cat]) cats[m.cat] = []; cats[m.cat].push(m); });
  let html = '';
  for (const [cat, items] of Object.entries(cats)) {
    html += `<div class="sas-cat">${cat}</div>`;
    items.forEach(m => {
      html += `<button class="sas-macro-item" data-id="${m.id}" title="${m.desc}">
        <span class="sas-macro-name">${m.name}</span>
        <span class="sas-macro-desc">${m.desc}</span>
      </button>`;
    });
  }
  listEl.innerHTML = html;

  listEl.addEventListener('click', async (e) => {
    const item = e.target.closest('.sas-macro-item');
    if (!item) return;
    const macroId = item.dataset.id;
    if (currentMacroId === macroId) return;
    currentMacroId = macroId;
    listEl.querySelectorAll('.sas-macro-item').forEach(b => b.classList.remove('active'));
    item.classList.add('active');
    editor.value = '加载中...';
    infoEl.textContent = '';

    try {
      const res = await fetch(`/api/sas-load/${macroId}`);
      const data = await res.json();
      if (data.error) { editor.value = `/* ${data.error} */`; return; }
      const macro = data.macro;
      infoEl.innerHTML = `<strong>${macro.name}</strong> — ${macro.desc}<br><span style="font-size:11px;color:#999;">参数: %macro ${macro.name}(${macro.params || '无'})</span>`;

      let code = `/* ═══════════════════════════════════════════════════════════════
   CDM Toolkit SAS Runner
   宏: ${macro.name} (${macro.id})
   说明: ${macro.desc}
   参数: ${macro.params || '无'}
   配置: ${document.getElementById('sas-config').value === 'zh' ? '中文 (zh)' : 'UTF-8 (u8)'}
   ═══════════════════════════════════════════════════════════════ */

${sasEnvTemplate()}`;

      if (data.source) code += `/* ── 宏定义 ── */\n${data.source}\n\n`;
      if (data.test) {
        code += `/* ── 测试调用（请根据实际数据修改参数） ── */\n${data.test}\n`;
      } else if (data.source) {
        code += `/* ── 调用宏（请修改参数） ── */\n%${macro.name}(${macro.params || ''})\n`;
      }

      editor.value = code;
      logEl.textContent = '准备就绪 — 修改参数后点击"执行 SAS 程序"';
    } catch (err) {
      editor.value = `/* 加载失败: ${err.message} */`;
    }
  });

  // ── Run ──
  runBtn.addEventListener('click', async () => {
    const code = editor.value.trim();
    if (!code) { showToast('请先加载或输入 SAS 代码', 'error'); return; }

    const config = document.getElementById('sas-config').value;
    runBtn.textContent = '⏳ 执行中...';
    logEl.textContent = '提交 SAS 任务...';
    statusBadge.textContent = '';

    try {
      const res = await fetch('/api/sas-run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code, config }),
      });
      const data = await res.json();
      if (data.error) {
        logEl.textContent = `❌ ${data.error}`;
        runBtn.textContent = '▶ 执行 SAS 程序';
        return;
      }

      const jobId = data.job_id;
      statusBadge.textContent = `任务: ${jobId}`;
      logEl.textContent = 'SAS 执行中...\n';

      if (pollTimer) clearInterval(pollTimer);
      pollTimer = setInterval(async () => {
        try {
          const sr = await fetch(`/api/sas-status/${jobId}`);
          const sj = await sr.json();
          logEl.textContent = sj.log || '(无输出)';
          logEl.scrollTop = logEl.scrollHeight;

          if (sj.status === 'completed') {
            clearInterval(pollTimer); pollTimer = null;
            const elapsed = sj.elapsed || 0;
            statusBadge.textContent = `✅ 完成 (${elapsed}秒)`;
            runBtn.textContent = '▶ 执行 SAS 程序';
            logEl.textContent += `\n\n/* ══ SAS 执行完成 (${elapsed}秒) ══ */`;
            const hint = document.getElementById('sas-workspace-hint');
            if (hint) hint.innerHTML = '<div style="background:#E8F5E9;padding:12px;border-radius:8px;font-size:13px;">💡 SAS 执行完毕。如果生成了 Excel 文件，请前往 <a href="/" data-route="/" style="color:var(--color-primary);">仪表盘</a> 上传到工作区继续处理。</div>';
          } else if (sj.status === 'error') {
            clearInterval(pollTimer); pollTimer = null;
            statusBadge.textContent = `❌ 错误`;
            runBtn.textContent = '▶ 执行 SAS 程序';
            logEl.textContent += '\n\n/* ══ SAS 执行出错 ══ */';
          } else if (sj.status === 'timeout') {
            clearInterval(pollTimer); pollTimer = null;
            statusBadge.textContent = `⏰ 超时`;
            runBtn.textContent = '▶ 执行 SAS 程序';
          }
        } catch (_) {}
      }, 1000);
    } catch (err) {
      logEl.textContent = `❌ 请求失败: ${err.message}`;
      runBtn.textContent = '▶ 执行 SAS 程序';
    }
  });

  document.getElementById('sas-clear-log').addEventListener('click', () => {
    logEl.textContent = '';
  });
}

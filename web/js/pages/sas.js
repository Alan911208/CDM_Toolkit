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
        <!-- Output Panel -->
        <div class="sas-output" id="sas-output" style="margin-top:12px;display:none;">
          <!-- Summary Bar -->
          <div class="sas-summary" id="sas-summary"></div>
          <!-- Output Files -->
          <div id="sas-output-files" style="margin-top:12px;"></div>
          <!-- Log Sections -->
          <div class="sas-sections" id="sas-sections" style="margin-top:12px;"></div>
          <!-- Raw Log Toggle -->
          <div style="margin-top:8px;">
            <button id="sas-toggle-raw" class="btn btn-outline" style="font-size:11px;padding:4px 10px;">📄 查看原始日志</button>
          </div>
          <pre id="sas-raw-log" class="sas-log" style="display:none;margin-top:8px;"></pre>
        </div>
        <div id="sas-waiting" style="margin-top:12px;padding:24px;text-align:center;color:#999;background:var(--color-surface);border-radius:var(--radius);">等待执行...</div>
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
      .sas-log { background: #0C0C0C; color: #CCCCCC; padding: 16px; border-radius: var(--radius); font-family: 'Consolas', 'Courier New', monospace; font-size: 12px; line-height: 1.4; max-height: 400px; overflow-y: auto; white-space: pre-wrap; word-break: break-all; margin: 8px 0 0 0; }
      .sas-summary { display: flex; gap: 12px; flex-wrap: wrap; }
      .sas-stat { padding: 8px 16px; border-radius: 8px; font-size: 13px; font-weight: 600; display: flex; align-items: center; gap: 6px; }
      .sas-stat.ok { background: #E8F5E9; color: #2E7D32; }
      .sas-stat.warn { background: #FFF3E0; color: #E65100; }
      .sas-stat.err { background: #FFEBEE; color: #C62828; }
      .sas-stat.info { background: #E3F2FD; color: #1565C0; }
      .sas-output-files { display: flex; flex-wrap: wrap; gap: 8px; }
      .sas-output-file { display: inline-flex; align-items: center; gap: 6px; padding: 6px 12px; background: #E8F5E9; border-radius: 8px; font-size: 12px; font-family: monospace; }
      .sas-section { margin-top: 8px; border: 1px solid var(--color-border); border-radius: 8px; overflow: hidden; }
      .sas-section-header { padding: 8px 14px; font-size: 12px; font-weight: 600; cursor: pointer; display: flex; justify-content: space-between; align-items: center; user-select: none; }
      .sas-section-header.err { background: #FFEBEE; color: #C62828; }
      .sas-section-header.warn { background: #FFF3E0; color: #E65100; }
      .sas-section-header.note { background: #F5F5F5; color: #666; }
      .sas-section-body { padding: 8px 14px; font-family: 'Consolas', monospace; font-size: 11px; line-height: 1.4; max-height: 200px; overflow-y: auto; white-space: pre-wrap; word-break: break-all; background: #FAFAFA; }
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
${libnamePath ? `libname raw "${libnamePath}";\n%let rawData = ${libnamePath};` : '/* 未上传 SAS 数据集 — 上传后可在此处自动生成 libname */'}
`;
}

export async function init() {
  const listEl = document.getElementById('sas-macro-list');
  const editor = document.getElementById('sas-editor');
  const runBtn = document.getElementById('sas-run-btn');
  const outputPanel = document.getElementById('sas-output');
  const waitingEl = document.getElementById('sas-waiting');
  const summaryEl = document.getElementById('sas-summary');
  const outputFilesEl = document.getElementById('sas-output-files');
  const sectionsEl = document.getElementById('sas-sections');
  const rawLogEl = document.getElementById('sas-raw-log');
  const toggleRawBtn = document.getElementById('sas-toggle-raw');
  const infoEl = document.getElementById('sas-macro-info');
  const statusBadge = document.getElementById('sas-status-badge');
  const freeModeBtn = document.getElementById('sas-free-mode');

  // ── Output Renderer ──
  function renderOutput(parsed, status, elapsed) {
    outputPanel.style.display = '';
    waitingEl.style.display = 'none';

    // Summary stats
    const hasErr = parsed && parsed.has_error;
    const hasWarn = parsed && parsed.has_warning;
    const nErr = parsed ? parsed.errors.length : 0;
    const nWarn = parsed ? parsed.warnings.length : 0;
    const nNote = parsed ? parsed.notes.length : 0;
    const nFiles = parsed ? parsed.output_files.length : 0;
    const realTime = parsed ? parsed.real_time : '';

    let statusClass = 'ok', statusIcon = '✅', statusText = '执行成功';
    if (status === 'error' || hasErr) { statusClass = 'err'; statusIcon = '❌'; statusText = '执行出错'; }
    else if (hasWarn) { statusClass = 'warn'; statusIcon = '⚠️'; statusText = '执行完成（有警告）'; }

    summaryEl.innerHTML = `
      <div class="sas-stat ${statusClass}">${statusIcon} ${statusText}</div>
      <div class="sas-stat info">⏱ ${elapsed || '?'} 秒${realTime ? ' · ' + realTime : ''}</div>
      ${nErr > 0 ? `<div class="sas-stat err">❌ ${nErr} 错误</div>` : ''}
      ${nWarn > 0 ? `<div class="sas-stat warn">⚠️ ${nWarn} 警告</div>` : ''}
      ${nNote > 0 ? `<div class="sas-stat info">📝 ${nNote} 条记录</div>` : ''}
      ${nFiles > 0 ? `<div class="sas-stat ok">📄 ${nFiles} 个输出文件</div>` : ''}`;

    // Output files
    if (nFiles > 0) {
      outputFilesEl.innerHTML = '<div style="font-size:12px;font-weight:600;margin-bottom:4px;">📂 生成的文件</div><div class="sas-output-files">' +
        parsed.output_files.map(f => `<span class="sas-output-file">📄 ${f.name}<span style="color:#999;font-size:10px;">${f.path}</span></span>`).join('') +
        '</div>';
    } else {
      outputFilesEl.innerHTML = '';
    }

    // Log sections (errors first, then warnings, then notes)
    let secHtml = '';
    if (nErr > 0) {
      secHtml += `<div class="sas-section"><div class="sas-section-header err" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'':'none'">❌ 错误 (${nErr}) <span>▶</span></div><div class="sas-section-body">${parsed.errors.map(e => h(e)).join('\n\n')}</div></div>`;
    }
    if (nWarn > 0) {
      secHtml += `<div class="sas-section"><div class="sas-section-header warn" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'':'none'">⚠️ 警告 (${nWarn}) <span>▶</span></div><div class="sas-section-body">${parsed.warnings.map(e => h(e)).join('\n\n')}</div></div>`;
    }
    if (nNote > 0 && nNote <= 10) {
      secHtml += `<div class="sas-section"><div class="sas-section-header note" onclick="this.nextElementSibling.style.display=this.nextElementSibling.style.display==='none'?'':'none'">📝 关键记录 (${nNote}) <span>▶</span></div><div class="sas-section-body">${parsed.notes.map(e => h(e)).join('\n\n')}</div></div>`;
    }
    sectionsEl.innerHTML = secHtml;
  }

  function h(str) { return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;'); }

  toggleRawBtn.addEventListener('click', () => {
    const show = rawLogEl.style.display === 'none';
    rawLogEl.style.display = show ? '' : 'none';
    toggleRawBtn.textContent = show ? '📄 隐藏原始日志' : '📄 查看原始日志';
  });

  // ── Init Session (client-managed UUID) ──
  if (!localStorage.getItem('cdm_session')) {
    localStorage.setItem('cdm_session', crypto.randomUUID());
  }

  // ── Dataset Upload ──
  const datasetInput = document.getElementById('sas-dataset-input');
  const datasetDrop = document.getElementById('sas-dataset-drop');
  const datasetList = document.getElementById('sas-dataset-list');
  const datasetCount = document.getElementById('sas-dataset-count');

  function _sasSessionHeaders() {
    const sid = localStorage.getItem('cdm_session') || '';
    return sid ? { 'X-CDM-Session': sid } : {};
  }

  async function loadDatasets() {
    try {
      const res = await fetch('/api/sas/datasets', { headers: _sasSessionHeaders() });
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
      const res = await fetch('/api/sas/upload-datasets', {
        method: 'POST',
        headers: _sasSessionHeaders(),
        body: fd,
      });
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
    waitingEl.textContent = '就绪 — 粘贴代码后点击"执行 SAS 程序"';
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
      const isStd = macro.type === 'standalone';
      infoEl.innerHTML = `<strong>${macro.name}</strong> — ${macro.desc}<br><span style="font-size:11px;color:#999;">${isStd ? '独立 SAS 程序（直接运行）' : '参数: %' + macro.name + '(' + (macro.params || '无') + ')'}</span>`;

      const isStandalone = macro.type === 'standalone';
      let code = `/* ═══════════════════════════════════════════════════════════════
   CDM Toolkit SAS Runner
   ${isStandalone ? '程序' : '宏'}: ${macro.name} (${macro.id})
   说明: ${macro.desc}
   ${isStandalone ? '' : '参数: ' + (macro.params || '无')}
   配置: ${document.getElementById('sas-config').value === 'zh' ? '中文 (zh)' : 'UTF-8 (u8)'}
   ═══════════════════════════════════════════════════════════════ */

${sasEnvTemplate()}`;

      if (data.source) {
        if (isStandalone) {
          code += `/* ── 独立 SAS 程序（可直接运行） ── */\n${data.source}\n`;
        } else {
          code += `/* ── 宏定义 ── */\n${data.source}\n\n`;
        }
      }
      if (!isStandalone) {
        if (data.test) {
          code += `/* ── 测试调用（请根据实际数据修改参数） ── */\n${data.test}\n`;
        } else if (data.source) {
          code += `/* ── 调用宏（请修改参数） ── */\n%${macro.name}(${macro.params || ''})\n`;
        }
      }

      editor.value = code;
      waitingEl.textContent = '准备就绪 — 修改参数后点击"执行 SAS 程序"';
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
    waitingEl.textContent = '提交 SAS 任务...';
    statusBadge.textContent = '';

    try {
      const res = await fetch('/api/sas-run', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ code, config }),
      });
      const data = await res.json();
      if (data.error) {
        rawLogEl.textContent = `❌ ${data.error}`;
        runBtn.textContent = '▶ 执行 SAS 程序';
        return;
      }

      const jobId = data.job_id;
      statusBadge.textContent = `任务: ${jobId}`;
      waitingEl.style.display = '';
      waitingEl.textContent = 'SAS 执行中...';
      outputPanel.style.display = 'none';

      if (pollTimer) clearInterval(pollTimer);
      pollTimer = setInterval(async () => {
        try {
          const sr = await fetch(`/api/sas-status/${jobId}`);
          const sj = await sr.json();
          rawLogEl.textContent = sj.log || '(无输出)';

          // Show structured output when completed
          if (sj.status === 'completed' || sj.status === 'error') {
            clearInterval(pollTimer); pollTimer = null;
            const elapsed = sj.elapsed || 0;
            statusBadge.textContent = sj.status === 'completed' ? `✅ 完成 (${elapsed}秒)` : `❌ 错误`;
            runBtn.textContent = '▶ 执行 SAS 程序';
            if (sj.parsed) {
              renderOutput(sj.parsed, sj.status, elapsed);
            } else {
              waitingEl.style.display = 'none';
              outputPanel.style.display = '';
              summaryEl.innerHTML = `<div class="sas-stat ${sj.status === 'completed' ? 'ok' : 'err'}">${sj.status === 'completed' ? '✅' : '❌'} 执行${sj.status === 'completed' ? '完成' : '出错'} · ${elapsed}秒</div>`;
            }
          } else if (sj.status === 'timeout') {
            clearInterval(pollTimer); pollTimer = null;
            statusBadge.textContent = `⏰ 超时`;
            runBtn.textContent = '▶ 执行 SAS 程序';
            waitingEl.style.display = 'none';
            outputPanel.style.display = '';
            summaryEl.innerHTML = '<div class="sas-stat err">⏰ SAS 执行超时（300 秒）</div>';
          }
        } catch (_) {}
      }, 1000);
    } catch (err) {
      rawLogEl.textContent = `❌ 请求失败: ${err.message}`;
      runBtn.textContent = '▶ 执行 SAS 程序';
    }
  });

  const clearBtn = document.getElementById('sas-clear-log');
  if (clearBtn) {
    clearBtn.addEventListener('click', () => {
      outputPanel.style.display = 'none';
      waitingEl.style.display = '';
      waitingEl.textContent = '日志已清空';
      rawLogEl.textContent = '';
    });
  }
}

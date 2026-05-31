// Workspace Panel — file management + preview for pipeline flow
export class WorkspacePanel extends HTMLElement {
  constructor() {
    super();
    this.sessionId = localStorage.getItem('cdm_session') || '';
    this.attachShadow({ mode: 'open' });
  }

  connectedCallback() {
    this.render();
    this.refresh();
  }

  render() {
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; background: var(--color-surface); border: 1px solid var(--color-border); border-radius: var(--radius); margin-bottom: 16px; }
        .ws-header { padding: 10px 16px; background: #F0F4FF; border-bottom: 1px solid var(--color-border); border-radius: var(--radius) var(--radius) 0 0; display: flex; justify-content: space-between; align-items: center; font-size: 13px; font-weight: 600; }
        .ws-body { padding: 16px; }
        .ws-empty { text-align: center; padding: 24px; color: #999; }
        .ws-file-info { display: flex; gap: 16px; font-size: 12px; color: #666; flex-wrap: wrap; }
        .ws-file-info span { background: #F5F7FA; padding: 4px 8px; border-radius: 4px; }
        .ws-upload-area { border: 2px dashed #ccc; border-radius: 8px; padding: 20px; text-align: center; cursor: pointer; transition: all 0.2s; margin-top: 8px; }
        .ws-upload-area:hover { border-color: var(--color-primary); background: #F5F7FA; }
        .ws-preview { margin-top: 12px; }
        .ws-preview table { width: 100%; border-collapse: collapse; font-size: 12px; }
        .ws-preview th { background: #4472C4; color: white; padding: 6px 8px; text-align: left; position: sticky; top: 0; white-space: nowrap; }
        .ws-preview td { padding: 4px 8px; border-bottom: 1px solid #E0E4E8; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .ws-preview-wrapper { max-height: 300px; overflow: auto; border: 1px solid var(--color-border); border-radius: 4px; }
        .ws-actions { display: flex; gap: 8px; margin-top: 12px; flex-wrap: wrap; }
        .ws-btn { padding: 6px 14px; border: 1px solid var(--color-border); border-radius: 4px; background: white; cursor: pointer; font-size: 12px; font-family: var(--font); transition: all 0.15s; }
        .ws-btn:hover { background: #E3F2FD; border-color: var(--color-primary); }
        .ws-btn.primary { background: var(--color-primary); color: white; border-color: var(--color-primary); }
        .ws-history { margin-top: 12px; font-size: 11px; color: #999; }
        .ws-history-item { padding: 2px 0; border-bottom: 1px dotted #eee; }
        .ws-history-item:last-child { border-bottom: none; }
        .ws-tabs { display: flex; gap: 0; border-bottom: 1px solid var(--color-border); }
        .ws-tab { padding: 6px 16px; cursor: pointer; font-size: 12px; border: none; background: none; color: #666; }
        .ws-tab.active { color: var(--color-primary); border-bottom: 2px solid var(--color-primary); font-weight: 600; }
      </style>
      <div class="ws-header">
        <span>📂 工作区</span>
        <span style="font-weight:400;font-size:11px;color:#999;" id="ws-status-text">加载中...</span>
      </div>
      <div class="ws-body" id="ws-body">
        <div class="ws-empty">拖拽或点击上传 Excel 文件开始</div>
        <div class="ws-upload-area" id="ws-upload-area">
          <input type="file" hidden accept=".xlsx,.xls" id="ws-file-input">
          <div style="font-size:28px;">📁</div>
          <div style="font-size:13px;">点击选择文件或拖拽到此处</div>
        </div>
      </div>`;

    // Upload handlers
    const input = this.shadowRoot.getElementById('ws-file-input');
    const area = this.shadowRoot.getElementById('ws-upload-area');
    area.addEventListener('click', () => input.click());
    area.addEventListener('dragover', e => { e.preventDefault(); area.style.borderColor = '#4472C4'; });
    area.addEventListener('dragleave', () => { area.style.borderColor = '#ccc'; });
    area.addEventListener('drop', e => {
      e.preventDefault();
      area.style.borderColor = '#ccc';
      if (e.dataTransfer.files.length) this.upload(e.dataTransfer.files[0]);
    });
    input.addEventListener('change', () => {
      if (input.files.length) this.upload(input.files[0]);
      input.value = '';
    });
  }

  async refresh() {
    if (!this.sessionId) return;
    try {
      const res = await fetch('/api/workspace/info', { headers: { 'X-CDM-Session': this.sessionId } });
      if (!res.ok) return;
      const info = await res.json();
      this._renderState(info);
    } catch (_) {}
  }

  async upload(file) {
    this.shadowRoot.getElementById('ws-status-text').textContent = '上传中...';
    const fd = new FormData();
    fd.append('file', file);
    try {
      const res = await fetch('/api/workspace/upload', {
        method: 'POST',
        headers: { 'X-CDM-Session': this.sessionId },
        body: fd,
      });
      const data = await res.json();
      if (data.session_id) {
        this.sessionId = data.session_id;
        localStorage.setItem('cdm_session', data.session_id);
      }
      this._renderState(data);
      this.dispatchEvent(new CustomEvent('workspace-changed', { detail: data, bubbles: true }));
    } catch (err) {
      this.shadowRoot.getElementById('ws-status-text').textContent = '上传失败';
    }
  }

  _renderState(data) {
    const body = this.shadowRoot.getElementById('ws-body');
    const statusEl = this.shadowRoot.getElementById('ws-status-text');
    const area = this.shadowRoot.getElementById('ws-upload-area');

    if (!data || (!data.has_file && !data.preview)) {
      body.innerHTML = `<div class="ws-empty">拖拽或点击上传 Excel 文件开始</div>
        <div class="ws-upload-area" id="ws-upload-area2">
          <input type="file" hidden accept=".xlsx,.xls" id="ws-file-input2">
          <div style="font-size:28px;">📁</div>
          <div style="font-size:13px;">点击选择文件或拖拽到此处</div>
        </div>`;
      statusEl.textContent = '无文件';
      return;
    }

    // Get preview info
    const fname = data.file_name || 'workspace_output.xlsx';
    let infoHtml = '';
    if (data.has_file !== undefined) {
      // From /api/workspace/info
      infoHtml = `<div class="ws-file-info">
        <span>📄 ${fname}</span>`;
      if (data.sheets) infoHtml += `<span>📊 ${data.sheets} 个工作表</span>`;
      if (data.rows) infoHtml += `<span>📝 ${data.rows} 行</span>`;
      if (data.cols) infoHtml += `<span>📐 ${data.cols} 列</span>`;
      infoHtml += `</div>`;
      // Fetch preview
      this._fetchPreview();
    }

    // Preview
    let previewHtml = '<div class="ws-preview" id="ws-preview-content"><p style="color:#999;font-size:12px;">加载预览中...</p></div>';

    // Actions
    let actionsHtml = `
      <div class="ws-actions">
        <button class="ws-btn primary" id="ws-download-btn">📥 下载</button>
        <button class="ws-btn" id="ws-refresh-btn">🔄 刷新预览</button>
      </div>`;

    // History
    let historyHtml = '';
    if (data.history && data.history.length) {
      historyHtml = '<div class="ws-history"><strong>操作历史</strong>' +
        data.history.map(h => `<div class="ws-history-item">${h.time} — ${h.action}</div>`).join('') +
        '</div>';
    }

    body.innerHTML = infoHtml + previewHtml + actionsHtml + historyHtml;

    // Update upload area
    area.style.display = 'none';
    statusEl.textContent = fname;

    // Wire buttons
    body.querySelector('#ws-download-btn')?.addEventListener('click', () => {
      const a = document.createElement('a');
      a.href = `/api/workspace/download?sid=${this.sessionId}`;
      a.download = fname; a.click();
    });
    body.querySelector('#ws-refresh-btn')?.addEventListener('click', () => this._fetchPreview());

    if (data.preview) this._showPreview(data.preview);

    // Re-bind upload
    const area2 = body.querySelector('#ws-upload-area2');
    const input2 = body.querySelector('#ws-file-input2');
    if (area2 && input2) {
      area2.addEventListener('click', () => input2.click());
      area2.addEventListener('dragover', e => { e.preventDefault(); area2.style.borderColor = '#4472C4'; });
      area2.addEventListener('dragleave', () => { area2.style.borderColor = '#ccc'; });
      area2.addEventListener('drop', e => {
        e.preventDefault(); area2.style.borderColor = '#ccc';
        if (e.dataTransfer.files.length) this.upload(e.dataTransfer.files[0]);
      });
      input2.addEventListener('change', () => {
        if (input2.files.length) this.upload(input2.files[0]);
        input2.value = '';
      });
    }
  }

  async _fetchPreview() {
    try {
      const res = await fetch('/api/workspace/preview?rows=50', { headers: { 'X-CDM-Session': this.sessionId } });
      const data = await res.json();
      if (data && data.sheets) this._showPreview(data);
    } catch (_) {}
  }

  _showPreview(data) {
    const el = this.shadowRoot.getElementById('ws-preview-content');
    if (!el) return;
    if (!data.sheets || !data.sheets.length) {
      el.innerHTML = '<p style="color:#999;">无数据</p>';
      return;
    }
    let html = '';
    if (data.sheets.length > 1) {
      html += '<div class="ws-tabs">' + data.sheets.map((s, i) =>
        `<button class="ws-tab ${i === 0 ? 'active' : ''}" data-sheet="${i}">${s.name}</button>`
      ).join('') + '</div>';
    }
    const s = data.sheets[0];
    html += `<div class="ws-preview-wrapper"><table>
      <thead><tr>${(s.headers || []).map(h => `<th>${h}</th>`).join('')}</tr></thead>
      <tbody>${(s.rows || []).map(row => `<tr>${row.map(v => `<td>${v}</td>`).join('')}</tr>`).join('')}</tbody>
    </table></div>`;
    html += `<div style="font-size:11px;color:#999;margin-top:4px;">显示前 ${(s.rows || []).length} 行</div>`;
    el.innerHTML = html;

    // Tab switching for multi-sheet
    if (data.sheets.length > 1) {
      el.querySelectorAll('.ws-tab').forEach(tab => {
        tab.addEventListener('click', () => {
          el.querySelectorAll('.ws-tab').forEach(t => t.classList.remove('active'));
          tab.classList.add('active');
          const idx = parseInt(tab.dataset.sheet);
          this._renderSheetTable(el, data.sheets[idx]);
        });
      });
    }
  }

  _renderSheetTable(el, sheet) {
    const wrapper = el.querySelector('.ws-preview-wrapper');
    if (wrapper) {
      wrapper.innerHTML = `<table>
        <thead><tr>${(sheet.headers || []).map(h => `<th>${h}</th>`).join('')}</tr></thead>
        <tbody>${(sheet.rows || []).map(row => `<tr>${row.map(v => `<td>${v}</td>`).join('')}</tr>`).join('')}</tbody>
      </table>`;
    }
  }
}

customElements.define('workspace-panel', WorkspacePanel);

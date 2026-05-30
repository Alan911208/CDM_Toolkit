export class DataTable extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }
  set data(value) { this._data = value; this.render(); }
  get data() { return this._data || []; }
  connectedCallback() { this._columns = JSON.parse(this.getAttribute('columns') || '[]'); this._sortCol = null; this._sortDir = 1; this.render(); }

  sort(colKey) {
    if (this._sortCol === colKey) { this._sortDir *= -1; } else { this._sortCol = colKey; this._sortDir = 1; }
    this._data.sort((a, b) => { const va = a[colKey] ?? '', vb = b[colKey] ?? ''; return String(va).localeCompare(String(vb)) * this._sortDir; });
    this.renderBody();
  }

  render() {
    const cols = this._columns;
    this.shadowRoot.innerHTML = `<style>:host { display: block; overflow-x: auto; } table { width: 100%; border-collapse: collapse; font-size: 13px; } th { background: #4472C4; color: white; padding: 8px 12px; text-align: left; cursor: pointer; user-select: none; white-space: nowrap; } th:hover { background: #2E5BA0; } td { padding: 8px 12px; border-bottom: 1px solid #E0E4E8; max-width: 300px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; } tr:hover { background: #F5F7FA; } .empty { text-align: center; padding: 24px; color: #999; } .badge { display: inline-block; padding: 2px 6px; border-radius: 4px; font-size: 11px; } .badge-modified { background: #E3F2FD; color: #1565C0; } .badge-added { background: #E8F5E9; color: #2E7D32; } .badge-deleted { background: #FFEBEE; color: #C62828; }</style>
      <table><thead><tr>${cols.map(c => `<th data-col="${c.key}">${c.label}${this._sortCol === c.key ? (this._sortDir === 1 ? ' ▲' : ' ▼') : ''}</th>`).join('')}</tr></thead><tbody></tbody></table>`;
    this.renderBody();
    this.shadowRoot.querySelectorAll('th').forEach(th => { th.addEventListener('click', () => this.sort(th.dataset.col)); });
  }

  renderBody() {
    const tbody = this.shadowRoot.querySelector('tbody');
    if (!this._data || this._data.length === 0) { tbody.innerHTML = '<tr><td class="empty" colspan="99">无数据</td></tr>'; return; }
    tbody.innerHTML = this._data.map(row => `<tr>${this._columns.map(c => `<td>${this.formatCell(row[c.key], c.key)}</td>`).join('')}</tr>`).join('');
  }

  formatCell(value, key) {
    if (key === 'type' || key === 'change_type') { const cls = { '修改': 'modified', 'modified': 'modified', '新增': 'added', 'added': 'added', '删除': 'deleted', 'deleted': 'deleted' }; return `<span class="badge badge-${cls[value] || 'modified'}">${value || ''}</span>`; }
    if (value === null || value === undefined) return '';
    const str = String(value);
    return str.length > 100 ? str.slice(0, 100) + '...' : str;
  }
}
customElements.define('data-table', DataTable);

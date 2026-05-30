export class FolderTree extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }
  set data(value) { this._data = value; this.render(); }
  render() {
    this.shadowRoot.innerHTML = `<style>:host { display: block; font-size: 13px; font-family: monospace; } .tree { padding: 4px 0; } .item { padding: 2px 0; white-space: nowrap; } .folder { color: #F57F17; cursor: pointer; } .file { color: #1565C0; cursor: pointer; } .file:hover, .folder:hover { text-decoration: underline; } .size { color: #999; font-size: 11px; margin-left: 8px; }</style><div class="tree"></div>`;
    if (this._data) this.renderTree();
  }
  renderTree() {
    const container = this.shadowRoot.querySelector('.tree');
    container.innerHTML = this._data.items?.map(item => `<div class="item ${item.is_dir ? 'folder' : 'file'}">${item.is_dir ? '📁' : '📄'} ${item.name}${!item.is_dir ? `<span class="size">${this.formatSize(item.size)}</span>` : ''}</div>`).join('') || '<div class="item">空文件夹</div>';
  }
  formatSize(bytes) { if (!bytes) return ''; if (bytes < 1024) return `${bytes} B`; if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`; return `${(bytes / (1024 * 1024)).toFixed(1)} MB`; }
}
customElements.define('folder-tree', FolderTree);

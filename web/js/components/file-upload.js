export class FileUpload extends HTMLElement {
  constructor() {
    super();
    this.files = [];
    this.attachShadow({ mode: 'open' });
  }

  connectedCallback() {
    this.accept = this.getAttribute('accept') || '.xlsx,.xls';
    this.multiple = this.hasAttribute('multiple');
    this.render();
  }

  render() {
    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; }
        .drop-zone { border: 2px dashed #ccc; border-radius: 8px; padding: 32px; text-align: center; cursor: pointer; transition: all 0.2s; }
        .drop-zone.dragover { border-color: #4472C4; background: #E3F2FD; }
        .drop-icon { font-size: 36px; }
        .file-list { margin-top: 12px; }
        .file-item { display: flex; justify-content: space-between; align-items: center; padding: 6px 12px; background: #f5f5f5; border-radius: 4px; margin-bottom: 4px; font-size: 13px; }
        .file-item button { border: none; background: none; cursor: pointer; color: #C62828; font-size: 16px; }
      </style>
      <div class="drop-zone"><div class="drop-icon">📂</div><p>拖拽文件到此处，或点击选择</p><p style="font-size:12px;color:#999;">支持: ${this.accept}</p><input type="file" hidden ${this.multiple ? 'multiple' : ''} accept="${this.accept}"></div>
      <div class="file-list"></div>`;

    const zone = this.shadowRoot.querySelector('.drop-zone');
    const input = this.shadowRoot.querySelector('input');
    zone.addEventListener('click', () => input.click());
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', e => { e.preventDefault(); zone.classList.remove('dragover'); this.addFiles(Array.from(e.dataTransfer.files)); });
    input.addEventListener('change', () => { this.addFiles(Array.from(input.files)); input.value = ''; });
  }

  addFiles(newFiles) {
    this.files = this.multiple ? [...this.files, ...newFiles] : newFiles;
    this.updateFileList();
    this.dispatchEvent(new CustomEvent('files-changed', { detail: { files: this.files }, bubbles: true }));
  }

  updateFileList() {
    const list = this.shadowRoot.querySelector('.file-list');
    list.innerHTML = this.files.map((f, i) => `<div class="file-item"><span>📄 ${f.name} (${(f.size / 1024).toFixed(1)} KB)</span><button data-idx="${i}">✕</button></div>`).join('');
    list.querySelectorAll('button').forEach(btn => { btn.addEventListener('click', (e) => { e.stopPropagation(); this.files.splice(parseInt(btn.dataset.idx), 1); this.updateFileList(); this.dispatchEvent(new CustomEvent('files-changed', { detail: { files: this.files }, bubbles: true })); }); });
  }

  getFormData() { const fd = new FormData(); this.files.forEach(f => fd.append('files', f)); return fd; }
}
customElements.define('file-upload', FileUpload);

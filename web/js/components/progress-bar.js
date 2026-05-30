export class ProgressBar extends HTMLElement {
  constructor() { super(); this.attachShadow({ mode: 'open' }); }
  connectedCallback() { this.render(); }
  render() {
    this.shadowRoot.innerHTML = `<style>:host { display: block; margin: 12px 0; } .container { background: #E0E4E8; border-radius: 4px; height: 8px; overflow: hidden; } .bar { background: #4472C4; height: 100%; width: 0%; transition: width 0.3s; border-radius: 4px; } .label { font-size: 12px; color: #666; margin-top: 4px; } .indeterminate .bar { animation: indeterminate 1.5s ease-in-out infinite; width: 30%; } @keyframes indeterminate { 0% { transform: translateX(-100%); } 100% { transform: translateX(400%); } }</style><div class="container"><div class="bar"></div></div><div class="label"></div>`;
  }
  setProgress(value, label) {
    const bar = this.shadowRoot.querySelector('.bar');
    const lbl = this.shadowRoot.querySelector('.label');
    if (value < 0) { this.shadowRoot.querySelector('.container').classList.add('indeterminate'); bar.style.width = ''; lbl.textContent = label || '处理中...'; }
    else { this.shadowRoot.querySelector('.container').classList.remove('indeterminate'); bar.style.width = `${value}%`; lbl.textContent = label || `${value}%`; }
  }
  hide() { this.style.display = 'none'; }
  show() { this.style.display = ''; }
}
customElements.define('progress-bar', ProgressBar);

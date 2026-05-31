// Pipeline Navigation — "Next Steps" recommended tools after an operation
export class PipelineNav extends HTMLElement {
  constructor() {
    super();
    this.attachShadow({ mode: 'open' });
  }

  connectedCallback() {
    this.render();
  }

  render() {
    const context = this.getAttribute('context') || '';
    const steps = this._getRecommendations(context);

    this.shadowRoot.innerHTML = `
      <style>
        :host { display: block; margin-top: 16px; padding: 16px; background: #FFF8E1; border-radius: var(--radius); border: 1px solid #F57F17; }
        .pn-title { font-size: 13px; font-weight: 600; color: #F57F17; margin-bottom: 8px; }
        .pn-steps { display: flex; gap: 8px; flex-wrap: wrap; }
        .pn-step { padding: 8px 16px; background: white; border: 1px solid var(--color-border); border-radius: 20px; cursor: pointer; font-size: 13px; text-decoration: none; color: var(--color-text); transition: all 0.15s; display: inline-flex; align-items: center; gap: 6px; }
        .pn-step:hover { background: var(--color-primary); color: white; border-color: var(--color-primary); }
        .pn-flow { font-size: 20px; color: #F57F17; line-height: 36px; }
      </style>
      <div class="pn-title">👉 下一步建议操作</div>
      <div class="pn-steps">
        ${steps.map((s, i) => {
          const sep = i < steps.length - 1 ? '<span class="pn-flow">→</span>' : '';
          return `<a class="pn-step" href="${s.route}" data-route="${s.route}">${s.icon} ${s.label}</a>${sep}`;
        }).join('')}
      </div>`;
  }

  _getRecommendations(context) {
    const all = {
      toc:      { icon: '📑', label: '生成 TOC 目录', route: '/toc' },
      clean:    { icon: '🧹', label: '数据清理', route: '/clean' },
      changes:  { icon: '📝', label: '修改痕迹跟踪', route: '/changes' },
      scan:     { icon: '🔍', label: '特殊字符扫描', route: '/scan' },
      sheets:   { icon: '📋', label: '工作表管理', route: '/sheets' },
      dates:    { icon: '📅', label: '日期计算', route: '/dates' },
      units:    { icon: '⚖️', label: '单位转换', route: '/units' },
      merge:    { icon: '📑', label: '工作簿合并', route: '/merge' },
      split:    { icon: '✂️', label: '工作表拆分', route: '/split' },
      medical:  { icon: '🩺', label: '医学计算器', route: '/creatinine' },
      download: { icon: '📥', label: '下载最终文件', route: '#' },
    };

    // Context-based recommendation chains
    const chains = {
      'merge':    [all.toc, all.clean, all.changes, all.download],
      'split':    [all.toc, all.clean, all.download],
      'toc':      [all.clean, all.changes, all.scan, all.download],
      'clean':    [all.toc, all.scan, all.changes, all.download],
      'changes':  [all.toc, all.download],
      'scan':     [all.clean, all.toc, all.download],
      'dates':    [all.clean, all.toc, all.download],
      'units':    [all.clean, all.toc, all.download],
      'sheets':   [all.toc, all.clean, all.download],
      'medical':  [all.toc, all.clean, all.download],
      'sas':      [all.clean, all.toc, all.scan, all.download],
      'default':  [all.toc, all.clean, all.scan, all.changes, all.download],
    };

    return chains[context] || chains['default'];
  }
}

customElements.define('pipeline-nav', PipelineNav);

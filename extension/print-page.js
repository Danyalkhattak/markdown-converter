(async function() {
  const data = await chrome.storage.local.get('exportData');
  if (!data.exportData) {
    document.getElementById('content').innerText = 'No data found.';
    return;
  }
  const { markdown, font, codeTheme, pageBreakHeadings, mermaidEnabled, mathEnabled } = data.exportData;

  // Apply font class
  const fontMap = {
    'Inter': 'font-inter', 'Roboto': 'font-roboto', 'Open Sans': 'font-opensans',
    'Times New Roman': 'font-times', 'Georgia': 'font-georgia', 'Garamond': 'font-garamond',
    'JetBrains Mono': 'font-jetbrains', 'Fira Code': 'font-firacode', 'Consolas': 'font-consolas',
    'SF Pro': 'font-sfpro'
  };
  document.body.classList.add(fontMap[font] || 'font-inter');

  // Switch code theme
  const lightTheme = document.getElementById('codeThemeLight');
  const darkTheme = document.getElementById('codeThemeDark');
  if (codeTheme === 'dark') {
    lightTheme.disabled = true;
    darkTheme.disabled = false;
    document.body.classList.add('code-dark');
  } else {
    lightTheme.disabled = false;
    darkTheme.disabled = true;
    document.body.classList.remove('code-dark');
  }

  // Parse markdown
  marked.setOptions({ gfm: true, breaks: false, tables: true });
  let html = marked.parse(markdown);
  if (mathEnabled) html = renderMath(html);
  document.getElementById('content').innerHTML = html;

  // Syntax highlight
  document.querySelectorAll('pre code').forEach(block => hljs.highlightElement(block));

  // Mermaid rendering
  if (mermaidEnabled) {
    mermaid.initialize({ startOnLoad: false, theme: codeTheme === 'dark' ? 'dark' : 'default' });
    const mermaidEls = document.querySelectorAll('pre code.language-mermaid');
    for (const el of mermaidEls) {
      try {
        const code = el.textContent;
        const id = 'm' + Math.random().toString(36).slice(2, 9);
        const { svg } = await mermaid.render(id, code);
        const div = document.createElement('div');
        div.className = 'mermaid';
        div.innerHTML = svg;
        el.parentElement.replaceWith(div);
      } catch (e) { console.warn('Mermaid error:', e); }
    }
  }

  // Page break handling – only for H1, and only if explicitly enabled
  if (pageBreakHeadings) {
    const h1s = document.querySelectorAll('h1');
    h1s.forEach(h => h.classList.add('page-break-before'));
  }

  // Auto-print after render
  setTimeout(() => window.print(), 400);
})();

function renderMath(html) {
  // Block math $$ ... $$
  html = html.replace(/\$\$([^$]+)\$\$/g, (_, formula) => {
    try {
      return katex.renderToString(formula.trim(), { displayMode: true, throwOnError: false });
    } catch {
      return `<pre>${formula}</pre>`;
    }
  });
  // Inline math $ ... $
  html = html.replace(/\$([^$]+)\$/g, (_, formula) => {
    try {
      return katex.renderToString(formula.trim(), { displayMode: false, throwOnError: false });
    } catch {
      return formula;
    }
  });
  return html;
}
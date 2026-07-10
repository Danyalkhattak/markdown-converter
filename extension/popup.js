// MD Converter v2.0.0 — Popup Logic

// --- Toast ---
function showNotification(message, type = 'info') {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.className = `toast ${type} show`;
  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => toast.classList.remove('show'), 3000);
}

// --- Init ---
document.addEventListener('DOMContentLoaded', async () => {
  await applyTheme();
  await loadSettings();

  // Header actions
  document.getElementById('darkModeToggle').addEventListener('click', toggleDarkMode);
  document.getElementById('settingsToggle').addEventListener('click', toggleSettings);
  document.getElementById('settingsClose').addEventListener('click', toggleSettings);

  // Input actions
  document.getElementById('uploadBtn').addEventListener('click', () => {
    document.getElementById('fileInput').click();
  });
  document.getElementById('fileInput').addEventListener('change', handleFileUpload);
  document.getElementById('getSelectionBtn').addEventListener('click', getSelection);
  document.getElementById('clearBtn').addEventListener('click', () => {
    document.getElementById('markdownInput').value = '';
    document.getElementById('markdownInput').focus();
  });

  // Export
  document.getElementById('exportPdf').addEventListener('click', exportPdf);
  document.getElementById('exportDocx').addEventListener('click', exportDocx);

  // Auto-save settings on change
  ['defaultFontSelect', 'defaultFormatSelect', 'codeThemeSelect', 'pageBreakToggle', 'mermaidToggle', 'mathToggle']
    .forEach(id => document.getElementById(id).addEventListener('change', saveSettings));
});

// --- Theme ---
const THEME_KEY = 'md_theme';

async function applyTheme() {
  const data = await chrome.storage.local.get(THEME_KEY);
  const theme = data[THEME_KEY] || 'light';
  document.body.className = theme;
  updateThemeIcons(theme);
}

async function toggleDarkMode() {
  const isDark = document.body.classList.contains('dark');
  const newTheme = isDark ? 'light' : 'dark';
  document.body.className = newTheme;
  updateThemeIcons(newTheme);
  await chrome.storage.local.set({ [THEME_KEY]: newTheme });
}

function updateThemeIcons(theme) {
  document.getElementById('sunIcon').style.display = theme === 'light' ? 'none' : 'block';
  document.getElementById('moonIcon').style.display = theme === 'light' ? 'block' : 'none';
}

// --- File Upload ---
function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = (e) => {
    document.getElementById('markdownInput').value = e.target.result;
    showNotification('File loaded', 'success');
  };
  reader.readAsText(file);
  event.target.value = '';
}

// --- Settings ---
const SETTINGS_KEY = 'md_settings';
const DEFAULTS = {
  defaultFont: 'Inter',
  defaultFormat: 'pdf',
  codeTheme: 'light',
  pageBreakHeadings: false,
  mermaidEnabled: true,
  mathEnabled: true
};

async function loadSettings() {
  const data = await chrome.storage.local.get(SETTINGS_KEY);
  let s = data[SETTINGS_KEY] || DEFAULTS;
  if (s.pageBreakHeadings === true) {
    s.pageBreakHeadings = false;
    await chrome.storage.local.set({ [SETTINGS_KEY]: s });
  }
  document.getElementById('fontSelect').value = s.defaultFont;
  document.getElementById('defaultFontSelect').value = s.defaultFont;
  document.getElementById('defaultFormatSelect').value = s.defaultFormat;
  document.getElementById('codeThemeSelect').value = s.codeTheme;
  document.getElementById('pageBreakToggle').checked = s.pageBreakHeadings;
  document.getElementById('mermaidToggle').checked = s.mermaidEnabled;
  document.getElementById('mathToggle').checked = s.mathEnabled;
}

async function saveSettings() {
  const settings = {
    defaultFont: document.getElementById('defaultFontSelect').value,
    defaultFormat: document.getElementById('defaultFormatSelect').value,
    codeTheme: document.getElementById('codeThemeSelect').value,
    pageBreakHeadings: document.getElementById('pageBreakToggle').checked,
    mermaidEnabled: document.getElementById('mermaidToggle').checked,
    mathEnabled: document.getElementById('mathToggle').checked
  };
  await chrome.storage.local.set({ [SETTINGS_KEY]: settings });
}

function toggleSettings() {
  const panel = document.getElementById('settingsPanel');
  if (panel.hidden) {
    panel.hidden = false;
  } else {
    panel.hidden = true;
  }
}

// --- Get Selection ---
async function getSelection() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab) return;
  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => window.getSelection().toString()
    });
    const text = results[0]?.result || '';
    if (text) {
      document.getElementById('markdownInput').value = text;
      showNotification('Selection inserted', 'success');
    } else {
      showNotification('No text selected', 'error');
    }
  } catch (e) {
    showNotification('Cannot access page', 'error');
  }
}

// --- Export ---
async function getExportSettings() {
  return {
    font: document.getElementById('fontSelect').value,
    codeTheme: document.getElementById('codeThemeSelect').value,
    pageBreakHeadings: document.getElementById('pageBreakToggle').checked,
    mermaidEnabled: document.getElementById('mermaidToggle').checked,
    mathEnabled: document.getElementById('mathToggle').checked
  };
}

async function exportPdf() {
  const markdown = document.getElementById('markdownInput').value.trim();
  if (!markdown) {
    showNotification('Enter some Markdown first', 'error');
    return;
  }
  const settings = await getExportSettings();
  await chrome.storage.local.set({ exportData: { markdown, ...settings } });
  const printUrl = chrome.runtime.getURL('print.html');
  chrome.tabs.create({ url: printUrl });
}

async function exportDocx() {
  const markdown = document.getElementById('markdownInput').value.trim();
  if (!markdown) {
    showNotification('Enter some Markdown first', 'error');
    return;
  }
  const settings = await getExportSettings();
  try {
    const blob = await generateDocx(markdown, settings);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'document.docx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showNotification('DOCX exported!', 'success');
  } catch (e) {
    console.error(e);
    showNotification('Export failed: ' + e.message, 'error');
  }
}
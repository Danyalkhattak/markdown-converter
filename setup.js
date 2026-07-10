const axios = require('axios');
const fs = require('fs');
const path = require('path');

const LIBS = {
  'marked.min.js': 'https://cdn.jsdelivr.net/npm/marked/marked.min.js',
  'katex.min.js': 'https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.js',
  'katex.min.css': 'https://cdn.jsdelivr.net/npm/katex@0.16.9/dist/katex.min.css',
  'mermaid.min.js': 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js',
  'highlight.min.js': 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/highlight.min.js',
  'github.min.css': 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/github.min.css',
  'atom-one-dark.min.css': 'https://cdnjs.cloudflare.com/ajax/libs/highlight.js/11.9.0/styles/atom-one-dark.min.css',
  'docx.umd.js': 'https://unpkg.com/docx@8.5.0/build/index.umd.js'
};

const KATEX_VERSION = '0.16.9';
const KATEX_FONT_NAMES = [
  'KaTeX_AMS-Regular',
  'KaTeX_Caligraphic-Bold',
  'KaTeX_Caligraphic-Regular',
  'KaTeX_Fraktur-Bold',
  'KaTeX_Fraktur-Regular',
  'KaTeX_Main-Bold',
  'KaTeX_Main-BoldItalic',
  'KaTeX_Main-Italic',
  'KaTeX_Main-Regular',
  'KaTeX_Math-BoldItalic',
  'KaTeX_Math-Italic',
  'KaTeX_Math-Regular',
  'KaTeX_SansSerif-Bold',
  'KaTeX_SansSerif-Italic',
  'KaTeX_SansSerif-Regular',
  'KaTeX_Script-Regular',
  'KaTeX_Size1-Regular',
  'KaTeX_Size2-Regular',
  'KaTeX_Size3-Regular',
  'KaTeX_Size4-Regular',
  'KaTeX_Typewriter-Regular'
];
const KATEX_FONT_EXTS = ['woff2', 'woff', 'ttf'];

const libDir = path.join(__dirname, 'extension', 'libs');
const fontDir = path.join(libDir, 'fonts');
if (!fs.existsSync(libDir)) fs.mkdirSync(libDir, { recursive: true });
if (!fs.existsSync(fontDir)) fs.mkdirSync(fontDir, { recursive: true });

function download(url) {
  return axios.get(url, { responseType: 'arraybuffer' }).then(function (res) {
    return res.data;
  });
}

(async () => {
  // 1. Download library files
  for (const [filename, url] of Object.entries(LIBS)) {
    console.log('Downloading ' + filename + '...');
    const data = await download(url);
    fs.writeFileSync(path.join(libDir, filename), data);
  }

  // 2. Download KaTeX fonts
  console.log('Downloading KaTeX fonts...');
  const cdnBase = 'https://cdn.jsdelivr.net/npm/katex@' + KATEX_VERSION + '/dist/fonts/';
  let fontCount = 0;

  for (const name of KATEX_FONT_NAMES) {
    // Download woff2 + ttf (woff is optional fallback)
    const exts = ['woff2', 'ttf'];
    for (const ext of exts) {
      const filename = name + '.' + ext;
      const url = cdnBase + filename;
      try {
        const data = await download(url);
        fs.writeFileSync(path.join(fontDir, filename), data);
        fontCount++;
      } catch (e) {
        console.warn('  Skipped ' + filename + ' (' + e.message + ')');
      }
    }
  }

  console.log('All libraries ready. (' + fontCount + ' font files downloaded)');
})();
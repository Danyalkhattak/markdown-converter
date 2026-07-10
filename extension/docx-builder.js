// docx-builder.js v2.0.0 — Full Markdown → DOCX converter
// Requires: global docx (docx.umd.js), global marked (marked.min.js)
// Auto-loads: katex.min.js, katex.min.css, mermaid.min.js

async function generateDocx(markdown, settings) {
  if (typeof docx === 'undefined') throw new Error('DOCX library not loaded.');
  if (typeof marked === 'undefined') throw new Error('Markdown parser not loaded.');

  await ensureLib('katex.min.css', 'css');
  await ensureLib('katex.min.js', 'js');
  await ensureLib('mermaid.min.js', 'js');

  const FONT = settings.font || 'Inter';
  const DARK = (settings.codeTheme || 'light') === 'dark';
  const MERMAID_ON = settings.mermaidEnabled !== false;
  const MATH_ON = settings.mathEnabled !== false;

  marked.setOptions({ gfm: true, breaks: false });
  const tokens = marked.lexer(markdown);
  const children = [];

  // ══════════════════════════════════════════════
  //  Dynamic library loader
  // ══════════════════════════════════════════════
  async function ensureLib(file, type) {
    const sel = type === 'css'
      ? 'link[href*="' + file + '"]'
      : 'script[src*="' + file + '"]';
    if (document.querySelector(sel)) return;
    try {
      const url = chrome.runtime.getURL('libs/' + file);
      if (type === 'css') {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = url;
        document.head.appendChild(link);
        await new Promise(r => setTimeout(r, 150));
      } else {
        await new Promise((resolve, reject) => {
          const s = document.createElement('script');
          s.src = url;
          s.onload = resolve;
          s.onerror = reject;
          document.head.appendChild(s);
        });
      }
    } catch (e) {
      console.warn('Failed to load ' + file + ':', e);
    }
  }

  // ══════════════════════════════════════════════
  //  TextRun factory
  // ══════════════════════════════════════════════
  function R(text, o = {}) {
    return new docx.TextRun({
      text: text || '',
      bold: !!o.bold,
      italics: !!o.italics,
      strike: !!o.strike,
      font: o.font || FONT,
      size: o.size || 24,
      color: o.color || undefined,
      underline: o.underline || undefined,
      shading: o.shading || undefined
    });
  }

  // ══════════════════════════════════════════════
  //  LaTeX → docx Math objects parser
  //  Uses the built-in docx Math API (v8.5.0)
  //  Produces native OMML equations in Word
  // ══════════════════════════════════════════════

  var GREEK = {
    alpha:'α',beta:'β',gamma:'γ',delta:'δ',epsilon:'ε',varepsilon:'ε',
    zeta:'ζ',eta:'η',theta:'θ',vartheta:'ϑ',iota:'ι',kappa:'κ',
    lambda:'λ',mu:'μ',nu:'ν',xi:'ξ',pi:'π',varpi:'ϖ',
    rho:'ρ',varrho:'ϱ',sigma:'σ',varsigma:'ς',tau:'τ',
    upsilon:'υ',phi:'φ',varphi:'φ',chi:'χ',psi:'ψ',omega:'ω',
    Gamma:'Γ',Delta:'Δ',Theta:'Θ',Lambda:'Λ',Xi:'Ξ',
    Pi:'Π',Sigma:'Σ',Phi:'Φ',Psi:'Ψ',Omega:'Ω'
  };

  var SYMBOLS = {
    infty:'∞',partial:'∂',nabla:'∇',hbar:'ℏ',ell:'ℓ',
    Re:'ℜ',Im:'ℑ',aleph:'ℵ',
    forall:'∀',exists:'∃',nexists:'∄',
    in:'∈',notin:'∉',subset:'⊂',supset:'⊃',subseteq:'⊆',supseteq:'⊇',
    cup:'∪',cap:'∩',setminus:'\\',emptyset:'∅',varnothing:'∅',
    rightarrow:'→',to:'→',leftarrow:'←',gets:'←',
    Rightarrow:'⇒',Leftarrow:'⇐',Leftrightarrow:'⇔',iff:'⇔',
    pm:'±',mp:'∓',times:'×',div:'÷',cdot:'·',ast:'∗',
    ldots:'…',cdots:'⋯',vdots:'⋮',ddots:'⋱',
    leq:'≤',geq:'≥',neq:'≠',approx:'≈',equiv:'≡',
    sim:'∼',simeq:'≃',cong:'≅',propto:'∝',
    langle:'⟨',rangle:'⟩',langle:'⟨',rangle:'⟩',
    lceil:'⌈',rceil:'⌉',lfloor:'⌊',rfloor:'⌋',
    parallel:'∥',perp:'⊥',angle:'∠',triangle:'△',
    prime:'′',dprime:'″',backslash:'\\',
    quad:'  ',qquad:'    ',space:' ',
    neg:'¬',wedge:'∧',vee:'∨',oplus:'⊕',otimes:'⊗',
    circ:'∘',bullet:'•',
    prec:'≺',succ:'≻',preceq:'⪯',succeq:'⪰',
    ll:'≪',gg:'≫',
    sum:'∑',prod:'∏',coprod:'∐',
    int:'∫',iint:'∬',iiint:'∭',oint:'∮',
    bigcup:'⋃',bigcap:'⋂',bigsqcup:'⨆',
    bigoplus:'⨁',bigotimes:'⨂',biguplus:'⨄'
  };

  var ACCENTS = {
    hat:'\u0302',tilde:'\u0303',bar:'\u0304',vec:'\u20D7',
    dot:'\u0307',ddot:'\u0308',dddot:'\u20DB',
    check:'\u030C',breve:'\u0306',acute:'\u0301',grave:'\u0300',
    widehat:'\u0302',widetilde:'\u0303',overline:'\u0304',
    overbrace:'',underbrace:''
  };

  var BRACKET_MAP = {
    '(': 'round', ')': 'round',
    '[': 'square', ']': 'square',
    '{': 'curly', '}': 'curly',
    '<': 'angle', '>': 'angle',
    '|': 'round', '.': 'none'
  };

  // Helper: flatten arrays
  function flat(arr) {
    var out = [];
    for (var i = 0; i < arr.length; i++) {
      if (Array.isArray(arr[i])) out.push.apply(out, arr[i]);
      else if (arr[i] !== null && arr[i] !== undefined) out.push(arr[i]);
    }
    return out;
  }

  // Helper: ensure array
  function arr(v) {
    if (Array.isArray(v)) return v;
    return v != null ? [v] : [];
  }

  // Main LaTeX math parser
  // Returns a docx.Math object, or null on failure
  function latexToMath(formula) {
    if (!formula || !formula.trim()) return null;
    try {
      var parser = new LatexParser(formula.trim());
      var items = parser.parseExpression();
      if (items.length === 0) return null;
      return new docx.Math({ children: items });
    } catch (e) {
      console.warn('LaTeX math parse error:', e.message, 'for:', formula);
      return null;
    }
  }

  function LatexParser(str) {
    this.s = str;
    this.p = 0;
  }

  var LP = LatexParser.prototype;

  LP.peek = function() { return this.p < this.s.length ? this.s[this.p] : ''; };
  LP.next = function() { return this.p < this.s.length ? this.s[this.p++] : ''; };
  LP.eat = function(ch) { if (this.peek() === ch) { this.p++; return true; } return false; };
  LP.atEnd = function() { return this.p >= this.s.length; };

  LP.parseExpression = function() {
    var items = [];
    while (!this.atEnd() && this.peek() !== '}') {
      var ch = this.peek();
      if (ch === '&') { this.next(); continue; }
      if (ch === '\\' && this.p + 1 < this.s.length && this.s[this.p + 1] === '\\') {
        this.p += 2; continue; // line break in equation
      }
      var result = this.parseItem();
      if (result) {
        var flat_items = flat([result]);
        for (var i = 0; i < flat_items.length; i++) items.push(flat_items[i]);
      }
    }
    return items;
  };

  LP.parseItem = function() {
    // Skip whitespace
    while (this.peek() === ' ' || this.peek() === '\t') this.next();
    if (this.atEnd() || this.peek() === '}') return null;

    // Braced group
    if (this.peek() === '{') {
      this.next(); // skip {
      var items = this.parseExpression();
      if (this.peek() === '}') this.next();
      return this.applySubSup(items);
    }

    // Backslash command
    if (this.peek() === '\\') {
      var cmd = this.parseCommand();
      if (cmd === null) return null;
      return this.applySubSup(Array.isArray(cmd) ? cmd : [cmd]);
    }

    // Single character
    var ch = this.next();
    return this.applySubSup([new docx.MathRun(ch)]);
  };

  LP.applySubSup = function(baseItems) {
    if (!baseItems || baseItems.length === 0) return null;
    var sub = null, sup = null;

    if (this.peek() === '_') {
      this.next();
      sub = this.parseGroup();
    }
    if (this.peek() === '^') {
      this.next();
      sup = this.parseGroup();
    }

    if (sub && sup) {
      return new docx.MathSubSuperScript({
        children: baseItems,
        subScript: arr(sub),
        superScript: arr(sup)
      });
    }
    if (sub) {
      return new docx.MathSubScript({
        children: baseItems,
        subScript: arr(sub)
      });
    }
    if (sup) {
      return new docx.MathSuperScript({
        children: baseItems,
        superScript: arr(sup)
      });
    }
    // No sub/superscript — if single item, return it directly
    return baseItems.length === 1 ? baseItems[0] : baseItems;
  };

  LP.parseGroup = function() {
    if (this.peek() === '{') {
      this.next();
      var items = this.parseExpression();
      if (this.peek() === '}') this.next();
      return items.length === 1 ? items[0] : items;
    }
    // Single token (character or command)
    if (this.peek() === '\\') {
      var cmd = this.parseCommand();
      return cmd;
    }
    return new docx.MathRun(this.next());
  };

  LP.parseCommand = function() {
    this.next(); // skip \
    if (this.atEnd()) return new docx.MathRun('\\');

    // Check for special single-char commands: \{ \} \_ \^ \  \& \# \% \!
    var sp = '{}_^\\&#%!';
    if (sp.indexOf(this.peek()) >= 0) {
      return new docx.MathRun(this.next());
    }

    // Read command name
    var cmd = '';
    while (!this.atEnd() && /[a-zA-Z]/.test(this.peek())) {
      cmd += this.next();
    }
    if (!cmd) return new docx.MathRun('\\');

    // ── Structural commands ──

    if (cmd === 'frac' || cmd === 'dfrac' || cmd === 'tfrac') {
      var num = this.parseGroup();
      var den = this.parseGroup();
      return new docx.MathFraction({
        numerator: arr(num),
        denominator: arr(den)
      });
    }

    if (cmd === 'sqrt') {
      var hasDegree = this.peek() === '[';
      var degreeItems = null;
      if (hasDegree) {
        this.next(); // skip [
        degreeItems = this.parseExpression();
        if (this.peek() === ']') this.next();
      }
      var radicand = this.parseGroup();
      var opts = { children: arr(radicand) };
      if (hasDegree && degreeItems && degreeItems.length > 0) {
        opts.properties = new docx.MathRadicalProperties({ hideDegree: false });
        opts.degree = degreeItems;
      } else {
        opts.properties = new docx.MathRadicalProperties({ hideDegree: true });
      }
      return new docx.MathRadical(opts);
    }

    // ── Nary operators (sum, prod, int, etc.) ──
    var naryMap = { sum:'sum', prod:'prod', coprod:'sum', int:'integral',
      iint:'integral', iiint:'integral', oint:'integral' };
    if (naryMap[cmd]) {
      return this.parseNary(cmd);
    }
    // Also handle bigcup, bigcap, etc.
    if (cmd.indexOf('big') === 0 && SYMBOLS[cmd]) {
      return this.parseNary(cmd);
    }

    // ── Accents ──
    if (ACCENTS[cmd] !== undefined) {
      // Check for {} group or single char
      var target = this.parseGroup();
      var chr = ACCENTS[cmd];
      if (cmd === 'overbrace' || cmd === 'underbrace') {
        // These are special — just render as the base with a brace character
        var braceItems = arr(target);
        if (cmd === 'overbrace') {
          braceItems.push(new docx.MathRun('\u23DE')); // top brace
        } else {
          braceItems.push(new docx.MathRun('\u23DF')); // bottom brace
        }
        return braceItems.length === 1 ? braceItems[0] : braceItems;
      }
      if (chr) {
        return new docx.MathAccentCharacter({
          character: chr,
          children: arr(target)
        });
      }
      return target;
    }

    // ── Left/Right delimiters ──
    if (cmd === 'left' || cmd === 'right') {
      // Skip optional space after \left/\right
      while (this.peek() === ' ') this.next();
      var delim = this.next();
      if (cmd === 'left') {
        // Parse content until matching \right
        var inner = this.parseUntilRight();
        var bracketResult = this.wrapInBracket(delim, inner);
        return bracketResult;
      }
      // \right just returns nothing (the content was already parsed by \left)
      return null;
    }

    // ── Text ──
    if (cmd === 'text' || cmd === 'textit' || cmd === 'textbf' || cmd === 'mathrm' || cmd === 'mathbf' || cmd === 'mathit' || cmd === 'mathcal' || cmd === 'mathbb' || cmd === 'mathsf' || cmd === 'mathtt') {
      var textContent = this.parseGroup();
      // Convert to array of MathRuns
      if (Array.isArray(textContent)) {
        return textContent; // Already MathRun objects
      }
      if (textContent instanceof docx.MathRun) {
        return textContent;
      }
      return new docx.MathRun(String(textContent));
    }

    // ── Spaces ──
    if (cmd === ',' || cmd === ';' || cmd === ':' || cmd === '!') {
      return new docx.MathRun(' ');
    }

    // ── Greek letters ──
    if (GREEK[cmd] !== undefined) {
      return new docx.MathRun(GREEK[cmd]);
    }

    // ── Symbols ──
    if (SYMBOLS[cmd] !== undefined) {
      return new docx.MathRun(SYMBOLS[cmd]);
    }

    // ── Binomial coefficient ──
    if (cmd === 'binom') {
      var top = this.parseGroup();
      var bottom = this.parseGroup();
      return new docx.MathFraction({
        numerator: arr(top),
        denominator: arr(bottom)
      });
    }

    // ── Underline/overline ──
    if (cmd === 'underline') {
      var ulTarget = this.parseGroup();
      return new docx.MathRun('_'); // Simplified
    }
    if (cmd === 'overline' || cmd === 'bar') {
      var olTarget = this.parseGroup();
      return new docx.MathAccentCharacter({
        character: '\u0304',
        children: arr(olTarget)
      });
    }

    // ── Known functions (render as upright text) ──
    var funcNames = ['sin','cos','tan','sec','csc','cot','arcsin','arccos','arctan',
      'sinh','cosh','tanh','coth','log','ln','exp','lim','sup','inf','min','max',
      'det','dim','ker','hom','deg','arg','gcd','Pr'];
    if (funcNames.indexOf(cmd) >= 0) {
      return new docx.MathRun(cmd);
    }

    // ── Unknown command: try to use as symbol ──
    return new docx.MathRun('\\' + cmd);
  };

  LP.parseNary = function(cmd) {
    var chr = SYMBOLS[cmd] || '∑';
    var sub = null, sup = null;

    // Parse optional limits
    if (this.peek() === '_') {
      this.next();
      sub = this.parseGroup();
    }
    if (this.peek() === '^') {
      this.next();
      sup = this.parseGroup();
    }

    var opts = {
      children: [new docx.MathRun(' ')],
      subScript: arr(sub),
      superScript: arr(sup)
    };

    if (cmd === 'int' || cmd === 'iint' || cmd === 'iiint' || cmd === 'oint') {
      return new docx.MathIntegral(opts);
    }
    return new docx.MathSum(opts);
  };

  LP.parseUntilRight = function() {
    // Parse expression until we hit \right
    var items = [];
    var depth = 0;
    while (!this.atEnd()) {
      // Check for \right
      if (this.peek() === '\\') {
        var savedP = this.p;
        this.next(); // skip \
        var cmdName = '';
        while (!this.atEnd() && /[a-zA-Z]/.test(this.peek())) {
          cmdName += this.next();
        }
        if (cmdName === 'right') {
          return items;
        }
        // Not \right, restore position and parse normally
        this.p = savedP;
      }
      if (this.peek() === '{') depth++;
      if (this.peek() === '}' && depth > 0) depth--;
      if (this.peek() === '}' && depth === 0) break;

      var item = this.parseItem();
      if (item) {
        var flat_item = flat([item]);
        for (var i = 0; i < flat_item.length; i++) items.push(flat_item[i]);
      }
    }
    return items;
  };

  LP.wrapInBracket = function(delim, items) {
    var type = BRACKET_MAP[delim];
    if (!type || type === 'none' || items.length === 0) {
      return items.length === 1 ? items[0] : items;
    }
    switch (type) {
      case 'round': return new docx.MathRoundBrackets({ children: items });
      case 'square': return new docx.MathSquareBrackets({ children: items });
      case 'curly': return new docx.MathCurlyBrackets({ children: items });
      case 'angle': return new docx.MathAngledBrackets({ children: items });
      default: return items.length === 1 ? items[0] : items;
    }
  };

  // ══════════════════════════════════════════════
  //  Recursive inline token → Run[]  (ASYNC)
  // ══════════════════════════════════════════════
  async function il(tokens, o = {}) {
    if (!Array.isArray(tokens)) return [];
    const out = [];
    for (const t of tokens) {
      switch (t.type) {

        case 'text': {
          if (t.tokens && t.tokens.length > 0) {
            out.push(...await il(t.tokens, o));
          } else {
            if (MATH_ON && /\$/.test(t.text)) {
              out.push(...await processMathText(t.text, o));
            } else {
              out.push(R(t.text, o));
            }
          }
          break;
        }

        case 'strong':
          out.push(...await il(t.tokens, { ...o, bold: true }));
          break;

        case 'em':
          out.push(...await il(t.tokens, { ...o, italics: true }));
          break;

        case 'del':
          out.push(...await il(t.tokens, { ...o, strike: true }));
          break;

        case 'codespan': {
          var codeText = t.text || '';
          var mathMatch = codeText.match(/^\$(.+)\$$/);
          if (MATH_ON && mathMatch) {
            var mathEl = latexToMath(mathMatch[1]);
            if (mathEl) { out.push(mathEl); break; }
          }
          out.push(R(t.text, {
            font: 'JetBrains Mono',
            size: (o.size || 24) - 3,
            shading: { type: 'solid', fill: DARK ? '2a2a3c' : 'f0f0f0' }
          }));
          break;
        }

        case 'link': {
          const lo = { ...o, color: '2563EB', underline: { type: 'single', color: '2563EB' } };
          const inner = t.tokens || [];
          const imgTk = inner.length === 1 && inner[0].type === 'image' ? inner[0] : null;
          if (imgTk) {
            out.push(R('[' + (imgTk.text || 'Image') + ']', lo));
          } else if (inner.length > 0) {
            out.push(...await il(inner, lo));
          } else {
            out.push(R(t.text || t.href, lo));
          }
          break;
        }

        case 'image':
          out.push(R('[' + (t.text || 'Image') + ']', { ...o, italics: true, color: '6b7280' }));
          break;

        case 'br':
          out.push(new docx.TextRun({ break: 1 }));
          break;

        default:
          if (t.raw) out.push(R(t.raw, o));
          break;
      }
    }
    return out;
  }

  // ══════════════════════════════════════════════
  //  Math text processor: $...$ and $$...$$
  //  Primary: native OMML via LaTeX parser
  //  Fallback: KaTeX → PNG image
  // ══════════════════════════════════════════════
  async function processMathText(text, o) {
    const out = [];
    const re = /\$\$([\s\S]+?)\$\$|\$([^\$\n]+?)\$/g;
    let last = 0, m, found = false;
    while ((m = re.exec(text)) !== null) {
      found = true;
      if (m.index > last) out.push(R(text.slice(last, m.index), o));
      const formula = (m[1] || m[2]).trim();

      // 1) Try native OMML equation via LaTeX parser
      var mathEl = latexToMath(formula);
      if (mathEl) {
        out.push(mathEl);
      } else {
        // 2) Fallback: KaTeX → PNG image
        var img = await renderMathImg(formula, !!m[1]);
        if (img) {
          out.push(img);
        } else {
          // 3) Last resort: italic text
          out.push(R(formula, {
            ...o, font: 'Times New Roman', italics: true, size: (o.size || 24) - 1
          }));
        }
      }
      last = m.index + m[0].length;
    }
    if (!found) out.push(R(text, o));
    else if (last < text.length) out.push(R(text.slice(last), o));
    return out;
  }

  // ══════════════════════════════════════════════
  //  KaTeX → PNG ImageRun  (fallback only)
  // ══════════════════════════════════════════════
  async function renderMathImg(formula, display) {
    if (typeof katex === 'undefined') return null;
    try {
      // Wait for fonts to be fully loaded
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }

      var htmlStr = katex.renderToString(formula, {
        displayMode: !!display,
        throwOnError: false,
        output: 'html',
        trust: true
      });

      var el = document.createElement('div');
      el.style.cssText = 'position:fixed;left:-9999px;top:0;visibility:hidden;background:#fff;padding:16px 20px;line-height:1.4;';
      el.innerHTML = htmlStr;
      document.body.appendChild(el);

      // Force layout computation
      var _ = el.offsetHeight;

      inlineAllStyles(el);

      var rect = el.getBoundingClientRect();
      var w = Math.ceil(rect.width) + 40;
      var h = Math.ceil(rect.height) + 32;

      // Safety: ensure dimensions are reasonable
      w = Math.max(w, 40);
      h = Math.max(h, 20);

      var clone = el.cloneNode(true);
      document.body.removeChild(el);

      var svgStr =
        '<svg xmlns="http://www.w3.org/2000/svg" width="' + w + '" height="' + h + '">' +
        '<foreignObject width="100%" height="100%">' +
        '<div xmlns="http://www.w3.org/1999/xhtml" style="background:#fff;margin:0;padding:4px 6px;">' +
        clone.innerHTML +
        '</div></foreignObject></svg>';

      var url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgStr);
      var img = await loadImg(url);

      var sc = 3;
      var canvas = document.createElement('canvas');
      canvas.width = w * sc;
      canvas.height = h * sc;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(sc, sc);
      ctx.drawImage(img, 0, 0, w, h);

      var b64 = canvas.toDataURL('image/png').split(',')[1];
      var bin = atob(b64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

      var dw = Math.min(w * 0.75, 540);
      return new docx.ImageRun({
        data: bytes,
        type: 'png',
        transformation: { width: Math.round(dw), height: Math.round(dw * h / w) }
      });
    } catch (e) {
      console.warn('Math image render failed:', e);
      return null;
    }
  }

  function inlineAllStyles(el) {
    var children = el.children;
    for (var i = 0; i < children.length; i++) {
      var child = children[i];
      var cs = window.getComputedStyle(child);
      var style = '';
      var props = [
        'position','display','vertical-align','text-align',
        'font-family','font-size','font-style','font-weight',
        'line-height','margin','margin-top','margin-bottom',
        'margin-left','margin-right','padding','padding-top',
        'padding-bottom','padding-left','padding-right',
        'border','width','height','top','bottom','left','right',
        'white-space','letter-spacing','color','background',
        'transform','transform-origin','box-sizing'
      ];
      for (var p = 0; p < props.length; p++) {
        var val = cs.getPropertyValue(props[p]);
        if (val && val !== 'none' && val !== 'normal' && val !== 'auto' && val !== '0px') {
          style += props[p] + ':' + val + ';';
        }
      }
      if (style) {
        child.setAttribute('style', (child.getAttribute('style') || '') + style);
      }
      inlineAllStyles(child);
    }
  }

  // ══════════════════════════════════════════════
  //  Mermaid → PNG ImageRun
  // ══════════════════════════════════════════════
  async function renderMermaidImg(code) {
    if (typeof mermaid === 'undefined') return null;
    try {
      // Create a hidden container for mermaid rendering
      var container = document.getElementById('mermaid-temp-container');
      if (!container) {
        container = document.createElement('div');
        container.id = 'mermaid-temp-container';
        container.style.cssText = 'position:fixed;left:-9999px;top:0;width:1200px;height:900px;visibility:hidden;overflow:visible;z-index:-1;';
        document.body.appendChild(container);
      }
      container.innerHTML = '';

      // Initialize mermaid — disable max-width so SVG gets real dimensions
      mermaid.initialize({
        startOnLoad: false,
        theme: 'default',
        securityLevel: 'loose',
        fontFamily: 'sans-serif',
        flowchart: { useMaxWidth: false, htmlLabels: true },
        sequence: { useMaxWidth: false, diagramMarginX: 50, diagramMarginY: 10 },
        pie: { useMaxWidth: false }
      });

      var id = 'mdc' + Date.now() + Math.random().toString(36).slice(2, 8);

      // mermaid v10 needs a DOM element with the render ID present
      var tempDiv = document.createElement('div');
      tempDiv.id = id;
      container.appendChild(tempDiv);

      var result = await mermaid.render(id, code.trim());

      // Handle both v9 (string) and v10 ({svg:string}) API
      var svgStr;
      if (typeof result === 'string') {
        svgStr = result;
      } else if (result && typeof result.svg === 'string') {
        svgStr = result.svg;
      } else {
        console.warn('Mermaid returned unexpected result:', result);
        return null;
      }

      // Clean up temp DOM elements
      var oldEl = document.getElementById(id);
      if (oldEl && oldEl !== tempDiv) oldEl.remove();
      tempDiv.remove();

      // Parse SVG to extract accurate dimensions
      var svgDoc = new DOMParser().parseFromString(svgStr, 'image/svg+xml');
      var svgRoot = svgDoc.documentElement;
      var svgW = 0, svgH = 0;

      // Method 1: viewBox (most reliable for mermaid)
      var vb = svgRoot.getAttribute('viewBox');
      if (vb) {
        var vbParts = vb.split(/[\s,]+/).map(Number);
        if (vbParts.length === 4 && vbParts[2] > 0 && vbParts[3] > 0) {
          svgW = vbParts[2];
          svgH = vbParts[3];
        }
      }

      // Method 2: width/height attributes
      if (!svgW || !svgH) {
        var attrW = parseFloat(svgRoot.getAttribute('width'));
        var attrH = parseFloat(svgRoot.getAttribute('height'));
        if (attrW > 0 && attrH > 0) {
          svgW = attrW;
          svgH = attrH;
        }
      }

      // Method 3: render to DOM and measure
      if (!svgW || !svgH) {
        container.innerHTML = svgStr;
        var renderedEl = container.querySelector('svg');
        if (renderedEl) {
          svgW = renderedEl.getBoundingClientRect().width || 400;
          svgH = renderedEl.getBoundingClientRect().height || 300;
        }
      }

      // Ensure minimum usable dimensions
      svgW = Math.max(svgW || 400, 100);
      svgH = Math.max(svgH || 300, 60);

      // Set explicit dimensions and clean up the SVG
      svgRoot.setAttribute('width', String(svgW));
      svgRoot.setAttribute('height', String(svgH));
      svgRoot.setAttribute('xmlns', 'http://www.w3.org/2000/svg');
      svgRoot.removeAttribute('style');

      svgStr = new XMLSerializer().serializeToString(svgRoot);

      // Convert SVG → Canvas → PNG
      var dataUrl = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svgStr);
      var img = await loadImg(dataUrl);

      var pad = 24;
      var canvasW = Math.ceil(svgW + pad * 2);
      var canvasH = Math.ceil(svgH + pad * 2);

      var sc = 2;
      var canvas = document.createElement('canvas');
      canvas.width = canvasW * sc;
      canvas.height = canvasH * sc;
      var ctx = canvas.getContext('2d');
      ctx.fillStyle = '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.scale(sc, sc);
      ctx.drawImage(img, pad, pad, svgW, svgH);

      var b64 = canvas.toDataURL('image/png').split(',')[1];
      var bin = atob(b64);
      var bytes = new Uint8Array(bin.length);
      for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);

      // Scale to fit DOCX page width (~520px)
      var maxW = 520;
      var dw = Math.min(canvasW, maxW);
      var dh = Math.round(dw * canvasH / canvasW);

      return new docx.ImageRun({
        data: bytes,
        type: 'png',
        transformation: { width: dw, height: dh }
      });
    } catch (e) {
      console.warn('Mermaid render failed:', e);
      return null;
    }
  }

  function loadImg(src) {
    return new Promise(function (res, rej) {
      var i = new Image();
      i.onload = function () { res(i); };
      i.onerror = rej;
      i.src = src;
    });
  }

  // ══════════════════════════════════════════════
  //  Code block → Paragraph[]
  // ══════════════════════════════════════════════
  function codeBlock(text, lang) {
    var paras = [];
    if (lang) {
      paras.push(new docx.Paragraph({
        children: [R(lang, { font: 'JetBrains Mono', size: 17, color: '6b7280', italics: true })],
        spacing: { before: 200, after: 40 }
      }));
    }
    var fill = DARK ? '1e1e1e' : 'f6f8fa';
    var col = DARK ? 'd4d4d4' : undefined;
    var lines = text.split('\n');
    for (var i = 0; i < lines.length; i++) {
      paras.push(new docx.Paragraph({
        children: [R(lines[i] || ' ', { font: 'JetBrains Mono', size: 20, color: col })],
        spacing: { before: 0, after: 0, line: 276 },
        shading: { type: 'solid', fill: fill },
        indent: { left: 240, right: 240 }
      }));
    }
    paras.push(new docx.Paragraph({ children: [], spacing: { after: 80 } }));
    return paras;
  }

  // ══════════════════════════════════════════════
  //  Block token processor
  // ══════════════════════════════════════════════
  for (var ti = 0; ti < tokens.length; ti++) {
    var tk = tokens[ti];
    try {
      switch (tk.type) {

        case 'heading': {
          var sz = [0, 40, 34, 28, 26, 24, 22][tk.depth] || 24;
          children.push(new docx.Paragraph({
            heading: 'Heading' + tk.depth,
            children: await il(tk.tokens, { bold: true, size: sz }),
            spacing: { before: tk.depth <= 2 ? 400 : 240, after: 120 }
          }));
          break;
        }

        case 'paragraph': {
          // Check if the entire paragraph is standalone display math $$...$$
          var rawText = tk.raw || '';
          var dmMatch = rawText.match(/^\s*\$\$([\s\S]+?)\$\$\s*$/);
          if (MATH_ON && dmMatch) {
            var dformula = dmMatch[1].trim();
            var dMathEl = latexToMath(dformula);
            if (dMathEl) {
              children.push(new docx.Paragraph({
                children: [dMathEl],
                alignment: 'center',
                spacing: { before: 240, after: 240 }
              }));
              break;
            }
            // Fallback to KaTeX image for display math
            var dImg = await renderMathImg(dformula, true);
            if (dImg) {
              children.push(new docx.Paragraph({
                children: [dImg],
                alignment: 'center',
                spacing: { before: 240, after: 240 }
              }));
              break;
            }
          }
          var runs = await il(tk.tokens, {});
          if (runs.length > 0) {
            children.push(new docx.Paragraph({ children: runs, spacing: { after: 160 } }));
          }
          break;
        }

        case 'code': {
          var lang = (tk.lang || '').toLowerCase();

          // Mermaid → rendered diagram PNG
          if (lang === 'mermaid' && MERMAID_ON) {
            var mImg = await renderMermaidImg(tk.text);
            if (mImg) {
              children.push(new docx.Paragraph({
                children: [mImg],
                spacing: { before: 200, after: 200 }
              }));
              break;
            }
          }

          // Math code blocks (```latex, ```math, ```katex, ```tex) → native equation
          if (MATH_ON && ['latex', 'math', 'katex', 'tex'].includes(lang)) {
            var formula = tk.text.replace(/^\$\$?/, '').replace(/\$\$?$/, '').trim();
            var blockMathEl = latexToMath(formula);
            if (blockMathEl) {
              children.push(new docx.Paragraph({
                children: [blockMathEl],
                alignment: 'center',
                spacing: { before: 240, after: 240 }
              }));
              break;
            }
            // Fallback to KaTeX image
            var blockImg = await renderMathImg(formula, true);
            if (blockImg) {
              children.push(new docx.Paragraph({
                children: [blockImg],
                alignment: 'center',
                spacing: { before: 240, after: 240 }
              }));
              break;
            }
            // Last resort: formatted text
            children.push(new docx.Paragraph({
              children: [R(formula, { font: 'Times New Roman', italics: true, size: 28 })],
              alignment: 'center',
              spacing: { before: 240, after: 240 },
              shading: { type: 'solid', fill: 'f9fafb' }
            }));
            break;
          }

          // Regular code block
          var cbParas = codeBlock(tk.text, lang);
          for (var ci = 0; ci < cbParas.length; ci++) children.push(cbParas[ci]);
          break;
        }

        case 'table': {
          var rows = [];
          var hCells = [];
          for (var hi = 0; hi < tk.header.length; hi++) {
            hCells.push(new docx.TableCell({
              children: [new docx.Paragraph({ children: await il(tk.header[hi].tokens, { bold: true, size: 22 }) })],
              shading: { fill: DARK ? '2a2a3c' : 'e5e7eb' },
              verticalAlign: 'center'
            }));
          }
          rows.push(new docx.TableRow({ children: hCells, tableHeader: true }));
          for (var ri = 0; ri < tk.rows.length; ri++) {
            var bCells = [];
            for (var bi = 0; bi < tk.rows[ri].length; bi++) {
              bCells.push(new docx.TableCell({
                children: [new docx.Paragraph({ children: await il(tk.rows[ri][bi].tokens, { size: 22 }) })],
                verticalAlign: 'center'
              }));
            }
            rows.push(new docx.TableRow({ children: bCells }));
          }
          children.push(new docx.Table({ rows: rows, width: { size: 100, type: 'pct' } }));
          children.push(new docx.Paragraph({ children: [], spacing: { after: 120 } }));
          break;
        }

        case 'list': {
          var ordered = tk.ordered;

          async function buildItem(item, lvl) {
            var result = [];
            var lastTk = item.tokens && item.tokens[item.tokens.length - 1];
            var hasNested = lastTk && lastTk.type === 'list';
            var inTokens = hasNested ? item.tokens.slice(0, -1) : (item.tokens || []);
            var itemRuns = await il(inTokens, { size: 24 });

            if (itemRuns.length > 0) {
              var para = {
                children: itemRuns,
                spacing: { before: 40, after: 40 },
                indent: { left: 360 + lvl * 360, hanging: 240 }
              };
              if (ordered) {
                para.numbering = { reference: 'docx-num', level: lvl };
              } else {
                para.bullet = { level: lvl };
              }
              result.push(new docx.Paragraph(para));
            }

            if (hasNested) {
              for (var si = 0; si < lastTk.items.length; si++) {
                var sub = await buildItem(lastTk.items[si], lvl + 1);
                for (var sj = 0; sj < sub.length; sj++) result.push(sub[sj]);
              }
            }
            return result;
          }

          for (var li = 0; li < tk.items.length; li++) {
            var built = await buildItem(tk.items[li], 0);
            for (var lj = 0; lj < built.length; lj++) children.push(built[lj]);
          }
          break;
        }

        case 'blockquote': {
          for (var qi = 0; qi < (tk.tokens || []).length; qi++) {
            var bq = tk.tokens[qi];
            if (bq.type === 'paragraph') {
              var bRuns = await il(bq.tokens, { color: '4b5563', size: 24 });
              if (bRuns.length) {
                children.push(new docx.Paragraph({
                  children: bRuns,
                  indent: { left: 720 },
                  border: { left: { style: 'single', size: 6, color: 'd1d5db', space: 8 } },
                  spacing: { before: 80, after: 80 }
                }));
              }
            } else if (bq.type === 'blockquote') {
              for (var ii = 0; ii < (bq.tokens || []).length; ii++) {
                var inner = bq.tokens[ii];
                if (inner.type === 'paragraph') {
                  var iRuns = await il(inner.tokens, { color: '6b7280', size: 24 });
                  if (iRuns.length) {
                    children.push(new docx.Paragraph({
                      children: iRuns,
                      indent: { left: 1080 },
                      border: { left: { style: 'single', size: 6, color: 'e5e7eb', space: 8 } },
                      spacing: { before: 60, after: 60 }
                    }));
                  }
                }
              }
            }
          }
          break;
        }

        case 'hr': {
          children.push(new docx.Paragraph({
            children: [],
            border: { bottom: { style: 'single', size: 6, color: 'd1d5db', space: 1 } },
            spacing: { before: 200, after: 200 }
          }));
          break;
        }

        case 'html': {
          var plain = tk.text.replace(/<[^>]*>/g, '').trim();
          if (plain) {
            children.push(new docx.Paragraph({ children: [R(plain)], spacing: { after: 120 } }));
          }
          break;
        }

        case 'space':
          break;

        default:
          if (tk.text || tk.raw) {
            children.push(new docx.Paragraph({ children: [R(tk.text || tk.raw)], spacing: { after: 120 } }));
          }
      }
    } catch (err) {
      console.warn('Token error:', tk.type, err);
    }
  }

  if (children.length === 0) throw new Error('No content to generate.');

  return await docx.Packer.toBlob(new docx.Document({
    styles: {
      default: {
        document: { run: { font: FONT, size: 24 } }
      }
    },
    numbering: {
      config: [{
        reference: 'docx-num',
        levels: [
          { level: 0, format: 'decimal', text: '%1.', alignment: 'start', style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
          { level: 1, format: 'lowerLetter', text: '%2)', alignment: 'start', style: { paragraph: { indent: { left: 1080, hanging: 360 } } } },
          { level: 2, format: 'lowerRoman', text: '%3.', alignment: 'start', style: { paragraph: { indent: { left: 1440, hanging: 360 } } } }
        ]
      }]
    },
    sections: [{ properties: {}, children: children }]
  }));
}
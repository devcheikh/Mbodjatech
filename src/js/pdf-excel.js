/* ==========================================================================
   CONVERTISSEUR PDF -> EXCEL
   100 % côté navigateur : aucun fichier n'est envoyé sur un serveur.

   Objectif : le classeur Excel doit ressembler au PDF d'origine.

   Pipeline :
     1. pdf.js fournit deux choses pour chaque page :
        - les fragments de texte avec leurs coordonnées et leur police
        - la liste des opérateurs graphiques, d'où l'on extrait les TRAITS
          du tableau et les aplats de couleur
     2. Si la page contient une grille de traits, elle sert directement de
        structure (mode « grille ») : c'est la reproduction la plus fidèle.
        Sinon, les colonnes sont déduites des gouttières (mode « espaces »).
     3. Chaque cellule conserve son style : gras, italique, taille, alignement,
        couleur de fond et bordures réellement présentes dans le PDF.
     4. ExcelJS écrit le classeur avec ces styles, les largeurs de colonnes et
        les hauteurs de lignes reprises du document.
   ========================================================================== */

(function () {
  'use strict';

  // ---------------------------------------------------------------- Config
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

  var MAX_PREVIEW_ROWS = 300;   // lignes affichées dans l'aperçu
  var MAX_COLUMNS = 60;         // garde-fou sur la détection de colonnes
  var RULE_TOL = 2.5;           // tolérance (pt) : épaisseur max d'un trait
  var MIN_RULE = 8;             // longueur (pt) minimale d'un trait retenu
  var CLUSTER_TOL = 3.5;        // tolérance (pt) de regroupement des traits
  var PT_PER_CHAR = 5.5;        // conversion points -> largeur Excel

  // ---------------------------------------------------------------- État
  var state = {
    files: [],        // { id, file, name, size, result, error, status }
    activeId: null,
    activePage: 0,
    busy: false
  };

  var el = {};

  // ================================================================
  //  UTILITAIRES
  // ================================================================

  function $(id) { return document.getElementById(id); }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function formatSize(bytes) {
    if (bytes < 1024) return bytes + ' o';
    if (bytes < 1048576) return (bytes / 1024).toFixed(0) + ' Ko';
    return (bytes / 1048576).toFixed(1) + ' Mo';
  }

  function baseName(name) {
    return name.replace(/\.pdf$/i, '') || 'document';
  }

  /** Numéro de colonne -> lettre Excel (0 -> A, 26 -> AA). */
  function colLetter(n) {
    var s = '';
    n = n + 1;
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  /** Analyse « 1-3,7,10- » -> table des pages autorisées. */
  function parsePageRange(text, total) {
    var raw = (text || '').trim();
    if (!raw) return null; // toutes les pages
    var allowed = {};
    var parts = raw.split(',');
    var valid = false;
    for (var i = 0; i < parts.length; i++) {
      var chunk = parts[i].trim();
      if (!chunk) continue;
      var m = chunk.match(/^(\d+)\s*-\s*(\d*)$/);
      if (m) {
        var from = parseInt(m[1], 10);
        var to = m[2] ? parseInt(m[2], 10) : total;
        for (var p = Math.max(1, from); p <= Math.min(total, to); p++) { allowed[p] = true; valid = true; }
      } else if (/^\d+$/.test(chunk)) {
        var n = parseInt(chunk, 10);
        if (n >= 1 && n <= total) { allowed[n] = true; valid = true; }
      }
    }
    return valid ? allowed : null;
  }

  var CURRENCIES = [
    { re: /(f\s?cfa|fcfa|xof|cfa)/i, sym: 'FCFA' },
    { re: /(€|eur)/i, sym: '€' },
    { re: /(\$|usd)/i, sym: '$' }
  ];

  /**
   * Convertit une chaîne en nombre si elle en est un (formats FR et EN).
   * Retourne { v, pct, dec, cur } ou null.
   *   « 1 234,56 »     -> 1234.56 (2 décimales)
   *   « (1 200) »      -> -1200   (notation comptable)
   *   « 12,5 % »       -> 0.125 + format pourcentage
   *   « 18 500 FCFA »  -> 18500  + format monétaire
   */
  function parseNumber(raw) {
    var s = String(raw).trim();
    if (!s) return null;

    var currency = null;
    for (var c = 0; c < CURRENCIES.length; c++) {
      if (CURRENCIES[c].re.test(s)) {
        var stripped = s.replace(CURRENCIES[c].re, ' ').trim();
        if (/\d/.test(stripped)) { currency = CURRENCIES[c].sym; s = stripped; }
        break;
      }
    }

    var negative = false;
    if (/^\(.*\)$/.test(s)) { negative = true; s = s.slice(1, -1).trim(); }

    var percent = /%$/.test(s.replace(/\s/g, ''));
    s = s.replace(/%/g, '').trim();

    if (/^-/.test(s)) { negative = true; s = s.slice(1).trim(); }
    if (/^\+/.test(s)) s = s.slice(1).trim();

    // Espaces (normaux, insécables, fins) ou apostrophes = séparateurs de milliers,
    // mais uniquement s'ils forment un vrai groupement par 3 : « 1 234 567,89 ».
    // Sans ce contrôle, « 77 123 45 67 » (téléphone) deviendrait un nombre.
    var SPACE_RE = /[\s   ']/;
    var GROUPED_RE = /^\d{1,3}([\s   ']\d{3})+([.,]\d+)?$/;
    if (SPACE_RE.test(s)) {
      if (!GROUPED_RE.test(s)) return null;
      s = s.replace(/[\s   ']/g, '');
    }

    if (!/^\d[\d.,]*$/.test(s)) return null;

    var lastComma = s.lastIndexOf(',');
    var lastDot = s.lastIndexOf('.');
    var decSep = lastComma > lastDot ? ',' : (lastDot > -1 ? '.' : '');
    var decimals = 0;

    if (decSep) {
      decimals = s.length - s.lastIndexOf(decSep) - 1;
      var occurrences = s.split(decSep).length - 1;
      var others = s.replace(new RegExp('\\' + decSep, 'g'), '');
      // 3 chiffres après un séparateur unique => plutôt un millier (1.234)
      if (decimals === 3 && occurrences === 1 && !/[.,]/.test(others)) {
        s = others;
        decSep = '';
        decimals = 0;
      }
    }

    var cleaned;
    if (decSep === ',') cleaned = s.replace(/\./g, '').replace(',', '.');
    else if (decSep === '.') cleaned = s.replace(/,/g, '');
    else cleaned = s.replace(/[.,]/g, '');

    var value = parseFloat(cleaned);
    if (!isFinite(value)) return null;
    if (negative) value = -value;
    if (percent) value = value / 100;

    return { v: value, pct: percent, dec: Math.min(decimals, 6), cur: currency };
  }

  /** Code de format Excel correspondant au nombre analysé. */
  function numberFormat(num) {
    if (num.pct) return num.dec ? '0.' + repeat('0', Math.min(num.dec, 4)) + '%' : '0%';
    var base = '#,##0' + (num.dec ? '.' + repeat('0', num.dec) : '');
    if (num.cur) return base + ' "' + num.cur + '"';
    return base;
  }

  function repeat(s, n) {
    var out = '';
    for (var i = 0; i < n; i++) out += s;
    return out;
  }

  // ---------------------------------------------------- Matrices PDF
  function matMul(a, b) {
    return [
      a[0] * b[0] + a[1] * b[2],
      a[0] * b[1] + a[1] * b[3],
      a[2] * b[0] + a[3] * b[2],
      a[2] * b[1] + a[3] * b[3],
      a[4] * b[0] + a[5] * b[2] + b[4],
      a[4] * b[1] + a[5] * b[3] + b[5]
    ];
  }

  function matApply(m, x, y) {
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }

  function toHex(v) {
    var n = Math.max(0, Math.min(255, Math.round(v)));
    return (n < 16 ? '0' : '') + n.toString(16);
  }

  function rgbToHex(r, g, b) {
    // pdf.js fournit soit 0-255, soit 0-1 selon l'espace colorimétrique
    if (r <= 1 && g <= 1 && b <= 1) { r *= 255; g *= 255; b *= 255; }
    return (toHex(r) + toHex(g) + toHex(b)).toUpperCase();
  }

  // ================================================================
  //  EXTRACTION : TEXTE
  // ================================================================

  /** Fragments de texte positionnés, origine en haut à gauche. */
  function toFragments(textContent, viewport, page) {
    var frags = [];
    var items = textContent.items || [];
    var fontCache = {};

    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (!it.str || !it.str.trim()) continue;
      var t = it.transform;
      var height = Math.abs(Math.hypot(t[1], t[3])) || it.height || 10;
      var width = it.width || it.str.length * height * 0.5;
      var style = fontStyle(page, textContent.styles, it.fontName, fontCache);
      frags.push({
        x: t[4],
        y: viewport.height - t[5],   // axe Y inversé : 0 = haut de page
        w: width,
        h: height,
        text: it.str,
        bold: style.bold,
        italic: style.italic,
        size: height
      });
    }
    return frags;
  }

  /** Gras / italique déduits du nom de la police embarquée. */
  function fontStyle(page, styles, fontName, cache) {
    if (!fontName) return { bold: false, italic: false };
    if (cache[fontName]) return cache[fontName];

    var name = '';
    try {
      var f = page.commonObjs.get(fontName);
      if (f) name = f.name || f.loadedName || '';
    } catch (e) { /* police pas encore résolue */ }
    if (!name && styles && styles[fontName]) name = styles[fontName].fontFamily || '';

    var res = {
      bold: /bold|black|heavy|semib|demib|[-_]bd\b/i.test(name),
      italic: /italic|oblique/i.test(name)
    };
    cache[fontName] = res;
    return res;
  }

  // ================================================================
  //  EXTRACTION : TRAITS ET APLATS
  // ================================================================

  /**
   * Parcourt les opérateurs de dessin pour récupérer :
   *   - hRules : traits horizontaux { y, x0, x1 }
   *   - vRules : traits verticaux   { x, y0, y1 }
   *   - fills  : aplats de couleur  { x0, y0, x1, y1, color }
   * Tout est ramené dans le même repère que le texte (origine en haut).
   */
  function extractGraphics(opList, viewport) {
    var OPS = pdfjsLib.OPS;
    var H = viewport.height;
    var ctm = [1, 0, 0, 1, 0, 0];
    var stack = [];
    var fillColor = '000000';
    var strokeColor = '000000';
    var hRules = [], vRules = [], fills = [];
    var pending = { segs: [], rects: [] };

    function flip(pt) { return [pt[0], H - pt[1]]; }

    function pushRect(x, y, w, h) {
      var p1 = flip(matApply(ctm, x, y));
      var p2 = flip(matApply(ctm, x + w, y + h));
      pending.rects.push({
        x0: Math.min(p1[0], p2[0]), x1: Math.max(p1[0], p2[0]),
        y0: Math.min(p1[1], p2[1]), y1: Math.max(p1[1], p2[1])
      });
    }

    function pushSeg(x0, y0, x1, y1) {
      var a = flip(matApply(ctm, x0, y0));
      var b = flip(matApply(ctm, x1, y1));
      pending.segs.push({ x0: a[0], y0: a[1], x1: b[0], y1: b[1] });
    }

    /** Classe la géométrie accumulée en traits / aplats. */
    function paint(isFill) {
      var i;
      for (i = 0; i < pending.rects.length; i++) {
        var r = pending.rects[i];
        var w = r.x1 - r.x0, h = r.y1 - r.y0;
        if (h <= RULE_TOL && w >= MIN_RULE) {
          hRules.push({ y: (r.y0 + r.y1) / 2, x0: r.x0, x1: r.x1 });
        } else if (w <= RULE_TOL && h >= MIN_RULE) {
          vRules.push({ x: (r.x0 + r.x1) / 2, y0: r.y0, y1: r.y1 });
        } else if (isFill && w >= MIN_RULE && h >= 4) {
          fills.push({ x0: r.x0, y0: r.y0, x1: r.x1, y1: r.y1, color: fillColor });
        }
      }
      for (i = 0; i < pending.segs.length; i++) {
        var s = pending.segs[i];
        var dx = Math.abs(s.x1 - s.x0), dy = Math.abs(s.y1 - s.y0);
        if (dy <= RULE_TOL && dx >= MIN_RULE) {
          hRules.push({ y: (s.y0 + s.y1) / 2, x0: Math.min(s.x0, s.x1), x1: Math.max(s.x0, s.x1) });
        } else if (dx <= RULE_TOL && dy >= MIN_RULE) {
          vRules.push({ x: (s.x0 + s.x1) / 2, y0: Math.min(s.y0, s.y1), y1: Math.max(s.y0, s.y1) });
        }
      }
      pending = { segs: [], rects: [] };
    }

    var fns = opList.fnArray, args = opList.argsArray;

    for (var k = 0; k < fns.length; k++) {
      var fn = fns[k], a = args[k];
      switch (fn) {
        case OPS.save: stack.push(ctm.slice()); break;
        case OPS.restore: ctm = stack.pop() || [1, 0, 0, 1, 0, 0]; break;
        case OPS.transform: ctm = matMul(a, ctm); break;

        case OPS.setFillRGBColor: fillColor = rgbToHex(a[0], a[1], a[2]); break;
        case OPS.setStrokeRGBColor: strokeColor = rgbToHex(a[0], a[1], a[2]); break;
        case OPS.setFillGray: fillColor = rgbToHex(a[0] * 255, a[0] * 255, a[0] * 255); break;
        case OPS.setStrokeGray: strokeColor = rgbToHex(a[0] * 255, a[0] * 255, a[0] * 255); break;

        case OPS.constructPath: readPath(a); break;
        case OPS.rectangle: pushRect(a[0], a[1], a[2], a[3]); break;

        case OPS.fill:
        case OPS.eoFill:
        case OPS.closeFillStroke:
        case OPS.fillStroke:
        case OPS.eoFillStroke:
          paint(true); break;
        case OPS.stroke:
        case OPS.closeStroke:
          paint(false); break;
        case OPS.endPath:
          pending = { segs: [], rects: [] }; break;
        default: break;
      }
    }
    paint(false);

    /** Décode un chemin composite (constructPath). */
    function readPath(a) {
      var ops = a[0] || [], coords = a[1] || [];
      var p = 0, cx = 0, cy = 0, sx = 0, sy = 0;
      for (var i = 0; i < ops.length; i++) {
        switch (ops[i]) {
          case OPS.moveTo:
            cx = coords[p++]; cy = coords[p++]; sx = cx; sy = cy; break;
          case OPS.lineTo:
            var nx = coords[p++], ny = coords[p++];
            pushSeg(cx, cy, nx, ny); cx = nx; cy = ny; break;
          case OPS.curveTo:
            p += 6; cx = coords[p - 2]; cy = coords[p - 1]; break;
          case OPS.curveTo2:
          case OPS.curveTo3:
            p += 4; cx = coords[p - 2]; cy = coords[p - 1]; break;
          case OPS.closePath:
            pushSeg(cx, cy, sx, sy); cx = sx; cy = sy; break;
          case OPS.rectangle:
            var rx = coords[p++], ry = coords[p++], rw = coords[p++], rh = coords[p++];
            pushRect(rx, ry, rw, rh); cx = rx; cy = ry; sx = rx; sy = ry; break;
          default: break;
        }
      }
    }

    return {
      h: mergeRules(hRules, 'y', 'x0', 'x1'),
      v: mergeRules(vRules, 'x', 'y0', 'y1'),
      fills: fills,
      strokeColor: strokeColor
    };
  }

  /** Fusionne les traits colinéaires qui se recouvrent. */
  function mergeRules(rules, axis, from, to) {
    rules.sort(function (a, b) { return a[axis] - b[axis] || a[from] - b[from]; });
    var out = [];
    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      var prev = out[out.length - 1];
      if (prev && Math.abs(prev[axis] - r[axis]) <= CLUSTER_TOL && r[from] <= prev[to] + 2) {
        prev[to] = Math.max(prev[to], r[to]);
        prev[axis] = (prev[axis] + r[axis]) / 2;
      } else {
        var copy = {};
        copy[axis] = r[axis]; copy[from] = r[from]; copy[to] = r[to];
        out.push(copy);
      }
    }
    return out;
  }

  // ================================================================
  //  REGROUPEMENT EN LIGNES
  // ================================================================

  function groupIntoLines(frags) {
    if (!frags.length) return [];

    var heights = frags.map(function (f) { return f.h; }).sort(function (a, b) { return a - b; });
    var medianH = heights[Math.floor(heights.length / 2)] || 10;
    var tol = Math.max(2, medianH * 0.55);

    var sorted = frags.slice().sort(function (a, b) { return a.y - b.y || a.x - b.x; });
    var lines = [];
    var current = null;

    for (var i = 0; i < sorted.length; i++) {
      var f = sorted[i];
      if (!current || Math.abs(f.y - current.y) > tol) {
        current = { y: f.y, frags: [f] };
        lines.push(current);
      } else {
        current.frags.push(f);
        current.y = (current.y * (current.frags.length - 1) + f.y) / current.frags.length;
      }
    }

    for (var l = 0; l < lines.length; l++) {
      lines[l].frags.sort(function (a, b) { return a.x - b.x; });
      lines[l].frags = mergeFragments(lines[l].frags);
    }
    return lines;
  }

  /** Recolle les fragments quasi collés (pdf.js découpe souvent un mot). */
  function mergeFragments(frags) {
    var out = [];
    for (var i = 0; i < frags.length; i++) {
      var f = frags[i];
      var prev = out[out.length - 1];
      if (prev) {
        var gap = f.x - (prev.x + prev.w);
        var space = Math.max(prev.h, f.h) * 0.32;
        if (gap < space && prev.bold === f.bold && prev.italic === f.italic) {
          var needSpace = gap > space * 0.35 && !/\s$/.test(prev.text) && !/^\s/.test(f.text);
          prev.text += (needSpace ? ' ' : '') + f.text;
          prev.w = (f.x + f.w) - prev.x;
          prev.h = Math.max(prev.h, f.h);
          prev.size = Math.max(prev.size, f.size);
          continue;
        }
      }
      out.push({
        x: f.x, y: f.y, w: f.w, h: f.h, text: f.text,
        bold: f.bold, italic: f.italic, size: f.size
      });
    }
    return out;
  }

  // ================================================================
  //  CONSTRUCTION DES CELLULES
  // ================================================================

  function makeCell() {
    return {
      t: '', frags: [], bold: false, italic: false, size: 0,
      align: null, fill: null, bd: null, span: 1
    };
  }

  /** Applique à la cellule le style de ses fragments. */
  function styleCell(cell, x0, x1) {
    var frags = cell.frags;
    if (!frags.length) return cell;

    var texts = [], bold = true, italic = true, size = 0;
    var minX = Infinity, maxX = -Infinity;

    for (var i = 0; i < frags.length; i++) {
      texts.push(frags[i].text);
      if (!frags[i].bold) bold = false;
      if (!frags[i].italic) italic = false;
      size = Math.max(size, frags[i].size);
      minX = Math.min(minX, frags[i].x);
      maxX = Math.max(maxX, frags[i].x + frags[i].w);
    }

    cell.t = texts.join(' ');
    cell.bold = bold;
    cell.italic = italic;
    cell.size = size;

    // Alignement déduit de la position du texte dans la cellule
    if (x1 > x0) {
      var leftGap = minX - x0;
      var rightGap = x1 - maxX;
      var width = x1 - x0;
      if (leftGap > width * 0.15 && Math.abs(leftGap - rightGap) < width * 0.12) cell.align = 'center';
      else if (rightGap < width * 0.08 && leftGap > width * 0.15) cell.align = 'right';
    }
    return cell;
  }

  /** Couleur de fond d'une zone, si un aplat la recouvre. */
  function fillAt(fills, x, y) {
    for (var i = fills.length - 1; i >= 0; i--) {
      var f = fills[i];
      if (x >= f.x0 && x <= f.x1 && y >= f.y0 && y <= f.y1) {
        if (f.color === 'FFFFFF') return null;   // blanc = pas de fond
        return f.color;
      }
    }
    return null;
  }

  // ---------------------------------------------------- Mode « grille »

  /** Regroupe des positions proches en frontières uniques. */
  function clusterPositions(values) {
    values.sort(function (a, b) { return a - b; });
    var out = [];
    for (var i = 0; i < values.length; i++) {
      if (!out.length || values[i] - out[out.length - 1] > CLUSTER_TOL) out.push(values[i]);
      else out[out.length - 1] = (out[out.length - 1] + values[i]) / 2;
    }
    return out;
  }

  /**
   * Construit une grille à partir des traits du PDF.
   * Retourne null si les traits ne dessinent pas un vrai tableau.
   */
  function buildGrid(graphics) {
    if (graphics.v.length < 2 || graphics.h.length < 3) return null;

    var xs = clusterPositions(graphics.v.map(function (r) { return r.x; }));
    var ys = clusterPositions(graphics.h.map(function (r) { return r.y; }));
    if (xs.length < 2 || ys.length < 3) return null;

    var width = xs[xs.length - 1] - xs[0];
    var height = ys[ys.length - 1] - ys[0];
    if (width < 40 || height < 20) return null;
    if (xs.length - 1 > MAX_COLUMNS) return null;

    return { xs: xs, ys: ys, x0: xs[0], x1: xs[xs.length - 1], y0: ys[0], y1: ys[ys.length - 1] };
  }

  /** Un trait couvre-t-il ce bord de cellule ? */
  function hasRule(rules, axis, pos, from, to, a0, a1) {
    var span = a1 - a0;
    for (var i = 0; i < rules.length; i++) {
      var r = rules[i];
      if (Math.abs(r[axis] - pos) > CLUSTER_TOL) continue;
      var overlap = Math.min(r[to], a1) - Math.max(r[from], a0);
      if (overlap > span * 0.5) return true;
    }
    return false;
  }

  /** Assemble les lignes de texte selon la grille détectée. */
  function latticeRows(lines, grid, graphics) {
    var xs = grid.xs, ys = grid.ys;
    var nCols = xs.length - 1;
    var rows = [];
    var r, c;

    // Une ligne de tableau par bande horizontale
    var bands = [];
    for (r = 0; r < ys.length - 1; r++) {
      var cells = [];
      for (c = 0; c < nCols; c++) cells.push(makeCell());
      bands.push({ top: ys[r], bottom: ys[r + 1], cells: cells, inGrid: true });
    }

    var outside = [];

    for (var l = 0; l < lines.length; l++) {
      var frags = lines[l].frags;
      for (var i = 0; i < frags.length; i++) {
        var f = frags[i];
        var cx = f.x + f.w / 2;
        var cy = f.y - f.h * 0.35;          // milieu approximatif du glyphe
        var band = null;
        for (r = 0; r < bands.length; r++) {
          if (cy >= bands[r].top - CLUSTER_TOL && cy <= bands[r].bottom + CLUSTER_TOL) { band = bands[r]; break; }
        }
        var col = -1;
        if (band && cx >= grid.x0 - CLUSTER_TOL && cx <= grid.x1 + CLUSTER_TOL) {
          for (c = 0; c < nCols; c++) {
            if (cx >= xs[c] - CLUSTER_TOL && cx <= xs[c + 1] + CLUSTER_TOL) { col = c; break; }
          }
        }
        if (band && col >= 0) band.cells[col].frags.push(f);
        else outside.push(f);
      }
    }

    // Styles + bordures des cellules de la grille
    for (r = 0; r < bands.length; r++) {
      var band = bands[r];
      for (c = 0; c < nCols; c++) {
        var cell = band.cells[c];
        styleCell(cell, xs[c], xs[c + 1]);
        cell.bd = {
          t: hasRule(graphics.h, 'y', band.top, 'x0', 'x1', xs[c], xs[c + 1]),
          b: hasRule(graphics.h, 'y', band.bottom, 'x0', 'x1', xs[c], xs[c + 1]),
          l: hasRule(graphics.v, 'x', xs[c], 'y0', 'y1', band.top, band.bottom),
          r: hasRule(graphics.v, 'x', xs[c + 1], 'y0', 'y1', band.top, band.bottom)
        };
        cell.fill = fillAt(graphics.fills, (xs[c] + xs[c + 1]) / 2, (band.top + band.bottom) / 2);
      }
      rows.push({
        cells: band.cells,
        height: band.bottom - band.top,
        y: band.top,
        grid: true
      });
    }

    // Texte hors tableau (titre, mentions) : une ligne pleine largeur
    var outLines = groupIntoLines(outside);
    for (var o = 0; o < outLines.length; o++) {
      var cell2 = makeCell();
      cell2.frags = outLines[o].frags;
      styleCell(cell2, grid.x0, grid.x1);
      cell2.span = nCols;
      var filler = [cell2];
      for (c = 1; c < nCols; c++) filler.push(makeCell());
      rows.push({ cells: filler, height: 0, y: outLines[o].y, grid: false });
    }

    rows.sort(function (a, b) { return a.y - b.y; });
    return { rows: rows, cols: columnWidths(xs) };
  }

  function columnWidths(xs) {
    var cols = [];
    for (var i = 0; i < xs.length - 1; i++) cols.push({ pt: xs[i + 1] - xs[i] });
    return cols;
  }

  // ---------------------------------------------------- Mode « espaces »

  /**
   * Détecte les frontières de colonnes via les gouttières verticales :
   * une bande qu'aucune (ou presque) ligne ne traverse.
   */
  function detectColumnBounds(lines, pageWidth, minGap) {
    var size = Math.ceil(pageWidth) + 2;
    var counts = new Uint16Array(size);
    var used = 0;

    for (var l = 0; l < lines.length; l++) {
      var frags = lines[l].frags;
      if (!frags.length) continue;
      used++;
      for (var i = 0; i < frags.length; i++) {
        var from = Math.max(0, Math.floor(frags[i].x));
        var to = Math.min(size - 1, Math.ceil(frags[i].x + frags[i].w));
        for (var x = from; x <= to; x++) counts[x]++;
      }
    }
    if (!used) return [];

    var tolerance = Math.floor(used * 0.06);
    var bounds = [];
    var runStart = -1;
    for (var p = 0; p < size; p++) {
      if (counts[p] <= tolerance) {
        if (runStart === -1) runStart = p;
      } else if (runStart !== -1) {
        if (runStart > 0 && (p - runStart) >= minGap) bounds.push((runStart + p) / 2);
        runStart = -1;
      }
    }
    return bounds;
  }

  /**
   * Lignes servant à détecter les colonnes : on écarte titres et paragraphes
   * pleine largeur, qui boucheraient les gouttières.
   */
  function detectionLines(lines, pageWidth) {
    var kept = [];
    for (var l = 0; l < lines.length; l++) {
      var frags = lines[l].frags;
      if (frags.length < 2) continue;
      var widest = 0;
      for (var i = 0; i < frags.length; i++) widest = Math.max(widest, frags[i].w);
      if (widest > pageWidth * 0.6) continue;
      kept.push(lines[l]);
    }
    return kept.length >= 3 ? kept : lines;
  }

  function streamRows(lines, pageWidth, graphics, opts) {
    var minGap = opts.minGap;
    var sample = detectionLines(lines, pageWidth);
    var bounds = detectColumnBounds(sample, pageWidth, minGap);

    var guard = 0;
    while (bounds.length + 1 > MAX_COLUMNS && guard < 12) {
      minGap = Math.ceil(minGap * 1.5);
      bounds = detectColumnBounds(sample, pageWidth, minGap);
      guard++;
    }

    // Étendue réelle du contenu, pour des largeurs de colonnes fidèles
    var minX = Infinity, maxX = -Infinity, l, i;
    for (l = 0; l < lines.length; l++) {
      for (i = 0; i < lines[l].frags.length; i++) {
        minX = Math.min(minX, lines[l].frags[i].x);
        maxX = Math.max(maxX, lines[l].frags[i].x + lines[l].frags[i].w);
      }
    }
    if (!isFinite(minX)) { minX = 0; maxX = pageWidth; }

    var edges = [minX].concat(bounds, [maxX]);
    var nCols = bounds.length + 1;
    var rows = [];

    for (l = 0; l < lines.length; l++) {
      var cells = [];
      for (var c = 0; c < nCols; c++) cells.push(makeCell());
      var frags = lines[l].frags;
      var height = 0;

      for (i = 0; i < frags.length; i++) {
        var f = frags[i];
        var center = f.x + f.w / 2;
        var idx = 0;
        while (idx < bounds.length && center > bounds[idx]) idx++;
        cells[idx].frags.push(f);
        height = Math.max(height, f.h);
      }
      for (c = 0; c < nCols; c++) {
        styleCell(cells[c], edges[c], edges[c + 1]);
        cells[c].fill = fillAt(graphics.fills, (edges[c] + edges[c + 1]) / 2, lines[l].y - height * 0.35);
      }
      rows.push({ cells: cells, height: height * 1.6, y: lines[l].y, grid: false });
    }
    return { rows: rows, cols: columnWidths(edges) };
  }

  /** Mode texte brut : une ligne du PDF = une cellule. */
  function rawRows(lines) {
    var rows = [];
    for (var l = 0; l < lines.length; l++) {
      var cell = makeCell();
      cell.frags = lines[l].frags;
      styleCell(cell, 0, 0);
      rows.push({ cells: [cell], height: 0, y: lines[l].y, grid: false });
    }
    return { rows: rows, cols: [{ pt: 400 }] };
  }

  // ================================================================
  //  NETTOYAGE
  // ================================================================

  function rowIsEmpty(row) {
    for (var i = 0; i < row.cells.length; i++) if (row.cells[i].t !== '') return false;
    return true;
  }

  function cleanTable(table, opts) {
    var rows = table.rows, cols = table.cols || [], i, j;

    for (i = 0; i < rows.length; i++) {
      for (j = 0; j < rows[i].cells.length; j++) {
        rows[i].cells[j].t = String(rows[i].cells[j].t || '').replace(/[\s ]+/g, ' ').trim();
      }
    }

    if (opts.dropEmptyRows) {
      rows = rows.filter(function (r) {
        // en mode grille, une ligne vide bordée reste une ligne du tableau
        return !rowIsEmpty(r) || (r.grid && opts.keepGridRows);
      });
    }

    var width = 0;
    for (i = 0; i < rows.length; i++) width = Math.max(width, rows[i].cells.length);
    for (i = 0; i < rows.length; i++) {
      while (rows[i].cells.length < width) rows[i].cells.push(makeCell());
    }
    while (cols.length < width) cols.push({ pt: 60 });
    cols = cols.slice(0, width);

    if (opts.dropEmptyCols && width > 1) {
      var keep = [];
      for (j = 0; j < width; j++) {
        var used = false;
        for (i = 0; i < rows.length; i++) {
          var row = rows[i];
          if (row.cells[j] && row.cells[j].t !== '') { used = true; break; }
          if (row.cells[0] && row.cells[0].span > 1) continue;   // ligne fusionnée
        }
        if (used) keep.push(j);
      }
      if (keep.length && keep.length < width) {
        for (i = 0; i < rows.length; i++) {
          var r2 = rows[i];
          if (r2.cells[0] && r2.cells[0].span > 1) {
            r2.cells = [r2.cells[0]].concat(keep.slice(1).map(makeCell));
            r2.cells[0].span = keep.length;
          } else {
            r2.cells = keep.map(function (k) { return r2.cells[k]; });
          }
        }
        cols = keep.map(function (k) { return cols[k]; });
      }
    }

    return { rows: rows, cols: cols };
  }

  // ================================================================
  //  LECTURE D'UN DOCUMENT
  // ================================================================

  function extractDocument(arrayBuffer, opts, onProgress) {
    var task = pdfjsLib.getDocument({
      data: new Uint8Array(arrayBuffer),
      isEvalSupported: false
    });

    task.onPassword = function (updatePassword, reason) {
      var msg = reason === pdfjsLib.PasswordResponses.INCORRECT_PASSWORD
        ? 'Mot de passe incorrect. Réessayez :'
        : 'Ce PDF est protégé. Entrez son mot de passe :';
      var pwd = window.prompt(msg);
      if (pwd === null) task.destroy();
      else updatePassword(pwd);
    };

    return task.promise.then(function (pdf) {
      var allowed = parsePageRange(opts.pageRange, pdf.numPages);
      var pages = [];
      var totalFrags = 0;
      var gridPages = 0;
      var targets = [];

      for (var n = 1; n <= pdf.numPages; n++) {
        if (!allowed || allowed[n]) targets.push(n);
      }

      var chain = Promise.resolve();
      targets.forEach(function (num, index) {
        chain = chain.then(function () {
          if (onProgress) onProgress(index, targets.length, num);
          return pdf.getPage(num).then(function (page) {
            var viewport = page.getViewport({ scale: 1, rotation: 0 });

            // L'ordre compte : getOperatorList charge les polices utilisées
            // par getTextContent (détection du gras).
            return page.getOperatorList().then(function (opList) {
              var graphics = extractGraphics(opList, viewport);
              return page.getTextContent().then(function (tc) {
                var frags = toFragments(tc, viewport, page);
                totalFrags += frags.length;

                var lines = groupIntoLines(frags);
                if (opts.skipLines > 0) lines = lines.slice(opts.skipLines);

                var table, usedGrid = false;
                if (opts.mode === 'raw') {
                  table = rawRows(lines);
                } else {
                  var grid = (opts.mode === 'grid' || opts.mode === 'auto') ? buildGrid(graphics) : null;
                  if (grid && opts.mode !== 'stream') {
                    table = latticeRows(lines, grid, graphics);
                    usedGrid = true;
                    gridPages++;
                  } else {
                    table = streamRows(lines, viewport.width, graphics, opts);
                  }
                }

                table = cleanTable(table, {
                  dropEmptyRows: opts.dropEmptyRows,
                  dropEmptyCols: opts.dropEmptyCols,
                  keepGridRows: usedGrid
                });

                pages.push({
                  num: num,
                  rows: table.rows,
                  cols: table.cols,
                  grid: usedGrid
                });
                page.cleanup();
              });
            });
          });
        });
      });

      return chain.then(function () {
        if (onProgress) onProgress(targets.length, targets.length, null);
        pdf.destroy();
        pages.sort(function (a, b) { return a.num - b.num; });
        return {
          pages: pages,
          numPages: pdf.numPages,
          totalFragments: totalFrags,
          gridPages: gridPages
        };
      });
    });
  }

  // ================================================================
  //  GÉNÉRATION DU CLASSEUR (ExcelJS)
  // ================================================================

  var BORDER = { style: 'thin', color: { argb: 'FF8A94A6' } };
  var HEADER_FILL = 'FFF1F5F9';

  /**
   * Indice de la ligne d'en-têtes : la première ligne appartenant au tableau.
   * Un titre ou une mention placée au-dessus ne doit pas être prise pour
   * l'en-tête, sinon le figeage et le filtre tombent au mauvais endroit.
   */
  function headerIndexOf(rows, opts) {
    if (!opts.headerRow || !rows.length) return -1;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].grid) return i;
    }
    return 0;
  }

  function ptToWidth(pt) {
    return Math.max(5, Math.min(80, Math.round(pt / PT_PER_CHAR)));
  }

  function argb(hex) { return 'FF' + hex; }

  /** Écrit une page (ou un ensemble de lignes) dans une feuille ExcelJS. */
  function writeSheet(ws, table, opts) {
    var rows = table.rows, cols = table.cols || [];
    var headerIndex = headerIndexOf(rows, opts);
    var i, j;

    for (j = 0; j < cols.length; j++) {
      ws.getColumn(j + 1).width = ptToWidth(cols[j].pt || 60);
    }

    for (i = 0; i < rows.length; i++) {
      var row = rows[i];
      var xlRow = ws.getRow(i + 1);
      if (opts.keepLayout && row.height > 6) {
        xlRow.height = Math.max(14, Math.min(120, Math.round(row.height)));
      }

      for (j = 0; j < row.cells.length; j++) {
        var cell = row.cells[j];
        var xl = xlRow.getCell(j + 1);
        var isHeader = i === headerIndex;
        var num = (opts.detectNumbers && cell.t && !isHeader) ? parseNumber(cell.t) : null;

        if (num) {
          xl.value = num.v;
          xl.numFmt = numberFormat(num);
        } else {
          xl.value = cell.t;
        }

        if (opts.keepLayout) {
          xl.font = {
            name: 'Calibri',
            size: cell.size ? Math.max(7, Math.min(22, Math.round(cell.size))) : 11,
            bold: !!cell.bold || isHeader,
            italic: !!cell.italic
          };
          xl.alignment = {
            horizontal: cell.align || (num ? 'right' : 'left'),
            vertical: 'middle',
            wrapText: true
          };
          if (cell.fill) {
            xl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(cell.fill) } };
          } else if (isHeader) {
            xl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
          }
          if (cell.bd) {
            var b = {};
            if (cell.bd.t) b.top = BORDER;
            if (cell.bd.b) b.bottom = BORDER;
            if (cell.bd.l) b.left = BORDER;
            if (cell.bd.r) b.right = BORDER;
            xl.border = b;
          }
        } else {
          xl.font = { name: 'Calibri', size: 11, bold: isHeader };
          xl.alignment = { horizontal: num ? 'right' : 'left', vertical: 'top' };
          if (isHeader) xl.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_FILL } };
        }
      }

      // Ligne pleine largeur (titre, mention hors tableau)
      if (row.cells[0] && row.cells[0].span > 1 && row.cells.length > 1) {
        try { ws.mergeCells(i + 1, 1, i + 1, Math.min(row.cells[0].span, row.cells.length)); }
        catch (e) { /* fusion impossible : on laisse tel quel */ }
      }
    }

    if (headerIndex >= 0 && rows.length > headerIndex + 1) {
      ws.views = [{ state: 'frozen', ySplit: headerIndex + 1 }];
      ws.autoFilter = {
        from: { row: headerIndex + 1, column: 1 },
        to: { row: rows.length, column: Math.max(1, cols.length) }
      };
    }
  }

  function sameRow(a, b) {
    if (!a || !b || a.cells.length !== b.cells.length) return false;
    for (var i = 0; i < a.cells.length; i++) {
      if (a.cells[i].t !== b.cells[i].t) return false;
    }
    return true;
  }

  function sheetName(base, used) {
    var name = String(base).replace(/[\\\/\?\*\[\]:]/g, '-').slice(0, 31) || 'Feuille';
    var n = 2;
    while (used[name]) {
      name = name.slice(0, 28) + '_' + n;
      n++;
    }
    used[name] = true;
    return name;
  }

  function buildWorkbook(result, opts) {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'Boost Sénégal — Convertisseur PDF vers Excel';
    wb.created = new Date();
    var used = {};

    if (opts.mergePages) {
      var merged = [];
      var cols = [];
      for (var i = 0; i < result.pages.length; i++) {
        var page = result.pages[i];
        if (!page.rows.length) continue;
        if (page.cols.length > cols.length) cols = page.cols;
        var rows = page.rows;
        if (i > 0 && opts.headerRow && merged.length && sameRow(rows[0], merged[0])) {
          rows = rows.slice(1);      // en-tête répété sur chaque page
        }
        merged = merged.concat(rows);
      }
      writeSheet(wb.addWorksheet(sheetName('Données', used)), { rows: merged, cols: cols }, opts);
    } else {
      var added = 0;
      for (var p = 0; p < result.pages.length; p++) {
        var pg = result.pages[p];
        if (!pg.rows.length) continue;
        writeSheet(wb.addWorksheet(sheetName('Page ' + pg.num, used)), pg, opts);
        added++;
      }
      if (!added) {
        wb.addWorksheet(sheetName('Données', used)).getCell('A1').value = '(aucun texte extrait)';
      }
    }
    return wb;
  }

  // ================================================================
  //  INTERFACE
  // ================================================================

  function readOptions() {
    return {
      mode: el.mode.value,
      pageRange: el.pageRange.value,
      minGap: parseInt(el.sensitivity.value, 10),
      skipLines: Math.max(0, parseInt(el.skipLines.value, 10) || 0),
      keepLayout: el.keepLayout.checked,
      headerRow: el.headerRow.checked,
      detectNumbers: el.detectNumbers.checked,
      dropEmptyRows: el.dropEmptyRows.checked,
      dropEmptyCols: el.dropEmptyCols.checked,
      mergePages: el.mergePages.checked
    };
  }

  function activeEntry() {
    for (var i = 0; i < state.files.length; i++) {
      if (state.files[i].id === state.activeId) return state.files[i];
    }
    return null;
  }

  function hasResult() {
    var e = activeEntry();
    if (!e || !e.result) return false;
    for (var i = 0; i < e.result.pages.length; i++) {
      if (e.result.pages[i].rows.length) return true;
    }
    return false;
  }

  function setBusy(on) {
    state.busy = on;
    el.convertBtn.disabled = on || !state.files.length;
    el.exportXlsx.disabled = on || !hasResult();
    el.exportCsv.disabled = on || !hasResult();
    el.exportAll.disabled = on || state.files.length < 2;
  }

  function showProgress(on, text, ratio) {
    el.progress.classList.toggle('is-on', !!on);
    if (text !== undefined) el.progressText.textContent = text;
    if (ratio !== undefined) el.progressFill.style.width = Math.round(ratio * 100) + '%';
  }

  function showAlert(kind, message) {
    el.alerts.innerHTML =
      '<div class="pe-alert pe-alert-' + kind + '">' +
        '<span class="pe-alert-icon">' + (kind === 'error' ? '⛔' : kind === 'warn' ? '⚠️' : 'ℹ️') + '</span>' +
        '<span>' + message + '</span>' +
      '</div>';
  }

  function clearAlert() { el.alerts.innerHTML = ''; }

  function addFiles(fileList) {
    var added = 0;
    for (var i = 0; i < fileList.length; i++) {
      var f = fileList[i];
      if (!/\.pdf$/i.test(f.name) && f.type !== 'application/pdf') continue;
      state.files.push({
        id: 'f' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
        file: f, name: f.name, size: f.size,
        result: null, error: null, status: 'en attente'
      });
      added++;
    }
    if (!added) {
      showAlert('error', 'Aucun fichier PDF valide. Sélectionnez des fichiers .pdf.');
      return;
    }
    clearAlert();
    state.activeId = state.files[state.files.length - added].id;
    state.activePage = 0;
    renderFiles();
    setBusy(false);
    convertActive();
  }

  function removeFile(id) {
    state.files = state.files.filter(function (f) { return f.id !== id; });
    if (state.activeId === id) {
      state.activeId = state.files.length ? state.files[0].id : null;
      state.activePage = 0;
    }
    renderFiles();
    renderPreview();
    setBusy(false);
  }

  function countRows(result) {
    var n = 0;
    for (var i = 0; i < result.pages.length; i++) n += result.pages[i].rows.length;
    return n;
  }

  function renderFiles() {
    if (!state.files.length) { el.fileList.innerHTML = ''; return; }
    var html = '';
    for (var i = 0; i < state.files.length; i++) {
      var f = state.files[i];
      var meta = f.error ? 'Erreur' :
        (f.result ? f.result.pages.length + ' page(s) · ' + countRows(f.result) + ' lignes' : f.status);
      html +=
        '<li class="pe-file' + (f.id === state.activeId ? ' is-active' : '') + '" data-id="' + f.id + '">' +
          '<span class="pe-file-icon">📄</span>' +
          '<span class="pe-file-info">' +
            '<span class="pe-file-name">' + escapeHtml(f.name) + '</span>' +
            '<span class="pe-file-meta">' + formatSize(f.size) + ' · ' + escapeHtml(meta) + '</span>' +
          '</span>' +
          '<button class="pe-file-remove" data-remove="' + f.id + '" title="Retirer">✕</button>' +
        '</li>';
    }
    el.fileList.innerHTML = html;
  }

  // ------------------------------------------------------- Conversion
  function convertActive() {
    var entry = activeEntry();
    if (!entry || state.busy) return;

    var opts = readOptions();
    setBusy(true);
    clearAlert();
    showProgress(true, 'Lecture du fichier…', 0.02);

    entry.file.arrayBuffer()
      .then(function (buf) {
        return extractDocument(buf, opts, function (done, total, pageNum) {
          var label = pageNum
            ? 'Analyse de la page ' + pageNum + ' (' + (done + 1) + '/' + total + ')…'
            : 'Finalisation…';
          showProgress(true, label, total ? done / total : 1);
        });
      })
      .then(function (result) {
        entry.result = result;
        entry.error = null;
        entry.status = 'converti';
        state.activePage = 0;

        if (result.totalFragments === 0) {
          showAlert('warn',
            '<b>Aucune couche de texte détectée.</b> Ce PDF est probablement un scan ou une image. ' +
            'Une reconnaissance optique (OCR) est nécessaire avant la conversion — ' +
            'par exemple via Adobe Acrobat, ou l\'outil gratuit OCRmyPDF.');
        } else if (result.pages.length === 0) {
          showAlert('warn', 'Aucune page ne correspond à la plage indiquée.');
        } else if (result.gridPages === 0 && opts.mode === 'auto') {
          showAlert('info',
            'Aucun trait de tableau détecté : les colonnes ont été déduites des espaces. ' +
            'Ajustez « Écart minimal entre colonnes » si le découpage ne convient pas.');
        }
        showProgress(false);
        renderFiles();
        renderPreview();
        setBusy(false);
        convertPending();
      })
      .catch(function (err) {
        entry.result = null;
        entry.error = err;
        entry.status = 'erreur';
        showProgress(false);
        var msg = (err && err.name === 'PasswordException')
          ? 'PDF protégé : mot de passe requis ou incorrect.'
          : 'Impossible de lire ce PDF : ' + escapeHtml((err && err.message) || String(err));
        showAlert('error', msg);
        renderFiles();
        renderPreview();
        setBusy(false);
        convertPending();
      });
  }

  /**
   * Si un fichier a été déposé pendant qu'une conversion tournait, il est
   * devenu actif sans être analysé : on le traite maintenant.
   */
  function convertPending() {
    var entry = activeEntry();
    if (entry && !entry.result && !entry.error && !state.busy) convertActive();
  }

  // ------------------------------------------------------- Aperçu
  function cellStyleAttr(cell, opts, isNum) {
    if (!opts.keepLayout) return isNum ? ' class="is-num"' : '';
    var css = [];
    if (cell.bold) css.push('font-weight:700');
    if (cell.italic) css.push('font-style:italic');
    if (cell.fill) css.push('background:#' + cell.fill);
    if (cell.align) css.push('text-align:' + cell.align);
    else if (isNum) css.push('text-align:right');
    if (cell.size) css.push('font-size:' + Math.max(9, Math.min(20, cell.size)) + 'px');
    if (cell.bd) {
      var none = '1px solid transparent';
      var line = '1px solid #64748b';
      css.push('border-top:' + (cell.bd.t ? line : none));
      css.push('border-bottom:' + (cell.bd.b ? line : none));
      css.push('border-left:' + (cell.bd.l ? line : none));
      css.push('border-right:' + (cell.bd.r ? line : none));
    }
    return (isNum ? ' class="is-num"' : '') + ' style="' + css.join(';') + '"';
  }

  function renderPreview() {
    var entry = activeEntry();
    var ok = entry && entry.result && entry.result.pages.length;

    if (!ok) {
      el.tabs.innerHTML = '';
      el.tabs.style.display = 'none';
      el.previewBody.innerHTML =
        '<div class="pe-preview-empty">' +
          '<div class="pe-preview-empty-icon">📊</div>' +
          '<h3>Aucun aperçu</h3>' +
          '<p>Déposez un PDF à gauche pour voir le tableau détecté.</p>' +
        '</div>';
      el.previewFoot.style.display = 'none';
      setBusy(state.busy);
      return;
    }

    var pages = entry.result.pages;
    if (state.activePage >= pages.length) state.activePage = 0;

    if (pages.length > 1) {
      var tabs = '';
      for (var i = 0; i < pages.length; i++) {
        tabs += '<button class="pe-tab' + (i === state.activePage ? ' is-active' : '') + '" data-page="' + i + '">' +
                'Page ' + pages[i].num + '</button>';
      }
      el.tabs.innerHTML = tabs;
      el.tabs.style.display = 'flex';
    } else {
      el.tabs.innerHTML = '';
      el.tabs.style.display = 'none';
    }

    var page = pages[state.activePage];
    var rows = page.rows;
    var opts = readOptions();

    if (!rows.length) {
      el.previewBody.innerHTML =
        '<div class="pe-preview-empty"><div class="pe-preview-empty-icon">🗒️</div>' +
        '<h3>Page vide</h3><p>Aucun texte extractible sur cette page.</p></div>';
    } else {
      var colCount = rows[0].cells.length;
      var head = '<thead><tr><th class="pe-rownum">#</th>';
      for (var c = 0; c < colCount; c++) {
        var label = opts.headerRow ? (rows[0].cells[c].t || colLetter(c)) : colLetter(c);
        head += '<th>' + escapeHtml(label) + '</th>';
      }
      head += '</tr></thead>';

      var start = opts.headerRow ? 1 : 0;
      var limit = Math.min(rows.length, start + MAX_PREVIEW_ROWS);
      var body = '<tbody>';
      for (var r = start; r < limit; r++) {
        var row = rows[r];
        body += '<tr><td class="pe-rownum">' + (r + 1) + '</td>';
        var span = row.cells[0] && row.cells[0].span > 1 ? Math.min(row.cells[0].span, colCount) : 1;
        for (var j = 0; j < row.cells.length; j++) {
          if (span > 1 && j > 0) break;
          var cell = row.cells[j];
          var num = opts.detectNumbers && cell.t ? parseNumber(cell.t) : null;
          body += '<td' + (span > 1 ? ' colspan="' + span + '"' : '') +
                  cellStyleAttr(cell, opts, !!num) + '>' + escapeHtml(cell.t) + '</td>';
        }
        body += '</tr>';
      }
      body += '</tbody>';

      var notes = '';
      if (rows.length > limit) {
        notes += '<div class="pe-alert pe-alert-info" style="margin:12px 18px 0">' +
          '<span class="pe-alert-icon">ℹ️</span><span>Aperçu limité aux ' + MAX_PREVIEW_ROWS +
          ' premières lignes. Le fichier Excel contiendra les ' + rows.length + ' lignes.</span></div>';
      }

      el.previewBody.innerHTML = notes +
        '<div class="pe-table-wrap"><table class="pe-table' +
        (opts.keepLayout ? ' pe-table-plain' : '') + '">' + head + body + '</table></div>';
    }

    var nbCols = rows.length ? rows[0].cells.length : 0;
    el.stats.innerHTML =
      '<b>' + rows.length + '</b> ligne' + (rows.length > 1 ? 's' : '') +
      ' · <b>' + nbCols + '</b> colonne' + (nbCols > 1 ? 's' : '') +
      ' · page ' + page.num + ' sur ' + entry.result.numPages +
      ' · <span class="pe-badge">' + (page.grid ? 'grille du PDF' : 'colonnes déduites') + '</span>';
    el.previewFoot.style.display = 'flex';
    setBusy(state.busy);
  }

  // ------------------------------------------------------- Exports
  function saveBlob(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 1500);
  }

  /**
   * CSV au format attendu par Excel en français : séparateur point-virgule
   * et virgule décimale (« 9750,5 » et non « 9750.5 »).
   */
  function tableToCsvFr(rows, opts) {
    var out = [];
    var headerIndex = headerIndexOf(rows, opts);
    for (var i = 0; i < rows.length; i++) {
      var cells = [];
      for (var j = 0; j < rows[i].cells.length; j++) {
        var text = rows[i].cells[j].t;
        var num = (opts.detectNumbers && text && i !== headerIndex) ? parseNumber(text) : null;
        if (num) {
          cells.push(String(num.v).replace('.', ','));
        } else {
          var t = text || '';
          if (/[";\r\n]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
          cells.push(t);
        }
      }
      out.push(cells.join(';'));
    }
    return out.join('\r\n');
  }

  function download(entry, format) {
    if (!entry || !entry.result) return Promise.resolve();
    var opts = readOptions();
    var name = baseName(entry.name);

    if (format === 'csv') {
      var rows = [];
      for (var i = 0; i < entry.result.pages.length; i++) {
        rows = rows.concat(entry.result.pages[i].rows);
      }
      var csv = tableToCsvFr(rows, opts);
      saveBlob(new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' }), name + '.csv');
      return Promise.resolve();
    }

    var wb = buildWorkbook(entry.result, opts);
    return wb.xlsx.writeBuffer().then(function (buffer) {
      saveBlob(new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }), name + '.xlsx');
    });
  }

  function exportActive(format) {
    var entry = activeEntry();
    if (!entry) return;
    setBusy(true);
    showProgress(true, 'Génération du fichier…', 0.6);
    download(entry, format)
      .catch(function (err) {
        showAlert('error', 'Échec de la génération : ' + escapeHtml((err && err.message) || String(err)));
      })
      .then(function () {
        showProgress(false);
        setBusy(false);
      });
  }

  function exportAll() {
    var opts = readOptions();
    var pending = state.files.slice();
    var index = 0;
    setBusy(true);

    function next() {
      if (index >= pending.length) {
        showProgress(false);
        setBusy(false);
        renderFiles();
        return;
      }
      var entry = pending[index++];
      showProgress(true, 'Conversion de ' + entry.name + ' (' + index + '/' + pending.length + ')…',
                   index / pending.length);

      var job = entry.result
        ? Promise.resolve(entry.result)
        : entry.file.arrayBuffer().then(function (buf) { return extractDocument(buf, opts); });

      job.then(function (result) {
        entry.result = result;
        var rows = 0;
        for (var i = 0; i < result.pages.length; i++) rows += result.pages[i].rows.length;
        entry.status = rows ? 'converti' : 'aucun texte';
        return rows ? download(entry, 'xlsx') : null;
      }).catch(function (err) {
        entry.error = err;
        entry.status = 'erreur';
      }).then(function () {
        setTimeout(next, 400);   // laisse le navigateur enchaîner les téléchargements
      });
    }
    next();
  }

  // ================================================================
  //  BRANCHEMENT DES ÉVÉNEMENTS
  // ================================================================

  function bind() {
    el = {
      drop: $('drop-zone'),
      input: $('file-input'),
      fileList: $('file-list'),
      alerts: $('alerts'),
      mode: $('opt-mode'),
      pageRange: $('opt-pages'),
      sensitivity: $('opt-sensitivity'),
      sensitivityVal: $('opt-sensitivity-val'),
      skipLines: $('opt-skip'),
      keepLayout: $('opt-layout'),
      headerRow: $('opt-header'),
      detectNumbers: $('opt-numbers'),
      dropEmptyRows: $('opt-empty-rows'),
      dropEmptyCols: $('opt-empty-cols'),
      mergePages: $('opt-merge'),
      convertBtn: $('btn-convert'),
      exportXlsx: $('btn-xlsx'),
      exportCsv: $('btn-csv'),
      exportAll: $('btn-all'),
      progress: $('progress'),
      progressFill: $('progress-fill'),
      progressText: $('progress-text'),
      tabs: $('page-tabs'),
      previewBody: $('preview-body'),
      previewFoot: $('preview-foot'),
      stats: $('preview-stats')
    };

    // Zone de dépôt
    el.drop.addEventListener('click', function () { el.input.click(); });
    el.drop.addEventListener('keydown', function (e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); el.input.click(); }
    });
    ['dragenter', 'dragover'].forEach(function (evt) {
      el.drop.addEventListener(evt, function (e) {
        e.preventDefault(); e.stopPropagation();
        el.drop.classList.add('is-over');
      });
    });
    ['dragleave', 'drop'].forEach(function (evt) {
      el.drop.addEventListener(evt, function (e) {
        e.preventDefault(); e.stopPropagation();
        el.drop.classList.remove('is-over');
      });
    });
    el.drop.addEventListener('drop', function (e) {
      if (e.dataTransfer && e.dataTransfer.files.length) addFiles(e.dataTransfer.files);
    });
    window.addEventListener('dragover', function (e) { e.preventDefault(); });
    window.addEventListener('drop', function (e) { e.preventDefault(); });

    el.input.addEventListener('change', function () {
      if (el.input.files.length) addFiles(el.input.files);
      el.input.value = '';
    });

    // Liste de fichiers
    el.fileList.addEventListener('click', function (e) {
      var rm = e.target.closest('[data-remove]');
      if (rm) { e.stopPropagation(); removeFile(rm.getAttribute('data-remove')); return; }
      var item = e.target.closest('.pe-file');
      if (!item) return;
      var id = item.getAttribute('data-id');
      if (id === state.activeId) return;
      state.activeId = id;
      state.activePage = 0;
      renderFiles();
      var entry = activeEntry();
      if (entry && !entry.result && !entry.error) {
        convertActive();
      } else {
        renderPreview();
        if (entry && entry.error) {
          showAlert('error', 'Ce fichier n\'a pas pu être lu : ' +
            escapeHtml((entry.error && entry.error.message) || String(entry.error)));
        } else {
          clearAlert();
        }
      }
    });

    // Onglets de page
    el.tabs.addEventListener('click', function (e) {
      var tab = e.target.closest('[data-page]');
      if (!tab) return;
      state.activePage = parseInt(tab.getAttribute('data-page'), 10);
      renderPreview();
    });

    // Options : ré-analyse ou simple re-rendu
    el.mode.addEventListener('change', convertActive);
    el.pageRange.addEventListener('change', convertActive);
    el.skipLines.addEventListener('change', convertActive);
    el.dropEmptyRows.addEventListener('change', convertActive);
    el.dropEmptyCols.addEventListener('change', convertActive);
    el.keepLayout.addEventListener('change', renderPreview);
    el.headerRow.addEventListener('change', renderPreview);
    el.detectNumbers.addEventListener('change', renderPreview);
    el.sensitivity.addEventListener('input', function () {
      el.sensitivityVal.textContent = el.sensitivity.value + ' pt';
    });
    el.sensitivity.addEventListener('change', convertActive);

    el.convertBtn.addEventListener('click', convertActive);
    el.exportXlsx.addEventListener('click', function () { exportActive('xlsx'); });
    el.exportCsv.addEventListener('click', function () { exportActive('csv'); });
    el.exportAll.addEventListener('click', exportAll);

    el.sensitivityVal.textContent = el.sensitivity.value + ' pt';
    renderPreview();
    setBusy(false);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', bind);
  else bind();
})();

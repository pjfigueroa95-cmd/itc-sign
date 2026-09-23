var SignCore = (function () {
  var DEFAULTS = {
    headings: ['client representative', "client's representative", 'principal representative',
               "principal's representative", 'owner representative', "owner's representative",
               'superintendent', 'client approval', 'client'],
    labels: { signature: ['signature', 'signed', 'sign'], name: ['name', 'print name'],
              company: ['company', 'organisation', 'organization'], date: ['date'] }
  };
  function opts(o) {
    o = o || {};
    return { headings: (o.headings || DEFAULTS.headings).map(norm),
             labels: o.labels || DEFAULTS.labels };
  }
  // Lower case, with Word's curly apostrophes made straight.
  function norm(s) { return s.toLowerCase().replace(/[‘’]/g, "'"); }
  function esc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  // Turns pdf.js getTextContent() output into items for findClientBlock.
  function textItems(textContent) {
    return textContent.items.filter(function (i) { return typeof i.str === 'string'; }).map(function (i) {
      var t = i.transform;
      return { str: i.str, x: t[4], y: t[5], w: i.width, h: i.height || Math.sqrt(t[2] * t[2] + t[3] * t[3]) };
    });
  }

  function buildLines(items) {
    var sorted = items.filter(function (i) { return i.str && i.str.trim(); })
      .slice().sort(function (a, b) { return b.y - a.y || a.x - b.x; });
    var lines = [];
    sorted.forEach(function (it) {
      var line = null;
      for (var k = 0; k < lines.length; k++) if (Math.abs(lines[k].y - it.y) <= 3) { line = lines[k]; break; }
      if (!line) { line = { y: it.y, items: [] }; lines.push(line); }
      line.items.push(it);
    });
    lines.forEach(function (l) {
      l.items.sort(function (a, b) { return a.x - b.x; });
      l.text = l.items.map(function (i) { return i.str; }).join(' ').replace(/\s+/g, ' ');
    });
    return lines;
  }

  // Finds the client signature block on one page.
  // items: [{str, x, y, w, h}] in PDF user space. box: {x0, y0, x1, y1}.
  function findClientBlock(items, box, options) {
    var o = opts(options);
    var lines = buildLines(items);
    var anchor = null;
    // Headings are tried in order; the first one found anywhere on the page wins.
    for (var hi = 0; hi < o.headings.length && !anchor; hi++) {
      var head = o.headings[hi], first = head.split(' ')[0];
      lines.forEach(function (l) {
        var t = norm(l.text);
        var at = t.indexOf(head);
        if (at < 0) return;
        if (head === 'client' && !/^\s*client\s*:?\s*$/i.test(l.text)) {
          // bare "client" only counts as a heading when it stands on its own in its column,
          // with no word right beside it (e.g. "Client" + "Company:" as separate items)
          var own = l.items.some(function (i) {
            if (!/^\s*client\s*:?\s*$/i.test(i.str)) return false;
            var gap = Math.max(6, (i.h || 8) * 1.5);
            return !l.items.some(function (n) { return n !== i && n.x + n.w > i.x - gap && n.x < i.x + i.w + gap; });
          });
          if (!own) return;
        }
        var it = l.items.filter(function (i) { return norm(i.str).indexOf(first) >= 0; })[0] || l.items[0];
        var off = norm(it.str).indexOf(first);
        var x = it.x + (off > 0 ? it.w * off / it.str.length : 0);
        if (!anchor || l.y < anchor.y) anchor = { x: x, y: l.y };
      });
    }
    if (!anchor) return null;

    var pageW = box.x1 - box.x0;
    var fieldRight = box.x1 - pageW * 0.05;
    var labels = {};
    Object.keys(o.labels).forEach(function (key) {
      var re = new RegExp('^\\s*(' + o.labels[key].map(esc).join('|') + ')\\s*:?', 'i');
      var best = null;
      items.forEach(function (it) {
        var m = it.str.match(re);
        if (!m) return;
        var rest = it.str.slice(m[0].length);
        if (/^[a-z]/i.test(rest)) return; // e.g. "Signatures", "Named"
        if (it.y >= anchor.y - 1 || it.y < anchor.y - 240) return;
        if (it.x < anchor.x - 45 || it.x > anchor.x + 90) return;
        if (!best || it.y > best.y) {
          var end = it.x + (it.str.length ? it.w * m[0].length / it.str.length : it.w);
          var filled = rest.trim().length > 0 || items.some(function (o) {
            return o !== it && o.str.trim() && Math.abs(o.y - it.y) < 3 && o.x > end + 1 && o.x < fieldRight;
          });
          best = { x: it.x, y: it.y, end: end, size: it.h || 10, filled: filled };
        }
      });
      if (best) labels[key] = best;
    });
    if (!labels.signature) return null;

    var sig = labels.signature;
    var colX = Math.max.apply(null, Object.keys(labels).map(function (k) { return labels[k].end; })) + 12;
    var rowH = labels.name ? sig.y - labels.name.y : 20;
    var h = Math.max(14, Math.min(36, rowH * 1.15, anchor.y - sig.y - 2));
    var sx = colX;
    var result = {
      anchor: anchor,
      sig: { x: sx, y: sig.y - 3, h: h, maxW: Math.max(40, fieldRight - sx) },
      name: null, date: null, sigFilled: sig.filled
    };
    ['name', 'date'].forEach(function (k) {
      var L = labels[k];
      if (L) result[k] = { x: colX, y: L.y, size: Math.max(7, Math.min(14, L.size)), filled: L.filled };
    });
    return result;
  }

  // Works out what to draw on one page.
  function placement(pg, name, date) {
    var box = pg.box;
    if (pg.manual) {
      var h = pg.auto ? pg.auto.sig.h : 26;
      var x = pg.manual.x, y = pg.manual.y - h / 2;
      var texts;
      if (pg.auto) texts = autoTexts(pg.auto, name, date);
      else texts = [{ str: [name, date].filter(Boolean).join('   '), x: x, y: y - 11, size: 9 }];
      return { sig: { x: x, y: y, h: h, maxW: Math.max(40, box.x1 - x - 10) }, texts: texts };
    }
    if (pg.auto) return { sig: pg.auto.sig, texts: autoTexts(pg.auto, name, date) };
    return null;
  }

  function autoTexts(a, name, date) {
    var t = [];
    if (a.name && !a.name.filled && name) t.push({ str: name, x: a.name.x, y: a.name.y, size: a.name.size });
    if (a.date && !a.date.filled && date) t.push({ str: date, x: a.date.x, y: a.date.y, size: a.date.size });
    return t;
  }

  function fitImage(sig, aspect) {
    var h = sig.h, w = h * aspect;
    if (w > sig.maxW) { w = sig.maxW; h = w / aspect; }
    return { x: sig.x, y: sig.y, w: w, h: h };
  }

  // Stamps one PDF. pages: [{index, place}], sigPng: Uint8Array, lib: PDFLib.
  async function stampPdf(lib, bytes, pages, sigPng, meta) {
    var doc = await lib.PDFDocument.load(bytes, { ignoreEncryption: true });
    var png = await doc.embedPng(sigPng);
    var font = await doc.embedFont(lib.StandardFonts.Helvetica);
    var aspect = png.width / png.height;
    pages.forEach(function (p) {
      var page = doc.getPage(p.index);
      var r = fitImage(p.place.sig, aspect);
      page.drawImage(png, { x: r.x, y: r.y, width: r.w, height: r.h });
      p.place.texts.forEach(function (t) {
        page.drawText(t.str, { x: t.x, y: t.y, size: t.size, font: font, color: lib.rgb(0.08, 0.12, 0.35) });
      });
    });
    var note = 'Client signed by ' + meta.name + ' on ' + meta.date;
    var kw = []; try { kw = (doc.getKeywords() || '').split(/[;,]\s*/).filter(Boolean); } catch (e) {}
    doc.setKeywords(kw.concat([note]));
    doc.setModificationDate(new Date());
    return await doc.save();
  }

  var KEYWORD = 'Client signed by';
  var NOTES_HEADING = 'Notes from the client';

  function isSigned(keywords) { return String(keywords || '').indexOf(KEYWORD) >= 0; }

  // Text the standard PDF fonts can draw (WinAnsi): smart quotes and dashes made plain, the rest dropped.
  function plainText(s) {
    return String(s == null ? '' : s)
      .replace(/[‘’‚′]/g, "'").replace(/[“”„″]/g, '"')
      .replace(/[–—−]/g, '-').replace(/…/g, '...').replace(/[  - ]/g, ' ')
      .replace(/[\r\n\t]+/g, ' ')
      .replace(/[^\x20-\x7e¡-ÿ€]/g, '');
  }

  // Why a dropped file is left out, or null to include it.
  // f: {path, error?, signed?, hasText?, hasBlock?}. Called with the path alone before the file is read.
  function leaveOutReason(f) {
    var parts = String(f.path).split(/[\\/]/).filter(Boolean);
    var base = parts[parts.length - 1] || '';
    if (parts.indexOf('__MACOSX') >= 0 || /^\._/.test(base)) return 'System file';
    if (!/\.pdf$/i.test(base)) return 'Not a PDF';
    if (parts.slice(0, -1).some(function (p) { return p.toLowerCase() === 'signed'; })) return 'In a folder called Signed';
    if (f.error) return f.error;
    if (f.signed) return 'Already signed through this page';
    if (f.hasText && !f.hasBlock) return 'No sign-off block found';
    return null;
  }

  // The QCMS entry in a sign pack's PDF info (JSON), or null.
  // {v: 1, project_no, transmittal, pages: [null | {doc_no, rev, page, pages, client: {x, y, w, h}}]}
  function parseQcms(str) {
    if (!str) return null;
    try {
      var o = JSON.parse(str);
      return o && o.v === 1 && Array.isArray(o.pages) ? o : null;
    } catch (e) { return null; }
  }

  // Placement from a QCMS hint: the client box as fractions of the page from the top-left.
  function hintPlacement(hint, box, name, date) {
    var c = hint && hint.client;
    if (!c) return null;
    var W = box.x1 - box.x0, H = box.y1 - box.y0;
    var bx = box.x0 + c.x * W, bw = c.w * W, bh = c.h * H, by = box.y1 - (c.y + c.h) * H;
    var sig = { x: bx + 4, y: by + 2, h: Math.max(10, Math.min(36, bh - 4)), maxW: Math.max(40, bw - 8) };
    return { sig: sig, texts: [{ str: [name, date].filter(Boolean).join('   '), x: sig.x, y: by - 11, size: 9 }] };
  }

  // Tesseract words -> text items in PDF units. The image is the page rendered at `scale` pixels
  // per point, cropped from `top` pixels down. Uses the word's baseline when there is one.
  function ocrItems(words, box, scale, top) {
    return words.filter(function (w) { return w.text && w.text.trim(); }).map(function (w) {
      var b = w.bbox, base = w.baseline && isFinite(w.baseline.y0) ? w.baseline.y0 : b.y1;
      return { str: w.text.trim(), x: box.x0 + b.x0 / scale, y: box.y1 - (top + base) / scale,
               w: (b.x1 - b.x0) / scale, h: (b.y1 - b.y0) / scale };
    });
  }

  // Groups a file's pages per ITC. marks: [{title, page}] (0-based, from bookmarks).
  // Pages before the first bookmark are the cover; no bookmarks -> one group named after the file.
  function groupPages(numPages, marks, fileTitle) {
    var m = (marks || []).filter(function (k) { return k.page >= 0 && k.page < numPages; })
      .sort(function (a, b) { return a.page - b.page; })
      .filter(function (k, i, arr) { return i === 0 || arr[i - 1].page !== k.page; });
    var all = [];
    for (var p = 0; p < numPages; p++) all.push(p);
    if (!m.length) return [{ title: fileTitle, pages: all }];
    var groups = [];
    if (m[0].page > 0) groups.push({ title: 'Cover', cover: true, pages: all.slice(0, m[0].page) });
    m.forEach(function (k, i) {
      groups.push({ title: k.title, pages: all.slice(k.page, i + 1 < m.length ? m[i + 1].page : numPages) });
    });
    return groups;
  }

  // "<original name> - signed.pdf", starting with the project number so intake accepts it.
  function outputName(source, project, date) {
    var base = String(source || '').replace(/\.(pdf|zip)$/i, '').trim();
    if (!base) base = 'ITCs ' + String(date || '').replace(/\//g, '.');
    if (project && base.toLowerCase().indexOf(String(project).toLowerCase()) !== 0) base = project + ' - ' + base;
    return base.replace(/[\\/:*?"<>|]/g, '_') + ' - signed.pdf';
  }

  /* ---------- uploaded signature clean-up (photos or scans of a signature on paper) ---------- */

  // EXIF orientation (1-8, 1 when there is none) of a JPEG, and its stored pixel size (before rotation).
  function jpegInfo(bytes) {
    var b = bytes, out = { orientation: 1, width: 0, height: 0 };
    if (b.length < 4 || b[0] !== 0xFF || b[1] !== 0xD8) return out;
    var i = 2;
    while (i + 4 <= b.length) {
      if (b[i] !== 0xFF) { i++; continue; }
      var m = b[i + 1];
      if (m === 0xFF) { i++; continue; }
      if (m === 0x01 || (m >= 0xD0 && m <= 0xD8)) { i += 2; continue; }
      if (m === 0xDA || m === 0xD9) break;
      var len = (b[i + 2] << 8) | b[i + 3];
      if (m === 0xE1 && b[i + 4] === 0x45 && b[i + 5] === 0x78 && b[i + 6] === 0x69 && b[i + 7] === 0x66) {
        var o = exifOrientation(b, i + 10, Math.min(b.length, i + 2 + len));
        if (o) out.orientation = o;
      }
      if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC) {
        out.height = (b[i + 5] << 8) | b[i + 6];
        out.width = (b[i + 7] << 8) | b[i + 8];
      }
      i += 2 + len;
    }
    return out;
  }
  // Orientation tag from a TIFF header starting at t; 0 when missing or unreadable.
  function exifOrientation(b, t, end) {
    var le = b[t] === 0x49;
    function u16(p) { return le ? b[p] | (b[p + 1] << 8) : (b[p] << 8) | b[p + 1]; }
    function u32(p) { return le ? (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16)) + b[p + 3] * 16777216
                                : b[p] * 16777216 + ((b[p + 1] << 16) | (b[p + 2] << 8) | b[p + 3]); }
    var ifd = t + u32(t + 4);
    if (ifd + 2 > end) return 0;
    var n = u16(ifd);
    for (var k = 0; k < n; k++) {
      var e = ifd + 2 + k * 12;
      if (e + 12 > end) return 0;
      if (u16(e) === 0x0112) { var v = u16(e + 8); return v >= 1 && v <= 8 ? v : 0; }
    }
    return 0;
  }

  // Canvas transform that turns a w x h image with EXIF orientation o upright.
  // Returns { w, h (upright size), m: [a, b, c, d, e, f] for ctx.transform }.
  function orientTransform(o, w, h) {
    switch (o) {
      case 2: return { w: w, h: h, m: [-1, 0, 0, 1, w, 0] };
      case 3: return { w: w, h: h, m: [-1, 0, 0, -1, w, h] };
      case 4: return { w: w, h: h, m: [1, 0, 0, -1, 0, h] };
      case 5: return { w: h, h: w, m: [0, 1, 1, 0, 0, 0] };
      case 6: return { w: h, h: w, m: [0, 1, -1, 0, h, 0] };
      case 7: return { w: h, h: w, m: [0, -1, -1, 0, h, w] };
      case 8: return { w: h, h: w, m: [0, -1, 1, 0, 0, w] };
      default: return { w: w, h: h, m: [1, 0, 0, 1, 0, 0] };
    }
  }

  // Luminance 0-255 of RGBA pixels, with see-through pixels counted as white paper.
  function greyscale(rgba) {
    var n = rgba.length / 4, g = new Uint8Array(n);
    for (var i = 0; i < n; i++) {
      var a = rgba[i * 4 + 3] / 255;
      var l = 0.299 * rgba[i * 4] + 0.587 * rgba[i * 4 + 1] + 0.114 * rgba[i * 4 + 2];
      g[i] = Math.round(l * a + 255 * (1 - a));
    }
    return g;
  }

  // Evens out the lighting: estimates the paper brightness everywhere and divides it out, so shadows
  // and grey paper come out white and the ink keeps its contrast. The estimate is the brightest pixel
  // per small cell (which drops the ink strokes), widened by a cell, then blurred over a large area.
  function evenLighting(grey, w, h) {
    var cell = Math.max(8, Math.round(Math.max(w, h) / 60));
    var cw = Math.ceil(w / cell), ch = Math.ceil(h / cell), x, y;
    var mx = new Float32Array(cw * ch);
    for (y = 0; y < h; y++) for (x = 0; x < w; x++) {
      var c = ((y / cell) | 0) * cw + ((x / cell) | 0), v = grey[y * w + x];
      if (v > mx[c]) mx[c] = v;
    }
    // Widened so strokes thicker than a cell don't darken the estimate.
    var wide = new Float32Array(cw * ch);
    for (var cy = 0; cy < ch; cy++) for (var cx = 0; cx < cw; cx++) {
      var best = 0;
      for (var dy = -1; dy <= 1; dy++) for (var dx = -1; dx <= 1; dx++) {
        var yy = cy + dy, xx = cx + dx;
        if (yy >= 0 && yy < ch && xx >= 0 && xx < cw && mx[yy * cw + xx] > best) best = mx[yy * cw + xx];
      }
      wide[cy * cw + cx] = best;
    }
    var bg = boxBlur(boxBlur(wide, cw, ch, 2), cw, ch, 2);
    var out = new Uint8Array(w * h);
    for (y = 0; y < h; y++) {
      var fy = Math.min(ch - 1, Math.max(0, (y + 0.5) / cell - 0.5)), y0 = fy | 0, y1 = Math.min(ch - 1, y0 + 1), ty = fy - y0;
      for (x = 0; x < w; x++) {
        var fx = Math.min(cw - 1, Math.max(0, (x + 0.5) / cell - 0.5)), x0 = fx | 0, x1 = Math.min(cw - 1, x0 + 1), tx = fx - x0;
        var b = (bg[y0 * cw + x0] * (1 - tx) + bg[y0 * cw + x1] * tx) * (1 - ty) +
                (bg[y1 * cw + x0] * (1 - tx) + bg[y1 * cw + x1] * tx) * ty;
        out[y * w + x] = Math.min(255, Math.round(grey[y * w + x] * 255 / Math.max(b, 1)));
      }
    }
    return out;
  }
  function boxBlur(src, w, h, r) {
    var tmp = new Float32Array(w * h), out = new Float32Array(w * h), x, y, k, s, n;
    for (y = 0; y < h; y++) for (x = 0; x < w; x++) {
      s = 0; n = 0;
      for (k = -r; k <= r; k++) if (x + k >= 0 && x + k < w) { s += src[y * w + x + k]; n++; }
      tmp[y * w + x] = s / n;
    }
    for (y = 0; y < h; y++) for (x = 0; x < w; x++) {
      s = 0; n = 0;
      for (k = -r; k <= r; k++) if (y + k >= 0 && y + k < h) { s += tmp[(y + k) * w + x]; n++; }
      out[y * w + x] = s / n;
    }
    return out;
  }

  // Otsu's threshold for 0-255 values: values below it are one class (ink), the rest the other (paper).
  function otsuThreshold(values) {
    var hist = new Float64Array(256), n = values.length, i;
    for (i = 0; i < n; i++) hist[values[i]]++;
    var sum = 0;
    for (i = 0; i < 256; i++) sum += i * hist[i];
    var sumB = 0, wB = 0, best = -1, t = 128;
    for (i = 0; i < 256; i++) {
      wB += hist[i];
      if (!wB) continue;
      var wF = n - wB;
      if (!wF) break;
      sumB += i * hist[i];
      var mB = sumB / wB, mF = (sum - sumB) / wF, between = wB * wF * (mB - mF) * (mB - mF);
      if (between > best) { best = between; t = i + 1; }
    }
    return t;
  }

  var INK = [20, 32, 90]; // same blue-black as the drawing pen

  // Turns evened-out grey values into the signature image: values at or above threshold are see-through,
  // darker ones are solid ink colour with an alpha by how dark they are, so edges stay smooth. Specks
  // (tiny blobs) are removed. Returns { data (RGBA), box (the ink plus a margin; null when blank) }.
  function inkImage(norm, w, h, threshold, color) {
    color = color || INK;
    var n = w * h, alpha = new Uint8Array(n), i, inkSum = 0, inkN = 0;
    for (i = 0; i < n; i++) if (norm[i] < threshold) { inkSum += norm[i]; inkN++; }
    var ramp = Math.max(24, (threshold - (inkN ? inkSum / inkN : 0)) * 0.8);
    for (i = 0; i < n; i++) {
      var d = threshold - norm[i];
      if (d > 0) alpha[i] = Math.round(255 * Math.min(1, 0.2 + d / ramp));
    }
    removeSpecks(alpha, w, h, Math.max(6, Math.round(n / 60000)));
    var data = new Uint8ClampedArray(n * 4);
    for (i = 0; i < n; i++) {
      if (!alpha[i]) continue;
      data[i * 4] = color[0]; data[i * 4 + 1] = color[1]; data[i * 4 + 2] = color[2]; data[i * 4 + 3] = alpha[i];
    }
    var pad = Math.max(4, Math.round(Math.max(w, h) / 100));
    return { data: data, box: trimBox(data, w, h, pad) };
  }

  // Clears blobs of ink (8-connected) smaller than minArea pixels, in place.
  function removeSpecks(alpha, w, h, minArea) {
    var seen = new Uint8Array(w * h), stack = new Int32Array(w * h), blob = [];
    for (var s = 0; s < w * h; s++) {
      if (!alpha[s] || seen[s]) continue;
      var top = 0; stack[top++] = s; seen[s] = 1; blob.length = 0;
      while (top) {
        var p = stack[--top], px = p % w, py = (p - px) / w;
        blob.push(p);
        for (var dy = -1; dy <= 1; dy++) for (var dx = -1; dx <= 1; dx++) {
          var x = px + dx, y = py + dy, q = y * w + x;
          if (x < 0 || y < 0 || x >= w || y >= h || seen[q] || !alpha[q]) continue;
          seen[q] = 1; stack[top++] = q;
        }
      }
      if (blob.length < minArea) for (var k = 0; k < blob.length; k++) alpha[blob[k]] = 0;
    }
  }

  // The whole clean-up for RGBA pixels (already upright and scaled down). adjust moves the threshold
  // from the automatic one: positive removes more background, negative keeps fainter ink.
  // Returns { data, box, threshold, auto, norm }; pass norm back as prepared to skip the slow steps.
  function cleanSignature(rgba, w, h, adjust, prepared) {
    var norm = prepared || evenLighting(greyscale(rgba), w, h);
    var auto = Math.min(235, Math.max(60, otsuThreshold(norm)));
    var threshold = Math.min(250, Math.max(20, auto - (adjust || 0)));
    var r = inkImage(norm, w, h, threshold);
    return { data: r.data, box: r.box, threshold: threshold, auto: auto, norm: norm };
  }

  // Smallest box around the non-transparent pixels, plus pad; null when there are none.
  function trimBox(data, w, h, pad) {
    var x0 = w, y0 = h, x1 = -1, y1 = -1;
    for (var y = 0; y < h; y++) for (var x = 0; x < w; x++) {
      if (data[(y * w + x) * 4 + 3] > 16) {
        if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
      }
    }
    if (x1 < 0) return null;
    pad = pad || 0;
    x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad);
    x1 = Math.min(w - 1, x1 + pad); y1 = Math.min(h - 1, y1 + pad);
    return { x: x0, y: y0, w: x1 - x0 + 1, h: y1 - y0 + 1 };
  }

  // Top-level bookmarks. entries: [{title, index}] (0-based page). Replaces any outline the doc has.
  function setOutline(lib, doc, entries) {
    if (!entries.length) return;
    var ctx = doc.context, N = lib.PDFName;
    var root = ctx.obj({ Type: 'Outlines', Count: entries.length });
    var rootRef = ctx.register(root);
    var refs = entries.map(function () { return ctx.nextRef(); });
    entries.forEach(function (e, i) {
      var item = ctx.obj({ Title: lib.PDFHexString.fromText(e.title), Parent: rootRef });
      item.set(N.of('Dest'), ctx.obj([doc.getPage(e.index).ref, N.of('Fit')]));
      if (i > 0) item.set(N.of('Prev'), refs[i - 1]);
      if (i < refs.length - 1) item.set(N.of('Next'), refs[i + 1]);
      ctx.assign(refs[i], item);
    });
    root.set(N.of('First'), refs[0]);
    root.set(N.of('Last'), refs[refs.length - 1]);
    doc.catalog.set(N.of('Outlines'), rootRef);
  }

  // Joins separate PDFs into one, with a bookmark per file named after it. files: [{title, bytes}].
  async function combinePdfs(lib, files) {
    var out = await lib.PDFDocument.create();
    var entries = [];
    for (var i = 0; i < files.length; i++) {
      var src = await lib.PDFDocument.load(files[i].bytes, { ignoreEncryption: true, updateMetadata: false });
      var pages = await out.copyPages(src, src.getPageIndices());
      entries.push({ title: files[i].title, index: out.getPageCount() });
      pages.forEach(function (p) { out.addPage(p); });
    }
    setOutline(lib, out, entries);
    return await out.save();
  }

  function wrap(text, font, size, maxW) {
    var words = text.split(' '), lines = [], line = '';
    words.forEach(function (w) {
      var next = line ? line + ' ' + w : w;
      if (line && font.widthOfTextAtSize(next, size) > maxW) { lines.push(line); line = w; } else line = next;
    });
    if (line) lines.push(line);
    return lines;
  }

  // Adds the "Notes from the client" page(s): one "<ITC>: <note>" entry per note. notes: [{label, note}].
  async function addNotesPage(lib, bytes, notes, meta) {
    var doc = await lib.PDFDocument.load(bytes, { ignoreEncryption: true, updateMetadata: false });
    var font = await doc.embedFont(lib.StandardFonts.Helvetica);
    var bold = await doc.embedFont(lib.StandardFonts.HelveticaBold);
    var W = 595.28, H = 841.89, M = 56, size = 10.5, lead = 15;
    var page, y;
    function newPage() {
      page = doc.addPage([W, H]);
      page.drawText(NOTES_HEADING, { x: M, y: H - M - 16, size: 16, font: bold });
      y = H - M - 50;
    }
    newPage();
    notes.forEach(function (n) {
      var lines = wrap(plainText(n.label) + ': ' + plainText(n.note), font, size, W - 2 * M - 14);
      if (y - lines.length * lead < M + 30) newPage();
      lines.forEach(function (l, i) {
        page.drawText(l, { x: M + (i ? 14 : 0), y: y, size: size, font: font });
        y -= lead;
      });
      y -= 6;
    });
    if (y < M + 30) newPage();
    page.drawText(plainText('Client: ' + meta.name + ', ' + meta.date), { x: M, y: y - 10, size: size, font: font });
    return await doc.save();
  }

  return { DEFAULTS: DEFAULTS, KEYWORD: KEYWORD, NOTES_HEADING: NOTES_HEADING,
           textItems: textItems, buildLines: buildLines, findClientBlock: findClientBlock, placement: placement,
           fitImage: fitImage, stampPdf: stampPdf, isSigned: isSigned, plainText: plainText,
           leaveOutReason: leaveOutReason, parseQcms: parseQcms, hintPlacement: hintPlacement, ocrItems: ocrItems,
           groupPages: groupPages, outputName: outputName, trimBox: trimBox,
           jpegInfo: jpegInfo, orientTransform: orientTransform, greyscale: greyscale, evenLighting: evenLighting, otsuThreshold: otsuThreshold,
           inkImage: inkImage, removeSpecks: removeSpecks, cleanSignature: cleanSignature,
           setOutline: setOutline, combinePdfs: combinePdfs, addNotesPage: addNotesPage };
})();
if (typeof module !== 'undefined') module.exports = SignCore;

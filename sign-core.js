var SignCore = (function () {
  var DEFAULTS = {
    headings: ['client representative', "client's representative", 'principal representative',
               "principal's representative", 'owner representative', "owner's representative",
               'superintendent', 'client approval', 'reviewed by client', 'accepted for client',
               'client witness', 'client'],
    // tried only when no heading above is found; a block found this way is optional (left unticked)
    fallback: ['test witnessed by'],
    // headings that can be signed on their own line, e.g. "TEST WITNESSED BY: ________  OF: ______"
    inline: ['test witnessed by', 'client witness'],
    labels: { signature: ['signature', 'signed', 'sign'], name: ['name', 'print name'],
              company: ['company', 'organisation', 'organization'], date: ['date'] },
    // The company (contractor) side, for QCMS's Sign as company (D-00); in order, first found wins.
    company: ['contractor representative', "contractor's representative", 'company representative',
              'accepted for company', 'verified (qc)', 'verified by', 'qc inspector', 'signed by',
              'authorised company official signatory', 'authorized company official signatory', 'company signatory']
  };
  function opts(o) {
    o = o || {};
    return { headings: (o.headings || DEFAULTS.headings).map(norm),
             fallback: (o.fallback || DEFAULTS.fallback).map(norm),
             inline: (o.inline || DEFAULTS.inline).map(norm),
             labels: o.labels || DEFAULTS.labels };
  }
  // Lower case, with Word's curly apostrophes made straight.
  function norm(s) { return s.toLowerCase().replace(/[‘’]/g, "'"); }
  function esc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  // A field with nothing written in it: blank, or only the line printed for writing on (____ or .....).
  function blank(s) { return /^[\s_.\-–—:·…]*$/.test(s); }

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

  // Every place a heading appears: [{x, y, end, size, line}], lowest on the page first.
  // end is where the heading's text stops (x), for signing on the heading's own line.
  function headingAnchors(lines, head) {
    var out = [];
    var re = new RegExp(head.split(' ').map(esc).join('\\s+'), 'g');
    lines.forEach(function (l) {
      // the line's text with where each item starts, so a match maps back to its items
      var t = '', starts = [];
      l.items.forEach(function (i) { if (t) t += ' '; starts.push(t.length); t += norm(i.str); });
      // x at character index c of the line text
      function xAt(c) {
        var k = starts.length - 1;
        while (k > 0 && starts[k] > c) k--;
        var i = l.items[k], s = i.str.length, off = Math.min(s, Math.max(0, c - starts[k]));
        return i.x + (s ? i.w * off / s : 0);
      }
      re.lastIndex = 0;
      var m = re.exec(t);
      if (!m) return;
      if (head === 'client' && !/^\s*client\s*:?\s*$/i.test(l.text)) {
        // bare "client" only counts as a heading when it stands on its own in its column,
        // with no word right beside it (e.g. "Client" + "Company:" as separate items)
        var own = l.items.some(function (i) {
          if (!/^\s*client\s*:?\s*$/i.test(i.str)) return false;
          var gap = Math.max(6, (i.h || 8) * 1.5);
          return !l.items.some(function (n) { return n !== i && n.x + n.w > i.x - gap && n.x < i.x + i.w + gap; });
        });
        if (!own) return;
        var alone = l.items.filter(function (i) { return /^\s*client\s*:?\s*$/i.test(i.str); })[0];
        var lead = alone.str.length - alone.str.replace(/^\s+/, '').length;
        m = { index: starts[l.items.indexOf(alone)] + lead, 0: 'client' };
      }
      // where the heading, and a colon after it, ends
      var e = m.index + m[0].length, j = e;
      while (t[j] === ' ') j++;
      if (t[j] === ':') e = j + 1;
      var k = starts.length - 1;
      while (k > 0 && starts[k] > m.index) k--;
      out.push({ x: xAt(m.index), y: l.y, end: xAt(e), size: l.items[k].h || 9, line: l });
    });
    return out.sort(function (a, b) { return a.y - b.y; });
  }

  // Finds the client signature block on one page.
  // items: [{str, x, y, w, h}] in PDF user space. box: {x0, y0, x1, y1}.
  // Three shapes are read:
  //  - labelled: heading, then Signature / Name / Company / Date labels in any order under it
  //  - boxed: heading with only a "Date :" line under it (e.g. "Reviewed by Client"); signed in the space between
  //  - inline: heading with room to sign on its own line ("Test Witnessed By: ______"), for o.inline headings only
  // Result: {kind, heading, optional, anchor, sig, name, date, sigFilled, area, already}. already says why
  // the sheet looks signed from its text (a name or date written in, something on the signature line).
  function findClientBlock(items, box, options) {
    var o = opts(options);
    var lines = buildLines(items);
    var lists = [{ heads: o.headings, optional: false }, { heads: o.fallback, optional: true }];
    for (var li = 0; li < lists.length; li++) {
      // Headings are tried in order; each place a heading appears is tried, lowest on the page first.
      for (var hi = 0; hi < lists[li].heads.length; hi++) {
        var head = lists[li].heads[hi];
        var anchors = headingAnchors(lines, head);
        for (var ai = 0; ai < anchors.length; ai++) {
          var a = anchors[ai];
          var r = labelledBlock(items, box, o, a) ||
            (head !== 'client' ? boxedBlock(items, box, o, a) : null) ||
            (o.inline.indexOf(head) >= 0 ? inlineBlock(items, box, a) : null);
          if (r) {
            r.heading = head;
            r.optional = lists[li].optional;
            r.anchor = { x: a.x, y: a.y };
            return r;
          }
        }
      }
    }
    return null;
  }

  // The company (contractor) signature block on one page: the same shapes as the client block, with the
  // company headings (options.headings are tried first, e.g. a pack's own words). Null when none is found.
  function findCompanyBlock(items, box, options) {
    options = options || {};
    var heads = (options.headings || []).concat(DEFAULTS.company);
    return findClientBlock(items, box, { headings: heads, fallback: [], inline: ['signed by'], labels: options.labels });
  }

  // Placement in a box given as fractions of the page from the top-left (signature_boxes, a layout map's
  // sign box): the signature in the box, name and date small under it.
  function boxPlacement(c, box, name, date) {
    if (!c) return null;
    var W = box.x1 - box.x0, H = box.y1 - box.y0;
    var bx = box.x0 + c.x * W, bw = c.w * W, bh = c.h * H, by = box.y1 - (c.y + c.h) * H;
    var sig = { x: bx + 4, y: by + 2, h: Math.max(10, Math.min(36, bh - 4)), maxW: Math.max(40, bw - 8) };
    return { sig: sig, texts: [{ str: underText(name, date), x: sig.x, y: by - 11, size: 9, role: 'under' }] };
  }

  // The line under a signature where the sheet has no Name or Date box of its own: name and date, three spaces apart.
  function underText(name, date) {
    return [name, date].filter(Boolean).join('   ');
  }

  // A text of a placement with the piece it is (name, date, company, label, or name and date under the signature), so
  // the page can move or take off each piece on its own (from QC Suite's signing page). Null stays null.
  function withRole(t, role) {
    if (t) t.role = role;
    return t;
  }

  // A layout map sign-off (QCMS's form map; a sign pack's client_blocks hint): {sign, name?, date?, company?},
  // boxes as fractions of the page from the top-left. The signature in its sign box; name, date and company in
  // their own boxes when it has them (each drawn smaller to fit its box). A block with a date box but no name
  // box (e.g. "Verified (QC)" over "Date :") gets the name small at the bottom right inside the sign box, the
  // signature kept clear of it: below the box is that date line. No company: its box is left blank.
  function layoutPlacement(s, box, name, date, company) {
    name = String(name || ''); date = String(date || '');
    var W = box.x1 - box.x0, H = box.y1 - box.y0;
    var inside = !s.name && !!s.date && !!name.trim();
    var p = boxPlacement(s.sign, box, s.name || inside ? '' : name, s.date ? '' : date);
    var texts = p.texts.filter(function (t) { return t.str.trim(); });
    var sig = p.sig;
    if (inside) {
      var bx = box.x0 + s.sign.x * W, bw = s.sign.w * W, by = box.y1 - (s.sign.y + s.sign.h) * H;
      var nameW = Math.min(bw / 2, name.trim().length * INSIDE_NAME_SIZE * 0.52); // Helvetica, about half an em a letter
      texts.push({ str: name.trim(), x: bx + bw - 4 - nameW, y: by + 3, size: INSIDE_NAME_SIZE, role: 'name' });
      sig = { x: sig.x, y: sig.y, h: sig.h, maxW: Math.max(40, sig.maxW - nameW - 6) };
    }
    [withRole(boxText(s.name, box, name), 'name'), withRole(boxText(s.date, box, date), 'date'), withRole(boxText(s.company, box, company), 'company')].forEach(function (t) { if (t) texts.push(t); });
    return { sig: sig, texts: texts };
  }

  // Whether the form prints this company inside a box (fractions of the page from the top-left) already: the words in
  // the box hold the company's name, so it isn't written there again. Only the company itself counts: a label
  // ("Company:"), a line to write on or other words leave the box to be filled. items as textItems gives them.
  function printedIn(items, c, box, company) {
    var want = String(company || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
    if (!c || !want) return false;
    var words = wordsIn(items, c, box).join(' ');
    return (' ' + words.toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim() + ' ').indexOf(' ' + want + ' ') >= 0;
  }
  // The words printed inside a box (fractions of the page from the top-left): each item whose middle is in it.
  function wordsIn(items, c, box) {
    var r = fracRect(c, box);
    return (items || []).filter(function (i) {
      var mx = i.x + (i.w || 0) / 2, my = i.y + (i.h || 0) * 0.35;
      return String(i.str || '').trim() && mx > r.x0 && mx < r.x1 && my > r.y0 && my < r.y1;
    }).map(function (i) { return String(i.str); });
  }

  // The small label under a signature the client adds anywhere (a photo page too), {name}, {date} and {company} filled in.
  var ADDED_LABEL = 'Signed by: {name}, {company}, Client Representative, {date}';
  function addedLabel(template, name, date, company) {
    var t = String(template || '').trim() || ADDED_LABEL;
    var v = { name: String(name || '').trim(), date: String(date || '').trim(), company: String(company || '').trim() };
    return plainText(t.replace(/\{(name|date|company)\}/g, function (all, k) { return v[k]; }))
      .replace(/\s+,/g, ',').replace(/,(\s*,)+/g, ',').replace(/[\s,]+$/, '').replace(/\s+/g, ' ').trim();
  }

  // A placement with the client's own changes to it, piece by piece (the safety net when a box isn't quite right):
  // adj.moves { sig | name | date | company | under | label: { dx, dy } } moves only that piece, adj.off lists the
  // pieces taken off with the ×. With the signature taken off nothing is stamped there: null.
  function adjustPlace(p, adj) {
    if (!p || !adj) return p;
    var off = adj.off || [], moves = adj.moves || {};
    if (off.indexOf('sig') >= 0) return null;
    var m = moves.sig;
    var sig = m ? { x: p.sig.x + m.dx, y: p.sig.y + m.dy, h: p.sig.h, maxW: p.sig.maxW } : p.sig;
    var texts = p.texts.filter(function (t) { return !t.role || off.indexOf(t.role) < 0; }).map(function (t) {
      var d = t.role && moves[t.role];
      if (!d) return t;
      var o = {};
      for (var k in t) o[k] = t[k];
      o.x = t.x + d.dx; o.y = t.y + d.dy;
      return o;
    });
    var out = { sig: sig, texts: texts };
    if (p.frame) out.frame = p.frame;
    return out;
  }

  // Each piece of a placement as a rectangle in PDF units (y up), for the page's drag and ×: the signature at the size
  // it is drawn (fitImage with the picture's shape, or its box before there is one), each text with a role from its
  // baseline at the size it is drawn. width(str, size) measures a text, else about half an em a letter (Helvetica).
  function pieceRects(p, aspect, width) {
    var r = aspect ? fitImage(p.sig, aspect) : { x: p.sig.x, y: p.sig.y, w: Math.min(p.sig.maxW, p.sig.h * 3.5), h: p.sig.h };
    var out = [{ role: 'sig', x0: r.x, y0: r.y, x1: r.x + r.w, y1: r.y + r.h }];
    p.texts.forEach(function (t) {
      if (!t.role) return;
      var size = width ? fitSize(t, width) : t.size;
      var w = width ? width(t.str, size) : t.str.length * size * 0.52;
      out.push({ role: t.role, x0: t.x, y0: t.y - size * 0.25, x1: t.x + w, y1: t.y + size * 0.85 });
    });
    return out;
  }

  // A box given as fractions of the page from the top-left, in PDF units {x0, y0, x1, y1}.
  function fracRect(c, box) {
    var W = box.x1 - box.x0, H = box.y1 - box.y0;
    var x0 = box.x0 + c.x * W, y1 = box.y1 - c.y * H;
    return { x0: x0, y0: y1 - c.h * H, x1: x0 + c.w * W, y1: y1 };
  }

  // One client block of a form with several (a sign pack's client_blocks hint: {heading, optional, sign, ...}),
  // as the text shows it: a block under its own heading whose signature area is near the mapped sign box, so
  // one client block is never taken for another (e.g. "Test Witnessed By" for "Accepted for Client"). Every
  // place the heading appears is tried, and every shape (labelled, boxed, inline) at each. Null when the text
  // doesn't show it there: the map's own boxes are used instead.
  function matchClientBlock(items, box, mapped, options) {
    var head = norm(String((mapped && mapped.heading) || '').replace(/[\s:]+$/, '').trim());
    if (!head || !mapped.sign) return null;
    var o = opts({ headings: [head], fallback: [], inline: DEFAULTS.inline.concat([head]), labels: options && options.labels });
    var r = fracRect(mapped.sign, box), pad = 30;
    var anchors = headingAnchors(buildLines(items), head);
    for (var ai = 0; ai < anchors.length; ai++) {
      var a = anchors[ai];
      var shapes = [labelledBlock(items, box, o, a), head !== 'client' ? boxedBlock(items, box, o, a) : null, inlineBlock(items, box, a)];
      for (var k = 0; k < shapes.length; k++) {
        var b = shapes[k];
        if (!b) continue;
        var s = b.area;
        if (s.x0 < r.x1 + pad && s.x1 > r.x0 - pad && s.y0 < r.y1 + pad && s.y1 > r.y0 - pad) {
          b.heading = head;
          b.optional = !!mapped.optional;
          b.anchor = { x: a.x, y: a.y };
          return b;
        }
      }
    }
    return null;
  }

  /* ---------- checklist rows (digital completion D-01, site mode) ---------- */

  // Column header words, lower case; a pack can add its own (options.headers). pass / fail / na are answer columns.
  var CHECK_HEADERS = {
    item: ['item', 'item no', 'item no.', 'no', 'no.', '#'],
    desc: ['description', 'inspection item', 'inspection activity', 'check', 'activity', 'requirement', 'details'],
    pass: ['completed', 'complete', 'yes', 'pass', 'ok', 'accepted', 'acceptable', 'satisfactory', 'done'],
    fail: ['no', 'fail', 'not ok', 'rejected'],
    na: ['n/a', 'na', 'not applicable'],
    ref: ['reference document / specification', 'reference document', 'reference', 'specification', 'ref doc', 'ref. doc.'],
    sign: ['sign & date', 'sign and date', 'sign/date', 'initials & date', 'initial & date', 'initials', 'sign'],
    remarks: ['remarks', 'comments', 'comment', 'notes']
  };
  var ITEM_NO = /^\d{1,3}(\.\d{1,2})?[a-z]?\.?$/i;

  function headerKind(text, headers) {
    var t = norm(text).replace(/\s+/g, ' ').replace(/\s*:$/, '').trim();
    var order = ['pass', 'na', 'fail', 'sign', 'ref', 'remarks', 'desc', 'item'];
    for (var i = 0; i < order.length; i++) if (headers[order[i]].indexOf(t) >= 0) return order[i];
    return null;
  }

  // Finds the checklist table on one page: its columns (by their header words, wrapped headers joined) and
  // its rows (item numbers in the item column under the header, gaps fine), each with its text and height.
  // items/box as findClientBlock. options.headers: extra words per column kind; options.columns: the columns
  // found on an earlier page, for a table that runs on with no header repeated.
  // Result: {columns: {kind: {x0, x1, cx, label}}, answers: [kind...], headerY, rows: [{item, text, y, top, bottom}],
  //          problem} (problem: why no table could be read, else null). PDF units, y up.
  function findChecklist(items, box, options) {
    options = options || {};
    var headers = {};
    Object.keys(CHECK_HEADERS).forEach(function (k) {
      headers[k] = CHECK_HEADERS[k].concat(((options.headers || {})[k] || []).map(norm));
    });
    var textItems = items.filter(function (i) { return i.str && i.str.trim(); });
    var columns = null, headerY = null, headerBottom = null;
    // header: a line with an answer word (Completed, Yes, N/A ...), its neighbours within 12 pt joined per column
    var anchors = textItems.filter(function (i) { var k = headerKind(i.str, headers); return k === 'pass' || k === 'na'; })
      .sort(function (a, b) { return b.y - a.y; });
    for (var a = 0; a < anchors.length && !columns; a++) {
      var y0 = anchors[a].y;
      var band = textItems.filter(function (i) { return Math.abs(i.y - y0) <= 12; }).sort(function (p, q) { return p.x - q.x; });
      var groups = [];
      band.forEach(function (it) {
        var g = groups[groups.length - 1];
        if (g && it.x < g.x1 + 3) { g.items.push(it); g.x1 = Math.max(g.x1, it.x + it.w); } else groups.push({ x0: it.x, x1: it.x + it.w, items: [it] });
      });
      var found = {}, kinds = [];
      groups.forEach(function (g) {
        var text = g.items.slice().sort(function (p, q) { return q.y - p.y || p.x - q.x; }).map(function (i) { return i.str.trim(); }).join(' ');
        var k = headerKind(text, headers) || headerKind(text.replace(/\s*\/\s*$/, ''), headers);
        if (!k) g.items.forEach(function (i) { k = k || headerKind(i.str, headers); });
        if (k && !found[k]) { found[k] = { x0: g.x0, x1: g.x1, label: text }; kinds.push(k); }
      });
      // one-letter answer headers beside N/A: A (accept) / R (reject), Y / N
      if (found.na) {
        groups.forEach(function (g) {
          var t = norm(g.items.map(function (i) { return i.str.trim(); }).join(' '));
          var k = (t === 'a' || t === 'y') ? 'pass' : (t === 'r' || t === 'n') ? 'fail' : null;
          if (k && !found[k]) { found[k] = { x0: g.x0, x1: g.x1, label: g.items.map(function (i) { return i.str.trim(); }).join(' ') }; kinds.push(k); }
        });
      }
      var answers = ['pass', 'fail', 'na'].filter(function (k) { return found[k]; });
      if (answers.length && kinds.length >= 3) {
        columns = found;
        headerY = Math.max.apply(null, band.map(function (i) { return i.y + (i.h || 8); }));
        headerBottom = Math.min.apply(null, band.map(function (i) { return i.y; }));
      }
    }
    if (!columns && options.columns) { columns = options.columns; headerY = box.y1; headerBottom = box.y1; }
    if (!columns) return { columns: null, answers: [], headerY: null, rows: [], problem: 'No checklist header found (a Completed / N/A or Yes / No / N/A column heading).' };

    // column edges: halfway between neighbouring headers, the first from the item numbers' left, the last to the margin
    var order = Object.keys(columns).sort(function (p, q) { return columns[p].x0 - columns[q].x0; });
    var cols = {};
    order.forEach(function (k, i) {
      var c = columns[k], prev = columns[order[i - 1]], next = columns[order[i + 1]];
      // cx: the header's centre (answer and sign headers sit centred over their cells, so a tick goes there)
      cols[k] = { x0: prev ? (prev.x1 + c.x0) / 2 : Math.max(box.x0, c.x0 - 6), x1: next ? (c.x1 + next.x0) / 2 : Math.max(c.x1 + 6, box.x1 - (box.x1 - box.x0) * 0.05), cx: (c.x0 + c.x1) / 2, label: c.label };
    });
    var firstAnswer = Math.min.apply(null, ['pass', 'fail', 'na'].filter(function (k) { return cols[k]; }).map(function (k) { return cols[k].x0; }));
    var itemCol = cols.item || cols[order[0]];
    var descX0 = cols.desc ? cols.desc.x0 : itemCol.x1;

    // Item numbers printed left of every header (Word forms often head the item and description together, or
    // leave the item column unheaded): they are the item column, and the description starts after them.
    var firstHeadX = columns[order[0]].x0;
    var leftNums = textItems.filter(function (i) {
      return i.y < headerBottom - 2 && ITEM_NO.test(i.str.trim()) && i.x + (i.w || 0) < firstHeadX - 2 && i.x >= box.x0;
    });
    if (leftNums.length >= 2) {
      var numsX0 = Math.min.apply(null, leftNums.map(function (i) { return i.x; }));
      var numsX1 = Math.max.apply(null, leftNums.map(function (i) { return i.x + (i.w || 0); }));
      itemCol = { x0: numsX0 - 2, x1: numsX1 + 2, cx: (numsX0 + numsX1) / 2, label: '' };
      descX0 = numsX1 + 1;
    }
    // rows: item numbers in the item column under the header, top down, until a big gap
    var nums = textItems.filter(function (i) {
      return i.y < headerBottom - 2 && ITEM_NO.test(i.str.trim()) && i.x >= itemCol.x0 - 4 && i.x < Math.max(itemCol.x1, descX0) + 2;
    }).sort(function (p, q) { return q.y - p.y; });
    var rows = [];
    for (var n = 0; n < nums.length; n++) {
      if (rows.length) {
        var gap = rows[rows.length - 1].y - nums[n].y;
        var usual = rows.length > 1 ? (rows[0].y - rows[rows.length - 1].y) / (rows.length - 1) : 30;
        if (gap > Math.max(45, usual * 3)) break;
      }
      rows.push({ item: nums[n].str.trim().replace(/\.$/, ''), y: nums[n].y, h: nums[n].h || 8 });
    }
    if (!rows.length) return { columns: cols, answers: [], headerY: headerY, rows: [], problem: 'A checklist header was found but no item numbers under it.' };
    var step = rows.length > 1 ? (rows[0].y - rows[rows.length - 1].y) / (rows.length - 1) : 20;
    rows.forEach(function (r, i) {
      r.top = i ? (rows[i - 1].y + r.y) / 2 + r.h / 2 : Math.min(headerBottom - 1, r.y + step / 2 + r.h / 2);
      r.bottom = i < rows.length - 1 ? (r.y + rows[i + 1].y) / 2 + rows[i + 1].h / 2 : r.y - step / 2 + r.h / 2;
      var words = textItems.filter(function (t) { return t.y < r.top && t.y >= r.bottom && t.x >= descX0 - 2 && t.x < firstAnswer - 2 && !(ITEM_NO.test(t.str.trim()) && t.x < descX0 + 2); });
      r.text = buildLines(words).sort(function (p, q) { return q.y - p.y; }).map(function (l) { return l.text; }).join(' ').replace(/\s+/g, ' ').trim();
      delete r.h;
    });
    var answers = ['pass', 'fail', 'na'].filter(function (k) { return cols[k]; });
    return { columns: cols, answers: answers, headerY: headerY, rows: rows, problem: null };
  }

  // The contractor (company) sign-off block, for digital completion: the company headings, Contractor
  // Representative first.
  function findContractorBlock(items, box, options) { return findCompanyBlock(items, box, options); }

  // Labels in the heading's column below it: {key: {x, y, end, size, filled}}, nearest to the heading.
  function findLabels(items, box, o, anchor) {
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
          // written in: text after the label, in the same item or beside it (a printed ____ line doesn't count)
          var filled = !blank(rest) || items.some(function (n) {
            return n !== it && !blank(n.str) && Math.abs(n.y - it.y) < 3 && n.x > end + 1 && n.x < fieldRight &&
              !re.test(n.str) && n.x < anchor.x + pageW * 0.45;
          });
          best = { x: it.x, y: it.y, end: end, size: it.h || 10, filled: filled };
        }
      });
      if (best) labels[key] = best;
    });
    return labels;
  }

  function alreadyReason(sigFilled, name, date) {
    if (sigFilled) return 'Something is already on the client signature line';
    if (name && name.filled) return 'The client name is already filled in';
    if (date && date.filled) return 'The client date is already filled in';
    return null;
  }

  function labelledBlock(items, box, o, anchor) {
    var labels = findLabels(items, box, o, anchor);
    if (!labels.signature) return null;
    var pageW = box.x1 - box.x0;
    var fieldRight = box.x1 - pageW * 0.05;
    var sig = labels.signature;
    var colX = Math.max.apply(null, Object.keys(labels).map(function (k) { return labels[k].end; })) + 12;
    // row height: the nearest other label, above or below (labels come in any order)
    var rowH = 20;
    Object.keys(labels).forEach(function (k) {
      if (k !== 'signature') { var d = Math.abs(labels[k].y - sig.y); if (d > 3 && (rowH === 20 || d < rowH)) rowH = d; }
    });
    // room above the signature line: the nearest text above it in the column (a label, a ____ line, the heading)
    var above = anchor.y;
    items.forEach(function (it) {
      if (it.str.trim() && it.y > sig.y + 3 && it.y < above && it.x < fieldRight && it.x + it.w > colX - 12) above = it.y;
    });
    var h = Math.max(14, Math.min(36, rowH * 1.15, above - sig.y - 2));
    var sx = colX;
    var result = {
      kind: 'labelled',
      sig: { x: sx, y: sig.y - 3, h: h, maxW: Math.max(40, fieldRight - sx) },
      name: null, date: null, sigFilled: sig.filled
    };
    ['name', 'date', 'company'].forEach(function (k) {
      var L = labels[k];
      if (L) result[k] = { x: colX, y: L.y, size: Math.max(7, Math.min(14, L.size)), filled: L.filled };
    });
    result.area = { x0: sx - 4, y0: sig.y - 3, x1: sx + Math.min(result.sig.maxW, 220), y1: sig.y - 3 + h };
    result.already = alreadyReason(sig.filled, result.name, result.date);
    return result;
  }

  // The next item to the right of x on the heading's line (the next column's heading or an "OF:" label).
  function rightEdge(box, anchor, fromX) {
    var pageW = box.x1 - box.x0, edge = box.x1 - pageW * 0.05;
    anchor.line.items.forEach(function (i) { if (i.x > fromX + 2 && i.x < edge) edge = i.x; });
    return edge;
  }

  // "Reviewed by Client:" with a "Date :" line under it and no Signature / Name labels: the signature goes in
  // the space between the heading and the Date line, the date on the Date line.
  function boxedBlock(items, box, o, anchor) {
    var labels = findLabels(items, box, o, anchor);
    if (labels.signature || labels.name || !labels.date) return null;
    var d = labels.date;
    var right = rightEdge(box, anchor, anchor.end);
    var top = anchor.y - Math.max(2, anchor.size * 0.3), bottom = d.y + d.size + 2;
    if (top - bottom < 14 || right - anchor.x < 60) return null;
    // anything written in the space means it's been signed already
    var filledIn = items.some(function (it) {
      return !blank(it.str) && it.y > bottom && it.y < top - 1 && it.x >= anchor.x - 4 && it.x < right - 2;
    });
    var sig = { x: anchor.x + 2, y: bottom + 1, h: Math.min(36, top - bottom - 2), maxW: Math.max(40, right - anchor.x - 8) };
    var date = { x: d.end + 6, y: d.y, size: Math.max(7, Math.min(14, d.size)), filled: d.filled };
    return {
      kind: 'boxed', sig: sig, name: null, date: date, sigFilled: filledIn,
      area: { x0: anchor.x, y0: bottom, x1: right - 4, y1: top },
      already: alreadyReason(filledIn, null, date)
    };
  }

  // "TEST WITNESSED BY: __________  OF: ________": the signature on the heading's line, name and date small under it.
  function inlineBlock(items, box, anchor) {
    var line = anchor.line;
    // the next label on the line ("OF:"); written-in text before it means it's filled in
    var right = box.x1 - (box.x1 - box.x0) * 0.05, filledIn = false;
    line.items.forEach(function (i) {
      if (i.x <= anchor.end + 1 || blank(i.str)) return;
      if (/:\s*$/.test(i.str)) { if (i.x < right) right = i.x; } else filledIn = true;
    });
    if (right - anchor.end < 60) return null;
    var above = anchor.y + 40;
    items.forEach(function (it) {
      if (it.str.trim() && it.y > anchor.y + 3 && it.y < above && it.x < right && it.x + it.w > anchor.end) above = it.y;
    });
    var h = Math.max(12, Math.min(24, above - anchor.y - 6));
    var sig = { x: anchor.end + 6, y: anchor.y - 3, h: h, maxW: Math.max(40, right - anchor.end - 14) };
    return {
      kind: 'inline', sig: sig, sigFilled: filledIn,
      name: null, date: null, under: { x: sig.x, y: anchor.y - 9, size: 6.5 },
      area: { x0: sig.x, y0: sig.y, x1: sig.x + sig.maxW, y1: sig.y + h },
      already: alreadyReason(filledIn, null, null)
    };
  }

  // Works out what to draw on one page.
  function placement(pg, name, date, company) {
    var box = pg.box;
    if (pg.manual) {
      var h = pg.auto ? pg.auto.sig.h : 26;
      var x = pg.manual.x, y = pg.manual.y - h / 2;
      var sig = { x: x, y: y, h: h, maxW: Math.max(40, box.x1 - x - 10) };
      if (pg.auto) return withInsideName(sig, autoTexts(pg.auto, name, date, company), pg.auto, name);
      return { sig: sig, texts: [{ str: underText(name, date), x: x, y: y - 11, size: 9, role: 'under' }] };
    }
    if (pg.auto) return withInsideName(pg.auto.sig, autoTexts(pg.auto, name, date, company), pg.auto, name);
    return null;
  }

  var INSIDE_NAME_SIZE = 7.5;

  // A block with no Name line of its own (a boxed "Reviewed by Client:" over "Date :"): the name small at
  // the bottom right inside the signature area, the signature kept clear of it, so the sheet still says who.
  function withInsideName(sig, texts, a, name) {
    name = String(name || '').trim();
    if (!name || a.name || a.under) return { sig: sig, texts: texts };
    var w = Math.min(sig.maxW / 2, name.length * INSIDE_NAME_SIZE * 0.52); // Helvetica, about half an em a letter
    return {
      sig: { x: sig.x, y: sig.y, h: sig.h, maxW: Math.max(40, sig.maxW - w - 6) },
      texts: texts.concat([{ str: name, x: sig.x + sig.maxW - w, y: sig.y + 1, size: INSIDE_NAME_SIZE, role: 'name' }]),
    };
  }

  function autoTexts(a, name, date, company) {
    var t = [];
    // (each with its piece, so the page moves or takes off each on its own)
    if (a.name && !a.name.filled && name) t.push({ str: name, x: a.name.x, y: a.name.y, size: a.name.size, role: 'name' });
    if (a.date && !a.date.filled && date) t.push({ str: date, x: a.date.x, y: a.date.y, size: a.date.size, role: 'date' });
    // the block's own Company line, when the caller gives a company (the client's signing page) and nothing is there
    if (a.company && !a.company.filled && company) t.push({ str: plainText(company), x: a.company.x, y: a.company.y, size: a.company.size, role: 'company' });
    // inline blocks have no Name or Date line: both go small under the signature
    if (a.under && (name || date)) t.push({ str: underText(name, date), x: a.under.x, y: a.under.y, size: a.under.size, role: 'under' });
    return t;
  }

  /* ---------- sheets with no client block: an added "Reviewed by Client" stamp ---------- */

  var STAMP = { w: 190, h: 64, label: 'Reviewed by Client' };
  // tried when the full-size stamp doesn't fit on a busy page
  var STAMP_SMALL = { w: 150, h: 48 };

  // What to draw for an added client sign-off whose frame's bottom-left corner is at (x, y):
  // a thin frame, the label, the signature, then name and date on one line.
  // at may carry its own size {w, h} (a smaller stamp from emptySpot); everything inside scales with it.
  function stampPlacement(at, box, name, date) {
    var w = at.w || STAMP.w, h = at.h || STAMP.h, k = h / STAMP.h;
    var x = Math.max(box.x0 + 4, Math.min(box.x1 - w - 4, at.x));
    var y = Math.max(box.y0 + 4, Math.min(box.y1 - h - 4, at.y));
    var texts = [{ str: STAMP.label, x: x + 5 * k, y: y + h - 11 * k, size: 7.5 * k, bold: true }];
    if (name) texts.push({ str: 'Name: ' + name, x: x + 5 * k, y: y + 6 * k, size: 7 * k });
    if (date) texts.push({ str: 'Date: ' + date, x: x + w - 62 * k, y: y + 6 * k, size: 7 * k });
    return { sig: { x: x + 6 * k, y: y + 16 * k, h: 30 * k, maxW: w - 12 * k }, texts: texts, frame: { x: x, y: y, w: w, h: h } };
  }

  // Luminance per pixel (0-255) of RGBA data, with transparent pixels as white.
  function lum(rgba, i) {
    var a = rgba[i + 3] / 255;
    return (0.299 * rgba[i] + 0.587 * rgba[i + 1] + 0.114 * rgba[i + 2]) * a + 255 * (1 - a);
  }

  // Finds empty space for the stamp near the bottom of the page, above the footer.
  // rgba: the page rendered at `scale` pixels per point (w x h pixels). items: its text items (for the footer).
  // Returns the frame's bottom-left corner and size {x, y, w, h} in PDF units: full size if it fits anywhere,
  // else the smaller stamp; null when neither fits.
  // Prefers the lowest spot, then the one furthest right (where client sign-offs usually sit).
  function emptySpot(rgba, w, h, scale, box, items) {
    var H = box.y1 - box.y0;
    // the footer: text in the bottom 10% of the page; the stamp goes above it
    var floor = box.y0 + 24;
    (items || []).forEach(function (it) {
      if (it.str.trim() && it.y < box.y0 + H * 0.1) floor = Math.max(floor, it.y + (it.h || 8) + 6);
    });
    var sat = new Uint32Array((w + 1) * (h + 1)); // summed-area table of dark pixels
    for (var y = 0; y < h; y++) {
      var run = 0;
      for (var x = 0; x < w; x++) {
        if (lum(rgba, (y * w + x) * 4) < 200) run++;
        sat[(y + 1) * (w + 1) + x + 1] = sat[y * (w + 1) + x + 1] + run;
      }
    }
    function dark(x0, y0, x1, y1) { // pixel rect, inclusive-exclusive
      return sat[y1 * (w + 1) + x1] - sat[y0 * (w + 1) + x1] - sat[y1 * (w + 1) + x0] + sat[y0 * (w + 1) + x0];
    }
    function search(size, pad) {
      var sw = size.w + 2 * pad, sh = size.h + 2 * pad;
      var pw = Math.ceil(sw * scale), ph = Math.ceil(sh * scale);
      var step = 4, mx0 = Math.round(18 * scale), mx1 = w - Math.round(18 * scale);
      // from the floor up to 40% of the way up the page ("near the bottom")
      for (var by = floor; by < box.y0 + H * 0.4; by += step) {
        var py1 = Math.round((box.y1 - (by - pad)) * scale), py0 = py1 - ph;
        if (py0 < 0 || py1 > h) continue;
        for (var bx = box.x1 - 24 - sw; bx >= box.x0 + 24; bx -= step * 2) {
          var px0 = Math.round((bx - box.x0) * scale), px1 = px0 + pw;
          if (px0 < 0 || px1 > w) continue;
          if (dark(px0, py0, px1, py1) !== 0) continue;
          // an empty table cell is a field on the form, not free space: skip spots with something beside them on
          // both sides (out to the page margins); real free space near the bottom of a form is open on one side
          var enclosed = dark(Math.min(mx0, px0), py0, px0, py1) > 0 && dark(px1, py0, Math.max(mx1, px1), py1) > 0;
          if (!enclosed) return { x: bx + pad, y: by, w: size.w, h: size.h };
        }
      }
      return null;
    }
    return search(STAMP, 6) || search(STAMP_SMALL, 3);
  }

  // Ink or an image in a signature area, e.g. a signature added in another PDF editor.
  // rgba: the area rendered (w x h pixels). skip: rects [{x0, y0, x1, y1}] in pixels to ignore (known text).
  // Ruled lines are ignored: long straight runs of dark pixels (cell borders, a ____ line), and the pixels
  // right beside them (anti-aliasing), wherever they start and stop in the area.
  function inkInBox(rgba, w, h, skip) {
    var dark = new Uint8Array(w * h);
    for (var i = 0; i < w * h; i++) if (lum(rgba, i * 4) < 160) dark[i] = 1;
    (skip || []).forEach(function (r) {
      for (var y2 = Math.max(0, Math.floor(r.y0)); y2 < Math.min(h, Math.ceil(r.y1)); y2++)
        for (var x2 = Math.max(0, Math.floor(r.x0)); x2 < Math.min(w, Math.ceil(r.x1)); x2++) dark[y2 * w + x2] = 0;
    });
    var line = new Uint8Array(w * h);
    var hMin = Math.max(24, Math.round(w * 0.2)), vMin = Math.max(16, Math.round(h * 0.5));
    function mark(i) { line[i] = 1; }
    for (var y = 0; y < h; y++) {
      for (var x = 0; x < w;) {
        if (!dark[y * w + x]) { x++; continue; }
        var s = x; while (x < w && dark[y * w + x]) x++;
        if (x - s >= hMin) for (var k = s; k < x; k++) { mark(y * w + k); if (y > 0) mark((y - 1) * w + k); if (y < h - 1) mark((y + 1) * w + k); }
      }
    }
    for (var cx = 0; cx < w; cx++) {
      for (var cy = 0; cy < h;) {
        if (!dark[cy * w + cx]) { cy++; continue; }
        var st = cy; while (cy < h && dark[cy * w + cx]) cy++;
        if (cy - st >= vMin) for (var j = st; j < cy; j++) { mark(j * w + cx); if (cx > 0) mark(j * w + cx - 1); if (cx < w - 1) mark(j * w + cx + 1); }
      }
    }
    var count = 0;
    for (var p = 0; p < w * h; p++) if (dark[p] && !line[p]) count++;
    var ratio = count / Math.max(1, w * h);
    return { count: count, ratio: ratio, inked: count >= 40 && ratio >= 0.004 };
  }

  /* ---------- client blocks found on a picture (a scan or image-only sheet: no words, no hints) ---------- */
  // Ported by hand from QC Suite's signing page (its sign-core's alignRows, inkIn): a sign-off block is a heading's
  // tinted band ("Contractor Representative | Client Representative") over rows, each row's foot a ruled line (the
  // line written on, or a cell's border). The page draws the sheet (look), finds the bands and ruled lines here
  // (pictureRows), reads only the headings and labels with OCR, and pictureClientBlock picks the client's column and
  // gives each piece its own box (sign, name, date, company) as a layout map does.
  //
  // look: the page drawn for looking at: { data (RGBA), w, h, scale (pixels a point), x0, y1 (its top-left corner in
  // the page's units, y up) }. Rectangles below are in the page's units, y up.

  var DARK = 160;
  // Runs along one pixel row where test(lum) holds, gaps of up to `gap` pixels bridged: [[start, end), ...]
  function rowRuns(look, y, test, gap) {
    var out = [], s = -1, last = -1;
    for (var x = 0; x <= look.w; x++) {
      var on = x < look.w && test(lum(look.data, (y * look.w + x) * 4));
      if (on) { if (s < 0) s = x; last = x; }
      else if (s >= 0 && x - last > gap) { out.push([s, last + 1]); s = -1; }
    }
    return out;
  }
  var px2x = function (look, px) { return look.x0 + px / look.scale; };
  var py2y = function (look, py) { return look.y1 - py / look.scale; };

  // The paper's own brightness: the commonest brightness of the light pixels (white paper about 255; a grey scan's
  // paper darker, about 215 to 235), so a tint is told from the paper of the sheet it is on.
  function paperLevel(look) {
    var hist = new Uint32Array(256), n = look.w * look.h;
    for (var i = 0; i < n; i += 3) hist[Math.round(lum(look.data, i * 4))]++;
    var best = 255;
    for (var v = 150; v < 256; v++) if (hist[v] > hist[best]) best = v;
    return best;
  }

  // Heading bands: rows of a light tint (a colour, or the grey a scan makes of it; darker than the paper) at least a
  // fifth of the page wide, 6 to 40 points tall. [{ top, bottom, x0, x1 }]
  function headingBands(look) {
    var minW = look.w * 0.2, rows = [], paper = Math.min(245, paperLevel(look) - 9);
    for (var y = 0; y < look.h; y++) {
      var seg = null;
      rowRuns(look, y, function (l) { return l < paper; }, 2).forEach(function (r) {
        if (r[1] - r[0] < minW) return;
        var n = 0;
        for (var x = r[0]; x < r[1]; x++) if (lum(look.data, (y * look.w + x) * 4) >= DARK) n++;
        if (n >= 0.6 * (r[1] - r[0]) && (!seg || r[1] - r[0] > seg[1] - seg[0])) seg = r;
      });
      rows.push(seg);
    }
    var bands = [], cur = null;
    for (var yy = 0; yy <= look.h; yy++) {
      var s = yy < look.h ? rows[yy] : null;
      if (s && cur && s[0] < cur.c1 && s[1] > cur.c0) { cur.end = yy; cur.c0 = Math.min(cur.c0, s[0]); cur.c1 = Math.max(cur.c1, s[1]); continue; }
      if (cur) {
        var hpt = (cur.end - cur.start + 1) / look.scale;
        if (hpt >= 6 && hpt <= 40) bands.push({ top: py2y(look, cur.start), bottom: py2y(look, cur.end + 1), x0: px2x(look, cur.c0), x1: px2x(look, cur.c1) });
      }
      cur = s ? { start: yy, end: yy, c0: s[0], c1: s[1] } : null;
    }
    return bands;
  }

  // Ruled lines: dark runs at least 40 points long and at most 4 points thick (runs on neighbouring pixel rows that
  // overlap are one line, so a scan turned a little still gives one). [{ y, x0, x1 }]
  function ruledLines(look) {
    var minW = 40 * look.scale, open = [], done = [];
    for (var y = 0; y <= look.h; y++) {
      var runs = y < look.h ? rowRuns(look, y, function (l) { return l < DARK; }, 2).filter(function (r) { return r[1] - r[0] >= minW; }) : [];
      var next = [];
      runs.forEach(function (r) {
        var l = open.filter(function (o) { return o.last >= y - 2 && r[0] < o.c1 && r[1] > o.c0; })[0];
        if (l) { l.last = y; l.c0 = Math.min(l.c0, r[0]); l.c1 = Math.max(l.c1, r[1]); l.ys += y; l.n++; }
        else { l = { first: y, last: y, c0: r[0], c1: r[1], ys: y, n: 1 }; open.push(l); }
        if (next.indexOf(l) < 0) next.push(l);
      });
      open = open.filter(function (o) {
        if (o.last >= y - 2 && y < look.h) return true;
        done.push(o);
        return false;
      });
    }
    return done.filter(function (o) { return (o.last - o.first + 1) / look.scale <= 4; })
      .map(function (o) { return { y: py2y(look, o.ys / o.n + 0.5), x0: px2x(look, o.c0), x1: px2x(look, o.c1) }; });
  }

  // Where a row's label ends (the page's units): a label cell's border (dark down most of the row, in the left half,
  // past the row's own left border), else the first stretch of 12 points with nothing dark after something dark (the
  // label's words); the row's left when nothing dark is in its first 40 per cent.
  function labelEnd(look, x0, x1, y0, y1) {
    var c0 = Math.max(0, Math.round((x0 - look.x0) * look.scale)), c1 = Math.min(look.w, Math.round((x1 - look.x0) * look.scale));
    var r0 = Math.max(0, Math.round((look.y1 - y1) * look.scale) + 2), r1 = Math.min(look.h, Math.round((look.y1 - y0) * look.scale) - 2);
    if (r1 - r0 < 3 || c1 - c0 < 8) return x0;
    var n = [];
    for (var x = c0; x < c1; x++) {
      var k = 0;
      for (var y = r0; y < r1; y++) if (lum(look.data, (y * look.w + x) * 4) < DARK) k++;
      n.push(k);
    }
    var border = function (i) { return n[i] >= 0.7 * (r1 - r0); };
    for (var i = Math.round(6 * look.scale); i < n.length / 2; i++) if (border(i)) return px2x(look, c0 + i + 2);
    var clear = 12 * look.scale, seen = false, gap = 0;
    for (var j = 0; j < n.length; j++) {
      if (n[j] && !border(j)) { seen = true; gap = 0; }
      else if (seen && ++gap >= clear) return px2x(look, c0 + j - gap + 3 * look.scale);
      if (!seen && j > 0.4 * n.length) return x0;
    }
    return x0;
  }

  // A column's parties side by side, [{ x0, x1 }] (the page's units): the cell borders inside it (dark down at least 70
  // per cent of every row, away from its ends) are a label's and its value's; three or more (an odd count: label | value
  // | label | value) split it at every second one. Otherwise the column is one party.
  function partyCuts(look, x0, x1, spans) {
    var whole = [{ x0: x0, x1: x1 }];
    if (spans.length < 2) return whole;
    var c0 = Math.max(0, Math.round((x0 - look.x0 + 4) * look.scale)), c1 = Math.min(look.w, Math.round((x1 - look.x0 - 4) * look.scale));
    var rows = spans.map(function (sp) {
      return [Math.max(0, Math.round((look.y1 - sp.top) * look.scale) + 2), Math.min(look.h, Math.round((look.y1 - sp.foot) * look.scale) - 2)];
    }).filter(function (r) { return r[1] - r[0] >= 3; });
    var xs = [], run = null;
    for (var x = c0; x < c1; x++) {
      var down = rows.length > 0 && rows.every(function (r) {
        var k = 0;
        for (var y = r[0]; y < r[1]; y++) if (lum(look.data, (y * look.w + x) * 4) < DARK) k++;
        return k >= 0.7 * (r[1] - r[0]);
      });
      if (down) run = run ? [run[0], x] : [x, x];
      else if (run) { xs.push(px2x(look, (run[0] + run[1]) / 2)); run = null; }
    }
    if (run) xs.push(px2x(look, (run[0] + run[1]) / 2));
    if (xs.length < 3 || xs.length % 2 === 0) return whole;
    var out = [], from = x0;
    for (var i = 1; i < xs.length; i += 2) { out.push({ x0: from, x1: xs[i] - 1, label: xs[i - 1] }); from = xs[i] + 1; }
    out.push({ x0: from, x1: x1, label: xs[xs.length - 1] });
    // only parties alike: widths within a quarter of each other, each label cell narrower than its value cell (a
    // checklist's columns, item | description | done | remarks, are not)
    var ws = out.map(function (o) { return o.x1 - o.x0; });
    var alike = Math.min.apply(null, ws) >= 0.75 * Math.max.apply(null, ws) &&
      out.every(function (o) { return o.label - o.x0 < o.x1 - o.label; });
    return alike ? out.map(function (o) { return { x0: o.x0, x1: o.x1 }; }) : whole;
  }

  // Ready for OCR: a strip of the sheet in plain black and white (a tint, a scan's grey paper gone) with its ruled
  // lines taken out (long dark runs across or down), so a heading in a bordered, tinted cell reads like plain words.
  // rgba: the strip drawn at `scale` pixels a point (w x h); changed in place.
  function cleanForOcr(rgba, w, h, scale) {
    var dark = new Uint8Array(w * h);
    for (var i = 0; i < w * h; i++) if (lum(rgba, i * 4) < DARK) dark[i] = 1;
    var keep = dark.slice(), across = Math.round(30 * scale), down = Math.round(14 * scale);
    for (var y = 0; y < h; y++) for (var x = 0; x < w;) {
      if (!dark[y * w + x]) { x++; continue; }
      var s = x; while (x < w && dark[y * w + x]) x++;
      if (x - s >= across) for (var k = s; k < x; k++) keep[y * w + k] = 0;
    }
    for (var cx = 0; cx < w; cx++) for (var cy = 0; cy < h;) {
      if (!dark[cy * w + cx]) { cy++; continue; }
      var t = cy; while (cy < h && dark[cy * w + cx]) cy++;
      if (cy - t >= down) for (var j = t; j < cy; j++) keep[j * w + cx] = 0;
    }
    for (var p = 0; p < w * h; p++) {
      var v = keep[p] ? 0 : 255;
      rgba[p * 4] = rgba[p * 4 + 1] = rgba[p * 4 + 2] = v; rgba[p * 4 + 3] = 255;
    }
    return rgba;
  }

  // Whether a page has words to read (a digital sheet, or a scan made searchable): any words but a label's. A scanned
  // sheet that carries only its ITC number label as words ("<ITC no.> Rev 0", beside its QR) is still a picture.
  function hasWords(items) {
    var words = (items || []).map(function (i) { return String(i.str || ''); }).join(' ').split(/\s+/).filter(function (w) {
      // a document number (letters and digits in two or more parts, a digit in it), "Rev", a number
      return w && !(/^[a-z0-9_.]+(-[a-z0-9_.]+)+$/i.test(w) && /\d/.test(w)) && !/^rev\.?$/i.test(w) && !/^[\d.\-\/]+$/.test(w);
    });
    return words.join('').length > 3;
  }

  // The sign-off blocks drawn on a sheet: each heading band with the ruled rows under it (within 160 points, above the
  // next band), in columns side by side. [{ band, cols: [{ head, x0, x1, rows: [{ box, label }] }] }], rectangles
  // { x0, y0, x1, y1 } in the page's units. A band with no rows under it is not a block.
  function pictureRows(look) {
    if (!look || !look.data) return [];
    var bands = headingBands(look), lines = ruledLines(look), out = [];
    bands.forEach(function (b) {
      var below = bands.filter(function (o) { return o !== b && o.top < b.bottom - 2; }).reduce(function (m, o) { return Math.max(m, o.top); }, -Infinity);
      // (the band's own ruled border just under it is left out: running the band's width, it would join its columns)
      var mine = lines.filter(function (l) {
        return l.y < b.bottom - 6 && l.y > b.bottom - 160 && l.y > below + 3 && l.x0 >= b.x0 - 6 && l.x1 <= b.x1 + 6;
      });
      // columns: lines whose spans overlap by most of the shorter one, the shorter lines first; a line running across
      // two columns (a footer's rule, a box round the block) joins neither
      var cols = [];
      mine.sort(function (p, q) { return (p.x1 - p.x0) - (q.x1 - q.x0) || p.x0 - q.x0; }).forEach(function (l) {
        var cs = cols.filter(function (k) { return Math.min(k.x1, l.x1) - Math.max(k.x0, l.x0) > 0.6 * Math.min(k.x1 - k.x0, l.x1 - l.x0); });
        if (cs.length > 1) return;
        var c = cs[0];
        if (c) { c.lines.push(l); c.x0 = Math.min(c.x0, l.x0); c.x1 = Math.max(c.x1, l.x1); }
        else cols.push({ x0: l.x0, x1: l.x1, lines: [l] });
      });
      cols = cols.filter(function (c) { return c.lines.length >= 2 || cols.length === 1; }).sort(function (p, q) { return p.x0 - q.x0; });
      var made = [];
      cols.forEach(function (c, ci) {
        var left = ci === 0 ? b.x0 : cols[ci - 1].x1 + 1, right = ci === cols.length - 1 ? Math.max(c.x1, b.x1) : c.x1;
        // rows from the band down, line by line; a gap of more than 50 points ends the block
        var ls = c.lines.sort(function (p, q) { return q.y - p.y; }), spans = [], top = b.bottom;
        for (var i = 0; i < ls.length && spans.length < 6; i++) {
          if (top - ls[i].y < 6) { top = ls[i].y; continue; } // a doubled line, or the band's own border
          if (top - ls[i].y > 50) break;
          spans.push({ top: top, foot: ls[i].y, x0: ls[i].x0 });
          top = ls[i].y;
        }
        // a ruled table whose row lines run on across the parties side by side (label | value | label | value): split
        // at every second cell border that runs down all its rows
        partyCuts(look, c.x0, c.x1, spans).forEach(function (p, pi, all) {
          var pl = pi === 0 ? left : p.x0, pr = pi === all.length - 1 ? right : p.x1, rows = [];
          spans.forEach(function (sp) {
            var end = labelEnd(look, pl, p.x1, sp.foot + 1, sp.top - 1), vx0 = Math.max(sp.x0, end);
            if (p.x1 - vx0 < 40) return;
            rows.push({ box: { x0: vx0, y0: sp.foot + 1, x1: p.x1, y1: sp.top - 1 }, label: { x0: pl, y0: sp.foot + 1, x1: vx0, y1: sp.top - 1 } });
          });
          if (rows.length >= 2) made.push({ head: { x0: pl, y0: b.bottom, x1: pr, y1: b.top }, x0: pl, x1: p.x1, rows: rows });
        });
      });
      if (made.length) out.push({ band: b, cols: made });
    });
    return out;
  }

  // The words read (OCR items, in the page's units) whose middle is in a rectangle, as one lower-case line.
  function wordsInRect(words, r) {
    return norm((words || []).filter(function (i) {
      var mx = i.x + (i.w || 0) / 2, my = i.y + (i.h || 0) * 0.35;
      return String(i.str || '').trim() && mx > r.x0 - 2 && mx < r.x1 + 2 && my > r.y0 - 2 && my < r.y1 + 2;
    }).sort(function (a, b) { return b.y - a.y || a.x - b.x; }).map(function (i) { return i.str; }).join(' ')).replace(/\s+/g, ' ').trim();
  }
  function headingOf(text, list) {
    var t = ' ' + text.replace(/[^a-z' ]+/g, ' ').replace(/\s+/g, ' ') + ' ';
    for (var i = 0; i < list.length; i++) if (t.indexOf(' ' + list[i] + ' ') >= 0) return list[i];
    return null;
  }
  var ROW_ORDER = { 2: ['sign', 'date'], 3: ['sign', 'name', 'date'], 4: ['sign', 'name', 'company', 'date'] };
  var LABEL_KEY = { signature: 'sign', name: 'name', date: 'date', company: 'company' };

  // The client's block among the blocks pictureRows found, from the words read on them (words: OCR items in the
  // page's units; none when the page couldn't read): the column under a client heading (options.headings first, then
  // the usual ones); where no client heading is read, of two columns side by side the one that isn't the company's (the
  // right one when neither is: guess, said on the card). Each row's piece from its label (Signature, Name, Date,
  // Company), else by its place (Signature, Name, Company, Date for four rows). Null when none is the client's.
  // Result: a layout map's block { heading, sign, name?, date?, company?, guess } (boxes as fractions of the page;
  // guess false, or 'left' / 'right': the column taken without a client heading read).
  function pictureClientBlock(found, words, box, options) {
    var o = opts(options), company = DEFAULTS.company.map(norm);
    var pick = null;
    (found || []).forEach(function (f) {
      if (pick && !pick.guess) return;
      var heads = f.cols.map(function (c) {
        var t = wordsInRect(words, c.head);
        return { col: c, text: t, client: headingOf(t, o.headings), company: headingOf(t, company) };
      });
      var client = heads.filter(function (h) { return h.client && !h.company; })[0] || heads.filter(function (h) { return h.client; })[0];
      if (client) { pick = { col: client.col, heading: client.client, guess: false }; return; }
      if (pick || heads.length !== 2) return;
      // (a heading read as nonsense counts as not read)
      var other = heads.filter(function (h) { return !h.company; });
      // guess: which of the two it is ('left' or 'right'), so the card names the one to check
      if (other.length === 1) pick = { col: other[0].col, heading: 'client', guess: other[0] === heads[0] ? 'left' : 'right' };
      else if (other.length === 2) pick = { col: heads[1].col, heading: 'client', guess: 'right' };
    });
    if (!pick) return null;
    var rows = pick.col.rows, keys = rows.map(function (r) {
      var t = wordsInRect(words, r.label).replace(/[^a-z ]+/g, ' ').trim();
      var hit = null;
      Object.keys(LABEL_KEY).forEach(function (k) {
        if (!hit && (o.labels[k] || []).some(function (w) { return t === w || t.indexOf(w + ' ') === 0 || t.slice(-w.length - 1) === ' ' + w; })) hit = LABEL_KEY[k];
      });
      return hit;
    });
    var order = ROW_ORDER[Math.min(4, rows.length)] || ROW_ORDER[4];
    if (keys.some(function (k) { return !k; })) {
      // labels not read: by place, unless what was read says otherwise
      var unread = order.filter(function (k) { return keys.indexOf(k) < 0; });
      keys = keys.map(function (k) { return k || unread.shift() || null; });
    }
    var W = box.x1 - box.x0, H = box.y1 - box.y0, m = { heading: pick.heading, guess: pick.guess };
    rows.forEach(function (r, i) {
      var k = keys[i];
      if (!k || m[k]) return;
      m[k] = { x: (r.box.x0 - box.x0) / W, y: (box.y1 - r.box.y1) / H, w: (r.box.x1 - r.box.x0) / W, h: (r.box.y1 - r.box.y0) / H };
    });
    return m.sign ? m : null;
  }

  // Whether something is already written or printed inside a box of the sheet drawn in look (r in the page's units):
  // dark marks in its middle, away from its edges, ruled lines not counted, and as tall as writing (3 points at
  // least: a dotted line to write on isn't). A signature in the sign box; the client's company printed in its
  // Company box (then it isn't written again). From QC Suite's inkIn.
  function inkIn(look, r) {
    if (!look || !r) return false;
    var x0 = r.x0 + 4, x1 = r.x1 - 4, y0 = r.y0 + 2.5, y1 = r.y1 - 2.5;
    var c0 = Math.max(0, Math.floor((x0 - look.x0) * look.scale)), c1 = Math.min(look.w, Math.ceil((x1 - look.x0) * look.scale));
    var r0 = Math.max(0, Math.floor((look.y1 - y1) * look.scale)), r1 = Math.min(look.h, Math.ceil((look.y1 - y0) * look.scale));
    var w = c1 - c0, h = r1 - r0;
    if (w < 4 || h < 4) return false;
    var data = new Uint8ClampedArray(w * h * 4);
    for (var y = 0; y < h; y++) data.set(look.data.subarray(((r0 + y) * look.w + c0) * 4, ((r0 + y) * look.w + c1) * 4), y * w * 4);
    if (!inkInBox(data, w, h).inked) return false;
    var tall = 0;
    for (var yy = 0; yy < h; yy++) {
      var n = 0;
      for (var xx = 0; xx < w && n < 2; xx++) if (lum(data, (yy * w + xx) * 4) < DARK) n++;
      if (n >= 2) tall++;
    }
    return tall >= 3 * look.scale;
  }

  // The lowest ruled line (a dark row running over half the area's width) in the bottom half of the area,
  // in pixels from the top, or null. For a boxed sign-off, the top of the Date row under the signature box.
  function bottomBorder(rgba, w, h) {
    for (var y = h - 1; y >= Math.floor(h / 2); y--) {
      var run = 0, best = 0;
      for (var x = 0; x < w; x++) {
        if (lum(rgba, (y * w + x) * 4) < 160) { run++; if (run > best) best = run; } else run = 0;
      }
      if (best >= w * 0.5) return y;
    }
    return null;
  }

  // The first cell border (a dark column running over half the area's height) at or right of fromX, in pixels,
  // or null. The text layer doesn't show where a table cell ends, so the page uses this to keep the
  // signature inside the client's cell.
  function rightBorder(rgba, w, h, fromX) {
    for (var x = Math.max(0, fromX); x < w; x++) {
      var run = 0, best = 0;
      for (var y = 0; y < h; y++) {
        if (lum(rgba, (y * w + x) * 4) < 160) { run++; if (run > best) best = run; } else run = 0;
      }
      if (best >= h * 0.55) return x;
    }
    return null;
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
    var bold = null;
    var aspect = png.width / png.height;
    var ink = lib.rgb(0.08, 0.12, 0.35);
    for (var i = 0; i < pages.length; i++) {
      var p = pages[i], page = doc.getPage(p.index);
      if (p.place.frame) {
        var f = p.place.frame;
        page.drawRectangle({ x: f.x, y: f.y, width: f.w, height: f.h, borderColor: ink, borderWidth: 0.6 });
      }
      var r = fitImage(p.place.sig, aspect);
      page.drawImage(png, { x: r.x, y: r.y, width: r.w, height: r.h });
      for (var k = 0; k < p.place.texts.length; k++) {
        var t = p.place.texts[k];
        if (t.bold && !bold) bold = await doc.embedFont(lib.StandardFonts.HelveticaBold);
        var f = t.bold ? bold : font;
        page.drawText(t.str, { x: t.x, y: t.y, size: fitSize(t, function (str, z) { return f.widthOfTextAtSize(str, z); }), font: f, color: ink });
      }
    }
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
  // f: {path, error?, signed?, already?, hasText?, hasBlock?}. already: every client block in it is signed already. Called with the path alone before the file is read.
  function leaveOutReason(f) {
    var parts = String(f.path).split(/[\\/]/).filter(Boolean);
    var base = parts[parts.length - 1] || '';
    if (parts.indexOf('__MACOSX') >= 0 || /^\._/.test(base)) return 'System file';
    if (!/\.pdf$/i.test(base)) return 'Not a PDF';
    if (parts.slice(0, -1).some(function (p) { return p.toLowerCase() === 'signed'; })) return 'In a folder called Signed';
    if (f.error) return f.error;
    if (f.signed) return 'Already signed through this page';
    if (f.already) return 'Already signed';
    if (f.hasText && !f.hasBlock) return 'No sign-off block found';
    return null;
  }

  // The QCMS entry in a sign pack's PDF info (JSON), or null.
  // {v: 1, project_no, transmittal, pages: [null | {doc_no, rev, page, pages, client_signoff, client: {x, y, w, h},
  //   client_company, client_blocks: [{heading, optional, sign, name, date, company}]}]}
  // client_blocks: a form with more than one client block (e.g. "Test Witnessed By" and "Accepted for Client"),
  // each on its page from the form's map: the page signs each ticked one. Older packs have none.
  function parseQcms(str) {
    if (!str) return null;
    try {
      var o = JSON.parse(str);
      return o && o.v === 1 && Array.isArray(o.pages) ? o : null;
    } catch (e) { return null; }
  }

  // Placement from a QCMS hint: the client box as fractions of the page from the top-left.
  function hintPlacement(hint, box, name, date) {
    return boxPlacement(hint && hint.client, box, name, date);
  }

  // Text written into a box of the form (fractions of the page from the top-left), e.g. a sign-off's Name,
  // Date or Company box: at its left, up to 10 pt (at least 7, less than the box's height), with maxW its
  // width so the drawing makes a long text smaller to fit (fitSize). null for no box or no text.
  function boxText(c, box, str) {
    str = plainText(str).trim();
    if (!c || !str) return null;
    var W = box.x1 - box.x0, H = box.y1 - box.y0, bh = c.h * H;
    return { str: str, x: box.x0 + c.x * W + 3, y: box.y1 - (c.y + c.h) * H + Math.max(2, (bh - 8) / 2),
             size: Math.max(7, Math.min(10, bh - 4)), maxW: Math.max(10, c.w * W - 6) };
  }

  // The size a placed text is drawn at: its own, smaller (down to 4 pt) when at that size it is wider than
  // its maxW. width(str, size) measures it in the font it is drawn in.
  function fitSize(t, width) {
    if (!t.maxW) return t.size;
    var w = width(t.str, t.size);
    return w > t.maxW ? Math.max(4, Math.floor(t.size * t.maxW / w * 4) / 4) : t.size;
  }

  // A placement with the signer's company written in the block's Company box (the sign pack's hint), when
  // the block has one and the company is known.
  function withCompanyBox(place, c, box, company) {
    var t = place && withRole(boxText(c, box, company), 'company');
    return t ? { sig: place.sig, texts: place.texts.concat([t]), frame: place.frame } : place;
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

  // The "Tags inspected" box (SITE-TAGS.md): found by its label in the text layer. The space right of the
  // label to the next text on the line (or the page's right margin) when it is wide enough, else the lines
  // under the label down to the next text below. filled: text already printed in it (e.g. tags known when
  // the sheet was made), end: where that text stops. Null when the label isn't on the page.
  function findTagsBox(items, box) {
    var lines = buildLines(items), W = box.x1 - box.x0;
    for (var k = 0; k < lines.length; k++) {
      var line = lines[k], m = /tags?\s+inspected\s*:?/i.exec(line.text);
      if (!m) continue;
      // the label's end: the item holding "inspected", measured in proportion when it has more text in it
      var labelEnd = null, first = null;
      for (var j = 0; j < line.items.length; j++) {
        var it = line.items[j], at = it.str.search(/inspected\s*:?/i);
        if (/^\s*tags?\b/i.test(it.str) && first === null) first = it;
        if (at >= 0) {
          var mm = /inspected\s*:?/i.exec(it.str);
          var cut = at + mm[0].length;
          labelEnd = it.x + it.w * (cut / Math.max(1, it.str.length));
          first = first || it;
          var rest = it.str.slice(cut).trim();
          var after = line.items.slice(j + 1).filter(function (x) { return x.x >= labelEnd - 1; });
          var nextLabel = after.filter(function (x) { return /:\s*$/.test(x.str.trim()); })[0];
          var filled = [rest].concat(after.filter(function (x) { return !nextLabel || x.x < nextLabel.x; }).map(function (x) { return x.str.trim(); }))
            .filter(Boolean).join(' ').replace(/[_.]{3,}/g, '').trim();
          var right = nextLabel ? nextLabel.x - 4 : box.x1 - W * 0.05;
          var printed = after.filter(function (x) { return !nextLabel || x.x < nextLabel.x; }).filter(function (x) { return !/^[_.\s]+$/.test(x.str); });
          var end = printed.length ? Math.max.apply(null, printed.map(function (x) { return x.x + x.w; })) : (rest ? it.x + it.w : labelEnd);
          var h = line.items.reduce(function (a, x) { return Math.max(a, x.h || 0); }, 6);
          var below = lines.filter(function (l) { return l.y < line.y - 2 && l.items.some(function (x) { return x.x < right && x.x + x.w > first.x; }); })
            .sort(function (a, b) { return b.y - a.y; })[0];
          var floor = below ? below.y + (below.items[0].h || 7) + 1 : line.y - 3 * h;
          floor = Math.max(floor, line.y - 3.5 * h);
          if (right - Math.max(labelEnd, end) >= W * 0.2) {
            return { page: null, x0: labelEnd + 3, x1: right, y1: line.y + h + 1, y0: Math.min(line.y - 2, floor), labelled: true, filled: filled, end: Math.max(end, labelEnd) };
          }
          return { page: null, x0: first.x, x1: right, y1: line.y - 2, y0: floor, labelled: true, filled: "", end: first.x };
        }
      }
    }
    return null;
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
           findCompanyBlock: findCompanyBlock, boxPlacement: boxPlacement, layoutPlacement: layoutPlacement,
           fracRect: fracRect, matchClientBlock: matchClientBlock,
           CHECK_HEADERS: CHECK_HEADERS, findChecklist: findChecklist, findContractorBlock: findContractorBlock, findTagsBox: findTagsBox,
           underText: underText, printedIn: printedIn, adjustPlace: adjustPlace, pieceRects: pieceRects, ADDED_LABEL: ADDED_LABEL, addedLabel: addedLabel,
           pictureRows: pictureRows, hasWords: hasWords, cleanForOcr: cleanForOcr, pictureClientBlock: pictureClientBlock, inkIn: inkIn, headingBands: headingBands, ruledLines: ruledLines,
           STAMP: STAMP, stampPlacement: stampPlacement, emptySpot: emptySpot, inkInBox: inkInBox, rightBorder: rightBorder, bottomBorder: bottomBorder, blank: blank,
           fitImage: fitImage, stampPdf: stampPdf, isSigned: isSigned, plainText: plainText,
           leaveOutReason: leaveOutReason, parseQcms: parseQcms, hintPlacement: hintPlacement, ocrItems: ocrItems,
           boxText: boxText, fitSize: fitSize, withCompanyBox: withCompanyBox,
           groupPages: groupPages, outputName: outputName, trimBox: trimBox,
           jpegInfo: jpegInfo, orientTransform: orientTransform, greyscale: greyscale, evenLighting: evenLighting, otsuThreshold: otsuThreshold,
           inkImage: inkImage, removeSpecks: removeSpecks, cleanSignature: cleanSignature,
           setOutline: setOutline, combinePdfs: combinePdfs, addNotesPage: addNotesPage };
})();
if (typeof module !== 'undefined') module.exports = SignCore;

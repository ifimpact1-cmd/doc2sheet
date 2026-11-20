/* Doc2Sheet engine: DOCX -> XLSX berbasis posisi (twips -> kolom Excel). Butuh JSZip + ExcelJS. */
(function (G) {
'use strict';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const el = n => n && n.nodeType === 1;
const kids = (e, n) => e ? [...e.childNodes].filter(c => el(c) && (!n || c.localName === n)) : [];
const kid = (e, n) => kids(e, n)[0];
const at = (e, n) => e ? e.getAttributeNS(W, n) : null;
const val = e => at(e, 'val');
const on = e => !!e && !['0', 'false', 'none'].includes(val(e));
const num = (e, n) => { const v = at(e, n); return v === null || v === '' ? undefined : +v; };
const desc = (e, n, o = []) => { for (const c of kids(e)) { if (c.localName === n) o.push(c); desc(c, n, o); } return o; };
const nb = b => !['nil', 'none'].includes(val(b));
const FIX = { 'â€“': '–', 'â€”': '—', 'â€™': '’', 'â€œ': '“', 'â€˜': '‘', 'â€¢': '•', 'Â ': ' ' };
const fix = t => t.includes('â') || t.includes('Â') ? t.replace(/â€[“”™œ˜¢]|â€\u009d|Â /g, m => FIX[m] || (m === 'â€\u009d' ? '”' : m)) : t;

const rprOf = r => {
  const o = {}; if (!r) return o;
  const b = kid(r, 'b'); if (b) o.b = on(b);
  const i = kid(r, 'i'); if (i) o.i = on(i);
  const u = kid(r, 'u'); if (u) o.u = val(u) !== 'none';
  const s = num(kid(r, 'sz'), 'val'); if (s) o.sz = s / 2;
  const c = val(kid(r, 'color')); if (c && c !== 'auto') o.color = c;
  const f = at(kid(r, 'rFonts'), 'ascii'); if (f) o.font = f;
  return o;
};
const bdOf = e => {
  const o = {};
  for (const s of ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']) {
    const b = kid(e, s);
    if (b) { const c = at(b, 'color'); o[s] = nb(b) ? { sz: +at(b, 'sz') || 4, color: c && c !== 'auto' ? c : '000000' } : null; }
  }
  return o;
};
const pprOf = p => {
  const o = {}; if (!p) return o;
  const j = val(kid(p, 'jc')); if (j) o.jc = j;
  const n = kid(p, 'ind');
  if (n) {
    const l = num(n, 'left') ?? num(n, 'start'); if (l !== undefined) o.left = l;
    const r = num(n, 'right') ?? num(n, 'end'); if (r !== undefined) o.right = r;
    if (num(n, 'firstLine') !== undefined) o.first = num(n, 'firstLine');
    if (num(n, 'hanging') !== undefined) o.hang = num(n, 'hanging');
  }
  const s = kid(p, 'spacing');
  if (s) { for (const k of ['before', 'after', 'line']) if (num(s, k) !== undefined) o[k] = num(s, k); if (at(s, 'lineRule')) o.rule = at(s, 'lineRule'); }
  const t = kid(p, 'tabs'); if (t) o.tabs = kids(t, 'tab').map(x => ({ pos: +at(x, 'pos'), v: val(x) }));
  const b = kid(p, 'pBdr'); if (b) o.bdr = bdOf(b);
  if (kid(p, 'numPr')) o.num = 1;
  const mk = rprOf(kid(p, 'rPr')); if (mk.sz) o.mark = mk.sz;
  return o;
};

// perkiraan lebar teks (twips)
const cw = (t, sz, b) => { let w = 0; for (const ch of t) w += ch === ' ' ? .28 : /[mwMW@]/.test(ch) ? .85 : /[Iil.,:;|!'j]/.test(ch) ? .27 : /[0-9]/.test(ch) ? .556 : /[A-Z]/.test(ch) ? .68 : /[ftr]/.test(ch) ? .33 : .5; return w * sz * 20 * (b ? 1.06 : 1); };
const segW = s => s.reduce((a, r) => a + cw(r.text, r.font.sz, r.font.b), 0);

async function build(file) {
  const zip = await G.JSZip.loadAsync(file);
  const rd = async p => { const f = zip.file(p); return f ? new DOMParser().parseFromString(await f.async('string'), 'application/xml') : null; };
  const doc = await rd('word/document.xml');
  if (!doc) throw Error('Bukan file .docx yang valid');
  const sty = await rd('word/styles.xml'), rels = await rd('word/_rels/document.xml.rels');
  const relMap = {};
  if (rels) for (const r of desc(rels.documentElement, 'Relationship')) relMap[r.getAttribute('Id')] = r.getAttribute('Target');

  // ---- styles
  const S = {}; let D = {}, dflt = 'Normal';
  if (sty) {
    const dd = desc(sty.documentElement, 'rPrDefault')[0]; if (dd) D = rprOf(kid(dd, 'rPr'));
    for (const s of desc(sty.documentElement, 'style')) {
      if (at(s, 'type') !== 'paragraph') continue;
      const id = at(s, 'styleId'); if (at(s, 'default') === '1') dflt = id;
      S[id] = { base: val(kid(s, 'basedOn')), p: pprOf(kid(s, 'pPr')), r: rprOf(kid(s, 'rPr')) };
    }
  }
  const res = (id, k, d = 0) => { const s = S[id]; return !s || d > 8 ? {} : { ...res(s.base, k, d + 1), ...s[k] }; };

  const body = kid(doc.documentElement, 'body');
  const secs = []; let cur = [];
  for (const c of kids(body)) {
    const n = c.localName;
    if (n === 'sectPr') { if (cur.length) secs.push({ sp: c, items: cur }); cur = []; continue; }
    if (n !== 'p' && n !== 'tbl') continue;
    cur.push(c);
    const sp = n === 'p' ? kid(kid(c, 'pPr'), 'sectPr') : null;
    if (sp) { secs.push({ sp, items: cur }); cur = []; }
  }
  if (cur.length) secs.push({ sp: kid(body, 'sectPr'), items: cur });

  const geo = sp => {
    const pg = kid(sp, 'pgSz'), m = kid(sp, 'pgMar');
    const w = +at(pg, 'w') || 11906, h = +at(pg, 'h') || 16838, l = num(m, 'left') ?? 1440, r = num(m, 'right') ?? 1440;
    const cs = kid(sp, 'cols'), n = +at(cs, 'num') || 1, gap = +at(cs, 'space') || 720, defs = kids(cs, 'col'), cols = [];
    if (defs.length > 1) { let x = l; defs.forEach(c => { const cwid = +at(c, 'w'); cols.push({ x0: x, x1: x + cwid }); x += cwid + (+at(c, 'space') || gap); }); }
    else if (n > 1) { const cwid = (w - l - r - (n - 1) * gap) / n; for (let i = 0; i < n; i++) cols.push({ x0: l + i * (cwid + gap), x1: l + i * (cwid + gap) + cwid }); }
    else cols.push({ x0: l, x1: w - r });
    return { w, h, l, r, t: num(m, 'top') ?? 1440, b: num(m, 'bottom') ?? 1440, cols };
  };

  const walkRuns = (e, o = []) => { for (const c of kids(e)) { if (c.localName === 'r') o.push(c); else if (['hyperlink', 'ins', 'smartTag', 'sdt', 'sdtContent', 'fldSimple'].includes(c.localName)) walkRuns(c, o); } return o; };

  // ---- layout satu paragraf pada bingkai [fx0,fx1]
  function layout(p, rn, fx0, fx1, mL = 0, mR = 0) {
    const pPr = kid(p, 'pPr'), sid = val(kid(pPr, 'pStyle')) || dflt;
    const pp = { ...res(dflt, 'p'), ...res(sid, 'p'), ...pprOf(pPr) };
    const rb = { ...D, ...res(dflt, 'r'), ...res(sid, 'r') };
    const L0 = fx0 + mL + (pp.left || 0), R0 = fx1 - mR - (pp.right || 0), st = L0 + (pp.first || 0) - (pp.hang || 0);
    const segs = [[]], last = () => segs[segs.length - 1];
    let fs = 0;
    for (const r of rn) {
      const f = { ...rb, ...rprOf(kid(r, 'rPr')) }; f.sz = f.sz || 11; f.font = f.font || 'Calibri';
      let t = ''; const flush = () => { if (t) { last().push({ text: fix(t), font: f }); fs = Math.max(fs, f.sz); t = ''; } };
      for (const c of kids(r)) {
        const n = c.localName;
        if (n === 't') t += c.textContent;
        else if (n === 'tab') { flush(); segs.push([]); }
        else if ((n === 'br' || n === 'cr') && at(c, 'type') !== 'column') t += '\n';
      }
      flush();
    }
    if (pp.num && segs[0].length) segs[0].unshift({ text: '• ', font: segs[0][0].font });
    const jc = pp.jc === 'both' ? 'justify' : pp.jc === 'center' ? 'center' : (pp.jc === 'right' || pp.jc === 'end') ? 'right' : 'left';
    const size = fs || pp.mark || rb.sz || 11;
    const clamp = x => Math.min(Math.max(x, fx0), fx1);
    const flat = segs.flat(), txt = flat.map(r => r.text).join('');
    let n = txt.split('\n').length;
    if (segs.length === 1) n = Math.max(n, Math.ceil(segW(flat) / (Math.max(500, R0 - st) * 1.08)));
    let pieces = [];
    if (segs.length === 1) {
      if (segs[0].length) pieces = [{ x0: fx0, x1: fx1, wrap: n > 1, runs: segs[0], al: jc, ind: Math.max(0, jc === 'right' ? mR + (pp.right || 0) : jc === 'center' ? 0 : mL + (pp.left || 0) + (pp.first || 0) - (pp.hang || 0)) }];
    } else {
      const stops = (pp.tabs || []).filter(t => t.v !== 'clear').map(t => ({ abs: fx0 + t.pos, v: t.v })).sort((a, b) => a.abs - b.abs);
      let c = st; const ps = [];
      segs.forEach((s, i) => {
        const w = segW(s); let x0 = c, x1 = null, al = 'left';
        if (i > 0) {
          let sp = stops.find(t => t.abs > c + 10);
          if (!sp) sp = { abs: fx0 + (Math.floor((c - fx0) / 720) + 1) * 720, v: 'left' };
          if (sp.v === 'right' || sp.v === 'end') { al = 'right'; x1 = sp.abs; x0 = Math.max(c, sp.abs - w - 60); }
          else if (sp.v === 'center') { al = 'center'; x0 = Math.max(c, sp.abs - w / 2 - 60); x1 = sp.abs + w / 2 + 60; }
          else x0 = sp.abs;
        }
        ps.push({ x0, x1, al, runs: s, w });
        c = al === 'left' ? x0 + w : x1;
      });
      const end = ps[ps.length - 1].x1 ?? (ps[ps.length - 1].x0 + ps[ps.length - 1].w);
      let sh = 0;
      if (jc === 'right') sh = R0 - end; else if (jc === 'center') sh = (L0 + R0) / 2 - (st + end) / 2;
      ps.forEach(q => { q.x0 += sh; if (q.x1 !== null) q.x1 += sh; });
      if (jc === 'right') { const q = ps[ps.length - 1]; if (q.al === 'left') { q.al = 'right'; q.x1 = R0; } }
      ps.forEach((q, i) => { if (i > 0 && q.al === 'right') { const p0 = ps[i - 1]; q.x0 = Math.max(p0.x0 + 100, (p0.al === 'left' ? p0.x0 + p0.w * 1.25 : p0.x1) + 60); } });
      ps.forEach((q, i) => { if (q.x1 === null) q.x1 = i < ps.length - 1 ? ps[i + 1].x0 : fx1; });
      pieces = ps.filter(q => q.runs.length).map(q => {
        const x0 = clamp(q.x0), x1 = clamp(q.x1);
        return { x0, x1: x1 <= x0 + 100 ? Math.min(fx1, x0 + 400) : x1, runs: q.runs, al: q.al, ind: 0 };
      });
    }
    if (pp.bdr) for (const q of pieces) q.bd = { top: pp.bdr.top || undefined, bottom: pp.bdr.bottom || undefined };
    const lh = pp.rule === 'exact' && pp.line ? pp.line / 20 : pp.rule === 'atLeast' && pp.line ? Math.max(pp.line / 20, size * 1.17) : size * 1.17 * ((pp.line || 240) / 240);
    return { pieces, flat, wrap: n > 1, al: jc, ind: pieces[0] ? pieces[0].ind : 0, size, st, pp, h: n * lh + ((pp.before || 0) + (pp.after || 0)) / 20 };
  }

  const imgs = [];
  function paraLine(p, runs, col, li) {
    const L = layout(p, runs, col.x0, col.x1);
    const line = { pieces: L.pieces, h: L.h, imgs: [], key: li };
    for (const r of runs) for (const d of desc(r, 'drawing')) {
      const a = kid(d, 'anchor') || kid(d, 'inline'); if (!a) continue;
      const ex = kid(a, 'extent'), cx = +ex.getAttribute('cx'), cy = +ex.getAttribute('cy');
      const blip = desc(a, 'blip')[0];
      if (blip) {
        const rid = blip.getAttributeNS(R, 'embed'); let x = L.st, dy = 0;
        if (a.localName === 'anchor') {
          const ph = kid(a, 'positionH'), pv = kid(a, 'positionV');
          const o = (+(kid(ph, 'posOffset') || {}).textContent || 0) / 635;
          x = ph && ph.getAttribute('relativeFrom') === 'page' ? o : col.x0 + o;
          dy = (+(kid(pv, 'posOffset') || {}).textContent || 0) / 12700;
        }
        line.imgs.push({ rid, x, dy, w: cx / 9525, h: cy / 9525 });
      } else if (cy <= 50800 && cx > 0) {
        const lw = desc(a, 'ln')[0]; const sz = lw ? Math.round((+lw.getAttribute('w') || 12700) / 1588) : 8;
        line.pieces.push({ x0: L.st, x1: Math.min(col.x1, L.st + cx / 635), runs: [], al: 'left', ind: 0, bd: { bottom: { sz, color: '000000' } } });
        line.h = 2.5;
      }
    }
    return line;
  }

  function tableLines(t, bx0) {
    const pr = kid(t, 'tblPr'), ind = num(kid(pr, 'tblInd'), 'w') || 0, mar = kid(pr, 'tblCellMar');
    const mL = num(kid(mar, 'left'), 'w') ?? 108, mR = num(kid(mar, 'right'), 'w') ?? 108, tbd = bdOf(kid(pr, 'tblBorders'));
    const trs = kids(t, 'tr');
    let g = kids(kid(t, 'tblGrid'), 'gridCol').map(c => +at(c, 'w'));
    if (!g.length) g = kids(trs[0], 'tc').map(c => +at(kid(kid(c, 'tcPr'), 'tcW'), 'w') || 1000);
    const cum = [0]; g.forEach(w => cum.push(cum[cum.length - 1] + w));
    return trs.map((tr, ri) => {
      const trp = kid(tr, 'trPr'), trh = kid(trp, 'trHeight'), th = (+at(trh, 'val') || 0) / 20;
      let gi = num(kid(trp, 'gridBefore'), 'val') || 0, hmax = 0; const pieces = [];
      for (const tc of kids(tr, 'tc')) {
        const pc = kid(tc, 'tcPr'), span = num(kid(pc, 'gridSpan'), 'val') || 1, e = Math.min(gi + span, g.length);
        const x0 = bx0 + ind + cum[gi], x1 = bx0 + ind + cum[e], vm = kid(pc, 'vMerge');
        const runs = []; let hh = 0, al = 'left', indent = 0, wrap = false;
        kids(tc, 'p').forEach((p, i) => {
          const L = layout(p, walkRuns(p), x0, x1, mL, mR);
          wrap = wrap || L.wrap || i > 0; if (i === 0) { al = L.al; indent = L.ind; } else if (runs.length) runs.push({ text: '\n', font: { sz: L.size, font: 'Calibri' } });
          L.pieces.length > 1 ? L.pieces.forEach((q, k) => { if (k) runs.push({ text: '  ', font: L.flat[0].font }); runs.push(...q.runs); }) : runs.push(...L.flat);
          hh += L.h;
        });
        const cb = bdOf(kid(pc, 'tcBorders')), pick = (k, fb) => cb[k] !== undefined ? cb[k] : fb;
        const last = ri === trs.length - 1, ed = e >= g.length;
        const fill = at(kid(pc, 'shd'), 'fill'), va = val(kid(pc, 'vAlign'));
        pieces.push({
          x0, x1, runs, al, wrap, ind: indent, isCell: true, fill: fill && fill !== 'auto' ? fill : null,
          va: va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'middle',
          vm: vm ? (val(vm) === 'restart' ? 'restart' : 'cont') : null,
          bd: { top: pick('top', ri === 0 ? tbd.top : tbd.insideH), bottom: pick('bottom', last ? tbd.bottom : tbd.insideH), left: pick('left', gi === 0 ? tbd.left : tbd.insideV), right: pick('right', ed ? tbd.right : tbd.insideV) }
        });
        hmax = Math.max(hmax, hh); gi = e;
      }
      return { pieces, h: at(trh, 'hRule') === 'exact' ? th : Math.max(th, hmax), imgs: [] };
    });
  }

  // ---- rakit baris dokumen
  const lines = []; let g0 = null, li = 0;
  for (const sec of secs) {
    const gm = geo(sec.sp); if (!g0) g0 = gm;
    const cols = gm.cols, buckets = cols.map(() => []); let ci = 0;
    for (const it of sec.items) {
      if (it.localName === 'tbl') { buckets[ci].push(...tableLines(it, cols[ci].x0)); continue; }
      const groups = []; let grp = [];
      for (const r of walkRuns(it)) {
        if (kids(r, 'br').some(b => at(b, 'type') === 'column')) { groups.push(grp); grp = []; }
        grp.push(r);
      }
      groups.push(grp);
      groups.forEach((gr, k) => {
        if (k > 0) ci = Math.min(ci + 1, cols.length - 1);
        const empty = !gr.length;
        if (k < groups.length - 1 && empty) return;
        buckets[ci].push(paraLine(it, gr, cols[ci], li++));
      });
    }
    const n = Math.max(0, ...buckets.map(b => b.length));
    for (let i = 0; i < n; i++) {
      const row = { pieces: [], h: 0, imgs: [] };
      for (const b of buckets) if (b[i]) { row.pieces.push(...b[i].pieces); row.imgs.push(...b[i].imgs); row.h = Math.max(row.h, b[i].h); }
      lines.push(row);
    }
  }
  return { lines, g0, zip, relMap, cols: secs.map(s => geo(s.sp).cols).flat() };
}

async function docx2xlsx(file) {
  const { lines, g0, zip, relMap, cols } = await build(file);
  const O = g0.l, PW = g0.w - g0.l - g0.r;
  let step = 60, bounds;
  for (const s of [60, 120, 200, 300, 450, 700]) {
    step = s; const set = new Set([0, Math.round(PW / s) * s]);
    const add = x => set.add(Math.min(PW, Math.max(0, Math.round((x - O) / s) * s)));
    cols.forEach(c => { add(c.x0); add(c.x1); });
    lines.forEach(l => l.pieces.forEach(p => { add(p.x0); add(p.x1); }));
    bounds = [...set].sort((a, b) => a - b);
    if (bounds.length <= 45) break;
  }
  const n = bounds.length - 1, cx = x => Math.max(0, bounds.indexOf(Math.min(PW, Math.max(0, Math.round((x - O) / step) * step)))) + 1;
  const wb = new G.ExcelJS.Workbook();
  const ws = wb.addWorksheet('Dokumen', { views: [{ showGridLines: false }] });
  const px = i => (bounds[i + 1] - bounds[i]) / 15;
  for (let i = 0; i < n; i++) ws.getColumn(i + 1).width = Math.max(0.3, (px(i) - 5) / 7);
  const argb = c => 'FF' + c;
  const bs = b => b ? { style: b.sz >= 12 ? 'medium' : b.sz <= 2 ? 'hair' : 'thin', color: { argb: argb(b.color) } } : undefined;
  const heights = [], open = {}, vlist = [];
  lines.forEach((ln, idx) => {
    const r = idx + 1; let lastEnd = 0;
    const ps = ln.pieces.sort((a, b) => a.x0 - b.x0);
    ps.forEach(p => { p.c1 = cx(p.x0); p.c2 = Math.max(p.c1, cx(p.x1) - 1); });
    ps.forEach((p, i) => {
      if (p.isCell || p.wrap || !p.runs.length) return;
      if (p.al === 'right') p.c1 = i ? Math.min(p.c1, ps[i - 1].c2 + 1) : 1;
      else if (p.al !== 'center') p.c2 = Math.max(p.c2, i < ps.length - 1 ? ps[i + 1].c1 - 1 : n);
    });
    ps.forEach(p => {
      let c1 = p.c1, c2 = p.c2;
      if (p.vm === 'cont') { const o = open[c1]; if (o) o.r2 = r; lastEnd = Math.max(lastEnd, c2); return; }
      if (c1 <= lastEnd) c1 = lastEnd + 1;
      if (c1 > c2 || c1 > n) return;
      lastEnd = c2;
      const cell = ws.getCell(r, c1);
      const runs = p.runs.filter(x => x.text).map(x => ({ text: x.text, font: { name: x.font.font, size: x.font.sz, bold: !!x.font.b, italic: !!x.font.i, underline: !!x.font.u, ...(x.font.color ? { color: { argb: argb(x.font.color) } } : {}) } }));
      if (runs.length) cell.value = { richText: runs };
      cell.alignment = { horizontal: p.al, vertical: p.va || 'bottom', wrapText: !!(runs.length && p.wrap), ...(p.ind > 60 && p.al !== 'center' ? { indent: Math.max(0, Math.min(15, Math.round(p.ind / 170), p.wrap ? 15 : Math.floor((p.x1 - p.x0 - segW(p.runs) * 1.2) / 200))) } : {}) };
      if (p.fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: argb(p.fill) } };
      if (p.bd) { const b = { top: bs(p.bd.top), left: bs(p.bd.left), bottom: bs(p.bd.bottom), right: bs(p.bd.right) }; if (b.top || b.left || b.bottom || b.right) cell.border = b; }
      if (p.vm === 'restart') { open[c1] = { r1: r, r2: r, c1, c2 }; vlist.push(open[c1]); }
      else { delete open[c1]; if (c2 > c1) ws.mergeCells(r, c1, r, c2); }
    });
    heights[idx] = Math.min(409, Math.max(2, ln.h));
    ws.getRow(r).height = heights[idx];
  });
  vlist.forEach(m => { if (m.r2 > m.r1 || m.c2 > m.c1) ws.mergeCells(m.r1, m.c1, m.r2, m.c2); });

  // gambar (logo dsb)
  let lineNo = -1;
  for (let idx = 0; idx < lines.length; idx++) for (const im of lines[idx].imgs) {
    const t = relMap[im.rid]; const f = t && zip.file('word/' + t.replace(/^\.?\//, ''));
    if (!f) continue;
    const ext = (t.split('.').pop() || 'png').toLowerCase().replace('jpg', 'jpeg');
    if (!['png', 'jpeg', 'gif'].includes(ext)) continue;
    let id; try { id = wb.addImage({ base64: await f.async('base64'), extension: ext }); } catch (_) { continue; }
    const xr = Math.min(PW, Math.max(0, im.x - O)); let ci = 0; while (ci < n - 1 && bounds[ci + 1] <= xr) ci++;
    const ext2 = { width: im.w, height: im.h }, fr = (xr - bounds[ci]) / (bounds[ci + 1] - bounds[ci]);
    try { ws.addImage(id, { tl: { nativeCol: ci, nativeColOff: Math.round((xr - bounds[ci]) / 15 * 9525), nativeRow: idx, nativeRowOff: Math.round(Math.max(0, im.dy) * 12700) }, ext: ext2, editAs: 'oneCell' }); }
    catch (_) { ws.addImage(id, { tl: { col: ci + fr, row: idx }, ext: ext2 }); }
  }

  ws.pageSetup = {
    paperSize: Math.abs(g0.w - 11906) < 120 ? 9 : 1, orientation: g0.h < g0.w ? 'landscape' : 'portrait',
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: g0.l / 1440, right: g0.r / 1440, top: g0.t / 1440, bottom: g0.b / 1440, header: .3, footer: .3 }
  };
  return new Blob([await wb.xlsx.writeBuffer()], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
}
G.docx2xlsx = docx2xlsx;
})(typeof window !== 'undefined' ? window : globalThis);

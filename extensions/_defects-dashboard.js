// MMO-specific: the "Defects Dashboard" slide (slide14 in MMO_Report_Template.pptx,
// placed right after the executive summary). All-workstream view:
//   - metric tiles (total / closed / open / open >14 days / open critical / unassigned)
//   - severity tiles (open defects per severity) with a proportional strip
//   - table of open Sev 1 (Critical) defects: WS, ADO id, owner (HE/Certify), comments
//   - all-workstream defect burndown (native chart, data written here each run)
//   - key-insights bullets (agent-written, ai-narrative key `defectsDashboardInsights`,
//     one bullet per line)
//
// Owner and Comments for the critical-defects table can't come from ADO (owner is a
// person's name, which never goes in the deck), so they're maintained in
// critical-defect-notes.json at the project root: { "<ADO id>": { owner, targetDate, comments } }.
// A critical defect with no entry still gets a row, with "TBD" cells.
//
// Runs in the early extension phase (see extensions/ppt.js), before the generic
// {{TOKEN}} replacement pass.
import fs   from 'fs';
import path from 'path';

const SLIDE        = 'ppt/slides/slide14.xml';
const BURNDOWN     = 'ppt/charts/chart14.xml';
const NOTES_PATH   = path.join(process.cwd(), 'critical-defect-notes.json');
const MAX_CRIT_ROWS = 5;
const DAY_MS = 24 * 60 * 60 * 1000;

const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const todayUtc = () => new Date(new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }) + 'T00:00:00Z');
const toDate = s => new Date(`${String(s).slice(0, 10)}T00:00:00Z`);

const allOpen   = d => Object.values(d.bugs ?? {}).flat();
const allClosed = d => Object.values(d.closedBugs ?? {}).flat();
const isCrit    = b => b.severity === '1 - Critical';

export function dashboardMetrics(d) {
  const open = allOpen(d), closed = allClosed(d);
  const total = open.length + closed.length;
  const t = todayUtc();
  const old = open.filter(b => b.createdDate && (t - toDate(b.createdDate)) / DAY_MS > 14).length;
  const pct = (n, of) => (of ? (Math.round((n / of) * 1000) / 10).toFixed(1) : '0.0');
  return {
    total, closed: closed.length, open: open.length, old,
    closedPct: pct(closed.length, total),
    openPct:   pct(open.length, total),
    oldPct:    pct(old, open.length),
    crit:      open.filter(isCrit).length,
    unassigned: open.filter(b => !b.owner || /^unassigned$/i.test(b.owner)).length,
  };
}

// ── Severity chart ───────────────────────────────────────────────────────────
const SEV = [['Critical', '1 - Critical'], ['High', '2 - High'], ['Medium', '3 - Medium'], ['Low', '4 - Low']];

export function severityCounts(d) {
  const open = allOpen(d), closed = allClosed(d);
  return SEV.map(([label, sev]) => ({
    label,
    open:     open.filter(b => b.severity === sev).length,
    resolved: closed.filter(b => b.severity === sev && b.state === 'Resolved').length,
    closed:   closed.filter(b => b.severity === sev && b.state === 'Closed').length,
  }));
}

const SEV_COLORS = { Critical: 'C62828', High: 'E67E22', Medium: 'F1C40F', Low: '2ECC71' };
const SEV_NUM_COLORS = { Critical: 'C62828', High: 'D35400', Medium: 'B7950B', Low: '1E8449' };

// Severity tiles (open defects per severity) + a proportional strip beneath, built as
// native shapes in the left column between the section head and the critical table.
const BOX = { x: 463611, y: 2850000, w: 5406762 };
export function buildSeverityShapes(rows) {
  const gap = 80000, tileW = Math.floor((BOX.w - 3 * gap) / 4), tileH = 800000, band = 55000;
  const totalOpen = rows.reduce((a, r) => a + r.open, 0);
  let id = 100, out = '';
  const sp = (name, x, y, w, h, fill, line, body = '') =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id++}" name="${name}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>${line ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${line}"/></a:solidFill></a:ln>` : '<a:ln><a:noFill/></a:ln>'}</p:spPr>`
    + `<p:txBody><a:bodyPr wrap="square" lIns="0" tIns="${band + 30000}" rIns="0" bIns="0" anchor="t"/><a:lstStyle/>${body || '<a:p><a:endParaRPr lang="en-US"/></a:p>'}</p:txBody></p:sp>`;
  const para = (text, sz, color, bold) =>
    `<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="${sz}" b="${bold ? 1 : 0}" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>${esc(text)}</a:t></a:r></a:p>`;
  rows.forEach((r, i) => {
    const x = BOX.x + i * (tileW + gap);
    const pct = totalOpen ? Math.round((r.open / totalOpen) * 100) : 0;
    out += sp(`SevTile${r.label}`, x, BOX.y, tileW, tileH, 'FFFFFF', 'D0D5DD',
      para(String(r.open), 2400, SEV_NUM_COLORS[r.label], true) + para(r.label, 1000, '313131', true) + para(`${pct}% of open`, 900, '667085', false));
    out += sp(`SevBand${r.label}`, x, BOX.y, tileW, band, SEV_COLORS[r.label], null);
  });
  // proportional strip (zero-count severities omitted)
  const barY = BOX.y + tileH + 70000, barH = 120000;
  let cx = BOX.x;
  const live = rows.filter(r => r.open > 0);
  live.forEach((r, i) => {
    const w = i === live.length - 1 ? BOX.x + BOX.w - cx : Math.round(BOX.w * r.open / totalOpen);
    out += sp(`SevBar${r.label}`, cx, barY, w, barH, SEV_COLORS[r.label], 'FFFFFF');
    cx += w;
  });
  return out;
}

// ── Burndown chart (rewrites the caches of the template's native chart) ─────
function burndownSeries(d) {
  const bugs = [...allOpen(d), ...allClosed(d)];
  const t = todayUtc();
  const points = [];
  for (let i = 7; i >= 0; i--) points.push(new Date(t.getTime() - i * 4 * DAY_MS));
  const rows = points.map(p => {
    const logged = bugs.filter(b => b.createdDate && toDate(b.createdDate) <= p).length;
    const closed = bugs.filter(b => b.closedDate && toDate(b.closedDate) <= p).length;
    return { label: p.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' }), logged, closed, backlog: logged - closed };
  });
  return rows;
}

const niceMax = n => { const step = n > 100 ? 100 : 10; return Math.ceil((n * 1.1) / step) * step; };

function rewriteBurndown(xml, rows) {
  const strCache = vals => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = vals => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numCache>`;
  const cats = rows.map(r => r.label);
  const series = [
    ['Closed base', rows.map(r => r.closed)],
    ['Open backlog', rows.map(r => r.backlog)],
    ['Cumulative logged', rows.map(r => r.logged)],
    ['Cumulative closed/resolved', rows.map(r => r.closed)],
  ];
  // Document order of caches: per series -> tx(str), cat(str), val(num).
  const queue = series.flatMap(([name, vals]) => [strCache([name]), strCache(cats), numCache(vals)]);
  let i = 0;
  xml = xml.replace(/<c:(?:str|num)Cache>[\s\S]*?<\/c:(?:str|num)Cache>/g, () => queue[i++]);
  if (i !== queue.length) throw new Error(`burndown chart cache count mismatch (${i} vs ${queue.length})`);
  const max = niceMax(Math.max(...rows.map(r => r.logged), 1));
  xml = xml.replace(/<c:max val="\d+"\/>/, `<c:max val="${max}"/>`)
           .replace(/<c:dLbl><c:idx val="\d+"\/>/g, `<c:dLbl><c:idx val="${rows.length - 1}"/>`)
           .replace(/<c:externalData[\s\S]*?<\/c:externalData>/, '');
  return xml;
}

// ── Critical defects table ───────────────────────────────────────────────────
function loadNotes() {
  try { return JSON.parse(fs.readFileSync(NOTES_PATH, 'utf8')); } catch { return {}; }
}

function fillRows(xml, data) {
  const rows = xml.match(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g) || [];
  const tpl = rows.find(r => r.includes('{{CRITID}}'));
  if (!tpl) return xml;
  const notes = loadNotes();
  const crit = Object.entries(data.bugs ?? {})
    .flatMap(([ws, arr]) => arr.filter(isCrit).map(b => ({ ws, b })));
  const shown = crit.slice(0, MAX_CRIT_ROWS);
  const out = shown.map(({ ws, b }, n) => {
    const note = notes[String(b.id)] || {};
    const vals = { CRITWS: ws, CRITID: `#${b.id}`, CRITOWNER: note.owner || 'TBD', CRITTARGET: note.targetDate || 'TBD', CRITCOMMENT: note.comments || 'TBD' };
    let r = tpl;
    for (const [k, v] of Object.entries(vals)) r = r.split(`{{${k}}}`).join(esc(v));
    return r.replace(/(a16:rowId[^>]*val=")\d+"/, `$1${700000 + n}"`);
  });
  let body = out.join('');
  if (crit.length > shown.length) {
    body += tpl.replace('{{CRITWS}}', '').replace('{{CRITID}}', '').replace('{{CRITOWNER}}', '').replace('{{CRITTARGET}}', '')
      .replace('{{CRITCOMMENT}}', esc(`+${crit.length - shown.length} more critical defect(s) not shown`))
      .replace(/(a16:rowId[^>]*val=")\d+"/, '$1799999"');
  }
  if (!shown.length) {
    body = tpl.replace('{{CRITWS}}', '--').replace('{{CRITID}}', '--').replace('{{CRITOWNER}}', '--').replace('{{CRITTARGET}}', '--')
      .replace('{{CRITCOMMENT}}', 'No open critical (Sev 1) defects.');
  }
  return xml.replace(tpl, body);
}

// ── Insights bullets ─────────────────────────────────────────────────────────
function fillInsights(xml, data) {
  const para = xml.match(/<a:p>(?:(?!<\/a:p>)[\s\S])*\{\{AI_DASH_INSIGHTS\}\}[\s\S]*?<\/a:p>/)?.[0];
  if (!para) return xml;
  const lines = String(data._aiNarratives?.defectsDashboardInsights ?? '').split('\n').map(s => s.trim()).filter(Boolean);
  const paras = (lines.length ? lines : ['']).map(l => para.replace('{{AI_DASH_INSIGHTS}}', esc(l)));
  return xml.replace(para, paras.join(''));
}

export function apply(zip, data) {
  if (!zip.file(SLIDE)) return;

  if (zip.file(BURNDOWN)) {
    zip.file(BURNDOWN, rewriteBurndown(zip.file(BURNDOWN).asText(), burndownSeries(data)));
  }
  let xml = zip.file(SLIDE).asText();
  xml = xml.replace(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*name="SeverityTiles"[\s\S]*?<\/p:sp>/, buildSeverityShapes(severityCounts(data)));
  xml = fillRows(xml, data);
  xml = fillInsights(xml, data);
  zip.file(SLIDE, xml);

  const m = dashboardMetrics(data);
  console.log(`  Defects dashboard: ${m.total} total / ${m.open} open / ${m.old} open >14d / ${m.crit} critical open / ${m.unassigned} unassigned`);
}

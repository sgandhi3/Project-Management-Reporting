// MMO-specific: renders the two "defect burndown" line charts (PDM, Benefits)
// that live on the PDM/Benefits Testing Defect Report slides, and swaps them
// into the generated deck each run.
//
// The chart plots cumulative defects logged vs. closed/resolved, by week,
// since the earliest defect's createdDate — both dates ADO tracks natively
// per-bug (System.CreatedDate / Microsoft.VSTS.Common.ClosedDate), so this
// is exact, not an approximation, and needs no separate history log.
//
// Rendered as native SVG embedded via the modern OOXML svgBlip extension
// (PowerPoint 2016+/365) rather than a rasterized PNG, so no canvas/image
// native dependency is needed. The template's original PNG placeholder
// (image9.png on the PDM defect slide, image10.png on the Benefits defect
// slide) stays in place as the fallback image for anything that doesn't
// understand svgBlip; only modern PowerPoint's fallback path would ever see
// it, and only as a static (stale) picture.
//
// Isolated here like _dynamic-benefits.js — MMO-specific slide surgery, not
// a generic reporting concern.

const CHART_TARGETS = [
  { slide: 'ppt/slides/slide4.xml', rels: 'ppt/slides/_rels/slide4.xml.rels', imageFile: 'image9.png',  workstream: 'PDM',      svgName: 'burndownPDM.svg',      title: 'PDM Defect Burndown' },
  { slide: 'ppt/slides/slide6.xml', rels: 'ppt/slides/_rels/slide6.xml.rels', imageFile: 'image10.png', workstream: 'Benefits', svgName: 'burndownBenefits.svg', title: 'Benefits Defect Burndown' },
];

const toDate  = s => new Date(`${s}T00:00:00Z`);
const addDays = (d, n) => { const r = new Date(d); r.setUTCDate(r.getUTCDate() + n); return r; };
const fmtTick = d => d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' });
const fmtLong = d => d.toLocaleDateString('en-US', { month: 'short', day: '2-digit', year: 'numeric', timeZone: 'UTC' });

function buildSeries(bugs) {
  const createdDates = bugs.map(b => b.createdDate).filter(Boolean).sort();
  if (!createdDates.length) return null;

  const start = toDate(createdDates[0]);
  const today = toDate(new Date().toLocaleDateString('en-CA', { timeZone: 'America/New_York' }));

  const ticks = [];
  for (let d = new Date(start); d < today; d = addDays(d, 7)) ticks.push(new Date(d));
  ticks.push(new Date(today));

  const points = ticks.map(t => ({
    date:   t,
    logged: bugs.filter(b => b.createdDate && toDate(b.createdDate) <= t).length,
    closed: bugs.filter(b => b.closedDate  && toDate(b.closedDate)  <= t).length,
  }));

  return { start, end: today, points };
}

function niceMax(n) {
  if (n <= 10) return 10;
  const magnitude = 10 ** Math.floor(Math.log10(n));
  const step = magnitude >= 100 ? magnitude / 2 : magnitude;
  return Math.ceil(n / step) * step;
}

function renderSvg(title, series) {
  const W = 900, H = 495;
  const padL = 60, padR = 60, padT = 90, padB = 80;
  const plotW = W - padL - padR, plotH = H - padT - padB;

  const maxVal = niceMax(Math.max(...series.points.map(p => p.logged), 1));
  const x = i => padL + (series.points.length > 1 ? (i / (series.points.length - 1)) * plotW : 0);
  const y = v => padT + plotH - (v / maxVal) * plotH;

  const loggedPts = series.points.map((p, i) => [x(i), y(p.logged)]);
  const closedPts = series.points.map((p, i) => [x(i), y(p.closed)]);
  const linePath  = pts => pts.map(([px, py], i) => `${i === 0 ? 'M' : 'L'}${px.toFixed(1)},${py.toFixed(1)}`).join(' ');
  const areaPath  = `${linePath(loggedPts)} L${closedPts[closedPts.length - 1][0].toFixed(1)},${closedPts[closedPts.length - 1][1].toFixed(1)} ${closedPts.slice().reverse().map(([px, py]) => `L${px.toFixed(1)},${py.toFixed(1)}`).join(' ')} Z`;

  const ySteps = 4;
  const gridLines = Array.from({ length: ySteps + 1 }, (_, i) => {
    const v = Math.round((maxVal / ySteps) * i);
    const yy = y(v);
    return `<line x1="${padL}" y1="${yy.toFixed(1)}" x2="${W - padR}" y2="${yy.toFixed(1)}" stroke="#e5e7eb" stroke-width="1"/>
      <text x="${padL - 12}" y="${(yy + 4).toFixed(1)}" text-anchor="end" font-size="13" fill="#374151" font-family="Arial, Helvetica, sans-serif">${v}</text>`;
  }).join('\n');

  const xLabels = series.points.map((p, i) => {
    const px = x(i);
    return `<text x="${px.toFixed(1)}" y="${H - padB + 22}" text-anchor="end" font-size="12" fill="#374151" font-family="Arial, Helvetica, sans-serif" transform="rotate(-30 ${px.toFixed(1)} ${H - padB + 22})">${fmtTick(p.date)}</text>`;
  }).join('\n');

  const dots = (pts, color) => pts.map(([px, py]) => `<circle cx="${px.toFixed(1)}" cy="${py.toFixed(1)}" r="4" fill="${color}"/>`).join('');

  const lastLogged = series.points[series.points.length - 1].logged;
  const lastClosed = series.points[series.points.length - 1].closed;
  const lastX = loggedPts[loggedPts.length - 1][0];

  return `<svg viewBox="0 0 ${W} ${H}" xmlns="http://www.w3.org/2000/svg" font-family="Arial, Helvetica, sans-serif">
    <rect x="0" y="0" width="${W}" height="${H}" fill="#ffffff"/>
    <text x="20" y="34" font-size="21" font-weight="bold" fill="#111827">${title} — ${fmtLong(series.start)} to ${fmtLong(series.end)}</text>
    <circle cx="30" cy="58" r="6" fill="#dc2626"/>
    <text x="42" y="63" font-size="14" fill="#374151">Cumulative Defects Logged</text>
    <circle cx="260" cy="58" r="6" fill="#16a34a"/>
    <text x="272" y="63" font-size="14" fill="#374151">Cumulative Defects Closed/Resolved</text>
    <rect x="530" y="52" width="14" height="14" fill="#fee2e2" stroke="#dc2626" stroke-width="1"/>
    <text x="550" y="63" font-size="14" fill="#374151">Open Backlog</text>
    <line x1="${padL}" y1="${padT + plotH}" x2="${W - padR}" y2="${padT + plotH}" stroke="#9ca3af" stroke-width="1.5"/>
    <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${padT + plotH}" stroke="#9ca3af" stroke-width="1.5"/>
    ${gridLines}
    <path d="${areaPath}" fill="#fee2e2" fill-opacity="0.7" stroke="none"/>
    <path d="${linePath(loggedPts)}" fill="none" stroke="#dc2626" stroke-width="2.5"/>
    <path d="${linePath(closedPts)}" fill="none" stroke="#16a34a" stroke-width="2.5"/>
    ${dots(loggedPts, '#dc2626')}
    ${dots(closedPts, '#16a34a')}
    <text x="${(lastX + 10).toFixed(1)}" y="${(y(lastLogged) + 5).toFixed(1)}" font-size="16" font-weight="bold" fill="#dc2626">${lastLogged}</text>
    <text x="${(lastX + 10).toFixed(1)}" y="${(y(lastClosed) + 5).toFixed(1)}" font-size="16" font-weight="bold" fill="#16a34a">${lastClosed}</text>
    ${xLabels}
    <text x="20" y="${padT + plotH / 2}" font-size="13" fill="#374151" transform="rotate(-90 20 ${padT + plotH / 2})" text-anchor="middle">Cumulative Defect Count</text>
  </svg>`;
}

// A 1x1 transparent PNG — used only as the pre-2016-Office fallback image
// behind the svgBlip extension; modern PowerPoint renders the SVG itself.
const BLANK_PNG_BASE64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

function ensureSvgContentType(zip) {
  const path = '[Content_Types].xml';
  let xml = zip.file(path).asText();
  if (xml.includes('Extension="svg"')) return;
  xml = xml.replace('</Types>', '<Default Extension="svg" ContentType="image/svg+xml"/></Types>');
  zip.file(path, xml);
}

function injectChart(zip, target, svgContent) {
  const relsPath = target.rels;
  if (!zip.file(relsPath) || !zip.file(target.slide)) return false;

  let relsXml = zip.file(relsPath).asText();
  const relMatch = relsXml.match(new RegExp(`<Relationship Id="(rId\\d+)"[^>]*Target="\\.\\./media/${target.imageFile}"[^>]*/>`));
  if (!relMatch) {
    console.warn(`  ⚠  _burndown-charts: couldn't find ${target.imageFile} relationship in ${relsPath} — skipping`);
    return false;
  }
  const pngRelId = relMatch[1];

  // Add a fresh relationship pointing at this run's rendered SVG.
  const existingIds = [...relsXml.matchAll(/Id="rId(\d+)"/g)].map(m => Number(m[1]));
  const svgRelId = `rId${Math.max(0, ...existingIds) + 1}`;
  const svgMediaName = target.svgName.replace(/\.svg$/, `-${Date.now()}.svg`);
  relsXml = relsXml.replace(
    '</Relationships>',
    `<Relationship Id="${svgRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${svgMediaName}"/></Relationships>`
  );
  zip.file(relsPath, relsXml);
  zip.file(`ppt/media/${svgMediaName}`, svgContent);

  // Point the existing PNG relationship at a blank placeholder — it's only
  // ever shown as the fallback for pre-2016 Office, and keeping the old
  // template's stale chart there would be misleading if it ever is.
  zip.file(`ppt/media/${target.imageFile}`, Buffer.from(BLANK_PNG_BASE64, 'base64'));

  let slideXml = zip.file(target.slide).asText();
  const blipRe = new RegExp(`<a:blip r:embed="${pngRelId}"\\s*/>`);
  if (!blipRe.test(slideXml)) {
    console.warn(`  ⚠  _burndown-charts: couldn't find <a:blip> for ${pngRelId} on ${target.slide} — skipping`);
    return false;
  }
  slideXml = slideXml.replace(
    blipRe,
    `<a:blip r:embed="${pngRelId}"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${svgRelId}"/></a:ext></a:extLst></a:blip>`
  );
  zip.file(target.slide, slideXml);
  return true;
}

export function apply(zip, data) {
  ensureSvgContentType(zip);

  for (const target of CHART_TARGETS) {
    const bugs = [...(data.bugs?.[target.workstream] ?? []), ...(data.closedBugs?.[target.workstream] ?? [])];
    const series = buildSeries(bugs);
    if (!series) {
      console.warn(`  ⚠  _burndown-charts: no dated ${target.workstream} defects — skipping chart`);
      continue;
    }
    const svg = renderSvg(target.title, series);
    if (injectChart(zip, target, svg)) {
      console.log(`  Burndown chart: ${target.workstream} (${series.points.length} weekly points, ${series.points.at(-1).logged} logged / ${series.points.at(-1).closed} closed)`);
    }
  }
}

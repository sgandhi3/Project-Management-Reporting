import fetch from 'node-fetch';
import ExcelJS from 'exceljs';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

dotenv.config({ path: path.join(path.dirname(fileURLToPath(import.meta.url)), '.env') });

const ORG     = process.env.ADO_ORG;
const PROJECT = process.env.ADO_PROJECT;
const TOKEN   = process.env.ADO_PAT;
const AUTH    = { 'Authorization': `Basic ${Buffer.from(':' + TOKEN).toString('base64')}`, 'Content-Type': 'application/json', 'Accept': 'application/json' };
const enc     = s => encodeURIComponent(s);
const base    = () => `https://dev.azure.com/${ORG}/${enc(PROJECT)}/_apis`;

async function get(url) {
  const res = await fetch(url, { headers: AUTH });
  if (!res.ok) throw new Error(`ADO ${res.status}: ${url}\n${(await res.text()).slice(0, 400)}`);
  return res.json();
}

async function post(url, body) {
  const res = await fetch(url, { method: 'POST', headers: AUTH, body: JSON.stringify(body) });
  if (!res.ok) throw new Error(`ADO ${res.status}: ${url}\n${(await res.text()).slice(0, 400)}`);
  return res.json();
}

function stripHtml(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

async function fetchLatestComment(id) {
  try {
    const data = await get(`${base()}/wit/workItems/${id}/comments?api-version=7.1-preview.3&$top=100`);
    const comments = data.comments || [];
    if (!comments.length) return '';
    // Sort by createdDate descending and take the latest
    comments.sort((a, b) => new Date(b.createdDate) - new Date(a.createdDate));
    const latest = comments[0];
    const author = latest.createdBy?.displayName || '';
    const date   = (latest.createdDate || '').slice(0, 10);
    const text   = stripHtml(latest.text || '');
    return `[${date} - ${author}] ${text}`;
  } catch (e) {
    console.warn(`  ⚠  Comments for ${id}: ${e.message.split('\n')[0]}`);
    return '';
  }
}

async function main() {
  // Query all open bugs project-wide (no area path filter)
  const wiql = `SELECT [System.Id], [System.Title], [System.State], [System.AreaPath]
    FROM WorkItems
    WHERE [System.WorkItemType] = 'Bug'
      AND [System.TeamProject] = '${PROJECT}'
      AND [System.State] <> 'Closed'
      AND [System.State] <> 'Resolved'
    ORDER BY [System.AreaPath] ASC, [System.Id] ASC`;

  console.log('Querying open bugs...');
  const queryData = await post(`${base()}/wit/wiql?api-version=7.0`, { query: wiql });
  const ids = (queryData.workItems || []).map(w => w.id);
  console.log(`Found ${ids.length} open bugs. Fetching details...`);

  // Fetch work item details in chunks of 200
  const items = [];
  for (let i = 0; i < ids.length; i += 200) {
    const chunk = ids.slice(i, i + 200).join(',');
    const res   = await get(`${base()}/wit/workitems?ids=${chunk}&api-version=7.0`);
    items.push(...(res.value || []));
  }

  // Fetch latest comment for each bug (with simple concurrency limit)
  const CONCURRENCY = 10;
  const rows = new Array(items.length);
  for (let i = 0; i < items.length; i += CONCURRENCY) {
    const batch = items.slice(i, i + CONCURRENCY);
    const results = await Promise.all(batch.map(async (wi, j) => {
      const idx = i + j;
      process.stdout.write(`\r  Fetching comments: ${idx + 1}/${items.length}`);
      const latestComment = await fetchLatestComment(wi.id);
      return {
        id:            wi.id,
        areaPath:      wi.fields['System.AreaPath']  || '',
        state:         wi.fields['System.State']     || '',
        title:         wi.fields['System.Title']     || '',
        triage:        '',
        triageComment: '',
        latestComment,
      };
    }));
    results.forEach((r, j) => { rows[i + j] = r; });
  }
  console.log('\nBuilding Excel...');

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Open Bugs');

  ws.columns = [
    { header: 'ADO ID',          key: 'id',            width: 10 },
    { header: 'Area Path',        key: 'areaPath',      width: 45 },
    { header: 'State',            key: 'state',         width: 15 },
    { header: 'Title',            key: 'title',         width: 60 },
    { header: 'Triage?',          key: 'triage',        width: 12 },
    { header: 'Triage Comment',   key: 'triageComment', width: 40 },
    { header: 'Latest Comment',   key: 'latestComment', width: 80 },
  ];

  // Header styling
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  headerRow.alignment = { vertical: 'middle' };

  for (const row of rows) {
    ws.addRow(row);
  }

  // Wrap and top-align text columns
  for (const col of ['areaPath', 'title', 'latestComment', 'triageComment']) {
    ws.getColumn(col).alignment = { wrapText: true, vertical: 'top' };
  }
  ws.getColumn('id').alignment    = { vertical: 'top' };
  ws.getColumn('state').alignment = { vertical: 'top' };
  ws.getColumn('triage').alignment = { vertical: 'top' };

  // Freeze header row
  ws.views = [{ state: 'frozen', ySplit: 1 }];

  const outPath = path.join(path.dirname(fileURLToPath(import.meta.url)), 'open-bugs.xlsx');
  await wb.xlsx.writeFile(outPath);
  console.log(`\nDone. ${rows.length} bugs exported to: ${outPath}`);
}

main().catch(e => { console.error(e.message); process.exit(1); });

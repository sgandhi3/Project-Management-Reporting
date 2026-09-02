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

const BUG_ID = 30974;

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

function extractIntroSentences(html) {
  const text = stripHtml(html);
  const tcPatterns = [
    /TC[\s\-#:]?\d+/i, /test case/i, /step\s+no/i, /step \d+/i,
    /expected result/i, /actual result/i, /bug filed on/i,
    /^\d{1,2}\/\d{1,2}\/\d{4}/m
  ];
  // Split on blank lines first (paragraphs), then sentences within each paragraph
  const paragraphs = text.split(/\n\n+/);
  const intro = [];
  for (const para of paragraphs) {
    if (tcPatterns.some(p => p.test(para))) break;
    const trimmed = para.trim();
    if (trimmed) intro.push(trimmed);
    if (intro.length >= 3) break;
  }
  if (!intro.length) {
    // fallback: take first 3 sentences
    const sentences = text.split(/(?<=[.!?])\s+/);
    for (const s of sentences) {
      if (tcPatterns.some(p => p.test(s))) break;
      intro.push(s);
      if (intro.length >= 3) break;
    }
  }
  return intro.join('\n\n').trim() || text.substring(0, 300);
}

async function main() {
  const url = `https://dev.azure.com/${ORG}/${enc(PROJECT)}/_apis/wit/workitems/${BUG_ID}?$expand=all&api-version=7.0`;
  console.log(`Fetching bug ${BUG_ID}...`);
  const res = await fetch(url, { headers: AUTH });
  if (!res.ok) {
    const err = await res.text();
    throw new Error(`ADO ${res.status}: ${err.slice(0, 400)}`);
  }
  const wi = await res.json();
  const f  = wi.fields;

  const id        = wi.id;
  const areaPath  = f['System.AreaPath'] || '';
  const title     = f['System.Title'] || '';
  const state     = f['System.State'] || '';
  const reproRaw  = f['Microsoft.VSTS.TCM.ReproSteps'] || '';
  const repro     = extractIntroSentences(reproRaw);

  console.log(`\nID:        ${id}`);
  console.log(`Area Path: ${areaPath}`);
  console.log(`Title:     ${title}`);
  console.log(`State:     ${state}`);
  console.log(`Repro:     ${repro}\n`);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Bug Details');

  ws.columns = [
    { header: 'ID',          key: 'id',        width: 10 },
    { header: 'Area Path',   key: 'areaPath',  width: 40 },
    { header: 'Title',       key: 'title',     width: 60 },
    { header: 'State',       key: 'state',     width: 15 },
    { header: 'Repro Steps', key: 'repro',     width: 80 },
  ];

  // Style header row
  ws.getRow(1).font = { bold: true };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF4472C4' } };
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

  ws.addRow({ id, areaPath, title, state, repro });

  // Wrap text in repro column
  ws.getColumn('repro').alignment = { wrapText: true, vertical: 'top' };
  ws.getColumn('areaPath').alignment = { vertical: 'top' };
  ws.getColumn('title').alignment = { vertical: 'top' };

  const outPath = path.join(path.dirname(fileURLToPath(import.meta.url)), `bug-${BUG_ID}.xlsx`);
  await wb.xlsx.writeFile(outPath);
  console.log(`Excel saved to: ${outPath}`);
}

main().catch(e => { console.error(e.message); process.exit(1); });

// PowerPoint output extension
//
// Reads the VARIABLE_MAP from variables.js and replaces every {{TOKEN}} in
// the template deck with its computed value, then writes the populated file.
//
// Config (via .env or CLI):
//   PPTX_TEMPLATE — path to the source .pptx file (default: temp.pptx in project root)
//   --out <path>  — where to write the output file (default: Report_YYYY-MM-DD.pptx)
//
// To create a different output extension (e.g. Excel), copy this file, implement
// the same generate(data) export, and set OUTPUT_FORMAT=excel in your .env.

import fs     from 'fs';
import { existsSync, readFileSync } from 'fs';
import path   from 'path';
import PizZip from 'pizzip';
import { fileURLToPath } from 'url';
import { VARIABLE_MAP } from '../variables.js';

const __dirname2 = path.dirname(fileURLToPath(import.meta.url));
const UI_CONFIG_PATH = path.join(__dirname2, '..', 'ui-config.json');

const args      = process.argv.slice(2);
const getArg    = flag => { const i = args.indexOf(flag); return i !== -1 ? args[i + 1] : null; };
const TEMPLATE  = process.env.PPTX_TEMPLATE || path.join(process.cwd(), 'temp.pptx');
const outputArg = getArg('--out') || path.join(process.cwd(), `Report_${new Date().toISOString().slice(0, 10)}.pptx`);

// PowerPoint sometimes splits a {{TOKEN}} across multiple XML text runs when
// spell-check or a capital letter mid-token triggers a run break. This heals
// those splits before the replacement pass runs.
function fixSplitTokens(xml) {
  const rb = '(?:<\\/a:t><\\/a:r>[\\s\\S]{0,400}?<a:r>[\\s\\S]{0,400}?<a:t[^>]*>)?';
  const pattern = new RegExp(
    '\\{' + rb + '\\{' +
    '(?:[A-Za-z0-9]' + rb + ')+' +
    '\\}' + rb + '\\}',
    'g'
  );
  return xml.replace(pattern, m =>
    m.replace(/<\/a:t><\/a:r>[\s\S]{0,400}?<a:r>[\s\S]{0,400}?<a:t[^>]*>/g, '')
  );
}

export async function generate(data) {
  if (!fs.existsSync(TEMPLATE)) {
    console.warn(`\n⚠  Template not found: ${TEMPLATE} — skipping report generation.`);
    return;
  }

  const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));

  // Use UI-configured mappings if present, otherwise fall back to variables.js VARIABLE_MAP
  let effectiveMap = VARIABLE_MAP;

  if (existsSync(UI_CONFIG_PATH)) {
    const uiConfig = JSON.parse(readFileSync(UI_CONFIG_PATH, 'utf8'));
    const uiMappings = uiConfig.variableMappings || [];
    if (uiMappings.length > 0) {
      effectiveMap = {};
      for (const { token, path: expr } of uiMappings) {
        effectiveMap[token] = new Function('d', `try { return ${expr}; } catch { return ''; }`);
      }
    }
  }

  // Project-specific row-injection hooks (e.g. _dynamic-benefits.js) mutate
  // the zip's slide XML directly and may add extra token getters to
  // effectiveMap — optional, so branches/projects without such a file are
  // unaffected. Must run before slide XML is read below.
  try {
    const dynamicBenefits = await import('./_dynamic-benefits.js');
    dynamicBenefits.apply(zip, data, effectiveMap);
  } catch (e) {
    if (e.code !== 'ERR_MODULE_NOT_FOUND') throw e;
  }

  const replacements = {};
  for (const [key, getter] of Object.entries(effectiveMap)) {
    try {
      replacements[`{{${key}}}`] = String(getter(data) ?? '');
    } catch {
      replacements[`{{${key}}}`] = '';
    }
  }

  const slideNames = Object.keys(zip.files).filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n));
  for (const name of slideNames) {
    let xml = zip.file(name).asText();
    xml = fixSplitTokens(xml);
    for (const [token, value] of Object.entries(replacements)) {
      xml = xml.split(token).join(value);
    }
    zip.file(name, xml);
  }

  const buf = zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' });
  fs.writeFileSync(outputArg, buf);
  console.log(`\nReport saved → ${outputArg}`);

  // Let downstream extensions (e.g. sharepoint) know what was generated
  data._generatedFiles = data._generatedFiles || [];
  data._generatedFiles.push(outputArg);
}

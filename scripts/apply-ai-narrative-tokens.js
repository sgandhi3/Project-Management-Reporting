// One-time (re-runnable) template maintenance script.
//
// Finds specific hardcoded narrative paragraphs in temp2.pptx (the ones with
// stale facts — dates, counts, plan names) and replaces each one with a single
// {{AI_TOKEN}} placeholder, collapsing multi-run paragraphs into one run so the
// normal token-replacement pass in extensions/ppt.js can fill it in.
//
// Run again after editing TARGETS below (e.g. if the template is regenerated
// and the hardcoded sentences need to be re-tokenized):
//   node scripts/apply-ai-narrative-tokens.js [path-to-pptx]
import fs from 'fs';
import PizZip from 'pizzip';

const TEMPLATE = process.argv[2] || 'temp2.pptx';

// Each target is matched by a short, unique leading substring (avoids
// hand-typing full paragraphs and hitting invisible-character mismatches like
// non-breaking spaces) — the whole paragraph is then replaced regardless of
// its exact trailing whitespace/unicode quirks.
const TARGETS = {
  'ppt/slides/slide2.xml': [
    ['Overall, SIT execution status is behind plan.', '{{AI_OVERALL_STATUS}}'],
    ['PDM Iteration 2.1 execution concluded on 8/7', '{{AI_PDM_ITERATION_UPDATE}}'],
    ['Majority of the PDM defects are configuration and data defects.', '{{AI_PDM_DEFECT_SUMMARY}}'],
    ['First round of testing for priority benefits has completed', '{{AI_BENEFITS_UPDATE}}'],
  ],
  'ppt/slides/slide3.xml': [
    ['Overall PDM Cursory Review testing is in Off-Track status', '{{AI_PDM_CURSORY_STATUS}}'],
  ],
  'ppt/slides/slide4.xml': [
    ['Overall Benefits SIT testing is At-Risk.', '{{AI_BENEFITS_STATUS}}'],
  ],
  'ppt/slides/slide5.xml': [
    ['Overall Enrollment SIT testing is on track.', '{{AI_ENROLLMENT_STATUS}}'],
  ],
  'ppt/slides/slide6.xml': [
    ['Overall EDI SIT testing is', '{{AI_EDI_STATUS}}'],
  ],
  'ppt/slides/slide7.xml': [
    ['2,054 invalid CPIMS orgs are directly cascading', '{{AI_PDM_DQ_DEFECT_DETAIL_1}}'],
    ['3,056 rejected org locations, 7,825 practitioner roles', '{{AI_PDM_DQ_DEFECT_DETAIL_2}}'],
    ['All 9 defects are tied to Rosters and related to data.', '{{AI_PDM_DM_DEFECT_DETAIL}}'],
  ],
};

function paraText(pXml) {
  return [...pXml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)].map(m => m[1]).join('');
}

// Collapses a paragraph's runs into one, using the first run's formatting,
// with the new token as its text.
function collapseParagraph(pXml, token) {
  const runs = [...pXml.matchAll(/<a:r>[\s\S]*?<\/a:r>/g)];
  if (!runs.length) return null;
  const firstRun = runs[0][0];
  const newFirstRun = firstRun.replace(/<a:t([^>]*)>[^<]*<\/a:t>/, `<a:t$1>${token}</a:t>`);
  let result = pXml.slice(0, runs[0].index) + newFirstRun;
  const afterFirst = pXml.slice(runs[0].index + firstRun.length);
  // strip any remaining <a:r>...</a:r> runs from the tail, keep everything else (e.g. closing </a:p>)
  result += afterFirst.replace(/<a:r>[\s\S]*?<\/a:r>/g, '');
  return result;
}

function applyToSlide(xml, targets) {
  let changed = 0;
  for (const [prefix, token] of targets) {
    const paras = [...xml.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)];
    const matches = paras.filter(m => paraText(m[0]).startsWith(prefix));
    if (matches.length === 0) {
      console.warn(`  ⚠  Not found (skipping): ${JSON.stringify(prefix)}...`);
      continue;
    }
    if (matches.length > 1) {
      console.warn(`  ⚠  Ambiguous — ${matches.length} paragraphs start with (skipping): ${JSON.stringify(prefix)}...`);
      continue;
    }
    const match = matches[0];
    const collapsed = collapseParagraph(match[0], token);
    if (!collapsed) {
      console.warn(`  ⚠  No runs in matched paragraph — skipping ${token}`);
      continue;
    }
    xml = xml.slice(0, match.index) + collapsed + xml.slice(match.index + match[0].length);
    changed++;
  }
  return { xml, changed };
}

const zip = new PizZip(fs.readFileSync(TEMPLATE, 'binary'));
let total = 0;
for (const [slideFile, targets] of Object.entries(TARGETS)) {
  const file = zip.file(slideFile);
  if (!file) { console.warn(`⚠  ${slideFile} not found in ${TEMPLATE}`); continue; }
  console.log(`${slideFile}:`);
  const { xml, changed } = applyToSlide(file.asText(), targets);
  zip.file(slideFile, xml);
  total += changed;
  console.log(`  ${changed}/${targets.length} tokenized`);
}

fs.writeFileSync(TEMPLATE, zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' }));
console.log(`\n${total} paragraph(s) tokenized in ${TEMPLATE}`);

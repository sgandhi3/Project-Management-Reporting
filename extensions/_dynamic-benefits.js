// MMO-specific: dynamically inserts PPTX table rows on the Benefits
// execution slide's per-plan breakdown table for Benefits plans that have
// started executing but aren't one of the 5 plans with a permanent template
// row (Signature HMO, Access PPO, Premium PPO, MMEGWP PPO CWRU MAPD,
// Classic HMO NEOH — see FIXED_BENEFIT_ROW_KEYS in variables.js). Plain
// {{TOKEN}} substitution can't do this — it only fills existing cells, it
// can't add table rows — so this runs BEFORE the normal substitution pass
// in extensions/ppt.js, mutating the raw slide XML directly.
//
// As of the 09/2026 template refresh this targets slide5 (previously
// slide4, before PDM Execution/Defects and Benefits Execution/Defects were
// split into their own slides) and clones the "Classic HMO NEOH" row — now
// the last of the 5 fixed rows, immediately before Grand Total.
//
// The executive summary (slide 2) does NOT get a dynamic row — it shows the
// combined Priority + active-2026 total directly via ACTIVEBEN* tokens, so
// it never needs new rows added.
//
// This is intentionally isolated here rather than folded into ppt.js/
// variables.js's generic engine — it's specific to this project's Priority
// vs. non-Priority Benefit plan structure and won't apply to other projects
// using this template. Lives on the mmo-benefits-active-suite-fix branch;
// don't merge it into a shared/generic base for other teams.
//
// KNOWN LIMITATION: PowerPoint doesn't auto-reflow a slide when a table
// grows. The shift thresholds below were measured by hand against the
// current template (see git history for the measurements). If the
// template is redesigned — shapes moved, resized, or added near this
// table — these hardcoded values will need re-measuring. This has only
// been verified at the XML/text level, not by rendering in PowerPoint —
// open a generated report and visually check the Benefits execution slide
// after this runs, especially with more than one or two extra active plans.
// A newly-inserted row also has no real Target Completion Date — it
// inherits whatever literal date string is on the cloned Classic HMO NEOH
// row, which will read wrong until someone edits it by hand.

import { getExtraActiveBenefitPlans, getNotStartedBenefitsRow } from '../variables.js';
import { benefitsDefectPlan, BENEFITS_PLAN_LABELS } from './_defect-classification.js';

const FIXED_DEFECT_PLANS = new Set(['Signature HMO', 'Access PPO', 'Premium PPO', 'Classic HMO NEOH', 'MMEGWP PPO CWRU MAPD']);
const norm = s => String(s).toLowerCase().replace(/[^a-z0-9]/g, '');

// The classifier label (e.g. "MMEGWP 3 Tier UH") matching an ADO suite plan
// label (e.g. "Mmegwp 3tier Uh New"): same letters/digits ignoring case,
// spacing and a trailing "New".
function defectPlanForExecLabel(label) {
  const n = norm(label);
  return BENEFITS_PLAN_LABELS.find(l => n.includes(norm(l))) ?? null;
}

function extractRows(frameXml) {
  return [...frameXml.matchAll(/<a:tr h="\d+">[\s\S]*?<\/a:tr>/g)].map(m => m[0]);
}

function rowHeight(rowXml) {
  return parseInt(rowXml.match(/<a:tr h="(\d+)">/)[1]);
}

function cloneRowWithTokens(templateRow, nameFrom, nameTo, tokenRenames) {
  let row = templateRow.split(`>${nameFrom}<`).join(`>${nameTo}<`);
  for (const [oldTok, newTok] of Object.entries(tokenRenames)) {
    row = row.split(oldTok).join(newTok);
  }
  return row;
}

function growExtent(xml, deltaY) {
  return xml.replace(
    /(<p:xfrm><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy=")(\d+)("\/><\/p:xfrm>)/,
    (_full, pre, cy, post) => `${pre}${parseInt(cy) + deltaY}${post}`
  );
}

// Shifts every top-level shape's <p:xfrm> whose y-offset is >= thresholdY
// down by deltaY. Grows (rather than shifts) any shape whose bottom edge
// falls within ~250000 EMU of a value in growPanelBottomsNear — a
// background panel that encloses the growing table and needs to get taller
// along with it instead of moving down.
function reflow(xml, { thresholdY, deltaY, growPanelBottomsNear = [] }) {
  // Regular shapes (<p:sp>) use <a:xfrm>; graphicFrames use <p:xfrm> — match both.
  return xml.replace(
    /(<[ap]:xfrm[^>]*>)<a:off x="(\d+)" y="(\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/g,
    (full, xfrmTag, x, y, cx, cy) => {
      const yNum = parseInt(y), cyNum = parseInt(cy);
      const bottom = yNum + cyNum;
      if (yNum >= thresholdY) {
        return `${xfrmTag}<a:off x="${x}" y="${yNum + deltaY}"/><a:ext cx="${cx}" cy="${cy}"/>`;
      }
      if (growPanelBottomsNear.some(target => Math.abs(bottom - target) < 250000)) {
        return `${xfrmTag}<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cyNum + deltaY}"/>`;
      }
      return full;
    }
  );
}

// ── Benefits execution slide: per-plan breakdown table ─────────────────────

function injectSlide5(xml, extraPlans, tokenMap, notStarted) {
  const anchorIdx = xml.indexOf('Classic HMO NEOH');
  const gfStart = xml.lastIndexOf('<p:graphicFrame>', anchorIdx);
  const gfEnd = xml.indexOf('</p:graphicFrame>', anchorIdx) + '</p:graphicFrame>'.length;
  const frame = xml.slice(gfStart, gfEnd);

  const rows = extractRows(frame);
  const templateRow = rows.find(r => r.includes('Classic HMO NEOH'));
  const grandTotalRow = rows.find(r => r.includes('Grand Total'));
  if (!templateRow || !grandTotalRow) {
    console.warn('  ⚠  _dynamic-benefits: template/Grand Total row not found — skipping row injection');
    return xml;
  }

  const addedHeight = rowHeight(templateRow) * (extraPlans.length + (notStarted ? 1 : 0));

  let newRowsXml = '';
  extraPlans.forEach((plan, i) => {
    const prefix = `DYNBENE${i}`;
    tokenMap[`${prefix}TTC`]  = () => plan.stats.planned;
    tokenMap[`${prefix}ETC`]  = () => plan.stats.executed;
    tokenMap[`${prefix}PTC`]  = () => plan.stats.passed;
    tokenMap[`${prefix}FTC`]  = () => plan.stats.failed;
    tokenMap[`${prefix}IPTC`] = () => plan.stats.inProgress;
    tokenMap[`${prefix}BTC`]  = () => plan.stats.blocked;
    tokenMap[`${prefix}NSTC`] = () => plan.stats.notStarted;
    const defPlan = defectPlanForExecLabel(plan.label);
    tokenMap[`${prefix}DEF`]  = d => defPlan ? (d.bugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === defPlan).length : 0;

    newRowsXml += cloneRowWithTokens(templateRow, 'Classic HMO NEOH', plan.label, {
      '{{BENENEOHTTC}}':  `{{${prefix}TTC}}`,
      '{{BENENEOHETC}}':  `{{${prefix}ETC}}`,
      '{{BENENEOHPTC}}':  `{{${prefix}PTC}}`,
      '{{BENENEOHFTC}}':  `{{${prefix}FTC}}`,
      '{{BENENEOHIPTC}}': `{{${prefix}IPTC}}`,
      '{{BENENEOHBTC}}':  `{{${prefix}BTC}}`,
      '{{BENENEOHNSTC}}': `{{${prefix}NSTC}}`,
      '{{BENENEOHDEFOPEN}}': `{{${prefix}DEF}}`,
    });
  });

  // Remainder row so the table foots to the full Benefits total.
  if (notStarted) {
    tokenMap.DYNBENENSTTC  = () => notStarted.planned;
    tokenMap.DYNBENENSNSTC = () => notStarted.notStarted;
    tokenMap.DYNBENEZERO   = () => 0;
    let row = cloneRowWithTokens(templateRow, 'Classic HMO NEOH', `Not Started (${notStarted.plansLeft} plans)`, {
      '{{BENENEOHTTC}}':  '{{DYNBENENSTTC}}',
      '{{BENENEOHETC}}':  '{{DYNBENEZERO}}',
      '{{BENENEOHPTC}}':  '{{DYNBENEZERO}}',
      '{{BENENEOHFTC}}':  '{{DYNBENEZERO}}',
      '{{BENENEOHIPTC}}': '{{DYNBENEZERO}}',
      '{{BENENEOHBTC}}':  '{{DYNBENEZERO}}',
      '{{BENENEOHNSTC}}': '{{DYNBENENSNSTC}}',
      '{{BENENEOHDEFOPEN}}': '{{DYNBENEZERO}}',
    });
    newRowsXml += row.replace('>11/30/2026<', '>TBD<');
  }

  const insertAt = frame.indexOf(grandTotalRow);
  let newFrame = frame.slice(0, insertAt) + newRowsXml + frame.slice(insertAt);
  newFrame = growExtent(newFrame, addedHeight);

  let newXml = xml.slice(0, gfStart) + newFrame + xml.slice(gfEnd);

  // Adding rows pushes the footnote boxes off the bottom of the slide, so
  // instead of shifting by a hand-measured offset, re-fit: shrink the data
  // rows as needed, then stack the footnotes directly under the table.
  newXml = fitBenefitsSlide(newXml);

  return newXml;
}

const SLIDE_H = 6858000;
const FOOT_GAP = 40000;

function setBoxY(xml, name, y, cy) {
  return xml.replace(new RegExp(`(<p:sp>(?:(?!</p:sp>).)*?name="${name}".*?<a:off x="\\d+" y=")\\d+("/><a:ext cx="\\d+" cy=")\\d+`, 's'),
    (_m, a, b) => `${a}${y}${b}${cy}`);
}

function fitBenefitsSlide(xml) {
  const gfStart = xml.lastIndexOf('<p:graphicFrame>', xml.indexOf('Benefits SIT Review Summary'));
  const gfEnd = xml.indexOf('</p:graphicFrame>', gfStart) + '</p:graphicFrame>'.length;
  let frame = xml.slice(gfStart, gfEnd);
  const top = parseInt(frame.match(/<p:xfrm><a:off x="\d+" y="(\d+)"/)[1]);
  const rows = extractRows(frame);
  const heights = rows.map(rowHeight);
  const fixed = heights[0] + heights[1];               // title + header rows
  const nData = rows.length - 2;
  const footH = 280000 + 492443 + 184666 + FOOT_GAP * 3;
  const H = Math.max(230000, Math.min(340000, Math.floor((SLIDE_H - 60000 - footH - top - fixed) / nData)));
  rows.forEach((r, i) => { if (i >= 2) frame = frame.replace(r, r.replace(/<a:tr h="\d+">/, `<a:tr h="${H}">`)); });
  const tableH = fixed + H * nData;
  frame = frame.replace(/(<p:xfrm><a:off x="\d+" y="\d+"\/><a:ext cx="\d+" cy=")\d+/, `$1${tableH}`);
  let out = xml.slice(0, gfStart) + frame + xml.slice(gfEnd);
  // Totals now include every Benefits case, so the old "pending plans are not
  // reflected" note is stale.
  out = out.replace(/Note: Additional 2026 Benefit plans[^<]*/, 'Note: Totals include all Benefits test cases; plans that have not started executing are grouped in the Not Started row.');
  let y = top + tableH + FOOT_GAP;
  out = setBoxY(out, 'TextBox 17', y, 280000); y += 280000 + FOOT_GAP;
  out = setBoxY(out, 'TextBox 15', y, 492443); y += 492443 + FOOT_GAP;
  out = setBoxY(out, 'TextBox 9',  y, 184666);
  return out;
}

// ── Benefits defect slide: per-plan defect table ─────────────────────────
// Plans beyond the 5 fixed rows (e.g. MMEGWP 3 Tier UH) get their own row once
// they have any defects. The table grows, so the severity/aging tables below
// it shift down and the Key Insights box gives up the same height (its bottom
// edge stays fixed). Measured against the 10/2026 template.
const EMU_PX = 9525;

function shapeBlock(xml, name) {
  const re = new RegExp(`<p:(sp|graphicFrame)>(?:(?!</p:\\1>)[\\s\\S])*?name="${name}"[\\s\\S]*?</p:\\1>`);
  const m = xml.match(re);
  return m ? m[0] : null;
}

function shiftShape(xml, name, dy, dh = 0) {
  const block = shapeBlock(xml, name);
  if (!block) return xml;
  const nb = block.replace(/(<(?:a|p):xfrm[^>]*><a:off x="\d+" y=")(\d+)("\/><a:ext cx="\d+" cy=")(\d+)/,
    (_m, a, y, b, cy) => `${a}${parseInt(y) + dy}${b}${parseInt(cy) + dh}`);
  return xml.replace(block, nb);
}

function injectDefectSlide(xml, extraLabels, tokenMap) {
  const anchor = xml.indexOf('MMEGWP PPO CWRU MAPD');
  if (anchor < 0 || !extraLabels.length) return xml;
  const gfStart = xml.lastIndexOf('<p:graphicFrame>', anchor);
  const gfEnd = xml.indexOf('</p:graphicFrame>', anchor) + '</p:graphicFrame>'.length;
  let frame = xml.slice(gfStart, gfEnd);
  const rows = extractRows(frame);
  const templateRow = rows.find(r => r.includes('MMEGWP PPO CWRU MAPD'));
  const totalRow = rows.find(r => r.includes('>TOTAL<'));
  if (!templateRow || !totalRow) {
    console.warn('  ⚠  _dynamic-benefits: defect table template/TOTAL row not found — skipping defect row injection');
    return xml;
  }
  let newRows = '';
  extraLabels.forEach((label, i) => {
    const p = `DYNDEF${i}`;
    tokenMap[`${p}TOTAL`]    = d => [...(d.bugs?.Benefits ?? []), ...(d.closedBugs?.Benefits ?? [])].filter(b => benefitsDefectPlan(b) === label).length;
    tokenMap[`${p}OPEN`]     = d => (d.bugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === label).length;
    tokenMap[`${p}CLOSED`]   = d => (d.closedBugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === label && b.state === 'Closed').length;
    tokenMap[`${p}RESOLVED`] = d => (d.closedBugs?.Benefits ?? []).filter(b => benefitsDefectPlan(b) === label && b.state === 'Resolved').length;
    newRows += cloneRowWithTokens(templateRow, 'MMEGWP PPO CWRU MAPD', label, {
      '{{BENEMMEGWPDEFTOTAL}}':    `{{${p}TOTAL}}`,
      '{{BENEMMEGWPDEFOPEN}}':     `{{${p}OPEN}}`,
      '{{BENEMMEGWPDEFCLOSED}}':   `{{${p}CLOSED}}`,
      '{{BENEMMEGWPDEFRESOLVED}}': `{{${p}RESOLVED}}`,
    });
  });
  const added = rowHeight(templateRow) * extraLabels.length;
  frame = frame.slice(0, frame.indexOf(totalRow)) + newRows + frame.slice(frame.indexOf(totalRow));
  frame = growExtent(frame, added);
  let out = xml.slice(0, gfStart) + frame + xml.slice(gfEnd);

  const shift = Math.max(0, added - 10 * EMU_PX * extraLabels.length);
  out = shiftShape(out, 'SeverityTable', shift);
  out = shiftShape(out, 'AgingTable', shift);
  out = shiftShape(out, 'NarrLeft', shift - 7 * EMU_PX * extraLabels.length, -(shift - 7 * EMU_PX * extraLabels.length));
  return out;
}

// Mutates zip's slide5.xml in place and adds any new per-plan token getters
// into effectiveMap. No-op if no plan outside the 5 fixed rows has started
// executing yet.
export function apply(zip, data, effectiveMap) {
  const extraPlans = getExtraActiveBenefitPlans(data);
  const notStarted = getNotStartedBenefitsRow(data);
  const slide5 = injectSlide5(zip.file('ppt/slides/slide5.xml').asText(), extraPlans, effectiveMap, notStarted.planned > 0 ? notStarted : null);
  zip.file('ppt/slides/slide5.xml', slide5);

  // Benefits defect slide (slide6.xml): one extra row per plan beyond the 5 fixed ones
  const defFile = zip.file('ppt/slides/slide6.xml');
  if (defFile && defFile.asText().includes('Testing Defect Summary by Severity')) {
    const all = [...(data.bugs?.Benefits ?? []), ...(data.closedBugs?.Benefits ?? [])];
    const extraLabels = BENEFITS_PLAN_LABELS.filter(l => !FIXED_DEFECT_PLANS.has(l) && all.some(b => benefitsDefectPlan(b) === l));
    if (extraLabels.length) {
      zip.file('ppt/slides/slide6.xml', injectDefectSlide(defFile.asText(), extraLabels, effectiveMap));
      console.log(`  Dynamic Benefits defect rows: added ${extraLabels.join(', ')}`);
    }
  }

  console.log(`  Dynamic Benefits rows: added ${extraPlans.length} plan(s) — ${extraPlans.map(p => p.label).join(', ')}${notStarted.planned > 0 ? ` + Not Started row (${notStarted.planned})` : ''}`);
}

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

import { getExtraActiveBenefitPlans } from '../variables.js';

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

function injectSlide5(xml, extraPlans, tokenMap) {
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

  const addedHeight = rowHeight(templateRow) * extraPlans.length;

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
    tokenMap[`${prefix}DEF`]  = () => 0;

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

  const insertAt = frame.indexOf(grandTotalRow);
  let newFrame = frame.slice(0, insertAt) + newRowsXml + frame.slice(insertAt);
  newFrame = growExtent(newFrame, addedHeight);

  let newXml = xml.slice(0, gfStart) + newFrame + xml.slice(gfEnd);

  // Measured against the 09/2026 template: the table's own graphicFrame
  // bottom sits at y=5632627, with the legend/status panels starting right
  // after at y=5633823 — no separate enclosing background panel to grow
  // here (unlike the pre-refresh template).
  newXml = reflow(newXml, { thresholdY: 5600000, deltaY: addedHeight });

  return newXml;
}

// Mutates zip's slide5.xml in place and adds any new per-plan token getters
// into effectiveMap. No-op if no plan outside the 5 fixed rows has started
// executing yet.
export function apply(zip, data, effectiveMap) {
  const extraPlans = getExtraActiveBenefitPlans(data);
  if (!extraPlans.length) return;

  const slide5 = injectSlide5(zip.file('ppt/slides/slide5.xml').asText(), extraPlans, effectiveMap);
  zip.file('ppt/slides/slide5.xml', slide5);

  console.log(`  Dynamic Benefits rows: added ${extraPlans.length} plan(s) — ${extraPlans.map(p => p.label).join(', ')}`);
}

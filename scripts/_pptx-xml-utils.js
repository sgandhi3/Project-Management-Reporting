// Shared raw-XML helpers for one-time template-patch scripts (scripts/add-*.js,
// scripts/remove-*.js). Pure string/regex surgery on PizZip-loaded slide XML —
// deliberately NOT python-pptx: a round-trip load+save through python-pptx on
// this deck was found to scramble some slides' .rels files (image/notesSlide
// relationships silently vanished or got reassigned to the wrong slide) even
// though the slide XML content itself came back correct. Every script in this
// directory must only ever touch the exact bytes it needs to change and leave
// everything else — including every .rels file — completely untouched.

// Matches each top-level occurrence of <tag ...>...</tag> or <tag .../>.
// Safe for tags that never nest within themselves (a:tc, a:tr, a:gridCol,
// p:sp) — not safe for tags that can self-nest (e.g. a:p).
function matchTopLevelTags(xml, tag) {
  const re = new RegExp(`<${tag}\\b[^>]*?(?:/>|>[\\s\\S]*?</${tag}>)`, 'g');
  const out = [];
  let m;
  while ((m = re.exec(xml))) {
    out.push({ start: m.index, end: m.index + m[0].length, text: m[0] });
  }
  return out;
}

function cellText(tcXml) {
  const texts = [...tcXml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map(m => m[1]);
  return texts.join('');
}

// Removes the col-th column from a <a:tbl>...</a:tbl> block: drops its
// <a:gridCol>, redistributes its width proportionally across the remaining
// columns, and removes the col-th <a:tc> from every row (decrementing a
// title-bar row's gridSpan on its first cell if that row spans the full
// width). Returns the modified tbl XML, or null if "In Progress" isn't found
// in the header row (row index 1).
function removeTableColumn(tblXml, columnHeaderText) {
  const gridMatch = tblXml.match(/<a:tblGrid>([\s\S]*?)<\/a:tblGrid>/);
  if (!gridMatch) return null;
  const cols = matchTopLevelTags(gridMatch[1], 'a:gridCol');

  const rows = matchTopLevelTags(tblXml, 'a:tr');
  if (rows.length < 2) return null;
  const headerCells = matchTopLevelTags(rows[1].text, 'a:tc');
  const colIndex = headerCells.findIndex(tc => cellText(tc.text).trim() === columnHeaderText);
  if (colIndex === -1) return null;

  // ── tblGrid: drop the column, redistribute its width proportionally ──────
  const widths = cols.map(c => Number(c.text.match(/w="(\d+)"/)[1]));
  const freed = widths[colIndex];
  const remaining = widths.filter((_, i) => i !== colIndex);
  const totalRemaining = remaining.reduce((a, b) => a + b, 0);
  let added = 0;
  const newWidths = remaining.map((w, i) => {
    const isLast = i === remaining.length - 1;
    const add = isLast ? freed - added : Math.round(freed * (w / totalRemaining));
    if (!isLast) added += add;
    return w + add;
  });
  const remainingColXml = cols.filter((_, i) => i !== colIndex).map(c => c.text);
  const newGridInner = remainingColXml
    .map((xml, i) => xml.replace(/w="\d+"/, `w="${newWidths[i]}"`))
    .join('');
  let out = tblXml.slice(0, gridMatch.index) + `<a:tblGrid>${newGridInner}</a:tblGrid>` + tblXml.slice(gridMatch.index + gridMatch[0].length);

  // ── Every row: drop the col-th <a:tc>, adjusting gridSpan if needed ──────
  // Re-match rows against the (unchanged-so-far) row region of `out` — the
  // tblGrid edit above was before the first <a:tr>, so row offsets below
  // need to be computed against `out`, not the original tblXml.
  const rowsNow = matchTopLevelTags(out, 'a:tr');
  // Process rows back-to-front so earlier offsets stay valid as we splice.
  for (let i = rowsNow.length - 1; i >= 0; i--) {
    const rowXml = rowsNow[i].text;
    const tcs = matchTopLevelTags(rowXml, 'a:tc');
    if (colIndex >= tcs.length) continue;
    let newRowXml = rowXml;
    const target = tcs[colIndex];
    const first = tcs[0];
    const spanMatch = first.text.match(/gridSpan="(\d+)"/);
    if (spanMatch && Number(spanMatch[1]) > 1 && target.start !== first.start) {
      const newSpan = Number(spanMatch[1]) - 1;
      const newFirst = first.text.replace(/gridSpan="\d+"/, `gridSpan="${newSpan}"`);
      // Replace first, then remove target — do whichever is later first to
      // keep the other's offset valid.
      if (target.start > first.start) {
        newRowXml = newRowXml.slice(0, target.start) + newRowXml.slice(target.end);
        newRowXml = newRowXml.slice(0, first.start) + newFirst + newRowXml.slice(first.end);
      } else {
        newRowXml = newRowXml.slice(0, first.start) + newFirst + newRowXml.slice(first.end);
        newRowXml = newRowXml.slice(0, target.start) + newRowXml.slice(target.end);
      }
    } else {
      newRowXml = newRowXml.slice(0, target.start) + newRowXml.slice(target.end);
    }
    out = out.slice(0, rowsNow[i].start) + newRowXml + out.slice(rowsNow[i].end);
  }

  return out;
}

export { matchTopLevelTags, cellText, removeTableColumn };

// Shared helpers for building new <a:tbl> table XML that visually matches
// the MMO report template's existing tables (SeverityTable, WorkstreamSummary,
// etc.): same table style GUID, teal title bar, light-gray bold header row,
// plain banded data rows. Used by the one-time template-patch scripts in
// this directory (scripts/add-*.js) — not part of the weekly pipeline.

export const TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
export const TEAL = '00827E';
export const HEADER_GRAY = 'F2F2F2';
export const DARK_TEXT = '313131';

export const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

function cell({ text = '', bold = false, fill = null, color = null, gridSpan = null, hMerge = false, align = 'ctr' }) {
  if (hMerge) {
    return `<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></a:txBody><a:tcPr/></a:tc>`;
  }
  const spanAttr = gridSpan ? ` gridSpan="${gridSpan}"` : '';
  const colorXml = color
    ? `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`
    : `<a:solidFill><a:schemeClr val="dk1"/></a:solidFill>`;
  const fillXml = fill ? `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>` : '';
  return `<a:tc${spanAttr}><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="0" marR="0" algn="${align}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="1000"${bold ? ' b="1"' : ''} kern="1200" dirty="0">${colorXml}<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:rPr><a:t>${esc(text)}</a:t></a:r></a:p></a:txBody><a:tcPr marL="9525" marR="9525" marT="9525" marB="0" anchor="ctr">${fillXml}</a:tcPr></a:tc>`;
}

export function titleRow(text, numCols, height = 215000) {
  const first = cell({ text, bold: true, fill: TEAL, color: 'FFFFFF', gridSpan: numCols });
  const rest = Array.from({ length: numCols - 1 }, () => cell({ hMerge: true })).join('');
  return `<a:tr h="${height}">${first}${rest}</a:tr>`;
}

export function headerRow(cells, height = 225000) {
  return `<a:tr h="${height}">${cells.map(text => cell({ text, bold: true, fill: HEADER_GRAY, color: DARK_TEXT })).join('')}</a:tr>`;
}

export function dataRow(cells, { height = 225000, boldFirst = true } = {}) {
  return `<a:tr h="${height}">${cells.map((text, i) => cell({ text, bold: boldFirst && i === 0 })).join('')}</a:tr>`;
}

export function buildTable({ id, name, x, y, cx, cy, colWidths, rowsXml }) {
  const grid = colWidths.map(w => `<a:gridCol w="${w}"/>`).join('');
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"><a:tableStyleId>${TABLE_STYLE_ID}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${rowsXml.join('')}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
}

// Locates a shape's <a:off>/<a:ext> by finding a {{TOKEN}} it contains and
// walking back to the nearest enclosing <p:sp> — same technique as
// scripts/fix-ai-status-box-position.js.
export function findShapeXfrm(xml, token) {
  const idx = xml.indexOf(token);
  if (idx === -1) return null;
  const spStart = xml.lastIndexOf('<p:sp>', idx);
  const spSnippet = xml.slice(spStart, spStart + 700);
  const m = spSnippet.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/><a:ext cx="(\d+)" cy="(\d+)"\/>/);
  if (!m) return null;
  return { x: Number(m[1]), y: Number(m[2]), cx: Number(m[3]), cy: Number(m[4]) };
}

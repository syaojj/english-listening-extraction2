/**
 * PDF 첫 4페이지 텍스트 추출 (앱과 동일한 itemsToLines/joinLineWithSpacing 로직)
 * 사용: node scripts/extract-4pages.js [PDF경로]
 * 기본 PDF: 21 수능만만[기본] 듣기모의 20회_본문(001-144).pdf
 */

const fs = require('fs');
const path = require('path');

const MAX_PAGES = 4;
const SPACE_GAP_THRESHOLD = 18;

function getPageTextItems(content) {
  const items = content.items.map(function (item) {
    const y = item.transform[5];
    const x = item.transform[4];
    const w = (item.width != null && item.width > 0) ? item.width : 0;
    return { str: item.str, y: y, x: x, width: w };
  });
  items.sort(function (a, b) {
    const lineA = Math.round(a.y / 4);
    const lineB = Math.round(b.y / 4);
    if (lineA !== lineB) return lineB - lineA;
    return a.x - b.x;
  });
  return items;
}

function joinLineWithSpacing(lineItems) {
  if (!lineItems.length) return '';
  lineItems.sort(function (a, b) { return a.x - b.x; });
  let out = lineItems[0].str || '';
  for (let k = 1; k < lineItems.length; k++) {
    const prevEnd = (lineItems[k - 1].x || 0) + (lineItems[k - 1].width || 0);
    const gap = (lineItems[k].x || 0) - prevEnd;
    out += (gap >= SPACE_GAP_THRESHOLD ? '  ' : ' ') + (lineItems[k].str || '');
  }
  return out.trim();
}

function itemsToLines(items, centerX) {
  function groupItemsIntoLines(sortedItems) {
    const out = [];
    let lastY = null;
    let line = [];
    for (let i = 0; i < sortedItems.length; i++) {
      const item = sortedItems[i];
      const y = Math.round(item.y / 3);
      if (lastY !== null && Math.abs(y - lastY) > 2) {
        out.push(joinLineWithSpacing(line));
        line = [];
      }
      line.push(item);
      lastY = y;
    }
    if (line.length) out.push(joinLineWithSpacing(line));
    return out.filter(Boolean);
  }
  if (centerX == null || centerX === undefined) {
    return groupItemsIntoLines(items);
  }
  const left = items.filter(function (item) { return (item.x || 0) < centerX; });
  const right = items.filter(function (item) { return (item.x || 0) >= centerX; });
  function sortTopToBottomThenLeftToRight(a, b) {
    const lineA = Math.round(a.y / 4);
    const lineB = Math.round(b.y / 4);
    if (lineA !== lineB) return lineB - lineA;
    return a.x - b.x;
  }
  left.sort(sortTopToBottomThenLeftToRight);
  right.sort(sortTopToBottomThenLeftToRight);
  return groupItemsIntoLines(left).concat(groupItemsIntoLines(right));
}

async function main() {
  const pdfPath = process.argv[2] || path.join(__dirname, '..', '21 수능만만[기본] 듣기모의 20회_본문(001-144).pdf');
  if (!fs.existsSync(pdfPath)) {
    console.error('파일 없음:', pdfPath);
    process.exit(1);
  }

  const pdfjsLib = require('pdfjs-dist');
  pdfjsLib.GlobalWorkerOptions.workerSrc = false;

  const buffer = fs.readFileSync(pdfPath);
  const doc = await pdfjsLib.getDocument({ data: buffer }).promise;
  const numPages = Math.min(doc.numPages, MAX_PAGES);

  for (let i = 1; i <= numPages; i++) {
    const page = await doc.getPage(i);
    const content = await page.getTextContent();
    const items = getPageTextItems(content);
    const viewport = page.getViewport({ scale: 1 });
    const centerX = viewport.width / 2;
    const lines = itemsToLines(items, centerX);
    console.log('========== p.' + i + ' (' + lines.length + '줄) ==========');
    console.log(lines.join('\n'));
    console.log('');
  }
}

main().catch(function (err) {
  console.error(err);
  process.exit(1);
});

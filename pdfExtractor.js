/**
 * pdfExtractor.js
 * -----------------------------------------------------------------
 * PDF から文字を抽出する最小構成の基本モジュール（Node.js版）。
 * doc-check-tool（ブラウザ版GUI）で使っている抽出ロジックを
 * デスクトップアプリ（Electron等）に載せ替えられるよう切り出したもの。
 *
 * 事前準備:
 *   npm install pdfjs-dist
 * -----------------------------------------------------------------
 */

const pdfjsLib = require('pdfjs-dist/legacy/build/pdf.js');
const fs = require('fs');

/**
 * PDF全ページのテキストをまとめて抽出する
 * @param {string} pdfPath
 * @returns {Promise<string>}
 */
async function extractFullText(pdfPath) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  let fullText = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    const pageText = content.items.map((item) => item.str).join('');
    fullText += pageText + '\n';
  }
  return fullText;
}

/**
 * 指定した座標範囲（box）に含まれるテキストだけを抽出する。
 * box の単位は PDF 標準の pt（scale:1 のビューポート基準）。
 * doc-check-tool のフィールド定義（x, y, w, h）とそのまま対応する。
 *
 * @param {string} pdfPath
 * @param {number} pageNum 1始まり
 * @param {{x:number, y:number, w:number, h:number}} box
 * @returns {Promise<string>}
 */
async function extractTextInBox(pdfPath, pageNum, box) {
  const data = new Uint8Array(fs.readFileSync(pdfPath));
  const pdf = await pdfjsLib.getDocument({ data }).promise;
  const page = await pdf.getPage(pageNum);
  const viewport = page.getViewport({ scale: 1 });
  const content = await page.getTextContent();

  const items = [];
  for (const item of content.items) {
    const tx = item.transform[4];
    const tyPdf = item.transform[5];
    const tyTop = viewport.height - tyPdf; // 左上原点に変換
    if (
      tx >= box.x - 2 && tx <= box.x + box.w + 2 &&
      tyTop >= box.y - 4 && tyTop <= box.y + box.h + 6
    ) {
      items.push({ x: tx, str: item.str });
    }
  }
  items.sort((a, b) => a.x - b.x);
  return items.map((i) => i.str).join('').trim();
}

/**
 * 2つの値を比較して判定を返す（doc-check-tool / VBA版と共通のロジック）
 * @param {string} a
 * @param {string} b
 * @param {'exact'|'numeric'|'date'|'contains'} type
 * @returns {'ok'|'ng'|'warn'}
 */
function judgeValues(a, b, type) {
  const norm = (s) => (s || '').replace(/[\s　,，]/g, '');
  const na = norm(a);
  const nb = norm(b);
  if (!na || !nb) return 'warn';

  switch (type) {
    case 'numeric': {
      const x = parseFloat(na.replace(/[^\d.-]/g, ''));
      const y = parseFloat(nb.replace(/[^\d.-]/g, ''));
      if (Number.isNaN(x) || Number.isNaN(y)) return 'warn';
      return x === y ? 'ok' : 'ng';
    }
    case 'date': {
      const norm2 = (s) => norm(s).replace(/[年月]/g, '/').replace(/日/g, '');
      return norm2(a) === norm2(b) ? 'ok' : 'ng';
    }
    case 'contains':
      return na.includes(nb) || nb.includes(na) ? 'ok' : 'ng';
    default: // exact
      return na === nb ? 'ok' : 'ng';
  }
}

module.exports = { extractFullText, extractTextInBox, judgeValues };

// CLIとして直接実行した場合のサンプル: node pdfExtractor.js sample.pdf
if (require.main === module) {
  const path = process.argv[2];
  if (!path) {
    console.log('使い方: node pdfExtractor.js <PDFファイルパス>');
    process.exit(1);
  }
  extractFullText(path).then((text) => console.log(text));
}

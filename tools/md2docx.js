// Markdown (このリポジトリのメモで使っている書き方だけ) を Word に変換する
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ExternalHyperlink, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, LevelFormat, Header, Footer,
  PageNumber, HeadingLevel, TableLayoutType,
} = require("docx");

const PAGE_W = 11906, MARGIN = 1134, CONTENT_W = PAGE_W - MARGIN * 2; // A4, 余白20mm
const FONT = { ascii: "Yu Gothic", hAnsi: "Yu Gothic", eastAsia: "Yu Gothic", cs: "Yu Gothic" };
const MONO = { ascii: "MS Gothic", hAnsi: "MS Gothic", eastAsia: "MS Gothic", cs: "MS Gothic" };
const NAVY = "1F3864", GRAY = "595959", LINK = "0563C1";

// Word版では、メモのファイル名を文書名に置き換える
const RENAME = [
  ["`notes/claude_api_and_contracts.md`", "「Claude APIと契約のまとめ」"],
  ["`notes/business_plan.md`", "「事業計画」"],
  ["`notes/goals.md`", "目標メモ(goals.md)"],
];

function inlineRuns(text, base = {}) {
  for (const [a, b] of RENAME) text = text.split(a).join(b);
  const out = [];
  const re = /(\*\*[^*]+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))/g;
  let last = 0, m;
  const plain = (t, extra = {}) => new TextRun({ text: t, font: FONT, ...base, ...extra });
  while ((m = re.exec(text))) {
    if (m.index > last) out.push(plain(text.slice(last, m.index)));
    const tok = m[0];
    if (tok.startsWith("**")) {
      // 太字の中のリンクやコードも扱う
      out.push(...inlineRuns(tok.slice(2, -2), { ...base, bold: true }));
    } else if (tok.startsWith("`")) {
      out.push(new TextRun({
        text: tok.slice(1, -1), font: MONO, ...base,
        size: base.size ? base.size - 1 : 19,
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "EEF1F5" },
      }));
    } else {
      const lm = tok.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
      out.push(new ExternalHyperlink({
        link: lm[2],
        children: [new TextRun({ text: lm[1], font: FONT, ...base, color: LINK, underline: {} })],
      }));
    }
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(plain(text.slice(last)));
  return out;
}

// 全角を2、半角を1として幅を数える
const visLen = (s) => [...s.replace(/\*\*|`|\[|\]\([^)]*\)/g, "")]
  .reduce((n, c) => n + (c.charCodeAt(0) > 0x2e7f ? 2 : 1), 0);

function buildTable(rows) {
  const parse = (l) => l.trim().replace(/^\||\|$/g, "").split("|").map((c) => c.trim());
  const header = parse(rows[0]);
  const body = rows.slice(2).map(parse);
  const n = header.length;
  const weights = header.map((h, i) => {
    const lens = [h, ...body.map((r) => r[i] || "")].map(visLen);
    return Math.min(Math.max(Math.max(...lens), 4), 46);
  });
  const total = weights.reduce((a, b) => a + b, 0);
  const widths = weights.map((w) => Math.floor((w / total) * CONTENT_W));
  widths[n - 1] += CONTENT_W - widths.reduce((a, b) => a + b, 0);
  const border = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cell = (text, i, isHead) => new TableCell({
    width: { size: widths[i], type: WidthType.DXA },
    borders,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    shading: isHead ? { type: ShadingType.CLEAR, color: "auto", fill: "D9E2F3" } : undefined,
    children: [new Paragraph({
      spacing: { before: 0, after: 0, line: 280 },
      children: inlineRuns(text || "", { size: 18, ...(isHead ? { bold: true, color: NAVY } : {}) }),
    })],
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: header.map((h, i) => cell(h, i, true)) }),
      ...body.map((r) => new TableRow({ cantSplit: true, children: header.map((_, i) => cell(r[i], i, false)) })),
    ],
  });
}

function convert(mdPath, docTitle) {
  const lines = fs.readFileSync(mdPath, "utf8").split("\n");
  const children = [];
  let i = 0, sawTitle = false, beforeFirstH2 = true;
  let stack = [];              // リストの字下げの深さ
  let orderedInst = {};        // 番号付きリストの通し番号(レベルごと)
  let lastKindAt = {};         // 各レベルで直前の項目が番号付きか
  let instCounter = 0;
  const orderedRefs = [];      // 番号付きリストごとに別の番号設定を使い、1から数え直す
  const resetList = () => { stack = []; orderedInst = {}; lastKindAt = {}; };
  const spacer = () => children.push(new Paragraph({ spacing: { before: 0, after: 60 }, children: [] }));

  while (i < lines.length) {
    const line = lines[i];
    if (!line.trim()) { i++; continue; }

    const h = line.match(/^(#{1,6})\s+(.*)$/);
    if (h) {
      resetList();
      const level = h[1].length, text = h[2];
      if (level === 1 && !sawTitle) {
        sawTitle = true;
        children.push(new Paragraph({
          spacing: { before: 0, after: 200 },
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 6 } },
          children: inlineRuns(text, { size: 36, bold: true, color: NAVY }),
        }));
      } else {
        if (level === 2) beforeFirstH2 = false;
        const map = { 2: HeadingLevel.HEADING_1, 3: HeadingLevel.HEADING_2, 4: HeadingLevel.HEADING_3 };
        children.push(new Paragraph({ heading: map[level] || HeadingLevel.HEADING_3, children: inlineRuns(text) }));
      }
      i++; continue;
    }

    if (line.startsWith("```")) {
      resetList();
      const code = [];
      i++;
      while (i < lines.length && !lines[i].startsWith("```")) code.push(lines[i++]);
      i++;
      code.forEach((c, k) => children.push(new Paragraph({
        spacing: { before: k === 0 ? 120 : 0, after: k === code.length - 1 ? 160 : 0, line: 300 },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "F2F4F7" },
        indent: { left: 200, right: 200 },
        children: [new TextRun({ text: c || " ", font: MONO, size: 19 })],
      })));
      continue;
    }

    if (line.trim().startsWith("|")) {
      resetList();
      const rows = [];
      while (i < lines.length && lines[i].trim().startsWith("|")) rows.push(lines[i++]);
      children.push(buildTable(rows));
      spacer();
      continue;
    }

    if (/^\s*>/.test(line)) {
      const indentSpaces = line.match(/^\s*/)[0].length;
      const quote = [];
      while (i < lines.length && /^\s*>/.test(lines[i])) quote.push(lines[i++].replace(/^\s*>\s?/, ""));
      const left = 360 + Math.min(indentSpaces, 6) * 120;
      quote.forEach((q, k) => children.push(new Paragraph({
        spacing: { before: k === 0 ? 80 : 0, after: k === quote.length - 1 ? 140 : 0, line: 320 },
        indent: { left, right: 200 },
        shading: { type: ShadingType.CLEAR, color: "auto", fill: "F5F7FA" },
        border: { left: { style: BorderStyle.SINGLE, size: 18, color: "8EA9DB", space: 8 } },
        children: q ? inlineRuns(q, { size: 20 }) : [new TextRun({ text: " ", font: FONT })],
      })));
      continue;
    }

    const li = line.match(/^(\s*)([-*]|\d+\.)\s+(.*)$/);
    if (li) {
      const ind = li[1].length, ordered = /\d/.test(li[2]);
      while (stack.length && stack[stack.length - 1] > ind) stack.pop();
      if (!stack.length || stack[stack.length - 1] < ind) stack.push(ind);
      const level = Math.min(stack.length - 1, 3);
      // 深いレベルの番号は、浅い項目が来たらリセット
      for (const k of Object.keys(orderedInst)) if (+k > level) { delete orderedInst[k]; delete lastKindAt[k]; }
      let numbering;
      if (ordered) {
        if (!orderedInst[level] || lastKindAt[level] === false) {
          orderedInst[level] = `numbers${++instCounter}`;
          orderedRefs.push(orderedInst[level]);
        }
        numbering = { reference: orderedInst[level], level };
      } else {
        numbering = { reference: "bullets", level };
      }
      lastKindAt[level] = ordered;
      children.push(new Paragraph({
        numbering,
        spacing: { before: 20, after: 40, line: 320 },
        children: inlineRuns(li[3]),
      }));
      i++; continue;
    }

    // ふつうの段落(1行ずつ)
    resetList();
    const meta = sawTitle && beforeFirstH2;
    children.push(new Paragraph({
      spacing: { before: 0, after: meta ? 40 : 120, line: 320 },
      children: inlineRuns(line.trim(), meta ? { size: 19, color: GRAY } : {}),
    }));
    i++;
  }

  const bulletChars = ["●", "○", "■", "・"];
  return new Document({
    creator: "Claude",
    title: docTitle,
    styles: {
      default: { document: { run: { font: FONT, size: 21 }, paragraph: { spacing: { line: 320 } } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 28, bold: true, color: NAVY },
          paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0,
            border: { left: { style: BorderStyle.SINGLE, size: 36, color: NAVY, space: 8 } }, indent: { left: 120 } } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 24, bold: true, color: NAVY },
          paragraph: { spacing: { before: 260, after: 100 }, keepNext: true, outlineLevel: 1,
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "8EA9DB", space: 2 } } } },
        { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { font: FONT, size: 22, bold: true, color: "2F5496" },
          paragraph: { spacing: { before: 200, after: 80 }, keepNext: true, outlineLevel: 2 } },
      ],
    },
    numbering: {
      config: [
        { reference: "bullets", levels: [0, 1, 2, 3].map((l) => ({
          level: l, format: LevelFormat.BULLET, text: bulletChars[l], alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 420 + l * 400, hanging: 280 } }, run: { font: FONT, size: l === 0 ? 14 : 18 } },
        })) },
        ...orderedRefs.map((reference) => ({ reference, levels: [0, 1, 2, 3].map((l) => ({
          level: l, format: LevelFormat.DECIMAL, text: `%${l + 1}.`, alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 420 + l * 400, hanging: 340 } } },
        })) })),
      ],
    },
    sections: [{
      properties: { page: { size: { width: PAGE_W, height: 16838 }, margin: { top: 1134, bottom: 1134, left: MARGIN, right: MARGIN, header: 567, footer: 567 } } },
      headers: { default: new Header({ children: [new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: docTitle, font: FONT, size: 16, color: "8C8C8C" })],
      })] }) },
      footers: { default: new Footer({ children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ children: ["- ", PageNumber.CURRENT, " -"], font: FONT, size: 18, color: "8C8C8C" })],
      })] }) },
      children,
    }],
  });
}

const [, , mdPath, outPath, title] = process.argv;
Packer.toBuffer(convert(mdPath, title)).then((buf) => {
  fs.writeFileSync(outPath, buf);
  console.log("wrote", outPath);
});

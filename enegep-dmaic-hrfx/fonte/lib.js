// Helpers de formatação ENEGEP (Times New Roman 12, 1,5, justificado; ilustrações TNR 10)
const fs = require("fs");
const path = require("path");
const D = require("docx");
const {
  Paragraph, TextRun, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle,
  ShadingType, ImageRun, Header, Footer, PageNumber, HorizontalPositionRelativeFrom,
  VerticalPositionRelativeFrom, TextWrappingType, VerticalAlign, HighlightColor,
} = D;

const BUILD = __dirname;
const FONT = "Times New Roman";
const LINE = 360; // 1,5
const TEXT_W = 9070; // A4 (11906) - 2 x 1418

// --- runs com marcação simples: **negrito**, *itálico*, [COMPLETAR ...] destacado
function runs(text, opts = {}) {
  const size = opts.size || 24;
  const out = [];
  const re = /(\*\*[^*]+\*\*|\*[^*]+\*|\[COMPLETAR[^\]]*\])/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size, bold: opts.bold, italics: opts.italics }));
    const tok = m[0];
    if (tok.startsWith("**")) out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size, bold: true, italics: opts.italics }));
    else if (tok.startsWith("[COMPLETAR")) out.push(new TextRun({ text: tok, font: FONT, size, bold: true, shading: { type: ShadingType.CLEAR, color: "auto", fill: "FFFF00" } }));
    else out.push(new TextRun({ text: tok.slice(1, -1), font: FONT, size, italics: true, bold: opts.bold }));
    last = m.index + tok.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size, bold: opts.bold, italics: opts.italics }));
  return out;
}

const P = (text, o = {}) => new Paragraph({
  children: runs(text, o), alignment: o.align || AlignmentType.JUSTIFIED,
  spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 }, keepNext: o.keepNext, keepLines: o.keepLines,
});
const blank = () => new Paragraph({ children: [new TextRun({ text: "", font: FONT, size: 24 })], spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 } });
// Título de seção: linha em branco antes
const H1 = (text) => [blank(), new Paragraph({ children: runs(text, { bold: true }), spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 }, keepNext: true, keepLines: true })];
// Subtítulo: sem linha em branco
const H2 = (text) => new Paragraph({ children: runs(text, { bold: true }), spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 }, keepNext: true, keepLines: true });

const bullet = (text) => new Paragraph({ children: runs(text), numbering: { reference: "lista", level: 0 }, alignment: AlignmentType.JUSTIFIED, spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 } });
let alineaInstance = 0;
function alineas(items) {
  alineaInstance++;
  return items.map((t) => new Paragraph({ children: runs(t), numbering: { reference: "alinea", level: 0, instance: alineaInstance }, alignment: AlignmentType.JUSTIFIED, spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 0 } }));
}

const caption = (text) => new Paragraph({ children: runs(text, { size: 20 }), alignment: AlignmentType.CENTER, spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 120, after: 0 }, keepNext: true, keepLines: true });
const fonte = (text) => new Paragraph({ children: runs(text, { size: 20 }), alignment: AlignmentType.CENTER, spacing: { lineRule: D.LineRuleType.AUTO, line: LINE, before: 0, after: 120 } });

function image(file, widthPx, heightPx) {
  return new Paragraph({
    alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 60, after: 0 },
    children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(BUILD, file)), transformation: { width: widthPx, height: heightPx } })],
  });
}

// --- tabela: widths em DXA (soma = TEXT_W ou menor), header = array de strings, rows = array de arrays
const border = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const borders = { top: border, bottom: border, left: border, right: border };
function cellPara(text, o) {
  return String(text).split("\n").map((line) => new Paragraph({
    children: runs(line, { size: o.size || 18, bold: o.bold }),
    alignment: o.align || AlignmentType.LEFT, keepNext: o.keepNext, spacing: { lineRule: D.LineRuleType.AUTO, line: 240, before: 0, after: 0 },
  }));
}
function table(widths, header, rows, o = {}) {
  const size = o.size || 18;
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (cells, isHeader, keep) => new TableRow({
    tableHeader: isHeader, cantSplit: true,
    children: cells.map((c, i) => new TableCell({
      borders, width: { size: widths[i], type: WidthType.DXA },
      shading: isHeader ? { fill: "D9D9D9", type: ShadingType.CLEAR, color: "auto" } : (o.shadeFirstCol && i === 0 ? { fill: "F2F2F2", type: ShadingType.CLEAR, color: "auto" } : undefined),
      margins: { top: 40, bottom: 40, left: 70, right: 70 }, verticalAlign: VerticalAlign.CENTER,
      children: cellPara(c, {
        size, bold: isHeader || (o.boldFirstCol && i === 0), keepNext: keep,
        align: isHeader || (o.centerCols && o.centerCols.includes(i)) ? AlignmentType.CENTER : AlignmentType.LEFT,
      }),
    })),
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths, alignment: AlignmentType.CENTER,
    rows: [mk(header, true, true), ...rows.map((r, k) => mk(r, false, k < rows.length - 1))],
  });
}

// --- cabeçalho e rodapé (modelo ENEGEP; dados da edição 2026)
function header() {
  const hRun = (t) => new TextRun({ text: t, font: "Trebuchet MS", size: 16, bold: true });
  const hp = (children) => new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 0, after: 0, line: 240, lineRule: D.LineRuleType.AUTO }, children });
  return new Header({
    children: [
      hp([
        new ImageRun({
          type: "jpg", data: fs.readFileSync(path.join(BUILD, "logo.jpg")), transformation: { width: 79, height: 62 },
          floating: {
            horizontalPosition: { relative: HorizontalPositionRelativeFrom.COLUMN, offset: -551179 },
            verticalPosition: { relative: VerticalPositionRelativeFrom.PARAGRAPH, offset: -151764 },
            wrap: { type: TextWrappingType.NONE }, allowOverlap: true, behindDocument: false,
          },
        }),
        hRun("XLVI ENCONTRO NACIONAL DE ENGENHARIA DE PRODUÇÃO"),
      ]),
      hp([hRun("\"Engenharia de Produção para a Transição Sustentável: Modelagem de Dados e Inovação\"")]),
      hp([hRun("São Paulo/SP, 20 a 23 de outubro de 2026.")]),
      new Paragraph({ spacing: { before: 0, after: 0, line: 240, lineRule: D.LineRuleType.AUTO }, children: [new TextRun({ text: "", size: 16 })] }),
    ],
  });
}
function footer() {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 20 })] })] });
}

const numbering = {
  config: [
    { reference: "lista", levels: [{ level: 0, format: D.LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: FONT } } }] },
    { reference: "alinea", levels: [{ level: 0, format: D.LevelFormat.LOWER_LETTER, text: "%1)", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { font: FONT, size: 24 } } }] },
  ],
};

function section(children) {
  return {
    properties: {
      page: { size: { width: 11906, height: 16838 }, margin: { top: 680, bottom: 680, left: 1418, right: 1418, header: 340, footer: 340 } },
    },
    headers: { default: header() }, footers: { default: footer() }, children,
  };
}

function save(doc, file) {
  return D.Packer.toBuffer(doc).then((b) => fs.writeFileSync(file, b));
}

const styles = { default: { document: { run: { font: FONT, size: 24 } } } };

module.exports = { P, blank, H1, H2, bullet, alineas, caption, fonte, image, table, section, save, numbering, styles, runs, TEXT_W, D, FONT, LINE };

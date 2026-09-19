// Генерация пояснительной записки в формате Word (.docx)
// на основе сводных показателей, полученных скриптом 02_obrabotka.py.

const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
} = require("docx");

const data = JSON.parse(
  fs.readFileSync(__dirname + "/output/svodnye_pokazateli.json", "utf-8")
);

const FONT = "Times New Roman";
const fmt = (n) => n.toLocaleString("ru-RU").replace(/ /g, " ") + " ₸";

function cell(text, { bold = false, width, shade, align = AlignmentType.LEFT } = {}) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade ? { type: ShadingType.CLEAR, fill: shade } : undefined,
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [
      new Paragraph({
        alignment: align,
        children: [new TextRun({ text, bold, font: FONT, size: 21 })],
      }),
    ],
  });
}

function summaryTable() {
  const rows = [
    ["Выручка за месяц", fmt(data.total_income)],
    ["Расходы за месяц", fmt(data.total_expense)],
    ["Чистая прибыль", fmt(data.net_profit)],
    ["Остаток денежных средств на начало месяца", fmt(data.opening_balance)],
    ["Остаток денежных средств на конец месяца", fmt(data.closing_balance)],
  ];
  return new Table({
    width: { size: 9000, type: WidthType.DXA },
    columnWidths: [6000, 3000],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
      left: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
      right: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" },
    },
    rows: [
      new TableRow({
        children: [
          cell("Показатель", { bold: true, width: 6000, shade: "305496" }),
          cell("Сумма", { bold: true, width: 3000, shade: "305496", align: AlignmentType.RIGHT }),
        ],
      }),
      ...rows.map(([label, value], i) =>
        new TableRow({
          children: [
            cell(label, { width: 6000, shade: i % 2 ? "F2F2F2" : undefined }),
            cell(value, { width: 3000, shade: i % 2 ? "F2F2F2" : undefined, align: AlignmentType.RIGHT }),
          ],
        })
      ),
    ],
  });
}

function bodyParagraphs(text) {
  return new Paragraph({
    spacing: { after: 200, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
    children: [new TextRun({ text, font: FONT, size: 24 })],
  });
}

function bullet(text) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { after: 120, line: 300 },
    children: [new TextRun({ text, font: FONT, size: 24 })],
  });
}

const doc = new Document({
  sections: [
    {
      properties: {
        page: { size: { width: 11906, height: 16838 } }, // A4
      },
      children: [
        new Paragraph({
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [
            new TextRun({ text: "Пояснительная записка к финансовой отчётности", bold: true, font: FONT, size: 32 }),
          ],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 300 },
          children: [
            new TextRun({ text: `${data.company} — ${data.period}`, italics: true, font: FONT, size: 24, color: "595959" }),
          ],
        }),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 100, after: 150 },
          children: [new TextRun({ text: "Ключевые показатели месяца", bold: true, font: FONT, size: 26 })],
        }),
        summaryTable(),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 350, after: 150 },
          children: [new TextRun({ text: "Пояснение к отчёту", bold: true, font: FONT, size: 26 })],
        }),
        ...data.explanation_paragraphs.map(bodyParagraphs),

        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          spacing: { before: 250, after: 150 },
          children: [new TextRun({ text: "Ошибки, найденные при автоматической проверке", bold: true, font: FONT, size: 26 })],
        }),
        bullet(
          "Задвоенный платёж — операция «Оплата аренды офиса за август 2026 по договору №5» " +
          "(60 000 ₸) была ошибочно внесена в таблицу дважды. ИИ-сервис сравнил все операции " +
          "и обнаружил две полностью идентичные строки (одинаковые дата, контрагент, назначение " +
          "и сумма) — один из платежей исключён из отчётности как ошибочный."
        ),
        bullet(
          "Сумма не соответствует назначению платежа — в операции «Зарплата за август 2026 " +
          "(5 сотрудников по 45 000 тенге)» указана сумма 205 000 ₸, хотя расчёт по тексту " +
          "самого назначения даёт 225 000 ₸ (5 × 45 000). Расхождение в 20 000 ₸ было найдено " +
          "автоматически и исправлено при составлении отчётности."
        ),

        new Paragraph({
          spacing: { before: 350 },
          border: { top: { style: BorderStyle.SINGLE, size: 6, color: "BFBFBF", space: 8 } },
          children: [
            new TextRun({
              text: "Документ подготовлен автоматически с помощью ИИ-сервиса обработки бухгалтерских данных. " +
                "Учебный демонстрационный пример для лекции «Составление финансовой отчётности с применением ИИ-сервисов».",
              italics: true,
              font: FONT,
              size: 18,
              color: "808080",
            }),
          ],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = __dirname + "/output/Poyasnitelnaya_zapiska.docx";
  fs.writeFileSync(outPath, buffer);
  console.log("Сохранено:", outPath);
});

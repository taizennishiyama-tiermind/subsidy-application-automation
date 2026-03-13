#!/usr/bin/env node

const fs = require("fs");
const { Document, Packer, Paragraph, HeadingLevel, AlignmentType } = require("docx");

const args = Object.fromEntries(
  process.argv.slice(2).map((arg) => {
    const [k, ...rest] = arg.replace(/^--/, "").split("=");
    return [k, rest.join("=")];
  }),
);

if (!args.ratios || !args.output) {
  console.error("Usage: node generate_word_report.js --ratios=FILE --output=FILE [--company=NAME]");
  process.exit(1);
}

const payload = JSON.parse(fs.readFileSync(args.ratios, "utf-8"));
const latest = payload.ratios[payload.ratios.length - 1] || {};
const company = args.company || payload.company_name || "会社名未設定";

const percent = (value) => (value == null ? "N/A" : `${(value * 100).toFixed(1)}%`);
const number = (value) => (value == null ? "N/A" : `${value.toFixed(1)}`);

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          text: `${company} 財務分析レポート`,
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ text: "1. 経営状況の概要", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          text: `最新期の自己資本比率は ${percent(latest.equity_ratio)}、営業利益率は ${percent(latest.operating_profit_margin)} です。補助金申請では、継続性と投資余力の根拠として使います。`,
        }),
        new Paragraph({ text: "2. 収益性・安全性", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          text: `ROA は ${percent(latest.roa)}、ROE は ${percent(latest.roe)}、流動比率は ${percent(latest.current_ratio)} です。`,
        }),
        new Paragraph({ text: "3. 返済能力", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          text: `債務償還年数は ${number(latest.debt_repayment_years)} 年、インタレストカバレッジレシオは ${number(latest.interest_coverage_ratio)} 倍です。`,
        }),
        new Paragraph({ text: "4. 補助金申請への示唆", heading: HeadingLevel.HEADING_1 }),
        new Paragraph({
          text: "強みは積極的に訴求し、弱い指標は補助事業でどう改善するかまで説明します。概算値がある場合はその旨を明記します。",
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc)
  .then((buffer) => {
    fs.writeFileSync(args.output, buffer);
    console.log(`OK: wrote ${args.output}`);
  })
  .catch((error) => {
    console.error(error);
    process.exit(1);
  });

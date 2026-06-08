import Anthropic from "@anthropic-ai/sdk";
import { google } from "googleapis";
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, WidthType, BorderStyle, ShadingType,
} from "docx";
import { readFileSync, createReadStream, writeFileSync, unlinkSync } from "fs";

const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

// ── Fetch one ticker's report from Claude ────────────────────────────────────
async function fetchReport(ticker) {
  console.log(`  Researching ${ticker}...`);
  const prompt = `You are an equity research analyst. Generate a concise equity report for ${ticker}.

Use web search to find current, real data. Respond ONLY with a JSON object — no markdown, no backticks, no extra text:

{
  "company_name": "Full company name",
  "sector": "Sector",
  "price": "Current stock price e.g. $182.50",
  "price_change": "1-day change e.g. +1.2%",
  "market_cap": "Market cap e.g. $2.8T",
  "pe_ratio": "P/E ratio or N/A",
  "52w_range": "52-week low – high",
  "analyst_consensus": "Buy / Hold / Sell",
  "analyst_target": "Average price target",
  "analyst_summary": "2-3 sentences on analyst sentiment and recent rating changes.",
  "news": [
    {"headline": "...", "source": "...", "date": "..."},
    {"headline": "...", "source": "...", "date": "..."},
    {"headline": "...", "source": "...", "date": "..."}
  ],
  "financials_summary": "2-3 sentences on recent revenue, earnings, and key metrics.",
  "bull_case": "2 sentences.",
  "bear_case": "2 sentences.",
  "sentiment": "bullish | bearish | neutral"
}`;

  const response = await client.messages.create({
    model: "claude-sonnet-4-5",
    max_tokens: 1000,
    tools: [{ type: "web_search_20250305", name: "web_search" }],
    messages: [{ role: "user", content: prompt }],
  });

  const text = response.content
    .filter((b) => b.type === "text")
    .map((b) => b.text)
    .join("");
  return JSON.parse(text.replace(/```json|```/g, "").trim());
}

// ── Build the .docx from all reports ────────────────────────────────────────
async function buildDocx(reports) {
  const b = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
  const borders = { top: b, bottom: b, left: b, right: b };

  const cell = (text, width, isHeader) =>
    new TableCell({
      borders,
      width: { size: width, type: WidthType.DXA },
      shading: isHeader ? { fill: "E8F0FE", type: ShadingType.CLEAR } : undefined,
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [
        new Paragraph({
          children: [
            new TextRun({
              text: String(text || "—"),
              bold: !!isHeader,
              size: isHeader ? 18 : 20,
              font: "Arial",
            }),
          ],
        }),
      ],
    });

  const children = [
    new Paragraph({
      heading: HeadingLevel.HEADING_1,
      children: [new TextRun({ text: "TSIG Portfolio — Equity Reports", bold: true, size: 32, font: "Arial" })],
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: `Generated: ${new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" })}`,
          size: 22,
          color: "666666",
          font: "Arial",
        }),
      ],
    }),
    new Paragraph({ children: [new TextRun("")] }),
  ];

  for (const { ticker, data: r, error } of reports) {
    if (error) {
      children.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_2,
          children: [new TextRun({ text: `${ticker} — Error`, bold: true, size: 28, font: "Arial" })],
        }),
        new Paragraph({ children: [new TextRun({ text: error, size: 22, color: "CC0000", font: "Arial" })] }),
        new Paragraph({ children: [new TextRun("")] })
      );
      continue;
    }

    children.push(
      new Paragraph({
        heading: HeadingLevel.HEADING_2,
        children: [new TextRun({ text: `${ticker} — ${r.company_name || ""}`, bold: true, size: 28, font: "Arial" })],
      }),
      new Paragraph({
        children: [new TextRun({ text: `${r.sector || ""} · Sentiment: ${r.sentiment || ""}`, size: 22, color: "666666", font: "Arial" })],
      }),
      new Paragraph({ children: [new TextRun("")] })
    );

    const cols = [1560, 1560, 1560, 2400, 2280];
    children.push(
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: cols,
        rows: [
          new TableRow({
            children: ["Price", "Market cap", "P/E ratio", "52-week range", "Analyst consensus"].map((h, i) =>
              cell(h, cols[i], true)
            ),
          }),
          new TableRow({
            children: [
              r.price,
              r.market_cap,
              r.pe_ratio,
              r["52w_range"],
              `${r.analyst_consensus || "—"} (target: ${r.analyst_target || "—"})`,
            ].map((v, i) => cell(v, cols[i], false)),
          }),
        ],
      }),
      new Paragraph({ children: [new TextRun("")] })
    );

    for (const [label, text] of [
      ["Analyst sentiment", r.analyst_summary],
      ["Financials", r.financials_summary],
      ["Bull case", r.bull_case],
      ["Bear case", r.bear_case],
    ]) {
      if (!text) continue;
      children.push(
        new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 22, font: "Arial" })] }),
        new Paragraph({ children: [new TextRun({ text, size: 22, font: "Arial" })] }),
        new Paragraph({ children: [new TextRun("")] })
      );
    }

    if (r.news?.length) {
      children.push(new Paragraph({ children: [new TextRun({ text: "Recent news", bold: true, size: 22, font: "Arial" })] }));
      for (const n of r.news) {
        children.push(
          new Paragraph({ indent: { left: 360 }, children: [new TextRun({ text: `• ${n.headline}`, size: 20, font: "Arial" })] }),
          new Paragraph({ indent: { left: 360 }, children: [new TextRun({ text: `  ${n.source} · ${n.date}`, size: 18, color: "888888", font: "Arial" })] })
        );
      }
      children.push(new Paragraph({ children: [new TextRun("")] }));
    }

    children.push(
      new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "DDDDDD" } }, children: [new TextRun("")] }),
      new Paragraph({ children: [new TextRun("")] })
    );
  }

  const doc = new Document({
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 32, bold: true, font: "Arial", color: "1a1a2e" }, paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true, run: { size: 28, bold: true, font: "Arial", color: "1a3a5c" }, paragraph: { spacing: { before: 300, after: 120 }, outlineLevel: 1 } },
      ],
    },
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      children,
    }],
  });

  return Packer.toBuffer(doc);
}

// ── Upload to Google Drive ───────────────────────────────────────────────────
async function uploadToDrive(buffer, filename) {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ["https://www.googleapis.com/auth/drive.file"],
  });

  const drive = google.drive({ version: "v3", auth });
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;

  const tmpPath = `/tmp/${filename}`;
  writeFileSync(tmpPath, buffer);

  const res = await drive.files.create({
    requestBody: {
      name: filename,
      parents: [folderId],
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    },
    media: {
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      body: createReadStream(tmpPath),
    },
  });

  unlinkSync(tmpPath);
  console.log(`  Uploaded to Drive: ${filename} (id: ${res.data.id})`);
}

// ── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  const { tickers } = JSON.parse(readFileSync("tickers.json", "utf8"));
  console.log(`\nGenerating reports for: ${tickers.join(", ")}\n`);

  const reports = [];
  for (const ticker of tickers) {
    try {
      const data = await fetchReport(ticker);
      reports.push({ ticker, data });
      console.log(`  ✓ ${ticker}`);
    } catch (e) {
      console.error(`  ✗ ${ticker}: ${e.message}`);
      reports.push({ ticker, error: e.message });
    }
  }

  console.log("\nBuilding .docx...");
  const buffer = await buildDocx(reports);

  const date = new Date().toISOString().slice(0, 7);
  const filename = `TSIG_Equity_Reports_${date}.docx`;

  console.log("Uploading to Google Drive...");
  await uploadToDrive(buffer, filename);

  console.log("\nDone! ✓");
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});

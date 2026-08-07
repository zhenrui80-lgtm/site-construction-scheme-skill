import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "C:/Users/Administrator/xwechat_files/wxid_bspmv0rjp1rf22_3149/msg/file/2026-08/报价清单.xlsx";
const outputDir = "C:/Users/Administrator/Desktop/Metallurgical_BOQ_AgentSkill_Fixed/agents/_work_quote_scheme/boq_preview";
await fs.mkdir(outputDir, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const sheetSummary = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 6000,
});
await fs.writeFile(path.join(outputDir, "sheets.ndjson"), sheetSummary.ndjson, "utf8");

const overview = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 16000,
  tableMaxRows: 12,
  tableMaxCols: 14,
  tableMaxCellChars: 120,
});
await fs.writeFile(path.join(outputDir, "overview.ndjson"), overview.ndjson, "utf8");

const results = [];
for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  const region = await workbook.inspect({
    kind: "region",
    sheetId: sheet.name,
    range: used.address,
    maxChars: 50000,
    tableMaxRows: 300,
    tableMaxCols: 30,
    tableMaxCellChars: 240,
  });
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(path.join(outputDir, `${safeName}.ndjson`), region.ndjson, "utf8");
  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1.25,
    format: "png",
  });
  const png = path.join(outputDir, `${safeName}.png`);
  await fs.writeFile(png, new Uint8Array(await preview.arrayBuffer()));
  results.push({ sheet: sheet.name, usedRange: used.address, png });
}

console.log(JSON.stringify(results, null, 2));

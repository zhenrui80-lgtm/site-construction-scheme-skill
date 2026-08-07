import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const baseDir = path.resolve("..");
const inputPath = path.join(baseDir, "boq.xlsx");
const previewDir = path.join(baseDir, "spreadsheet", "previews");
await fs.mkdir(previewDir, { recursive: true });

const input = await FileBlob.load(inputPath);
const workbook = await SpreadsheetFile.importXlsx(input);

const sheetIndex = await workbook.inspect({
  kind: "sheet",
  include: "id,name",
  maxChars: 12000,
});
console.log("SHEETS");
console.log(sheetIndex.ndjson);

const summary = await workbook.inspect({
  kind: "workbook,sheet,table",
  maxChars: 24000,
  tableMaxRows: 12,
  tableMaxCols: 12,
  tableMaxCellChars: 160,
});
console.log("SUMMARY");
console.log(summary.ndjson);

for (const sheet of workbook.worksheets.items) {
  const used = sheet.getUsedRange();
  console.log(`USED ${sheet.name} ${used?.address ?? "(none)"}`);

  const region = await workbook.inspect({
    kind: "region",
    sheetId: sheet.name,
    range: used?.address,
    maxChars: 50000,
    tableMaxRows: 200,
    tableMaxCols: 20,
    tableMaxCellChars: 240,
  });
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(
    path.join(previewDir, `${safeName}.ndjson`),
    region.ndjson,
    "utf8",
  );
  console.log(region.ndjson);

  const preview = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: 1,
    format: "png",
  });
  await fs.writeFile(
    path.join(previewDir, `${safeName}.png`),
    new Uint8Array(await preview.arrayBuffer()),
  );

  if (sheet.name === "表8-宝钢咨询清单") {
    const rows = sheet.getRange("A1:O124").values;
    const items = [];
    let category = "";
    let current = null;

    for (const row of rows) {
      const seq = row[0];
      const name = String(row[3] ?? "").trim();
      const feature = String(row[4] ?? "").trim();
      const work = String(row[7] ?? "").trim();

      if (!seq && name.endsWith("专业")) {
        category = name;
        continue;
      }

      if (typeof seq === "number" || /^\d+$/.test(String(seq ?? "").trim())) {
        current = {
          seq: Number(seq),
          code: String(row[1] ?? "").trim(),
          category,
          name,
          feature,
          work,
          unit: String(row[9] ?? "").trim(),
          quantity: row[10],
          note: String(row[14] ?? "").trim(),
        };
        items.push(current);
        continue;
      }

      if (
        current &&
        !name &&
        (feature || work) &&
        !["项目特征", "工作内容"].includes(feature) &&
        !["项目特征", "工作内容"].includes(work)
      ) {
        if (feature) current.feature = [current.feature, feature].filter(Boolean).join("\n");
        if (work) current.work = [current.work, work].filter(Boolean).join("\n");
      }
    }

    await fs.writeFile(
      path.join(baseDir, "boq_items.json"),
      JSON.stringify(items, null, 2),
      "utf8",
    );
    console.log(`ITEMS ${items.length}`);
    for (const item of items) {
      console.log(
        `${item.seq}\t${item.category}\t${item.name}\t${item.unit}\t${item.quantity}\t${item.feature.replaceAll("\n", " / ")}`,
      );
    }
  }
}

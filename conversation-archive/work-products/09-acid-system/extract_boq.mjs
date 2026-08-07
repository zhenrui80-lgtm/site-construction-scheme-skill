import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const source = "C:/Users/Administrator/xwechat_files/wxid_bspmv0rjp1rf22_3149/msg/file/2026-08/报价清单.xlsx";
const outputDir = "C:/Users/Administrator/Desktop/Metallurgical_BOQ_AgentSkill_Fixed/agents/_work_quote_scheme/boq_extract";
await fs.mkdir(outputDir, { recursive: true });

const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(source));
const sheet = workbook.worksheets.getItem("表8-宝钢咨询清单");
const used = sheet.getUsedRange();
const values = used.values;
await fs.writeFile(path.join(outputDir, "all_rows.json"), JSON.stringify(values, null, 2), "utf8");

const items = [];
const sections = [];
for (let index = 0; index < values.length; index += 1) {
  const row = values[index] ?? [];
  const serial = row[0];
  const name = row[3];
  const feature = row[4];
  const workContent = row[7];
  const unit = row[9];
  const quantity = row[10];
  const unitPrice = row[12];
  const total = row[13];
  if (typeof serial === "number" || /^\d+$/.test(String(serial ?? "").trim())) {
    items.push({
      row: index + 1,
      serial: Number(serial),
      code: row[1] ?? "",
      name: name ?? "",
      feature: feature ?? "",
      workContent: workContent ?? "",
      unit: unit ?? "",
      quantity: quantity ?? "",
      unitPrice: unitPrice ?? "",
      total: total ?? "",
      note: row[14] ?? "",
    });
  } else if (name && !unit && !quantity && !String(name).includes("项目名称") && !String(name).includes("分部小计")) {
    sections.push({ row: index + 1, name: String(name) });
  }
}
await fs.writeFile(path.join(outputDir, "items.json"), JSON.stringify(items, null, 2), "utf8");
await fs.writeFile(path.join(outputDir, "sections.json"), JSON.stringify(sections, null, 2), "utf8");

const previews = [];
for (let start = 1; start <= 580; start += 60) {
  const end = Math.min(start + 59, 580);
  const range = `A${start}:O${end}`;
  const preview = await workbook.render({
    sheetName: sheet.name,
    range,
    scale: 0.8,
    format: "png",
  });
  const png = path.join(outputDir, `rows_${String(start).padStart(3, "0")}_${String(end).padStart(3, "0")}.png`);
  await fs.writeFile(png, new Uint8Array(await preview.arrayBuffer()));
  previews.push(png);
}

console.log(JSON.stringify({ usedRange: used.address, items: items.length, sections, previews }, null, 2));

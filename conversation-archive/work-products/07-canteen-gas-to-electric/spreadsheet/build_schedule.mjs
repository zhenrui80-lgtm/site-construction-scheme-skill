import fs from "node:fs/promises";
import path from "node:path";
import { Workbook, SpreadsheetFile } from "@oai/artifact-tool";

const workDir = path.resolve("..");
const outputDir = path.join(workDir, "output");
const renderDir = path.join(workDir, "spreadsheet", "renders");
const boqPath = path.join(workDir, "boq_items.json");
const outputPath = path.join(outputDir, "质检中心食堂气改电施工进度计划.xlsx");

await fs.mkdir(outputDir, { recursive: true });
await fs.mkdir(renderDir, { recursive: true });

const boqItems = JSON.parse(await fs.readFile(boqPath, "utf8"));
const project = "（实事工程）质检中心食堂气改电";
const startDate = new Date("2026-08-10T00:00:00+08:00");
const finishDate = new Date("2026-12-31T00:00:00+08:00");
const oneDay = 24 * 60 * 60 * 1000;
const totalDays = Math.round((finishDate - startDate) / oneDay) + 1;

const parseDate = (value) => new Date(`${value}T00:00:00+08:00`);
const daysInclusive = (start, finish) =>
  Math.round((parseDate(finish) - parseDate(start)) / oneDay) + 1;
const displayDate = (value) => {
  const date = parseDate(value);
  return `${date.getMonth() + 1}月${date.getDate()}日`;
};
const isoDate = (date) =>
  `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;

const tasks = [
  ["A", "进场、图纸会审、清单复核、燃气介质及电源接口确认", "全场", "2026-08-10", "2026-08-16", "-", "项目经理/技术负责人", "关键"],
  ["B", "深化设计、设备技术核对、材料报审与甲供设备到货策划", "配电及设备接口", "2026-08-12", "2026-09-05", "A", "技术/材料/电气班", "关键"],
  ["C", "施工围挡、成品保护、临时用电、脚手及墙洞施工准备", "食堂分区", "2026-08-17", "2026-08-28", "A", "安全/土建班", "一般"],
  ["D", "停气审批、阀门关闭、可靠隔断、盲板、氮气置换与检测", "原燃气系统", "2026-08-24", "2026-08-28", "A,C", "业主能源单位/专业班组", "关键"],
  ["E", "原DN120燃气管道拆除、分类转运及封口标识", "原燃气系统", "2026-08-29", "2026-09-10", "D", "专业拆除班", "关键"],
  ["F", "墙体及地坪拆除、开孔、室内外电缆沟和排水沟土建", "土建作业面", "2026-08-24", "2026-09-30", "C", "土建班", "一般"],
  ["G", "C30设备基础、混凝土结构、沟体及盖板施工养护", "设备及沟槽区域", "2026-09-01", "2026-09-30", "F", "土建班", "关键"],
  ["H", "镀锌支架、桥架、钢管、接地干线和等电位端子安装", "配电线路", "2026-09-12", "2026-10-18", "B,F", "电气安装班", "关键"],
  ["I", "GGD柜、动力箱、封闭母线安装及机械电气连接检查", "配电间/用电区", "2026-09-28", "2026-10-20", "B,G", "盘柜母线班", "关键"],
  ["J", "动力电缆敷设、终端制作、穿管引下、编号和防火封堵", "桥架至设备", "2026-10-08", "2026-11-05", "H,I", "电缆班", "关键"],
  ["K", "甲供厨房设备就位、校正、固定及电气末端连接", "厨房设备区", "2026-10-25", "2026-11-20", "G,J", "设备/电气班", "关键"],
  ["L", "砌体、屋面防水、地坪、墙面、吊顶、门窗及排水恢复", "建筑恢复区域", "2026-09-20", "2026-11-30", "F,G", "土建装饰班", "一般"],
  ["M", "绝缘、接地、连续性、相序、保护动作及单机调试", "配电及设备系统", "2026-11-21", "2026-12-10", "J,K", "检测调试组", "关键"],
  ["N", "系统联调、负荷试运行、卫生清洁及使用培训", "全系统", "2026-12-11", "2026-12-22", "L,M", "联调组/使用单位", "关键"],
  ["P", "缺陷整改、竣工资料、联合验收和工程移交", "全场", "2026-12-23", "2026-12-31", "N", "项目经理/资料组", "关键"],
].map((task) => [...task.slice(0, 5), daysInclusive(task[3], task[4]), ...task.slice(5)]);

const taskControls = {
  A: "形成燃气介质确认单、电源接口确认单、设备功率核对表及施工边界清单。",
  B: "GGD柜、母线、电缆、桥架及甲供设备接口参数须经设计、业主和供货方共同确认。",
  C: "食堂运营区与施工区硬隔离，设置人员通道、材料通道、临时用电和消防点。",
  D: "未经能源管理单位批准不得操作阀门；采用可靠隔断，置换检测合格并签字后方可拆除。",
  E: "拆除前复测气体，禁用可能产生火花的机具；分段固定、分段拆除，管口及时封堵。",
  F: "开洞切割前探查暗埋管线；湿法降尘；沟槽设置临边防护并验槽。",
  G: "复核轴线、标高、预埋件和设备孔距；混凝土强度达到要求后方可安装设备。",
  H: "支架间距、桥架跨接、钢管连接和防腐补口逐段检查；接地标识连续清晰。",
  I: "盘柜基础槽钢、垂直度、母线相序、绝缘间距及扭矩值形成记录。",
  J: "电缆敷设前绝缘测试，敷设中控制弯曲半径和牵引力，端头挂牌并完成防火封堵。",
  K: "设备到货开箱、定位、固定和接口验收；未完成保护动作试验不得带负荷运行。",
  L: "防水基层、附加层、搭接、蓄水检查；恢复材料颜色和做法与原建筑协调。",
  M: "完成绝缘电阻、接地电阻、保护导体连续性、相序、漏电保护和断路器动作试验。",
  N: "按设备清单逐台试运行，记录工作电流、温升、异常声响及联锁保护状态。",
  P: "隐蔽、试验、调试、材料、变更、照片和竣工图资料同步闭合后办理移交。",
};

const resources = [
  ["项目管理", "项目经理、技术负责人、施工员、安全员、质量员、材料员、资料员", "7人", "全工期", "协调停气、停送电、食堂运营和分区施工", "项目经理"],
  ["燃气停用与拆除", "能源管理单位接口人员、持证燃气/管道作业人员、检测监护人员", "6人", "8月24日-9月10日", "阀门操作、隔断盲板、氮气置换、检测和拆除", "技术负责人/业主接口"],
  ["电气安装", "盘柜母线、电缆桥架、配管接地班组", "8-10人", "9月12日-11月20日", "柜、母线、桥架、电缆、钢管、接地和设备接线", "电气工长"],
  ["土建装饰", "拆除、钢筋模板混凝土、砌筑、防水、地坪及装饰班组", "10-12人", "8月17日-11月30日", "基础、沟槽、墙洞、屋面、地坪和饰面恢复", "土建工长"],
  ["厨房设备", "设备安装与电气末端接线人员", "6人", "10月25日-11月20日", "甲供设备开箱、就位、固定、接口和单机检查", "设备工长"],
  ["检测调试", "电气试验、设备调试及使用单位配合人员", "4-6人", "11月21日-12月22日", "绝缘、接地、相序、保护、单机和负荷试运行", "技术负责人"],
  ["主要机械", "25t汽车吊、10t平板车、12m登高作业平台", "2/2/5台班", "按吊装计划", "与清单台班匹配，吊装作业单独审批", "起重指挥"],
  ["燃气安全器材", "氮气瓶组、汇流排、减压器、软管、盲板、可燃气体和氧气检测仪", "按专项配置", "停气置换阶段", "检测仪在有效校准期，氮气量以置换计算和现场检测确定", "安全员"],
  ["电气工器具", "液压弯管器、压接钳、牵引机、扭矩扳手、绝缘电阻表、接地电阻表、相序表", "按两作业面配置", "电气施工及调试", "计量器具在检定/校准有效期内", "质量员"],
  ["安全环保", "围挡、灭火器、防火毯、吸附材料、集尘设备、警示牌和个人防护用品", "全场配置", "全工期", "覆盖动火、临电、吊装、高处、拆除、粉尘和食堂卫生隔离", "安全员"],
];

const standards = [
  ["GB 55024-2022", "建筑电气与智能化通用规范", "现行，强制性工程建设规范", "电气安装与验收总控", "https://ebook.chinabuilding.com.cn/c/2022-04-16/98159.shtml"],
  ["GB 50303-2015", "建筑电气工程施工质量验收规范", "现行，强制条文按GB 55024-2022调整", "盘柜、母线、桥架、电缆、接地验收", "住房城乡建设部公告第994号"],
  ["GB 50168-2018", "电气装置安装工程 电缆线路施工及验收标准", "现行", "电缆敷设、终端和试验", "住房城乡建设部工程建设标准"],
  ["GB 50169-2016", "电气装置安装工程 接地装置施工及验收规范", "现行", "接地与等电位", "住房城乡建设部工程建设标准"],
  ["GB 50254-2014", "电气装置安装工程 低压电器施工及验收规范", "现行", "低压电器和盘柜设备", "住房城乡建设部工程建设标准"],
  ["JGJ/T 46-2024", "建筑与市政工程施工现场临时用电安全技术标准", "2025-01-01实施，JGJ 46-2005同时废止", "施工临时用电", "https://www.ehscity.com/law/202409/JGJ/T46-2024《建筑与市政工程施工现场临时用电安全技术标准》自2025年1月1日起实施_1859.html"],
  ["GB 55009-2021", "燃气工程项目规范", "现行，2022-01-01实施", "天然气/城镇燃气系统停用及保护", "网络检索标准封面及住房城乡建设部公告"],
  ["GB 6222-2025", "工业企业煤气安全规范", "现行，2026-05-01实施，代替GB 6222-2005", "仅在介质确认为工业煤气时执行", "https://std.samr.gov.cn/gb/search/gbDetailed?id=42BD8BAE302308A1E06397BE0A0AB162"],
  ["GB 30871-2022", "危险化学品企业特殊作业安全规范", "现行", "适用范围满足时用于动火、盲板抽堵、高处、吊装等作业", "https://std.samr.gov.cn/gb/search/gbDetailed?id=DAB6B92C0762FC96E05397BE0A0A5F84"],
  ["GB 55037-2022", "建筑防火通用规范", "现行，强制性工程建设规范", "防火与施工消防", "住房城乡建设部工程建设标准"],
  ["GB 50204-2015", "混凝土结构工程施工质量验收规范", "现行", "设备基础和混凝土结构", "住房城乡建设部工程建设标准"],
  ["GB 50203-2011", "砌体结构工程施工质量验收规范", "现行", "墙体砌筑恢复", "住房城乡建设部工程建设标准"],
  ["GB 50207-2012", "屋面工程质量验收规范", "现行", "屋面防水恢复", "住房城乡建设部工程建设标准"],
  ["GB 50209-2010", "建筑地面工程施工质量验收规范", "现行", "防静电不发火地面及地砖恢复", "住房城乡建设部工程建设标准"],
  ["GB 50210-2018", "建筑装饰装修工程质量验收标准", "现行", "墙面、吊顶、门窗和饰面", "住房城乡建设部工程建设标准"],
  ["GB 50242-2002", "建筑给水排水及采暖工程施工质量验收规范", "现行", "排水管和沟槽恢复", "住房城乡建设部工程建设标准"],
  ["GB 50300-2013", "建筑工程施工质量验收统一标准", "现行", "单位工程质量验收", "住房城乡建设部工程建设标准"],
  ["JGJ 80-2016", "建筑施工高处作业安全技术规范", "现行", "脚手、登高平台及临边作业", "住房城乡建设部工程建设标准"],
  ["GB 6095-2021", "坠落防护 安全带", "现行", "高处作业个人防护", "全国标准信息公共服务平台"],
  ["GB/T 29639-2020", "生产经营单位生产安全事故应急预案编制导则", "现行", "应急预案和现场处置", "全国标准信息公共服务平台"],
];

const workbook = Workbook.create();
const gantt = workbook.worksheets.add("施工进度计划");
const docGantt = workbook.worksheets.add("文档版横道图");
const logic = workbook.worksheets.add("网络计划逻辑");
const resource = workbook.worksheets.add("主要资源计划");
const boq = workbook.worksheets.add("清单摘要");
const norm = workbook.worksheets.add("规范查新记录");

const colLetter = (n) => {
  let value = n;
  let result = "";
  while (value > 0) {
    const rem = (value - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
};

const applyBase = (sheet) => {
  sheet.showGridLines = false;
};

const setTitle = (sheet, rangeAddress, title) => {
  sheet.mergeCells(rangeAddress);
  const cell = sheet.getRange(rangeAddress.split(":")[0]);
  cell.values = [[title]];
  cell.format = {
    fill: "#1F4E78",
    font: { name: "宋体", size: 16, bold: true, color: "#FFFFFF" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
  };
  sheet.getRange(rangeAddress).format.rowHeight = 30;
};

const styleHeader = (range) => {
  range.format = {
    fill: "#D9EAF7",
    font: { name: "宋体", size: 10, bold: true, color: "#17365D" },
    horizontalAlignment: "center",
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#8EA9C1" },
  };
};

const styleBody = (range, align = "center") => {
  range.format = {
    font: { name: "宋体", size: 10, color: "#222222" },
    horizontalAlignment: align,
    verticalAlignment: "center",
    wrapText: true,
    borders: { preset: "all", style: "thin", color: "#B7C3CD" },
  };
};

for (const sheet of workbook.worksheets.items) applyBase(sheet);

const weeks = [];
for (let cursor = new Date(startDate); cursor <= finishDate; cursor = new Date(cursor.getTime() + 7 * oneDay)) {
  const end = new Date(Math.min(cursor.getTime() + 6 * oneDay, finishDate.getTime()));
  weeks.push({
    start: new Date(cursor),
    end,
    label: `${cursor.getMonth() + 1}.${cursor.getDate()}-${end.getMonth() + 1}.${end.getDate()}`,
  });
}

const ganttLastCol = colLetter(8 + weeks.length);
setTitle(gantt, `A1:${ganttLastCol}1`, `${project}施工进度计划（2026年8月10日-12月31日，共${totalDays}日历天）`);
gantt.getRange("A2:H2").values = [[
  "编码", "工作内容", "施工区域", "开始", "完成", "持续天数", "紧前工作", "责任班组",
]];
styleHeader(gantt.getRange("A2:H3"));
for (let c = 1; c <= 8; c++) {
  const letter = colLetter(c);
  gantt.mergeCells(`${letter}2:${letter}3`);
}
weeks.forEach((week, index) => {
  const col = colLetter(9 + index);
  gantt.getRange(`${col}2`).values = [[`第${index + 1}周`]];
  gantt.getRange(`${col}3`).values = [[week.label]];
});
styleHeader(gantt.getRange(`I2:${ganttLastCol}3`));

gantt.getRange(`A4:H${tasks.length + 3}`).values = tasks.map((task) => [
  task[0], task[1], task[2], displayDate(task[3]), displayDate(task[4]), task[5], task[6], task[7],
]);
styleBody(gantt.getRange(`A4:H${tasks.length + 3}`));
gantt.getRange(`B4:C${tasks.length + 3}`).format.horizontalAlignment = "left";
styleBody(gantt.getRange(`I4:${ganttLastCol}${tasks.length + 3}`));

tasks.forEach((task, rowIndex) => {
  const row = rowIndex + 4;
  const taskStart = parseDate(task[3]);
  const taskFinish = parseDate(task[4]);
  weeks.forEach((week, weekIndex) => {
    if (taskStart <= week.end && taskFinish >= week.start) {
      const cell = gantt.getRange(`${colLetter(9 + weekIndex)}${row}`);
      cell.format.fill = task[8] === "关键" ? "#ED7D31" : "#70AD47";
      if (taskStart >= week.start && taskStart <= week.end) {
        cell.values = [[task[0]]];
        cell.format.font = { name: "宋体", size: 9, bold: true, color: "#FFFFFF" };
      }
    }
  });
});

const milestoneRow = tasks.length + 4;
gantt.getRange(`A${milestoneRow}:H${milestoneRow}`).values = [[
  "里程碑", "联合验收并完成工程移交", "全场", "12月31日", "12月31日", 1, "P", "项目经理/业主",
]];
styleBody(gantt.getRange(`A${milestoneRow}:H${milestoneRow}`));
const finalWeekCell = gantt.getRange(`${ganttLastCol}${milestoneRow}`);
finalWeekCell.values = [["◆"]];
finalWeekCell.format = {
  fill: "#C00000",
  font: { name: "宋体", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#B7C3CD" },
};

const noteRow = milestoneRow + 1;
gantt.getRange(`A${noteRow}`).values = [["说明"]];
gantt.mergeCells(`B${noteRow}:H${noteRow}`);
gantt.getRange(`B${noteRow}`).values = [["橙色为关键工作，绿色为一般工作；停气、盲板隔断、置换检测、停送电和负荷试运行均须履行书面审批及交接手续。"]];
styleBody(gantt.getRange(`A${noteRow}:H${noteRow}`));
gantt.getRange(`B${noteRow}`).format.horizontalAlignment = "left";
gantt.getRange(`I${noteRow}:${ganttLastCol}${noteRow}`).merge();
gantt.getRange(`I${noteRow}`).values = [["计划以总日历工期控制，实际施工须结合食堂运营窗口和甲供设备到货动态调整。"]];
styleBody(gantt.getRange(`I${noteRow}:${ganttLastCol}${noteRow}`));
gantt.getRange(`I${noteRow}`).format.horizontalAlignment = "left";

[9, 42, 19, 10, 10, 10, 12, 24].forEach((width, index) => {
  const col = colLetter(index + 1);
  gantt.getRange(`${col}1:${col}${noteRow}`).format.columnWidth = width;
});
gantt.getRange(`I1:${ganttLastCol}${noteRow}`).format.columnWidth = 7.2;
gantt.getRange("2:3").format.rowHeight = 24;
gantt.getRange(`4:${milestoneRow}`).format.rowHeight = 36;
gantt.getRange(`${noteRow}:${noteRow}`).format.rowHeight = 42;
gantt.freezePanes.freezeRows(3);
gantt.freezePanes.freezeColumns(8);

const months = [
  { label: "8月", start: parseDate("2026-08-01"), end: parseDate("2026-08-31") },
  { label: "9月", start: parseDate("2026-09-01"), end: parseDate("2026-09-30") },
  { label: "10月", start: parseDate("2026-10-01"), end: parseDate("2026-10-31") },
  { label: "11月", start: parseDate("2026-11-01"), end: parseDate("2026-11-30") },
  { label: "12月", start: parseDate("2026-12-01"), end: parseDate("2026-12-31") },
];
setTitle(docGantt, "A1:G1", `${project}施工进度总控横道图（2026年8月10日-12月31日）`);
docGantt.getRange("A2:G2").values = [["编码", "主要工作", ...months.map((month) => month.label)]];
styleHeader(docGantt.getRange("A2:G2"));
docGantt.getRange(`A3:B${tasks.length + 2}`).values = tasks.map((task) => [task[0], task[1]]);
styleBody(docGantt.getRange(`A3:G${tasks.length + 2}`));
docGantt.getRange(`B3:B${tasks.length + 2}`).format.horizontalAlignment = "left";
tasks.forEach((task, rowIndex) => {
  const row = rowIndex + 3;
  const taskStart = parseDate(task[3]);
  const taskFinish = parseDate(task[4]);
  let marked = false;
  months.forEach((month, monthIndex) => {
    if (taskStart <= month.end && taskFinish >= month.start) {
      const cell = docGantt.getRange(`${colLetter(3 + monthIndex)}${row}`);
      cell.format.fill = task[8] === "关键" ? "#ED7D31" : "#70AD47";
      if (!marked) {
        cell.values = [[task[0]]];
        cell.format.font = { name: "宋体", size: 10, bold: true, color: "#FFFFFF" };
        marked = true;
      }
    }
  });
});
const docMilestoneRow = tasks.length + 3;
docGantt.getRange(`A${docMilestoneRow}:B${docMilestoneRow}`).values = [["◆", "12月31日联合验收并完成工程移交"]];
styleBody(docGantt.getRange(`A${docMilestoneRow}:G${docMilestoneRow}`));
docGantt.getRange(`G${docMilestoneRow}`).values = [["◆"]];
docGantt.getRange(`G${docMilestoneRow}`).format = {
  fill: "#C00000",
  font: { name: "宋体", size: 10, bold: true, color: "#FFFFFF" },
  horizontalAlignment: "center",
  verticalAlignment: "center",
  borders: { preset: "all", style: "thin", color: "#B7C3CD" },
};
const docNoteRow = docMilestoneRow + 1;
docGantt.mergeCells(`A${docNoteRow}:G${docNoteRow}`);
docGantt.getRange(`A${docNoteRow}`).values = [["橙色为关键工作，绿色为一般工作；停气置换、停送电及负荷试运行须履行书面审批。详细21周计划见“施工进度计划”工作表。"]];
styleBody(docGantt.getRange(`A${docNoteRow}:G${docNoteRow}`), "left");
[7, 48, 14, 14, 14, 14, 14].forEach((width, index) => {
  docGantt.getRange(`${colLetter(index + 1)}1:${colLetter(index + 1)}${docNoteRow}`).format.columnWidth = width;
});
docGantt.getRange(`3:${docMilestoneRow}`).format.rowHeight = 30;
docGantt.getRange(`${docNoteRow}:${docNoteRow}`).format.rowHeight = 36;
docGantt.freezePanes.freezeRows(2);

setTitle(logic, "A1:K1", `${project}网络计划逻辑及控制要点`);
logic.getRange("A2:K2").values = [[
  "编码", "工作内容", "区域", "开始", "完成", "持续天数", "紧前工作", "责任班组", "性质", "控制要点", "验收输出",
]];
styleHeader(logic.getRange("A2:K2"));
logic.getRange(`A3:I${tasks.length + 2}`).values = tasks.map((task) => [
  task[0], task[1], task[2], task[3], task[4], task[5], task[6], task[7], task[8],
]);
logic.getRange(`J3:J${tasks.length + 2}`).values = tasks.map((task) => [taskControls[task[0]]]);
logic.getRange(`K3:K${tasks.length + 2}`).values = tasks.map((task) => {
  const outputs = {
    A: "接口确认单/会审记录", B: "报审/到货计划", C: "开工条件验收", D: "停气交接/置换检测记录",
    E: "拆除交接记录", F: "隐蔽/验槽记录", G: "混凝土/基础验收", H: "桥架配管接地隐蔽",
    I: "盘柜母线安装记录", J: "绝缘/电缆敷设记录", K: "开箱/设备安装记录", L: "建筑分项验收",
    M: "电气试验记录", N: "试运行/培训记录", P: "联合验收/移交单",
  };
  return [outputs[task[0]]];
});
styleBody(logic.getRange(`A3:K${tasks.length + 2}`));
logic.getRange(`B3:C${tasks.length + 2}`).format.horizontalAlignment = "left";
logic.getRange(`J3:K${tasks.length + 2}`).format.horizontalAlignment = "left";
[8, 42, 22, 12, 12, 10, 14, 24, 10, 62, 26].forEach((width, index) => {
  logic.getRange(`${colLetter(index + 1)}1:${colLetter(index + 1)}${tasks.length + 2}`).format.columnWidth = width;
});
logic.getRange(`3:${tasks.length + 2}`).format.rowHeight = 46;
logic.freezePanes.freezeRows(2);

setTitle(resource, "A1:F1", `${project}主要资源计划`);
resource.getRange("A2:F2").values = [["资源类别", "人员/机械/器材", "计划投入", "使用阶段", "配置依据及用途", "责任岗位"]];
styleHeader(resource.getRange("A2:F2"));
resource.getRange(`A3:F${resources.length + 2}`).values = resources;
styleBody(resource.getRange(`A3:F${resources.length + 2}`));
resource.getRange(`A3:B${resources.length + 2}`).format.horizontalAlignment = "left";
resource.getRange(`D3:F${resources.length + 2}`).format.horizontalAlignment = "left";
[18, 58, 18, 24, 76, 24].forEach((width, index) => {
  resource.getRange(`${colLetter(index + 1)}1:${colLetter(index + 1)}${resources.length + 2}`).format.columnWidth = width;
});
resource.getRange(`3:${resources.length + 2}`).format.rowHeight = 48;
resource.freezePanes.freezeRows(2);

setTitle(boq, "A1:H1", `${project}工程量清单摘要（共${boqItems.length}项）`);
boq.getRange("A2:H2").values = [["专业", "序号", "项目编码", "项目名称", "单位", "工程量", "项目特征", "主要工作内容"]];
styleHeader(boq.getRange("A2:H2"));
boq.getRange(`A3:H${boqItems.length + 2}`).values = boqItems.map((item) => [
  item.category, item.seq, item.code, item.name, item.unit, item.quantity, item.feature, item.work,
]);
styleBody(boq.getRange(`A3:H${boqItems.length + 2}`));
boq.getRange(`A3:A${boqItems.length + 2}`).format.horizontalAlignment = "left";
boq.getRange(`C3:D${boqItems.length + 2}`).format.horizontalAlignment = "left";
boq.getRange(`G3:H${boqItems.length + 2}`).format.horizontalAlignment = "left";
[18, 8, 20, 34, 10, 14, 48, 58].forEach((width, index) => {
  boq.getRange(`${colLetter(index + 1)}1:${colLetter(index + 1)}${boqItems.length + 2}`).format.columnWidth = width;
});
boq.getRange(`3:${boqItems.length + 2}`).format.rowHeight = 42;
boq.freezePanes.freezeRows(2);
boq.freezePanes.freezeColumns(2);

setTitle(norm, "A1:E1", `${project}规范查新记录（查新日期：2026年7月31日）`);
norm.getRange("A2:E2").values = [["标准编号", "标准名称", "状态/版本核验", "本项目适用内容", "核验来源"]];
styleHeader(norm.getRange("A2:E2"));
norm.getRange(`A3:E${standards.length + 2}`).values = standards;
styleBody(norm.getRange(`A3:E${standards.length + 2}`));
norm.getRange(`A3:E${standards.length + 2}`).format.horizontalAlignment = "left";
[20, 50, 42, 54, 70].forEach((width, index) => {
  norm.getRange(`${colLetter(index + 1)}1:${colLetter(index + 1)}${standards.length + 2}`).format.columnWidth = width;
});
norm.getRange(`3:${standards.length + 2}`).format.rowHeight = 44;
norm.freezePanes.freezeRows(2);

const inspect = await workbook.inspect({
  kind: "table",
  sheetId: "施工进度计划",
  range: `A1:${ganttLastCol}${noteRow}`,
  maxChars: 20000,
  tableMaxRows: 30,
  tableMaxCols: 40,
  tableMaxCellChars: 180,
});
await fs.writeFile(path.join(workDir, "spreadsheet", "schedule_inspect.ndjson"), inspect.ndjson, "utf8");

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 100 },
  maxChars: 8000,
});
await fs.writeFile(path.join(workDir, "spreadsheet", "schedule_errors.ndjson"), errors.ndjson, "utf8");

for (const sheet of workbook.worksheets.items) {
  const rendered = await workbook.render({
    sheetName: sheet.name,
    autoCrop: "all",
    scale: sheet.name === "施工进度计划" ? 1.15 : sheet.name === "文档版横道图" ? 1.2 : 0.9,
    format: "png",
  });
  const safeName = sheet.name.replace(/[\\/:*?"<>|]/g, "_");
  await fs.writeFile(
    path.join(renderDir, `${safeName}.png`),
    new Uint8Array(await rendered.arrayBuffer()),
  );
}

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
console.log(`total_days=${totalDays} tasks=${tasks.length} boq_items=${boqItems.length} standards=${standards.length}`);

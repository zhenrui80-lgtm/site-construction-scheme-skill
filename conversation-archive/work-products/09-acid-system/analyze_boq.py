import json
from collections import defaultdict
from pathlib import Path


ROOT = Path(r"C:\Users\Administrator\Desktop\Metallurgical_BOQ_AgentSkill_Fixed\agents\_work_quote_scheme")
items = json.loads((ROOT / "boq_extract" / "items.json").read_text(encoding="utf-8"))

groups = [
    ("酸雾净化-土建", 8, 58),
    ("酸雾净化-设备", 60, 129),
    ("酸雾净化-管道", 130, 342),
    ("酸雾净化-电气设备", 343, 355),
    ("酸雾净化-电气材料", 356, 376),
    ("酸雾净化-仪表设备", 377, 402),
    ("酸雾净化-仪表材料", 403, 417),
    ("酸循环-设备", 419, 433),
    ("酸循环-管道", 434, 483),
    ("环境整治-101酸罐区域拆除", 485, 502),
    ("环境整治-101冷轧厂房拆除", 503, 515),
    ("环境整治-土建", 516, 580),
]


def group_name(row):
    for name, start, end in groups:
        if start <= row <= end:
            return name
    return "其他"


grouped = defaultdict(list)
for item in items:
    grouped[group_name(item["row"])].append(item)

lines = []
lines.append("项目名称：101机组酸循环系统和酸雾系统综合改造")
lines.append(f"清单项目数：{len(items)}")
grand_total = sum(float(item["total"] or 0) for item in items)
lines.append(f"清单合价（读取值）：{grand_total:,.2f}元")
lines.append("")

for name, _, _ in groups:
    group_items = grouped[name]
    total = sum(float(item["total"] or 0) for item in group_items)
    lines.append(f"## {name}（{len(group_items)}项，合价{total:,.2f}元）")
    for item in group_items:
        feature = str(item["feature"]).replace("\n", "；")
        if len(feature) > 120:
            feature = feature[:117] + "..."
        lines.append(
            f"- 行{item['row']} 序号{item['serial']} {item['name']}：{item['quantity']}{item['unit']}；{feature}"
        )
    lines.append("")

keywords = [
    "基础",
    "防腐",
    "净化塔",
    "风机",
    "烟囱",
    "泵",
    "储罐",
    "PPH",
    "风管",
    "管道",
    "法兰",
    "桥架",
    "电缆",
    "配电",
    "PLC",
    "仪表",
    "拆除净化设备",
    "拆除旧酸液储罐",
    "钢筋混凝土",
    "网格板",
]
lines.append("## 关键词汇总")
for keyword in keywords:
    matched = [item for item in items if keyword.lower() in (str(item["name"]) + str(item["feature"])).lower()]
    if matched:
        lines.append(f"### {keyword}（{len(matched)}项）")
        for item in matched[:30]:
            lines.append(f"- {group_name(item['row'])}｜{item['name']}｜{item['quantity']}{item['unit']}")

output = ROOT / "boq_summary.txt"
output.write_text("\n".join(lines), encoding="utf-8")
print(output)

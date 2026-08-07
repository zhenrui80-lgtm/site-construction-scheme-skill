from pathlib import Path
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE = Path(__file__).resolve().parent
OUT = BASE / "焦一作业区废气收集系统整改及环境整治施工方案.docx"
OUT_XLSX = BASE / "焦一作业区废气收集系统整改及环境整治100天进度计划.xlsx"


def set_font(run, name="宋体", size=10.5, bold=False, color=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn("w:eastAsia"), name)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)


def setup(doc):
    sec = doc.sections[0]
    sec.top_margin = Cm(1.9)
    sec.bottom_margin = Cm(1.7)
    sec.left_margin = Cm(2.2)
    sec.right_margin = Cm(2.0)
    sec.different_first_page_header_footer = True
    hp = sec.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run("焦一作业区废气收集系统整改及环境整治施工方案")
    set_font(r, size=9, color=(100, 100, 100))
    fp = sec.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("-  -")
    set_font(r, size=9)
    style = doc.styles["Normal"]
    style.font.name = "宋体"
    style._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    style.font.size = Pt(10.5)


def para(doc, text="", first=True, align=None, size=10.5, bold=False, font="宋体"):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.line_spacing = 1.08
    p.paragraph_format.space_after = Pt(2)
    if first:
        p.paragraph_format.first_line_indent = Pt(21)
    r = p.add_run(text)
    set_font(r, name=font, size=size, bold=bold)
    return p


def heading(doc, text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(7 if level == 1 else 3)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    set_font(r, name="黑体", size=14 if level == 1 else 12, bold=True, color=(31, 78, 121) if level == 1 else None)
    if level == 1:
        p_pr = p._p.get_or_add_pPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:fill"), "EAF3F8")
        p_pr.append(shd)
    return p


def cell(c, text, bold=False, align=WD_ALIGN_PARAGRAPH.CENTER, size=9):
    c.text = ""
    p = c.paragraphs[0]
    p.alignment = align
    r = p.add_run(str(text))
    set_font(r, size=size, bold=bold)
    c.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def shade(c, fill="D9EAF7"):
    tc_pr = c._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def table(doc, headers, rows):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        cell(t.rows[0].cells[i], h, True)
        shade(t.rows[0].cells[i])
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cell(cells[i], v, align=WD_ALIGN_PARAGRAPH.LEFT if i in (1, 3) else WD_ALIGN_PARAGRAPH.CENTER)
    return t


doc = Document()
setup(doc)

for _ in range(5):
    para(doc, "", first=False)
para(doc, "焦一作业区废气收集系统整改及环境整治", align=WD_ALIGN_PARAGRAPH.CENTER, first=False, size=22, bold=True, font="黑体")
para(doc, "施工方案", align=WD_ALIGN_PARAGRAPH.CENTER, first=False, size=26, bold=True, font="黑体")
para(doc, "（工期100天）", align=WD_ALIGN_PARAGRAPH.CENTER, first=False, size=14)
for _ in range(6):
    para(doc, "", first=False)
cover = doc.add_table(rows=5, cols=2)
cover.style = "Table Grid"
cover.alignment = WD_TABLE_ALIGNMENT.CENTER
for row, vals in zip(cover.rows, [
    ("工程名称", "焦一作业区废气收集系统整改及环境整治"),
    ("施工单位", "武汉兴达建筑安装工程有限公司"),
    ("计划工期", "100日历天"),
    ("编制日期", "2026年6月"),
    ("编制/审核/批准", "编制：              审核：              批准："),
]):
    cell(row.cells[0], vals[0], True)
    cell(row.cells[1], vals[1], align=WD_ALIGN_PARAGRAPH.LEFT, size=10.5)

doc.add_page_break()
para(doc, "目  录", align=WD_ALIGN_PARAGRAPH.CENTER, first=False, size=18, bold=True, font="黑体")
for item in [
    "一、编制说明",
    "二、工程概况",
    "三、施工部署",
    "四、施工准备",
    "五、主要施工方法",
    "六、100天施工进度计划",
    "七、质量保证措施",
    "八、安全文明施工措施",
    "九、高处、吊装、动火及有限空间专项措施",
    "十、环保、成品保护及现场恢复",
    "十一、资料管理与验收",
    "十二、应急处置措施",
]:
    p = para(doc, item, first=False, size=11)
    p.paragraph_format.left_indent = Cm(3)
doc.add_page_break()

heading(doc, "一、编制说明")
para(doc, "本施工方案依据《焦一作业区废气收集系统整改及环境整治工程量清单》、现行施工验收规范、焦化厂区安全管理要求、废气收集系统运行特点及类似工业厂区环保整治工程施工经验编制。方案适用于老生化废气收集系统风机恢复、HLA池旁备用风机搬迁及还原、好氧池旁新风机安装、玻璃钢风管安装、电气电缆和桥架敷设、光缆及消防点位图形化、集水坑和围堰防腐、污泥脱水间墙面整治、雨水沟和围墙修补等施工内容。")
para(doc, "本工程作业范围分散，专业交叉多，涉及风机、风管、电气、仪控、土建、防腐、保温、环境整治及高处机械作业。施工组织坚持“先确认停运隔离、后拆除；先支护防护、后高处作业；先设备基础和管线条件复核、后吊装安装；先单机试验、后系统联调；先隐蔽验收、后恢复”的原则。")

heading(doc, "二、工程概况")
para(doc, "工程名称为焦一作业区废气收集系统整改及环境整治，计划工期100日历天。清单范围共110项，主要分为老生化废气收集系统风机恢复、备用风机搬迁和还原、好氧池旁新风机安装、机械台班及安全措施、集水坑及围堰防腐、桥架电缆光缆系统、消防点位图形化和环境整治等内容。")
para(doc, "老生化废气收集系统风机恢复主要包括好氧池旁原有风机2台拆除、DN1200玻璃钢风管25米拆除、DN1200蝶阀4台拆除、就地操作箱6台拆除、动力电缆和控制电缆各类合计1200米拆除、建筑垃圾外运及人工杂物清理。")
para(doc, "HLA池旁备用风机搬迁及还原主要包括Q=45000m³/h、P=4500Pa、N=110kW玻璃钢防爆变频风机保护性拆除、搬迁和安装，DN1200阀门保护性拆除与安装，就地操作箱拆装，DN1200盲板封堵和拆除，DN1200玻璃钢风管、弯头、阀门安装，以及动力电缆、控制电缆敷设。")
para(doc, "好氧池旁新风机恢复主要包括新风机2台安装，DN1200玻璃钢风管50米、90°弯头3个、等径三通2个、DN1200蝶阀4台、就地操作箱6台安装，热镀锌梯级式桥架20米、支架1吨、三防投光灯4套、电线和控制线敷设、DN25线管80米、管道及支架除锈刷漆260平方米、混凝土池壁涂料层铲除及外墙涂料150平方米、生物滤池不锈钢保护层50平方米、保温岩棉更换150平方米及风机设备系统调试。")
para(doc, "环境整治及配套措施包括登高车16m台班、20m登高车台班、25T汽车吊、15T平板货车、施工围挡、安全走绳，集水坑土方开挖、C30 P6钢筋混凝土浇筑、三布四涂环氧树脂防腐、20mm耐酸砖防腐、钢格栅盖板、管道泵、DN80法兰、动力及控制电缆、SC DN25线管、就地控制箱；同时包含钢制槽式桥架120米、桥架支架0.5吨、穿线管90米、防火封堵12处、GYTA-12B1光缆7200米、光纤熔接、光纤收发器、ODF、计算机显示器、消防点位图形化编程和信号调试。")
para(doc, "土建及环境修复还包括污泥间排水沟40米、老生化围墙破损修补16立方米、雨水沟修补30米、污泥脱水间内墙墙皮铲除720平方米、内墙刮腻子720平方米、围堰耐酸砖面层拆除235平方米、围堰垃圾清理运输、环氧树脂防腐235平方米、耐酸砖防腐235平方米。")

heading(doc, "三、施工部署")
para(doc, "项目部实行项目经理负责制，设置技术负责人、施工员、安全员、质量员、材料员、资料员及专业班组长。根据工程内容配置设备安装班组、玻璃钢风管班组、电气仪控班组、光缆及消防系统调试班组、土建修复班组、防腐保温班组、吊装运输班组和清理恢复班组。")
para(doc, "施工总体顺序为：现场交接与安全隔离、原有设备风管和电缆拆除、备用风机保护性拆除及搬迁、HLA池旁功能恢复、好氧池旁新风机及风管安装、电气电缆桥架和线管施工、光缆及消防点位系统施工、集水坑和围堰防腐、污泥脱水间及排水沟围墙环境整治、单机调试、系统联调、现场恢复和竣工验收。")
para(doc, "本工程按100天组织，采用多专业分区流水施工。风机、DN1200玻璃钢风管和电气仪控为主线；土建修复、防腐、保温、光缆敷设和环境整治穿插进行。高处作业和吊装作业应避开生产高峰和不利天气，并服从建设单位对废气收集系统停运、切换和试运行的统一安排。")

heading(doc, "四、施工准备")
para(doc, "技术准备包括清单复核、现场踏勘、风机基础和接口复核、DN1200风管走向复核、设备编号和电缆回路核对、吊装路径确认、停电挂牌方案、风管胶衣连接工艺交底、电气调试交底、防腐保温做法交底、土建恢复做法交底和系统联调方案交底。")
para(doc, "材料准备包括风机、阀门、就地操作箱、玻璃钢风管、弯头、三通、盲板、热镀锌桥架、桥架支架、DN25线管、电力电缆、控制电缆、光缆、光纤收发器、ODF、三防投光灯、管道泵、耐酸砖、环氧树脂、玻璃布、保温岩棉、不锈钢保护层、钢格栅盖板、外墙涂料、腻子和混凝土等。甲供材料到场后应办理交接和外观验收。")
para(doc, "现场准备包括施工围挡、警戒标识、作业平台、安全走绳、登高车作业区域、汽车吊站位、材料堆放区、废料暂存区、临时用电、消防器材、照明通风、排水降水和道路保洁。涉及好氧池、HLA池、生物滤池、污泥脱水间、围堰和集水坑区域的作业，应先确认池边防护、地面承载、腐蚀介质、臭气和湿滑风险。")

heading(doc, "五、主要施工方法")
heading(doc, "1. 原有设备、风管、电缆拆除", 2)
para(doc, "拆除前由建设单位确认废气收集系统停运、设备断电挂牌、风管隔离、阀门关闭和现场安全条件。对原有风机、DN1200玻璃钢风管、蝶阀、就地操作箱、电力电缆和控制电缆进行编号标识，拆除范围经现场签认后实施。")
para(doc, "风机拆除前先解除电气连接、拆除附件和连接螺栓，采用汽车吊或倒链受控吊离，不得强行撬动基础和连接法兰。玻璃钢风管拆除采用分段切割或拆卸，切口和碎片及时收集，防止玻璃钢纤维划伤人员。电缆拆除须先停电验电，按回路拆除，揭盖板后及时恢复临时防护。")

heading(doc, "2. 备用风机搬迁、还原及新风机安装", 2)
para(doc, "备用风机保护性拆除和搬迁应先确认风机型号、重量、吊点、底座连接、进出口法兰和电机保护要求。吊装前检查吊具、钢丝绳、卸扣和吊车站位，试吊确认平衡后正式吊运。运输过程中对玻璃钢壳体、电机、隔音罩和变频附件采取防碰撞措施。")
para(doc, "风机安装前复核基础标高、地脚螺栓、进出口方向、检修空间和风管接口。风机就位后调整水平度和同心度，紧固地脚螺栓，安装阀门、就地操作箱和附件。安装完成后进行手动盘车、电机绝缘测试、空载试运行和系统联动调试。")

heading(doc, "3. DN1200玻璃钢风管、弯头、三通及盲板施工", 2)
para(doc, "玻璃钢风管安装前检查管节、弯头、三通、法兰、胶衣层和接口外观，破损、裂纹、分层和明显变形不得安装。架空敷设时先安装支架并复核标高和坡向，管节吊装就位后临时固定，再进行接口处理。")
para(doc, "胶衣连接部位应打磨、清洁、干燥，接口处按要求缠绕玻璃丝布并涂刷环氧树脂，搭接宽度、层数和固化时间满足施工要求。DN1200盲板制作安装应保证法兰对口平整，垫片完整，螺栓对称紧固。作业高度4-5m及以上时，登高车和安全走绳必须布置到位。")

heading(doc, "4. 电气、仪控、桥架及光缆施工", 2)
para(doc, "电气施工前核对电缆规格、走向、桥架容量和设备端子。ZRC-YJV动力电缆、ZRC-KVVP控制电缆、YJV-3*2.5、KVVP-14*1.5、YJV-4*6、KVVP-5*1.5等敷设时应开盘检查、排列整理、固定挂牌，转弯处满足弯曲半径要求。")
para(doc, "钢制桥架和支架安装前复核走向、支吊架间距和安装高度。桥架连接牢固，盖板齐全，跨接接地可靠。DN25镀锌线管、厚壁SC管敷设时应切管、套丝、煨弯、管口护口，作业高度2-15m时设置登高车、安全带和下方警戒。")
para(doc, "GYTA-12B1光缆敷设应按配盘和路径组织，穿放过程中控制牵引力和弯曲半径，管孔清理后穿放引线，盘留余长并做好标识。光纤熔接后进行测试，光纤收发器、收发器箱、光缆分配架、ODF、跳线和计算机显示器安装完成后进行通信测试、消防点位图形化编程和报警/故障/联动信号模拟调试。")

heading(doc, "5. 土建、集水坑、围堰及环境整治", 2)
para(doc, "集水坑施工包括围堰地面人工破除、土方开挖、废渣外运、C30 P6钢筋混凝土浇筑、环氧树脂三布四涂防腐、20mm耐酸砖防腐、钢格栅盖板、管道泵、法兰、电缆、线管和控制箱安装。基坑开挖深度约1.5m，应做好排水、边坡和临边防护。")
para(doc, "围堰耐酸砖面层拆除235平方米后，应清理垃圾并修补基层，环氧树脂防腐和耐酸砖施工前基层应坚实、干燥、无浮灰、无油污。三布四涂应分层施工，耐酸砖铺贴应控制坡向、缝宽、平整度和勾缝密实度。")
para(doc, "环境整治包括污泥间排水沟浇筑、老生化围墙破损修补、雨水沟修补、污泥脱水间墙皮铲除、刮腻子和池壁外墙涂料施工。施工中应控制扬尘、落灰、垃圾外运和成品保护，完成后恢复现场6S状态。")

heading(doc, "六、100天施工进度计划")
para(doc, "本工程计划总工期100日历天。具体开工日期以建设单位开工令为准，以下按第1天起算。施工采用风机系统主线推进、土建防腐和电气仪控穿插作业的方式组织。")
table(doc, ["序号", "施工阶段", "计划时间", "主要工作内容"], [
    ["1", "进场准备及系统交接", "第1-7天", "现场踏勘、作业票、停电挂牌、材料计划、围挡、安全走绳、吊装路线确认。"],
    ["2", "原有设备及管线拆除", "第8-18天", "原风机、风管、阀门、操作箱、电缆拆除，垃圾外运和现场清理。"],
    ["3", "备用风机搬迁及HLA池功能恢复", "第15-32天", "备用风机拆装移位、盲板封堵、DN1200风管、阀门、操作箱、电缆敷设。"],
    ["4", "HLA池和好氧池风机还原", "第28-45天", "备用风机还原、新订风机安装、风管拆装、阀门和操作箱安装。"],
    ["5", "好氧池旁新风机系统安装", "第40-62天", "DN1200风管、弯头、三通、桥架、支架、投光灯、线管、电线、除锈刷漆。"],
    ["6", "电气仪控、光缆及消防图形化", "第55-78天", "桥架、穿线管、动力和控制电缆、光缆、光纤设备、计算机显示器、软件授权和信号调试。"],
    ["7", "集水坑、围堰防腐及管道泵", "第50-82天", "破除、开挖、混凝土集水坑、环氧树脂、耐酸砖、钢格栅、泵和控制箱安装。"],
    ["8", "环境整治及土建恢复", "第70-92天", "排水沟、围墙、雨水沟、污泥脱水间墙面、池壁涂料、保温保护层。"],
    ["9", "系统调试、整改及验收", "第90-100天", "风机单机试车、系统联调、消防点位调试、资料整理、缺陷整改、竣工验收。"],
])

heading(doc, "七、质量保证措施")
para(doc, "质量目标为一次验收合格。项目部建立材料进场验收、样板先行、工序交接、隐蔽验收、试验调试和整改闭环制度。重点控制风机安装水平度、风管接口密封、DN1200法兰连接、胶衣连接固化质量、电缆绝缘测试、桥架接地跨接、光纤熔接测试、集水坑防渗、防腐层厚度、耐酸砖空鼓率和环境整治观感质量。")
para(doc, "甲供设备和材料到场后应由建设单位、施工单位共同验收，确认型号、数量、外观和附件完整性。乙供材料应具备合格证和检测资料，防腐材料、耐酸砖、混凝土、钢材、电缆桥架等应符合清单和规范要求。")
para(doc, "风机系统调试应按单机试车、风管检查、阀门动作、操作箱测试、联动试验顺序进行。电缆和光缆施工应形成绝缘测试、熔接测试和系统调试记录。隐蔽工程未经确认不得覆盖。")

heading(doc, "八、安全文明施工措施")
para(doc, "本工程安全风险主要包括高处坠落、物体打击、吊装伤害、动火火灾、触电、玻璃钢粉尘刺激、池边坠落、有限空间及生产区域交叉作业。所有人员进场前必须接受厂区安全教育和专项安全技术交底，严格执行高处、动火、吊装、临电、受限空间和停送电作业审批。")
para(doc, "高处作业使用登高车、脚手架或可靠平台，安全带高挂低用，安全走绳设置牢固。吊装作业由专人指挥，吊物下方禁止站人。动火作业执行一点一证一监护，设置防火毯、接火措施和灭火器，作业后进行火后检查。临时用电由专职电工接设，一机一闸一漏一箱。")
para(doc, "施工围挡按清单设置，高度1.8m，施工区域与生产区、道路和池边有效隔离。玻璃钢切割和打磨时作业人员佩戴护目镜、防尘口罩和手套，碎屑及时收集。现场材料分类堆放，做到工完料尽场地清。")

heading(doc, "九、高处、吊装、动火及有限空间专项措施")
para(doc, "作业高度2-15m的桥架、线管、风管、盲板和支架施工必须落实登高车、安全走绳、双钩安全带和下方警戒。登高车作业前检查支腿、地面承载、作业半径和上方障碍物，严禁超载和移动车辆时人员留在平台内。")
para(doc, "25T汽车吊吊装风机、DN1200风管和大型构件前，应核对吊重、吊点、吊装半径和站位，必要时铺设垫板。风机吊装应保护玻璃钢壳体、电机、隔音罩和变频附件，严禁碰撞池壁、风管、桥架和操作箱。")
para(doc, "池边、集水坑、管沟和污泥脱水间局部低洼潮湿区域作业前应检测通风、照明和积水情况，设置上下通道和监护人。发现异味、缺氧、积水上涨或人员不适时立即撤离。")

heading(doc, "十、环保、成品保护及现场恢复")
para(doc, "拆除垃圾、玻璃钢碎料、废电缆、废包装、涂料桶、建筑垃圾和施工废渣分类收集，按建设单位指定地点外运处置。墙皮铲除、破除、切割、打磨和喷刷涂料时采取洒水、遮挡、通风和清扫措施，减少扬尘和异味影响。")
para(doc, "已安装风机、风管、桥架、电缆、光纤设备、控制箱、防腐层、耐酸砖、保温保护层和墙面涂料应设置成品保护。防腐和涂料固化期间禁止踩踏、冲水和堆放材料；光缆余长、终端盒和收发器箱不得挤压碰撞。")
para(doc, "环境整治完成后恢复道路、排水沟、围墙、池壁、地坪和围堰的整洁状态。拆除临时围挡和警戒前，应确认无遗留工具、焊渣、玻璃钢碎片、线头、螺栓和尖锐废料。")

heading(doc, "十一、资料管理与验收")
para(doc, "资料员随工程同步收集施工方案、技术交底、安全交底、作业票、材料合格证、甲供材料交接单、吊装检查记录、设备安装记录、风管接口检查记录、电缆绝缘测试记录、光纤熔接测试记录、消防点位调试记录、防腐施工记录、耐酸砖检查记录、混凝土浇筑记录、隐蔽验收、试运行记录和竣工照片。")
para(doc, "验收按专业分项进行。风机系统验收包括设备安装、风管接口、阀门、盲板、操作箱、单机试车和系统联调；电气仪控验收包括桥架、线管、电缆、光缆、接地、绝缘和信号调试；土建防腐验收包括集水坑、围堰、耐酸砖、防腐层、排水沟、墙面和环境恢复。")

heading(doc, "十二、应急处置措施")
para(doc, "发生人员坠落、物体打击、触电、火情、吊装异常、玻璃钢粉尘刺激、池边落水、有限空间异常或风机试运行异常时，现场立即停止作业，切断相关能源，设置警戒并报告项目负责人和建设单位。")
para(doc, "触电人员应先脱离电源后施救；高处坠落和吊装伤害人员不得随意搬动，等待专业救援；火情初期使用灭火器扑救，无法控制时立即撤离并按厂区应急流程报警。试运行中发现异常振动、异响、过热、风管泄漏或电气报警，应立即停机检查，未经确认不得继续运行。")
para(doc, "项目部配备灭火器、急救箱、警戒带、应急照明、通讯工具、防尘口罩、护目镜、绝缘手套和应急排水设备。每日班前会明确当日高风险作业、监护人和应急联系人。")

heading(doc, "十三、资源配置计划")
para(doc, "施工资源按“主线保障、专业穿插、峰值受控”的原则配置。风机和DN1200玻璃钢风管施工期间，设备安装工、起重工、电工和安全监护人员应优先保障；土建防腐和环境整治期间，泥工、防腐工、普工和成品保护人员应分区组织，避免与风机吊装、电缆敷设交叉干扰。")
table(doc, ["资源类别", "配置内容", "使用阶段", "管理要求"], [
    ["管理人员", "项目经理、技术负责人、施工员、安全员、质量员、资料员", "全过程", "每日班前交底，关键作业旁站，资料随工程同步形成。"],
    ["设备安装", "钳工、起重工、焊工、普工", "第8-62天", "负责风机拆装、阀门、盲板、风管接口、支架和吊装配合。"],
    ["电气仪控", "电工、仪表工、光纤熔接人员、调试人员", "第35-90天", "负责桥架、线管、电缆、光缆、操作箱、消防点位和系统调试。"],
    ["土建防腐", "泥工、防腐工、油漆工、普工", "第45-95天", "负责集水坑、围堰、耐酸砖、环氧防腐、墙面及排水沟修复。"],
    ["主要机械", "25T汽车吊、15T平板车、16m及20m登高车、切割打磨机、绝缘测试仪、光纤熔接机", "按阶段投入", "进场前检查合格，特种设备和特种作业人员证件齐全。"],
    ["主要材料", "玻璃钢风管、阀门、风机、桥架、电缆、光缆、环氧树脂、耐酸砖、保温岩棉、钢格栅", "随进度分批进场", "材料进场验收，分区堆放，防潮、防撞、防污染。"],
])

heading(doc, "十四、关键验收控制点")
para(doc, "本工程涉及生产系统恢复和环保治理，验收应抓住拆除确认、设备安装、风管密封、电气绝缘、通信调试、防腐隐蔽和环境恢复七类控制点。每一控制点形成书面记录或影像资料，未经确认不得进入下道工序。")
table(doc, ["控制点", "检查内容", "验收资料"], [
    ["拆除前确认", "停电挂牌、系统隔离、拆除范围、保留设施保护、吊装通道", "安全交底、停电确认、现场签认照片"],
    ["风机安装", "基础、标高、水平度、地脚螺栓、进出口方向、试运行状态", "设备安装记录、单机试车记录"],
    ["DN1200风管", "法兰对口、支架间距、胶衣连接、盲板封堵、阀门启闭", "风管安装记录、接口检查记录"],
    ["电气仪控", "桥架接地、线管敷设、电缆标识、绝缘测试、端子接线", "绝缘测试记录、隐蔽验收记录、调试记录"],
    ["光缆消防", "光缆余量、熔接质量、收发器、ODF、点位图形化、报警信号", "光纤测试记录、消防点位调试记录"],
    ["土建防腐", "混凝土强度、防腐层遍数、耐酸砖平整度、勾缝、空鼓、坡向", "隐蔽验收、防腐施工记录、分项验收记录"],
    ["现场恢复", "垃圾清运、围挡撤除、道路排水、墙面观感、成品保护", "竣工照片、整改闭环记录"],
])

heading(doc, "十五、危险源辨识与控制")
para(doc, "施工前项目部应结合每日作业内容动态辨识危险源。涉及高处、吊装、动火、临电、池边、切割打磨、光纤施工和风机试运行时，必须执行专项交底和旁站监护。")
table(doc, ["危险源", "可能后果", "控制措施"], [
    ["高处作业", "人员坠落、工具坠落", "使用合格登高车或平台，安全带高挂低用，下方设置警戒区。"],
    ["风机和风管吊装", "吊物坠落、碰撞设备", "编制吊装交底，核实吊点和站位，专人指挥，试吊确认后起吊。"],
    ["临时用电和电缆施工", "触电、短路", "停电验电挂牌，一机一闸一漏一箱，绝缘测试合格后送电。"],
    ["动火切割打磨", "火灾、烫伤、粉尘刺激", "办理动火票，清理可燃物，配备灭火器，佩戴护目镜和防尘口罩。"],
    ["池边和基坑作业", "坠落、滑倒、积水伤害", "设置临边防护、上下通道、排水措施和监护人。"],
    ["风机试运行", "机械伤害、异常振动、风管泄漏", "试运行前清场，确认防护罩和紧固件，异常立即停机检查。"],
])

doc.save(OUT)


def make_schedule_xlsx():
    stages = [
        (1, "进场准备及系统交接", 1, 7, "现场踏勘、作业票、停电挂牌、材料计划、围挡、安全走绳、吊装路线确认。"),
        (2, "原有设备及管线拆除", 8, 18, "原风机、风管、阀门、操作箱、电缆拆除，垃圾外运和现场清理。"),
        (3, "备用风机搬迁及HLA池功能恢复", 15, 32, "备用风机拆装移位、盲板封堵、DN1200风管、阀门、操作箱、电缆敷设。"),
        (4, "HLA池和好氧池风机还原", 28, 45, "备用风机还原、新订风机安装、风管拆装、阀门和操作箱安装。"),
        (5, "好氧池旁新风机系统安装", 40, 62, "DN1200风管、弯头、三通、桥架、支架、投光灯、线管、电线、除锈刷漆。"),
        (6, "电气仪控、光缆及消防图形化", 55, 78, "桥架、穿线管、动力和控制电缆、光缆、光纤设备、计算机显示器、软件授权和信号调试。"),
        (7, "集水坑、围堰防腐及管道泵", 50, 82, "破除、开挖、混凝土集水坑、环氧树脂、耐酸砖、钢格栅、泵和控制箱安装。"),
        (8, "环境整治及土建恢复", 70, 92, "排水沟、围墙、雨水沟、污泥脱水间墙面、池壁涂料、保温保护层。"),
        (9, "系统调试、整改及验收", 90, 100, "风机单机试车、系统联调、消防点位调试、资料整理、缺陷整改、竣工验收。"),
    ]
    wb = Workbook()
    ws = wb.active
    ws.title = "100天进度计划"
    ws.merge_cells("A1:N1")
    ws["A1"] = "焦一作业区废气收集系统整改及环境整治100天施工进度计划"
    ws["A1"].font = Font(name="宋体", size=16, bold=True)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    headers = ["序号", "施工阶段", "开始天", "结束天", "主要工作内容"] + [f"第{i*10-9}-{i*10}天" for i in range(1, 11)]
    ws.append(headers)
    thin = Side(style="thin", color="999999")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    fill_header = PatternFill("solid", fgColor="D9EAF7")
    fill_bar = PatternFill("solid", fgColor="70AD47")
    for cell_obj in ws[2]:
        cell_obj.font = Font(name="宋体", bold=True)
        cell_obj.fill = fill_header
        cell_obj.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell_obj.border = border
    for seq, name, start, end, desc in stages:
        row = [seq, name, start, end, desc]
        for block in range(1, 11):
            b_start = block * 10 - 9
            b_end = block * 10
            row.append("■" if start <= b_end and end >= b_start else "")
        ws.append(row)
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for c in row:
            c.font = Font(name="宋体", size=10)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border
        row[1].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        row[4].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        for c in row[5:]:
            if c.value:
                c.fill = fill_bar
                c.font = Font(name="宋体", size=10, color="FFFFFF", bold=True)
    widths = [8, 28, 9, 9, 48] + [12] * 10
    for idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.row_dimensions[1].height = 30
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 36
    ws.freeze_panes = "F3"
    wb.save(OUT_XLSX)


make_schedule_xlsx()
print(OUT)
print(OUT_XLSX)

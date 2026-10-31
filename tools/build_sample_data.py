try:
    import openpyxl
except ImportError:
    import subprocess, sys
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


XL_PRIMARY = xl_color("#4472C4")
XL_LIGHT = xl_color("#D9E2F3")
XL_DARK = xl_color("#2F5597")
XL_BORDER = xl_color("#BFBFBF")
XL_WHITE = xl_color("#FFFFFF")
XL_WARN = xl_color("#FFF2CC")
XL_GREEN = xl_color("#C6EFCE")
XL_GREENTXT = xl_color("#006100")
XL_INPUT = xl_color("#E2EFDA")

thin = Side(style="thin", color=XL_BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
FONT = "微软雅黑"


def rmb_upper(amount: float) -> str:
    """人民币金额转大写。"""
    digits = "零壹贰叁肆伍陆柒捌玖"
    units = ["", "拾", "佰", "仟"]
    big = ["", "万", "亿", "万亿"]
    n = int(round(amount * 100))
    jiao, fen = (n // 10) % 10, n % 10
    yuan = n // 100

    def four(p: int) -> str:
        s = ""
        for i in range(3, -1, -1):
            d = (p // (10 ** i)) % 10
            if d:
                s += digits[d] + units[i]
            elif s and (p % (10 ** i)) != 0:
                s += "零"
        return s

    if yuan == 0:
        s = ""
    else:
        parts, t = [], yuan
        while t > 0:
            parts.append(t % 10000)
            t //= 10000
        s = ""
        for idx in range(len(parts) - 1, -1, -1):
            p = parts[idx]
            if p == 0:
                continue
            if s and p < 1000:
                s += "零"
            s += four(p) + big[idx]
        s += "元"

    if jiao == 0 and fen == 0:
        return s + "整"
    out = s
    if jiao:
        out += digits[jiao] + "角"
    elif fen:
        out += "零"
    if fen:
        out += digits[fen] + "分"
    else:
        out += "整"
    return out


def title_bar(ws, text, ncol, height=30):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
    c = ws.cell(row=1, column=1, value=text)
    c.font = Font(name=FONT, size=13, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=XL_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = height


def header_row(ws, names, widths, row=2, height=32):
    for i, (name, width) in enumerate(zip(names, widths), start=1):
        cell = ws.cell(row=row, column=i, value=name)
        cell.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
        cell.fill = PatternFill("solid", fgColor=XL_PRIMARY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[row].height = height


wb = Workbook()

# ================= Sheet 1 主表 =================
ws = wb.active
ws.title = "主表"

SUP = "供应商信息"
SUP_RANGE = f"{SUP}!$B$3:$M$11"

MAIN = [
    ("中标单位", 16, "input", None),
    ("项目名称", 40, "fixed", None),
    ("甲方名称", 26, "fixed", None),
    ("甲方地址", 30, "fixed", None),
    ("甲方联系人", 12, "fixed", None),
    ("甲方电话", 15, "fixed", None),
    ("乙方名称", 28, "lookup", 2),
    ("乙方地址", 34, "lookup", 7),
    ("乙方联系人", 11, "lookup", 8),
    ("乙方电话", 15, "lookup", 9),
    ("乙方法定代表人", 14, "lookup", 4),
    ("统一社会信用代码", 21, "lookup", 3),
    ("金额小写（元）", 14, "lookup", 10),
    ("金额大写", 26, "lookup", 11),
]

FIXED = {
    "项目名称": "长沙市天心区青园小学百川楼打卡文化墙采购项目",
    "甲方名称": "长沙市天心区青园小学",
    "甲方地址": "长沙市天心区万芙北路189号",
    "甲方联系人": "刘主任",
    "甲方电话": "0731-85894211",
}

title_bar(ws, "示例数据 · 长沙市天心区青园小学百川楼打卡文化墙采购项目（A3 选中标单位，其余字段自动联动）", len(MAIN))
header_row(ws, [m[0] for m in MAIN], [m[1] for m in MAIN])

for i, (name, width, kind, idx) in enumerate(MAIN, start=1):
    if kind == "input":
        cell = ws.cell(row=3, column=i, value="晨早")
    elif kind == "fixed":
        cell = ws.cell(row=3, column=i, value=FIXED[name])
    else:
        cell = ws.cell(row=3, column=i,
                       value=f'=IFERROR(VLOOKUP($A$3,{SUP_RANGE},{idx},FALSE),"")')
    cell.font = Font(name=FONT, size=10, bold=(kind == "input"))
    cell.border = BOX
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    if name in ("甲方电话", "乙方电话", "统一社会信用代码"):
        cell.number_format = "@"
    if name == "金额小写（元）":
        cell.number_format = "#,##0.00"
        cell.alignment = Alignment(horizontal="right", vertical="center")
    if kind == "input":
        cell.fill = PatternFill("solid", fgColor=XL_INPUT)
    elif kind == "lookup":
        cell.fill = PatternFill("solid", fgColor=XL_LIGHT)

ws.row_dimensions[3].height = 46

dv = DataValidation(type="list", formula1=f"'{SUP}'!$B$3:$B$11",
                    allow_blank=False, showDropDown=False)
dv.error = "请从下拉列表中选择中标单位"
dv.errorTitle = "取值无效"
ws.add_data_validation(dv)
dv.add(ws["A3"])

ws.freeze_panes = "B3"

# ================= Sheet 2 明细 =================
ws2 = wb.create_sheet("明细")
DH = [("序号", 6, "center"), ("名称", 16, "left"),
      ("型号规格、标准及配置", 48, "left"), ("数量", 8, "right"),
      ("单位", 7, "center"), ("单价（元）", 12, "right"), ("金额（元）", 13, "right")]
ROWS = [
    (1, "人物造型", "20mm厚结皮pvc板侧面刮灰烤漆，正uv打印，立体字", 1, "项", 2400.00),
    (2, "底板造型材料", "50*50*3mm厚方管焊接龙骨架，侧面和正面安装1.5mm厚热镀锌板", 75, "m2", 210.00),
    (3, "光纤激光切割", "1.5mm厚热镀锌板激光切割", 75, "m2", 14.00),
    (4, "焊接打磨", "按设计要求焊接50mm边，精细打磨，正面封板", 75, "m2", 110.00),
    (5, "辅材", "膨胀螺丝焊条等", 1, "项", 800.00),
    (6, "油漆喷涂", "刮汽车原子灰，正反两面喷环氧底漆2遍", 150, "m2", 78.00),
    (7, "运输", "6.8米货车", 1, "m2", 800.00),
    (8, "人员安装", "专业人员安装", 13, "人", 300.00),
]
title_bar(ws2, "示例数据 · 明细表（中标方「长沙晨早文化传播有限公司」报价，8 行物料）", len(DH))
header_row(ws2, [d[0] for d in DH], [d[1] for d in DH])

for idx, (no, name, spec, qty, unit, price) in enumerate(ROWS):
    r = 3 + idx
    for i, v in enumerate([no, name, spec, qty, unit, price], start=1):
        cell = ws2.cell(row=r, column=i, value=v)
        cell.font = Font(name=FONT, size=10)
        cell.border = BOX
        cell.alignment = Alignment(horizontal=DH[i - 1][2], vertical="center", wrap_text=(i == 3))
        if i == 4:
            cell.number_format = "0"
        if i == 6:
            cell.number_format = "#,##0.00"
    cell = ws2.cell(row=r, column=7, value=f'=IF(OR(D{r}="",F{r}=""),"",D{r}*F{r})')
    cell.font = Font(name=FONT, size=10)
    cell.border = BOX
    cell.number_format = "#,##0.00"
    cell.alignment = Alignment(horizontal="right", vertical="center")
    ws2.row_dimensions[r].height = 34

last = 3 + len(ROWS) - 1
tot = last + 1
ws2.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=5)
for i in range(1, 8):
    cc = ws2.cell(row=tot, column=i)
    cc.fill = PatternFill("solid", fgColor=XL_DARK)
    cc.border = BOX
c = ws2.cell(row=tot, column=1, value="合  计")
c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
c.alignment = Alignment(horizontal="center", vertical="center")
c = ws2.cell(row=tot, column=7, value=f"=SUM(G3:G{last})")
c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
c.number_format = "#,##0.00"
c.alignment = Alignment(horizontal="right", vertical="center")
ws2.row_dimensions[tot].height = 26
ws2.freeze_panes = "A3"

# ================= Sheet 3 供应商信息（9 家，含报价大写） =================
ws3 = wb.create_sheet(SUP)

COMPANIES = [
    ("晨早", "长沙晨早文化传播有限公司", "91430103MA4QXDUY22", "蒋美益", "100 万元", "2019-11-06",
     "湖南省长沙市天心区南二环282号3栋2116", "罗艳军", "13077327233", 44650.00, "真实报价"),
    ("泉旭", "长沙泉旭广告传媒有限公司", "91430121MACQMH1W7E", "张何", "20 万元", "2023-08-08",
     "长沙市雨花区高桥街道人民东路58号新三诚大厦1101-1102房F040号", "张何", "15111062831", 45875.00, "真实报价"),
    ("左右", "长沙左右文化传播有限公司", "91430103794720885W", "刘念", "200 万元", "2006-11-06",
     "湖南省长沙市天心区书院南路723号卓越江岸苑3栋1112号房", "刘念", "13707319090", 46713.00, "模拟报价"),
    ("五九", "湖南五九广告装饰工程有限公司", "91430104338468778X", "龙彬", "200 万元", "2015-04-23",
     "湖南省长沙市岳麓区咸嘉湖街道桐梓坡路519号高鑫麓城14栋308房", "龙彬", "15126584671", 47449.00, "模拟报价"),
    ("东方恒大", "湖南东方恒大广告传媒有限公司", "91430111MA4L425108", "张小海", "300 万元", "2016-04-27",
     "长沙市雨花区井莲路397号紫铭大厦1337号房", "张小海", "0731-85069118", 47624.00, "模拟报价"),
    ("大班威", "湖南大班威广告传媒有限公司", "91430111394064620T", "孙威", "200 万元", "2014-09-22",
     "长沙经济技术开发区盼盼路12号湖南湘联科技有限公司厂房401（4-001—4-019）", "孙威", "13507409923", 47799.00, "模拟报价"),
    ("天高", "长沙市天高文化传播有限公司", "9143011177228738X1", "李智", "100 万元", "2005-03-28",
     "湖南省长沙市天心区书院路2号森宇房地产1栋B座907房", "李智", "13307316919", 47859.00, "模拟报价"),
    ("上策", "长沙上策文化传播有限公司", "914301056639728634", "邓谋民", "101 万元", "2007-07-17",
     "湖南省长沙市开福区四方坪麦德龙北侧四季美景小区15栋103房", "邓谋民", "13549666136", 47925.00, "真实报价"),
    ("泽信", "长沙泽信文化传播有限公司", "91430105083553988Q", "晏乐杰", "100 万元", "2013-11-14",
     "湖南省长沙市天心区五一大道717号1801房", "晏乐杰", "15111261647", 48888.00, "模拟报价"),
]

SI = ["报价排名", "公司简称", "公司全称", "统一社会信用代码", "法定代表人", "注册资本",
      "成立日期", "注册地址", "报价联系人", "联系电话", "报价总额（元）", "报价大写", "数据性质"]
SI_W = [10, 12, 30, 22, 11, 11, 12, 42, 12, 15, 14, 28, 12]

title_bar(ws3, "九家供应商信息对照（按报价由低到高；主表 A3 的下拉即取自 B 列）", len(SI))
header_row(ws3, SI, SI_W)

for idx, comp in enumerate(COMPANIES):
    r = 3 + idx
    vals = [idx + 1] + list(comp[:10]) + [rmb_upper(comp[9]), comp[10]]
    for i, v in enumerate(vals, start=1):
        cell = ws3.cell(row=r, column=i, value=v)
        cell.font = Font(name=FONT, size=10, bold=(i == 2))
        cell.border = BOX
        cell.alignment = Alignment(
            horizontal="center" if i in (1, 5, 6, 7) else "left",
            vertical="center", wrap_text=(i in (3, 8, 12)))
        if i in (4, 10):
            cell.number_format = "@"
        if i == 11:
            cell.number_format = "#,##0.00"
            cell.alignment = Alignment(horizontal="right", vertical="center")
        if i == 2:
            cell.fill = PatternFill("solid", fgColor=XL_INPUT)
        if i == 13:
            cell.fill = PatternFill("solid", fgColor=XL_WARN if v == "模拟报价" else XL_GREEN)
    ws3.row_dimensions[r].height = 30 if len(comp[6]) < 30 else 42
ws3.freeze_panes = "C3"

# ================= Sheet 4 报价对比（9 家） =================
ws4 = wb.create_sheet("报价对比")

NAMES9 = [c[0] for c in COMPANIES]
QUOTE = [
    (1, "人物造型", 1, "项", 2400, 2450, 2440, 2583, 2481, 2451, 2459, 2500, 2585),
    (2, "底板造型材料", 75, "m2", 210, 215, 217, 218, 225, 232, 222, 220, 235),
    (3, "光纤激光切割", 75, "m2", 14, 14, 14, 16, 15, 14, 15, 15, 15),
    (4, "焊接打磨", 75, "m2", 110, 118, 119, 121, 119, 117, 125, 120, 118),
    (5, "辅材", 1, "项", 800, 700, 790, 756, 764, 734, 737, 800, 789),
    (6, "油漆喷涂", 150, "m2", 78, 76, 79, 84, 83, 79, 81, 80, 84),
    (7, "运输", 1, "m2", 800, 800, 846, 842, 896, 859, 891, 800, 881),
    (8, "人员安装", 13, "人", 300, 350, 349, 311, 316, 360, 344, 400, 341),
]
QH = ["序号", "项目名称", "数量", "单位"] + NAMES9 + ["最低价方"]
QW = [5, 18, 6, 6] + [9] * 4 + [10, 10] + [9] * 3 + [10]

title_bar(ws4, "九家报价单价对照（单价单位：元；标黄三家为模拟数据）", len(QH))
header_row(ws4, QH, QW)

first_price_col = 5
last_price_col = 4 + len(NAMES9)

for idx, row in enumerate(QUOTE):
    r = 3 + idx
    for i, v in enumerate(row, start=1):
        cell = ws4.cell(row=r, column=i, value=v)
        cell.font = Font(name=FONT, size=10)
        cell.border = BOX
        cell.alignment = Alignment(
            horizontal="center" if i in (1, 4) else ("right" if i >= first_price_col else "left"),
            vertical="center")
        if i >= first_price_col:
            cell.number_format = "#,##0.00"
    prices = {first_price_col + k: row[4 + k] for k in range(len(NAMES9))}
    low = min(prices.values())
    for col, v in prices.items():
        if v == low:
            cc = ws4.cell(row=r, column=col)
            cc.fill = PatternFill("solid", fgColor=XL_GREEN)
            cc.font = Font(name=FONT, size=10, bold=True, color=XL_GREENTXT)
    lc = get_column_letter(first_price_col)
    rc = get_column_letter(last_price_col)
    cell = ws4.cell(row=r, column=last_price_col + 1,
                    value=f'=INDEX(${lc}$2:${rc}$2,MATCH(MIN({lc}{r}:{rc}{r}),{lc}{r}:{rc}{r},0))')
    cell.font = Font(name=FONT, size=10)
    cell.border = BOX
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[r].height = 24

qlast = 3 + len(QUOTE) - 1
qtot = qlast + 1
ws4.merge_cells(start_row=qtot, start_column=1, end_row=qtot, end_column=4)
for i in range(1, len(QH) + 1):
    cc = ws4.cell(row=qtot, column=i)
    cc.fill = PatternFill("solid", fgColor=XL_DARK)
    cc.border = BOX
c = ws4.cell(row=qtot, column=1, value="合  计（数量 × 单价）")
c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
c.alignment = Alignment(horizontal="center", vertical="center")
for col in range(first_price_col, last_price_col + 1):
    letter = get_column_letter(col)
    c = ws4.cell(row=qtot, column=col,
                 value=f"=SUMPRODUCT($C$3:$C${qlast},{letter}3:{letter}{qlast})")
    c.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
    c.number_format = "#,##0.00"
    c.alignment = Alignment(horizontal="right", vertical="center")
ws4.row_dimensions[qtot].height = 26

note_r = qtot + 1
ws4.merge_cells(start_row=note_r, start_column=1, end_row=note_r, end_column=len(QH))
cell = ws4.cell(row=note_r, column=1,
                value="注 1：泉旭合计按「单价 × 数量」重算为 45,925 元，其报价清单原写 45,875 元——"
                      "序号 1 单价 2,450 × 1，总价却被填成 2,400，少 50 元。"
                      "注 2：左右、五九、东方恒大、大班威、天高、泽信六家的公司信息真实，"
                      "报价为随机模拟，仅用于模板填充测试，不得作为真实报价使用。")
cell.font = Font(name=FONT, size=9, color="FF9C0006")
cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
cell.fill = PatternFill("solid", fgColor=XL_WARN)
cell.border = BOX
ws4.row_dimensions[note_r].height = 40
ws4.freeze_panes = "E3"

# ================= Sheet 5 说明（字段映射，值列引用主表） =================
ws5 = wb.create_sheet("说明")
NOTE_H = ["模板红字占位", "主表列名", "供应商信息字段", "当前值（随中标单位联动）", "来源 / 备注"]
NOTE_W = [24, 20, 18, 40, 46]

NOTES = [
    ("某某甲方", "甲方名称", "—（固定值）", "C3", "签收单「采购人名称」；百度百科校址"),
    ("某某甲方地址", "甲方地址", "—（固定值）", "D3", "青园小学公开校址"),
    ("采购单位联系人", "甲方联系人", "—（固定值）", "E3", "长沙市政府采购网 · 青园小学采购公告「采购人信息」栏"),
    ("采购单位电话", "甲方电话", "—（固定值）", "F3", "同上，为采购公告登记电话"),
    ("某某甲方某某项目 / 某某文化墙", "项目名称", "—（固定值）", "B3", "三份报价清单表头"),
    ("某某公司", "乙方名称", "公司全称", "G3", "工商登记全称（无「市」字）"),
    ("某某公司地址", "乙方地址", "注册地址", "H3", "工商登记注册地址"),
    ("联系人", "乙方联系人", "报价联系人", "I3", "报价清单落款"),
    ("联系电话", "乙方电话", "联系电话", "J3", "报价清单落款 / 平台公开号码"),
    ("法定代表人", "乙方法定代表人", "法定代表人", "K3", "工商登记；与报价单联系人可能不是同一人"),
    ("—", "统一社会信用代码", "统一社会信用代码", "L3", "第三方平台，需在国家企业信用信息公示系统核对"),
    ("44650.00", "金额小写（元）", "报价总额（元）", "M3", "各供应商报价合计"),
    ("肆万肆仟陆佰伍拾元整", "金额大写", "报价大写", "N3", "按报价总额程序生成"),
]

title_bar(ws5, "字段映射表（当前值随主表 A3 选择的中标单位自动变化）", len(NOTE_H))
header_row(ws5, NOTE_H, NOTE_W, height=32)

for idx, row in enumerate(NOTES):
    r = 3 + idx
    for i, v in enumerate(row[:3], start=1):
        cell = ws5.cell(row=r, column=i, value=v)
        cell.font = Font(name=FONT, size=10, color="FF9C0006" if i == 1 else "FF000000")
        cell.border = BOX
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(i == 3))
    cell = ws5.cell(row=r, column=4, value=f"=主表!{row[3]}")
    cell.font = Font(name=FONT, size=10, bold=True, color=XL_DARK)
    cell.border = BOX
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    cell.fill = PatternFill("solid", fgColor=XL_LIGHT)
    if row[1] == "金额小写（元）":
        cell.number_format = "#,##0.00"
    if row[1] in ("甲方电话", "乙方电话", "统一社会信用代码"):
        cell.number_format = "@"
    cell = ws5.cell(row=r, column=5, value=row[4])
    cell.font = Font(name=FONT, size=10)
    cell.border = BOX
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws5.row_dimensions[r].height = 30 if len(str(row[4])) < 30 else 44

start = 3 + len(NOTES) + 1
ws5.merge_cells(start_row=start, start_column=1, end_row=start, end_column=5)
c = ws5.cell(row=start, column=1, value="使用说明与注意事项")
c.font = Font(name=FONT, size=11, bold=True, color=XL_DARK)
c.fill = PatternFill("solid", fgColor=XL_LIGHT)
c.alignment = Alignment(horizontal="left", vertical="center")
ws5.row_dimensions[start].height = 24

WARN = [
    "1. 【换中标方】主表 A3 是下拉框，从九家供应商里选任意一家，乙方名称 / 地址 / 联系人 / 电话 / 法定代表人 / 统一社会信用代码 / 金额小写 / 金额大写 八个字段全部自动联动，本表「当前值」列同步变化，不需要手动改任何一格。",
    "2. 【为什么不写死】模板填充脚本读的是主表第 3 行，改 A3 就等于换了整份合同的主体，晨早、泉旭、上策之外的六家同样可用。若选中的单位与「报价对比」里的最低价不一致，脚本不会阻止，需人工确认是否符合「最低价成交」原则。",
    "3. 【明细表不联动】明细表固定是晨早的 8 行报价，因为另外八家没有逐项报价单。换中标方时明细需另行提供，否则金额会与主表对不上。",
    "4. 九家里只有晨早、泉旭、上策三家真实参与本项目，金额来自其报价清单；其余六家（左右、五九、东方恒大、大班威、天高、泽信）公司信息真实可查，但报价为随机生成，纯属模拟，不得作为真实报价或竞争性磋商依据使用。",
    "5. 乙方名称以工商登记为准：报价清单写「长沙【市】晨早文化传播有限公司」，工商登记为「长沙晨早文化传播有限公司」，多一个「市」字（已修正）。签收单写的是后者。",
    "6. 泉旭报价清单算术有误：序号 1「人物造型」单价 2450、数量 1，总价却填 2400（少 50 元）。按其单价重算，泉旭合计应为 45925 而非报价单上的 45875；本表「报价总额」沿用报价单原值 45875。",
    "7. 甲方联系人与电话取自长沙市政府采购网青园小学采购公告「采购人信息」栏（刘主任 / 0731-85894211），比学校名录上的招生咨询电话更对口。九家供应商电话中，五九的为手工填写，其余为第三方平台或年报公示号码，使用前均需核实。",
    "8. 红字判定口径：红色系（R≥0x80 且 G≤0x60 且 B≤0x60），覆盖 EE0000 与 FF0000 两种模板标注色。",
]
for i, text in enumerate(WARN):
    r = start + 1 + i
    ws5.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    cell = ws5.cell(row=r, column=1, value=text)
    cell.font = Font(name=FONT, size=9)
    cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
    cell.fill = PatternFill("solid", fgColor=XL_WARN)
    cell.border = BOX
    ws5.row_dimensions[r].height = 44

wb.properties.title = "示例数据"
OUT = r"D:/git学习/Excel转指定位置word/项目文件/示例数据.tmp.xlsx"
wb.save(OUT)
print("saved:", OUT)

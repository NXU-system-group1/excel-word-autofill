# -*- coding: utf-8 -*-
"""为晨早 / 泉旭 / 左右三家各生成一份红字信息表（仅乙方字段，不含甲方）。"""
import sys
from pathlib import Path

try:
    import openpyxl
except ImportError:
    import subprocess
    subprocess.check_call([sys.executable, "-m", "pip", "install", "--quiet", "openpyxl>=3.1.0"])
    import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")


def xl_color(css_hex: str) -> str:
    value = css_hex.removeprefix("#").upper()
    if len(value) != 6:
        raise ValueError(f"Expected #RRGGBB, got: {css_hex}")
    return "FF" + value


XL_PRIMARY = xl_color("#4472C4")
XL_DARK = xl_color("#2F5597")
XL_BORDER = xl_color("#BFBFBF")
XL_WHITE = xl_color("#FFFFFF")
XL_WARN = xl_color("#FFF2CC")
XL_KEY = xl_color("#FCEBEB")

thin = Side(style="thin", color=XL_BORDER)
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
FONT = "微软雅黑"

PROJECT = "长沙市天心区青园小学百川楼打卡文化墙采购项目"

COMPANIES = [
    {
        "short": "晨早",
        "full": "长沙晨早文化传播有限公司",
        "rows": [
            ("某某甲方某某项目 / 某某文化墙", "项目名称", PROJECT, "三份报价清单表头"),
            ("某某公司", "乙方名称", "长沙晨早文化传播有限公司", "工商登记全称（无「市」字）"),
            ("某某A公司", "供应商名称（A）", "长沙晨早文化传播有限公司", "询价汇总表序号1 / 报价最低，中标"),
            ("某某公司地址", "乙方地址", "湖南省长沙市天心区南二环282号3栋2116", "工商登记注册地址"),
            ("联系人", "乙方联系人", "罗艳军", "报价清单落款"),
            ("联系电话", "乙方电话", "13077327233", "报价清单落款"),
            ("法定代表人", "乙方法定代表人", "蒋美益", "工商登记；与联系人不是同一人"),
            ("—", "统一社会信用代码", "91430103MA4QXDUY22", "第三方平台，需在国家企业信用信息公示系统核对"),
            ("44650.00", "金额小写（元）", 44650.00, "报价清单合计，与结算清单一致"),
            ("肆万肆仟陆佰伍拾元整", "金额大写", "肆万肆仟陆佰伍拾元整", "由金额小写生成"),
        ],
    },
    {
        "short": "泉旭",
        "full": "长沙泉旭广告传媒有限公司",
        "rows": [
            ("某某甲方某某项目 / 某某文化墙", "项目名称", PROJECT, "三份报价清单表头"),
            ("某某公司", "乙方名称", "长沙泉旭广告传媒有限公司", "工商登记全称"),
            ("某某B公司", "供应商名称（B）", "长沙泉旭广告传媒有限公司", "询价汇总表序号2"),
            ("某某公司地址", "乙方地址", "长沙市雨花区高桥街道人民东路58号新三诚大厦1101-1102房F040号", "工商登记注册地址"),
            ("联系人", "乙方联系人", "张何", "报价清单落款"),
            ("联系电话", "乙方电话", "15111062831", "报价清单落款"),
            ("法定代表人", "乙方法定代表人", "张何", "工商登记；与联系人为同一人"),
            ("—", "统一社会信用代码", "91430121MACQMH1W7E", "第三方平台，需在国家企业信用信息公示系统核对"),
            ("45875.00", "金额小写（元）", 45875.00, "报价清单原写金额；按单价重算为 45925，差 50 元（序号1单价2450×1被填成2400）"),
            ("肆万伍仟捌佰柒拾伍元整", "金额大写", "肆万伍仟捌佰柒拾伍元整", "按报价清单原写金额生成"),
        ],
    },
    {
        "short": "左右",
        "full": "长沙左右文化传播有限公司",
        "rows": [
            ("某某甲方某某项目 / 某某文化墙", "项目名称", PROJECT, "三份报价清单表头"),
            ("某某公司", "乙方名称", "长沙左右文化传播有限公司", "工商登记全称"),
            ("某某C公司", "供应商名称（C）", "长沙左右文化传播有限公司", "询价汇总表序号3"),
            ("某某公司地址", "乙方地址", "湖南省长沙市天心区书院南路723号卓越江岸苑3栋1112号房", "工商登记注册地址"),
            ("联系人", "乙方联系人", "刘念", "法定代表人兼任报价联系人"),
            ("联系电话", "乙方电话", "13707319090", "企查查公开号码，使用前需核实"),
            ("法定代表人", "乙方法定代表人", "刘念", "工商登记"),
            ("—", "统一社会信用代码", "91430103794720885W", "第三方平台，需在国家企业信用信息公示系统核对"),
            ("47925.00", "金额小写（元）", 47925.00, "报价清单合计（三家最高）"),
            ("肆万柒仟玖佰贰拾伍元整", "金额大写", "肆万柒仟玖佰贰拾伍元整", "由金额小写生成"),
        ],
    },
]

OWNER = {
    "short": "甲方",
    "full": "长沙市天心区青园小学",
    "tag": "仅甲方字段",
    "rows": [
        ("某某甲方", "甲方名称", "长沙市天心区青园小学", "政府采购公告 / 签收单「采购人名称」"),
        ("采购人名称", "甲方名称（签收单占位）", "长沙市天心区青园小学", "签收单 R2 的红字占位"),
        ("长沙市天心区青园小学", "甲方名称（模板已填真名）", "长沙市天心区青园小学",
         "部分模板直接写了真实校名，值不变，仅改黑"),
        ("某某甲方某某项目", "项目名称", PROJECT, "三份报价清单表头"),
        ("某某甲方项目名字", "项目名称（合同写法）", PROJECT, "合同正文 / 项目概况"),
        ("某某文化墙", "项目简称", "百川楼打卡文化墙", "用于合同标题「某某文化墙采购服务合同」"),
        ("某某甲方地址", "甲方地址", "空", "原青园小学地址只适用于旧项目；换项目请在界面里填，留空则输出「空」"),
        ("某某甲方内", "安装地点", "长沙市天心区青园小学校内", "合同「3、安装地点」"),
    ],
}

HEADERS = ["模板红字占位", "字段", "值"]
WIDTHS = [30, 18, 72]

OUT_DIR = Path(r"D:/git学习/Excel转指定位置word/项目文件/红字信息")
OUT_DIR.mkdir(parents=True, exist_ok=True)

def build_one(comp, skip_project=True):
    wb = Workbook()
    ws = wb.active
    ws.title = "红字信息"

    tag = comp.get("tag", "仅乙方字段，不含甲方")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3)
    c = ws.cell(row=1, column=1, value=f"红字信息 · {comp['full']}（{tag}）")
    c.font = Font(name=FONT, size=13, bold=True, color=XL_WHITE)
    c.fill = PatternFill("solid", fgColor=XL_DARK)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    for i, (name, width) in enumerate(zip(HEADERS, WIDTHS), start=1):
        cell = ws.cell(row=2, column=i, value=name)
        cell.font = Font(name=FONT, size=10, bold=True, color=XL_WHITE)
        cell.fill = PatternFill("solid", fgColor=XL_PRIMARY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[2].height = 28

    # 公司表不输出「项目名称」行（那是甲方信息）；甲方表全输出
    rows = [r for r in comp["rows"] if not (skip_project and r[1] == "项目名称")]
    for idx, (ph, field, val, note) in enumerate(rows):
        r = 3 + idx
        for i, v in enumerate([ph, field, val], start=1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.font = Font(
                name=FONT, size=10,
                bold=(i == 2),
                color="FF9C0006" if i == 1 else "FF000000",
            )
            cell.border = BOX
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(i == 3))
            if i == 1:
                cell.fill = PatternFill("solid", fgColor=XL_KEY)
            if field == "金额小写（元）" and i == 3:
                cell.number_format = "#,##0.00"
            if field in ("乙方电话", "统一社会信用代码") and i == 3:
                cell.number_format = "@"
        ws.row_dimensions[r].height = 30 if len(str(val)) < 34 else 44

    ws.freeze_panes = "A3"

    out = OUT_DIR / f"{comp['short']}.xlsx"
    wb.properties.title = f"红字信息_{comp['short']}"
    try:
        wb.save(out)
        print("saved:", out)
    except PermissionError:
        tmp = OUT_DIR / f"{comp['short']}.pending.xlsx"
        wb.save(tmp)
        print(f"LOCKED 跳过 {out.name}，已存为 {tmp.name}")


for comp in COMPANIES:
    build_one(comp, skip_project=True)

build_one(OWNER, skip_project=False)

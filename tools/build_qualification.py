# -*- coding: utf-8 -*-
"""为泉旭 / 晨早 / 左右三家生成《资质信息核查表》docx，放进 2.资质 对应文件夹。"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

if sys.stdout is not None:
    sys.stdout.reconfigure(encoding="utf-8")

BASE = Path(r"D:/git学习/Excel转指定位置word/项目文件/1、工程A三方询价 价格/2.资质")
CHECK_DATE = "2026年10月29日"

DATA = [
    {
        "folder": "1、晨早",
        "name": "长沙晨早文化传播有限公司",
        "code": "91430103MA4QXDUY22",
        "legal": "蒋美益",
        "capital": "100 万元",
        "founded": "2019-11-06",
        "address": "湖南省长沙市天心区南二环282号3栋2116",
        "scope": "广告设计；广告制作服务；企业形象策划服务；展览服务；承办展览展示；"
                 "文化活动服务；风景园林工程、室内装饰设计服务；园林绿化工程施工。",
        "fit": "较高。经营范围同时覆盖「广告制作服务」与「室内装饰设计服务」「园林绿化工程施工」，"
               "与本项目文化墙的设计、制作、现场安装环节基本对得上，是三家里面业务口径最贴近的一家。",
        "licenses": "公开渠道未检索到行政许可记录。",
        "marks": "未检索到商标信息。",
        "bids": "未检索到公开招投标记录。",
        "risk": "未检索到经营异常、行政处罚、严重违法失信记录。",
        "note": "注册资本 100 万元，认缴制，实缴情况未公示。成立 2019 年，经营年限约 6 年。",
    },
    {
        "folder": "2、泉旭",
        "name": "长沙泉旭广告传媒有限公司",
        "code": "91430121MACQMH1W7E",
        "legal": "张何",
        "capital": "20 万元",
        "founded": "2023-08-08",
        "address": "长沙市雨花区高桥街道人民东路58号新三诚大厦1101-1102房F040号",
        "scope": "组织文化艺术交流活动；广告制作；数字广告制作；广告发布；数字广告发布；"
                 "广告设计、代理；数字广告设计、代理；显示器件销售；专业设计服务；图文设计制作；"
                 "市政设施管理；平面设计；交通及公共管理用标牌销售。",
        "fit": "中等。有「广告制作」「图文设计制作」「交通及公共管理用标牌销售」「市政设施管理」，"
               "标牌类业务与本项目相关；但经营范围中没有装饰装修、钢结构焊接一类的施工类目。",
        "licenses": "第三方平台显示行政许可 0 条。",
        "marks": "第三方平台显示商标 0 条。",
        "bids": "第三方平台显示招投标记录 0 条。",
        "risk": "未检索到经营异常、行政处罚、严重违法失信记录。",
        "note": "注册资本仅 20 万元，2023 年 8 月成立，是三家里面规模最小、成立最晚的一家。"
                "本项目合同金额 45,875 元，约为其注册资本的 2.3 倍，履约能力需重点核实。",
    },
    {
        "folder": "3、左右",
        "name": "长沙左右文化传播有限公司",
        "code": "91430103794720885W",
        "legal": "刘念",
        "capital": "200 万元（实缴 51 万元）",
        "founded": "2006-11-06",
        "address": "湖南省长沙市天心区书院南路723号卓越江岸苑3栋1112号房",
        "scope": "一般项目：组织文化艺术交流活动；工艺美术品及礼仪用品销售（象牙及其制品除外）；"
                 "广告设计、代理；广告发布；广告制作；会议及展览服务；企业管理咨询；初级农产品收购；"
                 "地产中草药（不含中药饮片）购销；食用农产品批发；食用农产品零售；中草药收购；"
                 "健康咨询服务（不含诊疗服务）；新鲜蔬菜零售；新鲜蔬菜批发；化妆品零售；日用百货销售；"
                 "鲜肉零售；鲜肉批发。许可项目：食品销售。",
        "fit": "中等。有「广告设计、代理、发布、制作」，设计制作环节覆盖；"
               "但经营范围里没有装饰装修、焊接、钢结构等施工类目，现场安装部分无对应资质。",
        "licenses": "公开信息显示有 4 项行政许可（具体事项未公开明细）。",
        "marks": "2 个注册商标：秀客 SHOWKER、心流密钥。",
        "bids": "参与过招投标项目（次数未公开）。服务客户包括北辰三角洲、万科系列项目、比亚迪、"
                "东风日产、浏阳河酒等。",
        "risk": "存在历史经营异常记录（高风险提示），具体情形需在国家企业信用信息公示系统核实。",
        "note": "成立 2006 年，是三家里经营年限最长的一家；注册资本 200 万元，实缴 51 万元。"
                "曾多次迁址，2024 年 1 月迁至现址。",
    },
]


def set_font(run, size=12, bold=False, name="宋体"):
    run.font.name = name
    run.font.size = Pt(size)
    run.bold = bold
    run._element.get_or_add_rPr().get_or_add_rFonts().set(qn("w:eastAsia"), name)


def add_para(doc, text, size=12, bold=False, space_after=6, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_after = Pt(space_after)
    if indent:
        p.paragraph_format.first_line_indent = Pt(size * 2)
    r = p.add_run(text)
    set_font(r, size, bold)
    return p


def add_heading(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(6)
    r = p.add_run(text)
    set_font(r, 14, True, "黑体")
    return p


for item in DATA:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.top_margin = sec.bottom_margin = Cm(2.54)
    sec.left_margin = sec.right_margin = Cm(3.17)

    # 标题
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run("企业资质信息核查表")
    set_font(r, 18, True, "黑体")

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(14)
    r = p.add_run(f"核查日期：{CHECK_DATE}　　核查方式：公开渠道检索")
    set_font(r, 10.5)

    # 一、基本信息
    add_heading(doc, "一、工商登记信息")

    rows = [
        ("企业名称", item["name"]),
        ("统一社会信用代码", item["code"]),
        ("法定代表人", item["legal"]),
        ("注册资本", item["capital"]),
        ("成立日期", item["founded"]),
        ("注册地址", item["address"]),
        ("登记状态", "存续（在营、开业、在册）"),
    ]
    table = doc.add_table(rows=len(rows), cols=2)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (k, v) in enumerate(rows):
        c0, c1 = table.rows[i].cells
        c0.width = Cm(4.2)
        c1.width = Cm(10.5)
        p0 = c0.paragraphs[0]
        p0.paragraph_format.space_after = Pt(2)
        r0 = p0.add_run(k)
        set_font(r0, 11, True)
        p1 = c1.paragraphs[0]
        p1.paragraph_format.space_after = Pt(2)
        r1 = p1.add_run(v)
        set_font(r1, 11)

    # 二、经营范围
    add_heading(doc, "二、经营范围")
    add_para(doc, item["scope"], indent=True)

    # 三、匹配性
    add_heading(doc, "三、与本项目的匹配性")
    add_para(doc, item["fit"], indent=True)

    # 四、资质与信用
    add_heading(doc, "四、资质与信用情况")
    for label, key in (("行政许可", "licenses"), ("知识产权", "marks"),
                       ("招投标记录", "bids"), ("风险记录", "risk")):
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        r = p.add_run(f"· {label}：")
        set_font(r, 12, True)
        r2 = p.add_run(item[key])
        set_font(r2, 12)
        if key == "risk" and "异常" in item[key]:
            r2.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

    # 五、补充说明
    add_heading(doc, "五、补充说明")
    add_para(doc, item["note"], indent=True)

    # 六、核验建议
    add_heading(doc, "六、核验建议")
    for tip in [
        "以上信息来自企查查、爱企查、百度百科等公开渠道，仅供初步筛选，"
        "正式使用前须在「国家企业信用信息公示系统」（www.gsxt.gov.cn）按统一社会信用代码复核。",
        "营业执照、开户许可证等原件复印件应由投标人提供并加盖公章，公开渠道无法替代。",
        "本项目包含底板方管焊接、激光切割、油漆喷涂等现场施工内容，"
        "经营范围属「广告制作」的企业通常不具备建筑装饰装修施工资质，"
        "如需施工类资质，应在采购文件中明确要求并另行核验。",
    ]:
        p = doc.add_paragraph()
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.left_indent = Pt(12)
        r = p.add_run("· " + tip)
        set_font(r, 11)

    # 页脚说明
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(16)
    r = p.add_run("本表由公开信息整理，不构成对企业的资质认定。")
    set_font(r, 9)
    r.font.color.rgb = RGBColor(0x80, 0x80, 0x80)

    out = BASE / item["folder"] / "资质信息核查表.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    print("saved:", out)

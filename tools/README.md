# tools — 红字模板填充工具集

围绕「Excel 数据 → Word 模板红字替换」这套流程的辅助脚本。全部依赖项目自带的 `.venv`。

## 图形界面（推荐）

双击项目根目录的 **`run_gui.bat`**，或命令行 `python gui.py`。

界面里做四件事：选结算清单（价格与明细的唯一来源）→ 选中标方 → 填甲方信息（换项目时改）→ 点「开始生成」。日志实时滚动，完成后可一键打开输出目录。

**结算清单不需要任何颜色标记**——代码按表头文字定位列（找「序号」列当锚点），黑字白底照常识别。红色只是 Word 模板里的占位标记，跟上传的 Excel 无关。

## 命令行

项目根目录执行：

```bash
python start.py          # 全量：1.声明(三份) + 4~8 全部文档
python start.py --dry    # 只检查前置文件，不写任何东西
```

跑完输出：

```
1.声明/已填/          三份资格声明
4、询价汇总表_已填.docx
5、成交通知书_已填.docx
6、合同_已填.docx
7、签收单_已填.docx
8、验收书_已填.docx
```

规则：

- `红字信息/` 下所有表里存在的占位都会替换（三家公司 + 甲方）
- **金额（小写/大写）的唯一来源是 `9、结算清单 价格.xlsx`**——红字信息表里的金额字段只当「占位匹配用」，值会被结算清单覆盖
- 6/7/8 的明细表同样按结算清单同步（行数自动增删，合计重算）
- 源文件一律不改

改价流程：只改结算清单 → 重跑 `start.py`，所有文档的金额和明细一起变。

**红字信息目录**：

| 文件 | 内容 |
|---|---|
| `晨早.xlsx` / `泉旭.xlsx` / `左右.xlsx` | 各公司乙方字段（名称/地址/联系人/电话/法人/信用代码/金额） |
| `甲方.xlsx` | 甲方字段（校名/项目名/项目简称/甲方地址/安装地点） |

加一家新公司：在 `红字信息/` 放一份 `公司名.xlsx`（三列：模板红字占位 / 字段 / 值），`start.py` 会自动按文件名匹配 `1.声明` 里的文件。

改完报价清单后重跑 `start.py` 即可，不用手动同步表格。

## 脚本清单

| 脚本 | 用途 |
|---|---|
| `core.py` | **公共层**：路径常量、红色判定、文档遍历、映射加载、替换引擎 |
| `fillers.py` | **填充器**：每种文档一个函数（声明/汇总表/通知书/合同/签收单/验收书） |
| `scan_red.py` | 扫描 docx 里的红色系标记，列出位置、上下文、原文 |
| `inspect_docx.py` | 体检 docx：段落结构、表格矩阵、每个 run 的颜色、颜色分布统计 |
| `doc2docx.py` | 老格式 `.doc` / `.wps` 转 `.docx`（走 Word COM，只读源文件） |
| `fill_docx.py` | **通用红字填充**：任意 docx + 任意公司，一条命令 |
| `sync_quote_table.py` | **报价清单 → 文档表格同步**：行数随报价单自动增删 |
| `build_sample_data.py` | 生成 `示例数据.xlsx`（主表 / 明细 / 供应商信息 / 报价对比 / 说明） |
| `build_red_info.py` | 生成 `红字信息/` 下按公司拆分的红字信息表 |
| `build_qualification.py` | 生成 `2.资质/` 下各公司的资质信息核查表 |
| `fill_declaration.py` | 用红字信息填充 `1.声明` 里的资格声明 |
| `fill_notice.py` | 填充成交通知书 |
| `fill_contract.py` | 填充合同（含法定代表人栏占位纠错） |
| `fill_acceptance.py` | 填充验收书 |
| `fill_table.py` | 早期版本，按固定 8 行搬表格，已被 `sync_quote_table.py` 取代 |

## 用法

```bash
# 扫描单个文件
python tools/scan_red.py "项目文件/1、工程A三方询价 价格/5、成交通知书.docx"

# 批量扫描
python tools/scan_red.py 项目文件/1、工程A三方询价\ 价格/*.docx

# 体检（--quiet 只看红字和非黑颜色）
python tools/inspect_docx.py "5、成交通知书.docx" --quiet

# 老 doc 转 docx
python tools/doc2docx.py "6、合同.doc"

# 重新生成示例数据（会覆盖 示例数据.xlsx）
python tools/build_sample_data.py

# 通用红字填充（任意 docx + 任意公司）
python tools/fill_docx.py "5、成交通知书.docx" "5、成交通知书_已填.docx" 晨早

# 报价清单变了，同步表格到多份文档（行数自动增删，合计自动重算）
python tools/sync_quote_table.py \
  "项目文件/1、工程A三方询价 价格/3.报价/1、报价清单（晨早）44650.00.xlsx" \
  "项目文件/1、工程A三方询价 价格/6、合同_已填.docx" \
  "项目文件/1、工程A三方询价 价格/7、签收单_已填.docx" \
  "项目文件/1、工程A三方询价 价格/8、验收书_已填.docx"

# 重新生成三份红字信息表
python tools/build_red_info.py

# 填充资格声明
python tools/fill_declaration.py
```

## 红色系判定口径

`R >= 0x80 且 G <= 0x60 且 B <= 0x60`

可同时覆盖本项目遇到的两种标注色：`EE0000`（成交通知书、验收书、资格声明）和 `FF0000`（合同）。不要退回精确匹配单个色值，会漏。

## 两个必须知道的坑

**1. run 分裂。** Word 会把 `{{name}}` 或一个词组切成多个 run。替换前必须把「相邻且同色」的 run 合并成一个逻辑块再匹配，否则匹配不到。`scan_red.py` 和 `fill_declaration.py` 都已按这个规则处理。

**2. 合并单元格重复遍历。** `table.rows[i].cells` 对合并单元格会返回同一个 `tc` 多次。用 `cell._tc.getroottree().getpath(cell._tc)` 生成路径做去重，不要用 `id()`（lxml 代理对象 identity 不稳定）。

## 路径约定

`fill_declaration.py`、`build_sample_data.py`、`build_red_info.py` 里的输入输出路径是绝对路径，改项目位置时需要同步修改脚本顶部的 `BASE` / `OUT` 常量。

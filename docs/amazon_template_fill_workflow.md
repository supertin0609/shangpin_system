# Amazon Template Fill Workflow

这份文档用于下次在新设备或新对话里快速复用“亚马逊模板填表”流程。目标是：少走弯路、优先用项目已有经验、避免慢速全量渲染、避开中文路径和文件锁问题。

> 必须先读根目录 `PROJECT_RULES.md`。本文件是流程文档，实际字段写入以 `app/template_writer.py` 为准，模板自检以 `app/template_validator.py` 为准，成功样板默认值过滤以 `app/success_rule_defaults.py` 为准。

## 新对话启动提示

可以直接把下面这段发给 Codex：

```text
我要填亚马逊模板表格。请先读取 PROJECT_RULES.md、docs/amazon_template_fill_workflow.md、data/reference_docs/亚马逊上传表格_通用自检资料.md，并以 app/template_writer.py / app/template_validator.py 的实际规则为准。不要凭通用亚马逊填表经验直接填写。

流程要求：
1. 先读取最终价格表，确认 SKU、父体 SKU、子体 SKU、价格、尺寸、重量。
2. 读取 Amazon 模板的 Data Definitions 和 Template 第 5 行字段名，确认必填字段。
3. 读取 1688 详情和 Amazon 竞品详情，提炼标题、五点、描述、材质、颜色、卖点。
4. 默认走多变体路线；保留 Template 第 6 行示例行，从第 7 行开始写 Parent/Child 数据。
5. 图片字段默认不处理：不写入、不迁移、不检查图片 URL；带文字、水印、详情页排版或无法确认合规的图片链接一律不要填。
6. 优先使用 WPS 表格 COM 或 Excel/WPS 自动化接口写入 Template 页；如果不可用，再用 openpyxl 复制模板并只写目标字段。
7. 只做字段级校验：父子体、必填字段、价格尺寸、重量、五点描述 5 个单元格、危险品/电池/原产国。
8. 不要做大范围图片渲染，不要用慢速通用表格引擎处理整本 Amazon 模板，除非我明确要求视觉预览。
9. 输出文件只保留填写好的上传表格，命名为 `产品名_V.xlsx`；根据错误报告修正后的版本依次为 `产品名_V2.xlsx`、`产品名_V3.xlsx`。
```

## 推荐执行顺序

1. 准备文件

把每个产品的资料放在同一个文件夹：

- Amazon 模板 `.xlsm`
- 最终价格确认表 `.xls` / `.xlsx`
- 1688 商品详情 HTML
- Amazon 竞品详情 HTML
- 图片文件或 URL 只作为人工素材存放；默认不写入 Amazon 图片字段

2. 建立短路径工作区

中文路径可读性好，但自动化工具容易在命令行编码里变成 `????`。建议执行前复制到短英文路径：

```powershell
New-Item -ItemType Directory -Path C:\sp_work\phone_grip -Force
Copy-Item -LiteralPath "D:\dpan\桌面\手机握把\*" -Destination C:\sp_work\phone_grip -Force
```

填完后再复制最终文件回中文目录。

3. 读取模板结构

必须读取：

- `Template` 页第 5 行：字段名
- `Data Definitions`：必填字段
- `Valid Values` / `Browse Data`：Product Type、Item Type Keyword、有效值；枚举字段必须复制当前模板的精确值，包括大小写、斜杠和空格

本项目已有经验：

- `app/template_inspector.py`：扫描模板字段和必填项
- `app/template_writer.py`：已有字段映射 `FIELD_MAP`
- `app/success_rule_defaults.py`：成功模板里沉淀的安全默认值
- `data/success_template_rules.json`：成功模板规则

4. 生成产品资料行

字段来源优先级：

- SKU、价格、尺寸、重量：最终价格确认表
- Product Type、Browse / Item Type Keyword：模板本身
- 标题、五点、描述：1688 + Amazon 竞品参考后重写
- Brand / Manufacturer：无品牌路线填 `Generic`
- Country of Origin：通常填 `China`
- Batteries Required：无电池填 `No`
- Dangerous Goods Regulations：普通无危险品填 `Not Applicable`
- 图片字段：默认留空，不从任何来源继承或写入图片链接

5. 写入方式优先级

最快且最贴近实际打开状态：

```powershell
$app = New-Object -ComObject KET.Application
$wb = $app.Workbooks.Open("C:\sp_work\phone_grip\template.xlsm")
$ws = $wb.Worksheets.Item("Template")
# 按第 5 行字段名定位列，写第 7 行以后数据
$wb.SaveAs("C:\sp_work\phone_grip\output.xlsx", 51)
$wb.Close($false)
$app.Quit()
```

WPS 表格常见 COM 名称：

- `KET.Application`：WPS 表格
- `Excel.Application`：Microsoft Excel

如果 WPS COM 不可用，再退回 `openpyxl`：

- 复制原模板
- 只清空 `Template` 页数据行
- 按字段名写目标单元格
- 不要全量重建工作簿

避免作为首选：

- 通用表格渲染/导入引擎
- 大范围 `A1:HG10` 图片预览
- 整本 workbook 视觉渲染

这些会非常慢，而且对 Amazon 模板不一定更稳。

## 必填字段检查

每次交付前至少检查这些字段。

模板必填字段通常包括：

- `contribution_sku#1.value`
- `product_type#1.value`
- `item_name[marketplace_id=ATVPDKIKX0DER][language_tag=en_US]#1.value`
- `brand[marketplace_id=ATVPDKIKX0DER][language_tag=en_US]#1.value`
- `product_description[marketplace_id=ATVPDKIKX0DER][language_tag=en_US]#1.value`
- `bullet_point[marketplace_id=ATVPDKIKX0DER][language_tag=en_US]#1.value`
- `country_of_origin[marketplace_id=ATVPDKIKX0DER]#1.value`
- `batteries_required[marketplace_id=ATVPDKIKX0DER]#1.value`
- `supplier_declared_dg_hz_regulation[marketplace_id=ATVPDKIKX0DER]#1.value`

即使模板只标了 Bullet Point #1 必填，也建议五点全部填：

- `bullet_point...#1.value`
- `bullet_point...#2.value`
- `bullet_point...#3.value`
- `bullet_point...#4.value`
- `bullet_point...#5.value`

## Listing 文案增强规则

当前项目先不接第三方关键词工具分析路径。关键词来源优先使用人工提供资料、Amazon 自动补全、竞品页面和已有成功模板经验，后续再补卖家精灵 / Helium 10 / Jungle Scout / MerchantWords 等工具数据入口。

标题：

- 控制在 100-125 字符，尽量贴近 125 字符但不能超。
- 2026-07-27 起 Amazon 官方非媒体类标题方向为 75 字符以内；项目当前仍按可上传的 100-125 字符生成长标题，超过 75 字符时 `Item Highlight` / `title_differentiation` 必须留空。
- 核心关键词靠前，核心关键词之间用英文逗号隔开，并尽量覆盖更多相关关键词、长尾词、用途、场景和款式信息。
- 除介词、冠词、连词外，单词首字母大写。
- 除介词、冠词、连词外，同一单词不要超过两次。
- 标题允许资料可证明的普通材质词作为卖点，例如 `Velvet`、`Plastic`、`Silicone`。但材质安全/等级/环保/不含某物等合规宣称不要写进标题，例如 `Food Grade`、`Medical Grade`、`BPA Free`、`Phthalate Free`、`Lead Free`、`Latex Free`、`Flame Retardant`、`Waterproof`、`Biodegradable`、`Eco Friendly`、`Organic`、`Recycled`。
- 标题禁用 Amazon 高风险特殊字符：`!`、`$`、`?`、`_`、`{`、`}`、`^`、`¬`、`¦`；除明确属于品牌名外不要使用。
- 标题不使用装饰字符或网页噪音，例如 `~`、`#`、`<`、`>`、`*`、`|`、`;`、重复标点、HTML 标签、换行、制表符、连续多空格、中文或全角标点。
- 标题不写促销、平台背书、物流或售后信息，例如 `Free Shipping`、`Sale`、`Discount`、`Best Seller`、`Amazon Choice`、`Warranty`、`Refund`、`Fast Delivery`。
- 标题不写医疗、护理、清洁、防护或结果型功效表达，例如 `Dental`、`Teeth Cleaning`、`Oral Care`、`Plaque`、`Tartar`、`Health`、`Therapy`、`Treat`、`Cure`、`Prevent`、`Antibacterial`、`Disinfect`、`UV Protection`。
- 标题不写食用/摄入暗示或绝对化成分宣称，例如 `Flavor`、`Edible`、`Digestible`、`Food Grade`、`Non-Toxic`、`Natural`、`100% Natural`、`All Natural`、`Pure`、`Chemical Free`、`Safe`、`Hypoallergenic`。非食品商品表达气味用 `Scent`。
- 标题只表达商品类型、款式/形状、适用对象、使用场景和可感知属性；凡涉及“身体、皮肤、牙齿、空气、环境更健康/更干净/更安全”的结果型表达，默认不写进标题。

父子体标题：

- Parent 标题不继承第一条 Child 标题，而是从当前 Child 标题中提取共同商品关键词，生成能覆盖全部子体的总结性标题。
- Parent 标题不带具体单一颜色、尺寸或款式；多颜色变体写 `Multiple Colors Available`，多尺寸或多款式变体写 `Multiple Styles Available`。
- Child 标题使用通用标题骨架，并追加自己的颜色、尺寸或款式属性；Child 标题不写 `Multiple Colors Available`、`Multiple Styles Available` 等 Parent 总结词。
- 多件装或套装在 Parent 和 Child 标题中都要说明，并放在标题首位，例如 `50 Pcs`、`2 Set`。
- 父子体标题生成后仍按当前标题规则自检：100-125 字符、Title Case、关键词自然覆盖、不超字符上限、不含禁用词和敏感材质宣称。

Generic Keywords：

- 如果模板包含 `generic_keyword` 字段，控制在 220 字符以内。
- 使用长尾词，避免重复标题已有关键词。

五点：

- 每条控制在 200-250 字符。
- 每条以 3-5 个英文词的关键词或卖点开头，格式为 `Keyword Phrase: ...`。
- 五条开头和文意都不要重复。
- 覆盖尺寸、材质、颜色、使用场合等真实信息。
- 每条用英文句号结尾，不使用分号。

Product Description：

- 用 4 段商务英语描述。
- 总长度控制在 1500-1800 字符。
- 每段首单词首字母大写。
- 覆盖尺寸特点、使用场景、目标人群适配和材质优点。

风险词：

- 避免侵权品牌词、平台违禁词、儿童/孕妇相关描述、医疗/杀菌/防护类承诺、绝对化表达。
- 重点禁用 `durable`、`Elegant`、`Apple`、`Samsung`、`Velcro`、`Breathable`、`magnetic`、`antibacterial`、`anti-odor`、`UV protection` 等表达。

## 父子体检查

父体行：

- SKU：父体 SKU
- Parentage Level：`Parent`
- Variation Theme：必须按当前模板 `Valid Values` 精确填写，例如 PET_TOY 可能要求 `COLOR`，不要凭语义写成 `Color`
- Brand：`Generic`
- Item Name：基于子体标题生成通用总结标题，不继承第一条子体标题，不带具体单一颜色/尺寸/款式；多颜色写 `Multiple Colors Available`，多款式写 `Multiple Styles Available`
- 通常不填 Parent SKU
- 不填 Item Condition、Model Number、Model Name、Manufacturer、Part Number、Item Highlight、价格、报价日期、最低/最高价、包装尺寸、包装重量、颜色、尺寸等子体/报价/可售专属字段；图片字段全路线默认不处理

子体行：

- SKU：子体 SKU
- Parentage Level：`Child`
- Parent SKU：父体 SKU
- Child Relationship Type：`Variation`
- Variation Theme：与 Parent 保持一致，并且必须是当前模板 `Valid Values` 中的精确枚举值
- Item Highlight / Title Differentiation：标题按 100-125 字符时留空；只有标题不超过 75 字符时才可填写
- Color：例如 `Black`
- Item Name：基于通用标题骨架追加自身颜色、尺寸或款式属性，不写 Parent 的总结词或多色/多款式可选词
- 价格、尺寸、重量按最终价格表填写
- `List Price` 和 Haul/BZR `our_price` 必填；`minimum_seller_allowed_price` 默认留空，不做人工填写值限制

## Google Drive 怎么加速

Google Drive 插件主要加速“找资料”和“复用资料”，不直接替代 WPS 填表。

适合放到 Drive 的内容：

- 成功上传过的模板文件
- 每个 Product Type 的成功案例
- 1688 / Amazon 详情源文件
- 最终价格确认表
- 规则说明和历史报错修复记录

推荐用法：

- 新设备上先从 Google Drive 搜索产品文件夹或成功模板
- 下载/导出到本地短路径，例如 `C:\sp_work\产品名`
- 本地用 WPS COM 写表
- 输出完成后再上传或同步回 Drive

这样速度会更快，因为 Codex 不需要在本地到处找历史模板，也不用重新推断成功经验。

Google Drive 不适合直接做的事：

- 直接在线编辑 Amazon `.xlsm` 宏模板
- 依赖 Google Sheets 打开 Amazon 上传模板
- 保留复杂 Excel 数据验证和隐藏结构

Amazon 模板还是本地 WPS/Excel 写入最稳。

## 中文路径优化

推荐策略：

1. 用户资料可以继续放中文文件夹，方便人工管理。
2. 自动化处理时复制到短英文路径。
3. 脚本里不要硬编码中文文件名，优先用枚举匹配：

```python
files = os.listdir(".")
template = [f for f in files if f.endswith(".xlsm") and not f.startswith("~$")][0]
price = [f for f in files if "价格" in f and f.endswith((".xls", ".xlsx")) and not f.startswith("~$")][0]
```

4. PowerShell 操作文件时用 `-LiteralPath`，不要让特殊字符参与通配匹配：

```powershell
Copy-Item -LiteralPath "D:\dpan\桌面\手机握把\手机握把-v1.xlsx" -Destination "C:\sp_work\phone_grip\手机握把-v1.xlsx"
```

5. 过滤 WPS/Excel 临时锁文件：

```text
忽略 ~$ 开头的文件
```

6. 如果要覆盖文件，先检查是否被打开：

```powershell
Get-ChildItem -Force | Where-Object { $_.Name -like '~$*' }
```

看到 `~$文件名.xlsx` 时，说明文件大概率还在 WPS/Excel 中打开，覆盖会失败。

## 交付前最小校验

不要做慢速大范围渲染。交付前只校验关键字段：

- 文件能被 WPS/Excel 打开
- `Template` 页第 7 行起有数据
- 父体 SKU / 子体 SKU 正确
- 子体 Parent SKU 指向父体
- Product Type 正确
- Brand 是 `Generic`；子体 Manufacturer 是 `Generic`，父体 Manufacturer 留空
- 父体 Item Condition / Model Number / Model Name / Manufacturer / Part Number / Item Highlight / 价格 / 包装尺寸重量留空
- 子体 List Price / Haul Price 与最终价格表一致，卖方最低价格默认留空
- 标题超过 75 字符时，Item Highlight 留空
- 必填字段非空
- 五点描述 5 个单元格非空
- `Country of Origin = China`
- `Batteries Required = No`
- `Dangerous Goods Regulations = Not Applicable`

## 速度目标

正常单产品：

- 资料齐全、模板未锁：3 到 8 分钟
- 有 Drive 成功模板可复用：更快
- 需要新增类目字段规则：10 到 20 分钟
- 文件被锁、路径异常、模板损坏：另算

如果超过 10 分钟，应该先停下来检查：

- 是否误用了全量渲染
- 是否在中文路径里硬编码文件名
- 是否有 `~$` 锁文件
- 是否在用 Google Sheets 处理复杂 Excel 模板
- 是否没有复用项目里的 `FIELD_MAP` 和成功模板规则

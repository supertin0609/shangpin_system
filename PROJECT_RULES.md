# 上品系统项目规则总纲

这份文件是新任务、新同事、Codex 代理进入本项目时的第一入口。任何亚马逊模板填表、自动填表、上传报错修复、Listing 草稿、自检交付，都必须先按这里的规则对齐。

## 必读顺序

开始填表或修表前，先读：

1. `PROJECT_RULES.md`
2. `docs/amazon_template_fill_workflow.md`
3. `data/reference_docs/亚马逊上传表格_通用自检资料.md`
4. `app/template_writer.py`
5. `app/template_validator.py`
6. `app/success_rule_defaults.py`

代码优先级高于文档：实际写入以 `app/template_writer.py` 为准，父子体标题生成和标题专项自检以 `app/variation_title_rules.py` 为准，实际自检以 `app/template_validator.py` 为准，成功样板默认值过滤以 `app/success_rule_defaults.py` 为准。

## 当前默认路线

默认路线统一为父子体路线：

```text
Haul Generic Variation
```

默认行为：

- 生成 1 行 Parent 和多行 Child。
- Parent 行填写 `parentage_level = Parent` 和 `variation_theme`。
- Child 行填写 `parentage_level = Child`、`parent_sku`、`variation_theme`。
- Parent 行不填写价格、包装尺寸、包装重量、主图、颜色、尺寸、Item Condition、Model Number、Model Name、Manufacturer、Part Number、Item Highlight 等子体/报价/可售专属字段。

只有用户明确选择或资料中明确写明时，才走：

- `Haul Generic`：单链接子体，不建立 Parent/Child。
- `Haul Generic Set Bundle`：套装售卖，不建立 Parent/Child，数量字段按套装数量填写。
- `Brand`：品牌路线，Brand / Manufacturer / 文案 / 包装必须一致；图片字段仍默认不处理。

## 价格字段规则

当前默认填写：

```text
list_price[marketplace_id=ATVPDKIKX0DER]#1.value
purchasable_offer[marketplace_id=ATVPDKIKX0DER][audience=BZR]#1.our_price#1.schedule#1.value_with_tax
```

项目内字段名：

```text
list_price
haul_price
```

父子体 V4 默认卖方最低价格：

```text
minimum_seller_allowed_price 留空
```

不要默认填写：

```text
maximum_seller_allowed_price
```

最高价字段即使在成功样板中出现，也属于不安全默认字段，不能自动补。最低价默认留空，不要从成功样板继承最低价。

## 基础写入规则

基础字段映射在 `app/template_writer.py` 的 `FIELD_MAP`。当前覆盖：

- SKU、Product Type、Record Action。
- 标题、品牌、Manufacturer、Item Type Keyword。
- Parent/Child 变体字段。
- 描述、五点、材质、颜色、尺寸、件数、包装数量。
- 非 Parent 行 List Price、Haul/BZR Price。
- 商品尺寸、包装尺寸、包装重量及单位。
- 原产国、电池、危险品、Condition。
- 部分类目基础字段，例如 PROTECTIVE_GLOVE 的 glove/coating/palm 字段。

只有模板第 5 行存在对应字段列，且数据或默认值非空时才写入。

图片字段默认不写入：

- 不填 `main_product_image_locator`、`main_offer_image_locator`、`swatch_product_image_locator` 或其他 `image` / `media_location` 图片 URL 字段。
- 不从源表、成功样板、1688、供应商资料、竞品页面或草稿里继承图片链接。
- 带文字、水印、详情页排版、中文说明、促销信息或无法确认合规的图片链接一律不要填。
- 只有用户在当前任务中明确要求处理图片字段，并提供已确认合规的公开图片 URL 时，才按用户要求单独处理。

## 固定默认值

当前默认值：

```text
record_action = Create or Replace (Full Update)
非 Parent 行 condition_type = New，Parent 行留空
item_package_quantity = set_count 或 1
尺寸单位 = Inches
包装重量单位 = Pounds
country_of_origin = China
batteries_required = No
batteries_included = No
supplier_declared_dg_hz_regulation = Not Applicable
```

非电池商品不要填写：

```text
contains_battery_or_cell
```

该字段有效值是 `Battery` / `Cell`，不能填 `No`；普通非电池产品按默认留空。

Generic 路线：

```text
brand = Generic
manufacturer = Generic
```

## Data Definitions 必填兜底

系统会读取模板的 `Data Definitions`，凡是标记为 `Required` 的字段，会按稳定默认逻辑尝试补齐。

可兜底的字段类型包括：

- Brand / Manufacturer / Model Name / Model Number / Part Number。
- Material、Color Map、数量、Unit Count、Included Components。
- Uses、Theme、Pattern、Style、Care Instructions。
- Country of Origin、电池、危险品、Condition、Product Tax Code。
- 尺寸和重量数值/单位。

不能稳定判断的字段不硬填，交给自检报告或人工确认。

## Product Type 条件字段

类目条件字段在 `app/template_validator.py` 的 `PRODUCT_TYPE_CONDITIONAL_FIELDS`。

当前覆盖：

- `ANIMAL_COLLAR`
- `BOTTLE`
- `COSMETIC_CASE`
- `CLEANING_BRUSH`
- `PET_TOY`
- `PLANTER`
- `TOWEL`
- `PROTECTIVE_GLOVE`

这些规则只对对应 Product Type 生效，不能泛化到所有模板。

PET_TOY 注意：

- 当前上传实测 `Subject Character` 会成为条件必填；普通猫玩具可按产品主体填写 `Cat`。

ANIMAL_COLLAR 注意：

- `dog_breed_size` 模板有效值是 `Extra Small`、`Small`、`Medium`、`Large`、`Giant`、`All`。
- 不要填 `All Breed Sizes`；通用犬种尺寸用 `All`，明确小型宠物/小型犬时用 `Small`。

PLANTER 注意：

- 当前上传实测 `Model Name`、`Special Features`、`Mounting Type`、`Product Compliance Certificate` 会成为条件必填。
- 普通换盆垫/园艺操作垫可用 `Special Features = Foldable`、`Mounting Type = Tabletop`、`Product Compliance Certificate = Not Applicable`。
- 若 Amazon warning 提示 `item_length_width_height` 不适用于当前类目，应清空该组商品展开尺寸字段，只保留包装尺寸字段。

## 成功样板规则

成功样板来源：

```text
data/success_template_rules.json
```

写入前必须经过 `app/success_rule_defaults.py` 的安全过滤。

以下字段属于不安全默认，不能因为历史成功样板出现过就自动填：

- SKU、标题、描述、五点。
- 图片 URL。
- 价格字段，包括 `list_price`、`our_price`、`maximum_seller_allowed_price`，以及非项目固定值的 `minimum_seller_allowed_price`。
- 颜色、尺寸、变体字段。
- 包装尺寸/重量数值。
- `skip_offer`。

## 文案规则

Listing 文案必须英文填写。

标题：

- 100-125 字符，尽量贴近 125 字符但不能超。
- 2026-07-27 起 Amazon 官方非媒体类标题方向为 75 字符以内；项目当前仍按可上传的 100-125 字符生成长标题，但超过 75 字符时必须把 `Item Highlight` / `title_differentiation` 留空，并关注后续短标题迁移。
- 核心关键词靠前，并覆盖更多相关关键词、长尾词、用途、场景和款式信息。
- 不堆砌关键词，不全大写。
- 除介词、冠词、连词外，同一单词不要超过两次。
- Generic 路线不出现供应商品牌/公司名。
- 禁用 Amazon 标题高风险字符：`!`、`$`、`?`、`_`、`{`、`}`、`^`、`¬`、`¦`。除明确属于品牌名外，不出现在标题中。
- 标题不使用装饰或网页噪音字符/结构，例如 `~`、`#`、`<`、`>`、`*`、`|`、`;`、重复标点、HTML 标签、换行、制表符、连续多空格、中文或全角标点。
- 标题不写促销、平台背书、物流或售后表达，例如 `Free Shipping`、`Sale`、`Discount`、`Best Seller`、`Amazon Choice`、`Warranty`、`Refund`、`Fast Delivery`。
- 标题避免医疗、护理、清洁、防护、结果型功效或健康改善表达，例如 `Dental`、`Teeth Cleaning`、`Oral Care`、`Plaque`、`Tartar`、`Health`、`Therapy`、`Treat`、`Cure`、`Prevent`、`Antibacterial`、`Disinfect`、`UV Protection`。
- 标题避免食用/摄入暗示和绝对化成分宣称，例如 `Flavor`、`Edible`、`Digestible`、`Food Grade`、`Non-Toxic`、`Natural`、`100% Natural`、`All Natural`、`Pure`、`Chemical Free`、`Safe`、`Hypoallergenic`。普通非食品商品表达气味用 `Scent`。
- 标题允许资料可证明的普通材质词作为卖点，例如 `Velvet`、`Plastic`、`Silicone`。但材质安全/等级/环保/不含某物等合规宣称不要写进标题，例如 `Food Grade`、`Medical Grade`、`BPA Free`、`Phthalate Free`、`Lead Free`、`Latex Free`、`Flame Retardant`、`Waterproof`、`Biodegradable`、`Eco Friendly`、`Organic`、`Recycled`。
- 标题只描述商品类型、款式/形状、适用对象、使用场景和可感知属性；凡涉及“让身体、皮肤、牙齿、空气、环境更健康/更干净/更安全”的结果型表达，默认不写进标题。
- 标题按 100-125 字符时，不填写 `Item Highlight` / `title_differentiation`。Amazon 报错 100476 已确认：只要填写 Item Highlight，Item Name 必须 75 字符以内。

父子体标题：

- 父体标题不再继承第一条子体标题，系统会基于现有子体标题生成可覆盖全部子体的通用总结标题。
- 父体标题去掉具体颜色、尺寸等单一子体属性；多颜色变体加入 `Multiple Colors Available`，多尺寸或多款式变体加入 `Multiple Styles Available`。
- 子体标题基于通用标题生成，并追加自己的颜色、尺寸或款式属性；子体标题不能包含 `Multiple Colors Available`、`Multiple Styles Available` 等父体总结词。
- 多件装或套装如果在 `set_count` 或原标题中可识别，父体和子体标题都必须把数量/套装前缀放在首位，例如 `50 Pcs`、`2 Set`。
- 父子体标题仍必须遵守现有标题长度、关键词覆盖、Title Case、禁用词和敏感材质宣称规则，不能超过当前配置的标题字符上限。

五点：

- 尽量 5 条全部填写。
- 每条以 3-5 个英文词卖点开头，格式为 `Keyword Phrase: ...`。
- 每条末尾用英文句号。
- 不使用分号。

描述：

- 商务英语。
- 通常 4 段。
- 不写无法证明的承诺。

风险词：

- 避免侵权品牌词、医疗/杀菌/防护承诺、儿童/孕妇相关描述、绝对化表达。
- 项目历史风险词包括 `durable`、`Elegant`、`Apple`、`Samsung`、`Velcro`、`Breathable`、`magnetic`、`antibacterial`、`anti-odor`、`UV protection` 等。

## 自检规则

交付填好的 Amazon 模板前必须跑模板自检。

最低检查：

- 非 Parent 行 `Item Condition = New`，Parent 行可留空。
- `Skip Offer` 留空。
- 非 Parent 行 `List Price` 和 Haul/BZR `our_price` 已填写；`minimum_seller_allowed_price` 默认留空，不做人工填写值限制。
- 标题超过 75 字符时，`Item Highlight` / `title_differentiation` 必须留空。
- 尺寸数值和单位成对填写。
- Parent 行不填 Parent SKU；Child 行必须填 Parent SKU。
- Parent/Child 行必须有 Variation Theme。
- 枚举字段必须与当前模板 `Valid Values` 完全一致，包括大小写、斜杠和空格；不能把 Amazon 枚举值按人类习惯改写。重点核对 `variation_theme#1.name`、`parentage_level`、`child_relationship_type` 等字段，例如 PET_TOY 模板可能要求 `COLOR`，不是 `Color`。
- 标题、描述、五点不能含中文。
- `Data Definitions` 标 Required 的字段非空。
- 当前 Product Type 的条件字段非空。
- TOWEL 当前不允许的合规字段保持空。

常用命令：

```bash
python3 run.py auto-fill "data/projects/日期_产品名"
python3 run.py check-template "data/projects/日期_产品名/05_填表版本/产品名_v1.xlsx"
python3 run.py check-template "data/projects/日期_产品名/05_填表版本/产品名_v1.xlsx" --write-report
python3 run.py parse-report "processing-summary.xlsm"
```

默认不生成写入报告或模板自检报告；只有需要落地报告文件时才加 `--write-report` / `--write-reports` 或指定 `-o/--output`。

## 报错复盘规则

拿到 processing-summary 后：

- 先记录错误码、错误字段、受影响 SKU。
- 不直接修改 processing-summary 当上传文件。
- 回到源模板修字段。
- 如果报错字段是枚举值字段，先回到同一份模板的 `Valid Values` 页确认精确可接受值，再修源模板；不要只按语义判断字段值正确。
- 修复后的规则若属于某 Product Type 或某路线，必须补入整体自检或自动填表规则。
- 单品特殊情况只保留在单品学习记录，不泛化。

## 新任务开场要求

新开任务时，如果内容涉及填表或修表，先说明已经读取项目规则，再开始操作。没有读取上述规则时，不要直接填 Amazon 表格。

用户不需要在每次任务里重复发送填表规则。只要任务涉及 Amazon 模板填表、修表、自检、报错处理或 Listing 草稿，Codex 默认按本文件、`AGENTS.md` 和实际代码里的最新规则启动。

修改任何长期默认规则、字段写入规则、自检规则或报告生成行为时，必须同步更新 `AGENTS.md` 的入口摘要，避免新任务继续读取旧默认值。

# 上品流程本地 MVP

这是把“选品建档 -> 路线判断 -> 资料整理 -> 上传前自检 -> 报告解析”落地成可运行项目的第一版。

## 重要：先读项目规则

任何亚马逊模板填表、自动填表、上传报错修复、Listing 草稿、自检交付，必须先读根目录：

```bash
PROJECT_RULES.md
AGENTS.md
```

实际写入规则以 `app/template_writer.py` 为准，实际自检规则以 `app/template_validator.py` 为准，成功样板默认值过滤以 `app/success_rule_defaults.py` 为准。不要只凭通用亚马逊填表经验操作。

## 初始化

```bash
cd /Users/admin/Desktop/上品模版填写上传/shangpin_system
python3 run.py init
```

初始化后会生成：

- `templates/产品资料模板.xlsx`
- `data/projects/`
- `outputs/`

## 创建一个产品项目

```bash
python3 run.py new-project "黑色收纳袋"
```

每个项目会自动生成：

- `01_采购资料`
- `02_产品包装和定价`
- `03_产品详情页`
- `04_竞品参考`
- `05_模版原件`

产品资料表会放在 `03_产品详情页` 里。

## 查看项目状态

```bash
python3 run.py list-projects
```

系统会汇总显示：

- 项目目录
- 产品名
- 当前状态
- SKU 数
- 模板自检错误数
- 最新模板
- 最近更新时间

如果只想输出项目路径：

```bash
python3 run.py list-projects --paths
```

## 启动本地工作台

```bash
python3 run.py workbench
```

默认会打开：

```bash
http://127.0.0.1:8766
```

工作台正在重新设计中。旧工具界面已经清空，后续会按模块讨论后再生成新的网页预览。

如果只想启动服务，不自动打开浏览器：

```bash
python3 run.py workbench --no-open
```

## 上传前自检

```bash
python3 run.py validate "/完整路径/产品资料.xlsx"
```

系统会检查：

- 核心字段是否缺失
- Generic/品牌路线是否冲突
- Haul 价格是否超过配置上限
- 包装尺寸重量是否完整且大于 0
- List Price / Haul Price 是否冲突，卖方最低价格是否按默认留空
- 文案是否出现高风险宣称词
- 高风险类目关键词

模板自检默认只在终端输出结果，不生成报告；需要报告时使用 `check-template --write-report` 或指定 `-o/--output`。

## 自动提炼项目资料

先创建项目：

```bash
python3 run.py new-project "厨房硅胶刮刀"
```

然后把资料放进项目文件夹：

- `01_采购资料`：采购单、报价表、供应商参数表
- `02_产品包装和定价`：包装尺寸、重量、成本、定价资料
- `03_产品详情页`：标题、五点、描述、补充说明
- `04_竞品参考`：竞品链接、标题、五点、截图备注
- `05_模版原件`：Amazon 模板原件

目前支持读取：

- `.xlsx`
- `.xlsm`
- `.pdf`
- `.html`
- `.htm`
- `.csv`
- `.tsv`
- `.txt`
- `.md`

运行：

```bash
python3 run.py analyze-project "/完整路径/data/projects/日期_厨房硅胶刮刀"
```

系统会生成：

- `03_产品详情页/..._自动提炼草稿.xlsx`
- `outputs/..._资料提炼报告.md`

草稿会尽量提取：

- 产品名
- SKU
- 颜色
- 尺寸
- 材质
- 成本/价格
- 包装尺寸重量，并换算成 inches / pounds
- 供应商链接
- 竞品链接
- 默认多变体路线建议，自动设置 Parent/Child 和变体主题
- 标题、五点、描述草稿

注意：这一版是“提炼草稿”，不是最终上传文件。价格、类目、`item_type_keyword`、尺寸重量仍建议人工确认；图片字段默认不处理、不写入。

## 解析 processing-summary

```bash
python3 run.py parse-report "/完整路径/processing-summary.xlsx"
```

目前内置识别：

- `90041`
- `90220`
- `100643`
- `18320`
- `8541`
- `13013`
- `5887`
- `5882`
- `100521`

## 沉淀报错复盘库

每次上传失败后，把 `processing-summary` 加入复盘库：

```bash
python3 run.py learn-report "/完整路径/processing-summary.xlsm" --product "产品名" --note "本次修复备注"
```

也可以一次学习多个报告：

```bash
python3 run.py learn-report "/路径/v1-processing-summary.xlsm" "/路径/v2-processing-summary.xlsm" --product "彩色瑜伽砖"
```

系统会生成：

- `data/error_learnings.json`
- `outputs/上传报错复盘库.md`

复盘库会累计：

- 错误码
- 错误字段
- 受影响 SKU
- 原始报错信息
- 归纳出的上传前自检规则

注意区分规则作用范围：

- 通用规则：例如 `Brand = Generic`、普通非危险品 `Dangerous Goods Regulations = Not Applicable`。
- 模板字段规则：只有模板存在对应字段时才检查，例如 `item_depth_width_height` 的数值和单位配套。
- Product Type 规则：只对相同或相近 `product_type` 生效，不要直接套用到所有类目。

## 标记上传成功

Amazon 后台确认上传成功后，把项目状态沉淀下来：

```bash
python3 run.py mark-uploaded "data/projects/日期_产品名"
```

默认会自动取 `05_填表版本` 里的最新模板。也可以手动指定：

```bash
python3 run.py mark-uploaded "data/projects/日期_产品名" \
  --template "data/projects/日期_产品名/05_填表版本/成功模板.xlsx" \
  --sku-count 5 \
  --uploaded-at 2026-06-18 \
  --note "后台确认上传成功"
```

这会更新项目里的 `project_status.json`，后续 `auto-fill` 默认会跳过已成功上传的项目，避免误覆盖。

## 成功上传模板样板库

桌面 `亚马逊上传表格存档` 里的成功上传文件已整理到：

```bash
data/success_templates/
```

这些文件用于：

- 对照不同 Product Type 的真实成功字段。
- 提炼字段默认值、条件必填字段和类目差异。
- 作为后续自动填表逻辑的回归样本。

注意：这些是成功样板，不建议直接覆盖业务项目里的填表版本。新增产品仍应放到 `data/projects/` 下运行自动填表。

图片相关字段默认不整理、不检查、不写入；即使资料草稿或源表里存在图片 URL，自动填表也不会写入 Amazon 模板。

## 提炼成功填写规则

从成功上传模板样板库里提炼 Product Type 规则：

```bash
python3 run.py learn-success-templates
```

系统会生成：

- `data/success_template_rules.json`
- `outputs/成功模板规则提炼报告.md`

报告会按 Product Type 汇总：

- 成功样板数量和 SKU 数。
- 成功表格里稳定出现的固定默认值。
- 成功表格里经常填写、但当前自动填表还没覆盖的字段。
- 图片字段默认排除，不参与规则建议，不作为自动填表默认值。

`fill-template` 和 `auto-fill` 会读取这份规则库，保守补写低风险默认值：

- 只按相同 Product Type 应用。
- 只补模板里存在、当前为空的字段。
- 不覆盖产品资料草稿或原字段映射已经写入的内容。
- 默认会补 Parent/Child 变体关系和安全字段；不自动编造售价、标题、五点、描述或图片，也不写入任何图片 URL。

## 后续扩展方向

1. 在工作台里增加资料放置指引和文件打开按钮。
2. 继续扩大安全默认值范围，按 Product Type 补更多稳定字段。
3. 增加 Data Definitions/数据定义 解析，自动判断条件必填。
4. 保持图片字段默认留空；如后续要处理图片，单独建立人工确认后的图片流程。
5. 增加基础测试，覆盖自动填表、状态读写和模板自检。

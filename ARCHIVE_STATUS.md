# 上品自动填表 MVP 存档状态

存档日期：2026-06-17

## 当前能力

- 创建产品项目目录和产品资料表。
- 从采购资料、竞品资料、PDF、HTML、Excel、文本文件中生成产品资料草稿。
- 将产品资料草稿写入 Amazon 上传模板的 `Template` / `模板` 页。
- 对已填写模板执行上传前自检。
- 解析 processing-summary 常见错误。
- 将上传报错沉淀到复盘库。
- 通过 `project_status.json` 记录每个项目的流程状态。
- 通过 `auto-fill` 命令串联自动提炼、写模板和模板自检。
- 通过 `list-projects` 汇总显示项目状态、SKU 数、错误数和最新模板。
- 通过 `mark-uploaded` 命令在上传成功后更新项目状态。
- 通过 `data/success_templates/` 保存已成功上传表格样板，用于后续规则学习。
- 通过 `learn-success-templates` 从成功样板中提炼 Product Type 字段规则。
- `fill-template` / `auto-fill` 已开始保守接入成功规则，只补低风险默认值且不覆盖已有内容。
- 通过 `workbench` 启动本地网页工作台，查看项目状态、执行自动填表、标记上传成功和查看规则库。

## 已验证样品

- 花园艺手套：用户确认已成功上传。
- 花束卡片夹：用户确认已成功上传。
- 瑜伽砖：用户确认已成功上传。

## 当前项目状态

- `data/projects/20260606_花园艺手套/project_status.json`：`uploaded_success`
- `data/projects/20260604_花束卡片夹/project_status.json`：`uploaded_success`
- `data/projects/20260530_瑜伽砖/project_status.json`：`uploaded_success`
- 厨房硅胶刮刀是早期测试项目，已从项目库清理。
- `data/success_templates/`：已收录 13 个成功上传表格样板，覆盖 `ANIMAL_COLLAR`、`BRA`、`COFFEE_FILTER`、`EXERCISE_BAND`、`FUNNEL`、`GARDEN_SHEAR_SCISSORS`、`PET_TOY`、`RECREATION_BALL`、`UMBRELLA`。

## 常用命令

```bash
python3 run.py list-projects
python3 run.py auto-fill "data/projects/项目目录"
python3 run.py auto-fill "data/projects/项目目录" --force
python3 run.py mark-uploaded "data/projects/项目目录"
python3 run.py learn-success-templates
python3 run.py workbench
python3 run.py check-template "路径/已填写模板.xlsx"
python3 run.py learn-report "路径/processing-summary.xlsx" --product "产品名"
```

`check-template` 默认只输出终端结果；需要落地报告时加 `--write-report` 或 `-o`。

## 下一步建议

1. 在工作台里增加资料放置指引和文件打开按钮。
2. 继续扩大安全默认值范围，按 Product Type 补更多稳定字段。
3. 增加图片 URL 检查。
4. 增加基础测试，覆盖 `auto-fill`、状态读写和模板自检。

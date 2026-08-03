import argparse
from pathlib import Path

from .analyzer import analyze_project
from .auto_fill import auto_fill_project
from .error_learning import learn_reports
from .paths import TEMPLATES_DIR, ensure_base_dirs
from .project_manager import create_project, list_project_summaries, list_projects
from .project_status import infer_latest_template, infer_product_name, infer_sku_count, mark_uploaded_success
from .report_parser import parse_processing_summary
from .success_templates import learn_success_templates
from .template_inspector import inspect_template
from .template_validator import validate_template_file
from .template_writer import fill_template
from .validator import validate_intake, write_validation_report
from .workbench import run_workbench
from .workbook_io import create_intake_workbook


def cmd_init(_args):
    ensure_base_dirs()
    template_path = TEMPLATES_DIR / "产品资料模板.xlsx"
    create_intake_workbook(template_path)
    print(f"已初始化上品系统：{Path(__file__).resolve().parents[1]}")
    print(f"资料表模板：{template_path}")


def cmd_new_project(args):
    project_dir, intake_path = create_project(args.name)
    print(f"已创建项目：{project_dir}")
    print(f"产品资料表：{intake_path}")


def cmd_list_projects(args):
    summaries = list_project_summaries()
    if not summaries:
        print("还没有项目。")
        return
    if args.paths:
        for project in list_projects():
            print(project)
        return

    rows = []
    for item in summaries:
        rows.append([
            item["folder"],
            item["product_name"],
            _status_label(item["status"]),
            str(item["sku_count"] or "-"),
            _error_text(item["template_error_count"]),
            item["latest_template"] or "-",
            item["updated_at"][:10] or "-",
        ])
    _print_table(["项目目录", "产品", "状态", "SKU", "错误", "最新模板", "更新"], rows)


def cmd_validate(args):
    findings = validate_intake(args.intake)
    output = write_validation_report(args.intake, findings, args.output)
    print(f"自检完成：{output}")
    print(f"发现 {len(findings)} 个提示。")
    error_count = sum(1 for item in findings if item["severity"] == "error")
    if error_count:
        print(f"其中 {error_count} 个 Error 建议上传前先处理。")


def cmd_analyze_project(args):
    draft_path, report_path = analyze_project(args.project_dir, args.output)
    print(f"自动提炼草稿：{draft_path}")
    print(f"资料提炼报告：{report_path}")
    print("请先人工确认草稿里的价格、类目、item_type_keyword、品牌路线和尺寸重量；图片字段默认不处理。")


def cmd_parse_report(args):
    findings = parse_processing_summary(args.report)
    if not findings:
        print("没有识别到已配置的常见错误码。")
        return
    for item in findings:
        print(f"[{item['code']}] {item['type']}")
        print(f"含义：{item['meaning']}")
        print(f"建议：{item['fix']}")
        print(f"原文：{item['source'][:240]}")
        print("")


def cmd_inspect_template(args):
    findings, output_path = inspect_template(args.template, args.output)
    print(f"模板扫描完成：{output_path}")
    print(f"Product Type：{findings['product_type'] or '未识别'}")
    print(f"Browse Node / Item Type Keyword：{findings['browse_node'] or '未识别'}")
    print(f"必填字段数：{len(findings['required_fields'])}")


def cmd_fill_template(args):
    output_path, draft_path, template_path, written_fields = fill_template(
        args.project_dir,
        draft_path=args.draft,
        template_path=args.template,
        output_path=args.output,
        write_report=args.write_report,
    )
    print(f"已写入模板：{output_path}")
    print(f"使用草稿：{draft_path}")
    print(f"使用模板：{template_path}")
    print(f"写入字段数：{len(written_fields)}")


def cmd_learn_report(args):
    added, json_path, report_path = learn_reports(args.reports, product_name=args.product, note=args.note or "")
    print(f"新增复盘记录：{len(added)}")
    print(f"结构化记录：{json_path}")
    print(f"复盘报告：{report_path}")


def cmd_check_template(args):
    write_report = args.write_report or bool(args.output)
    findings, output_path = validate_template_file(args.template, args.output, write_report=write_report)
    if output_path:
        print(f"模板自检完成：{output_path}")
    else:
        print("模板自检完成：未生成自检报告。")
    print(f"发现 {len(findings)} 个问题。")


def cmd_auto_fill(args):
    result = auto_fill_project(
        args.project_dir,
        draft_path=args.draft,
        template_path=args.template,
        output_path=args.output,
        force=args.force,
        write_reports=args.write_reports,
    )
    if result["skipped"]:
        status = result["status"]
        print("项目已标记为上传成功，默认跳过自动填表。")
        print(f"项目：{result['project_dir']}")
        print(f"最新模板：{status.get('latest_template', '未记录')}")
        print("如需重新生成，请加 --force。")
        return

    if result.get("blocked"):
        print("自动填表暂停：项目资料还不完整。")
        print(f"项目：{result['project_dir']}")
        print(f"原因：{result['reason']}")
        print(f"项目状态：{result['status_path']}")
        return

    print(f"自动填表完成：{result['filled_path']}")
    print(f"使用草稿：{result['draft_path']}")
    print(f"使用模板：{result['template_path']}")
    print(f"项目状态：{result['status_path']}")
    print(f"SKU 数：{result['sku_count']}")
    print(f"写入字段数：{result['written_field_count']}")
    print(f"模板 Error：{result['error_count']}")
    if result["status"] == "ready_for_upload":
        print("结论：本地模板自检通过，可进入人工上传。")
    else:
        print("结论：还需要先处理自检报告里的字段。")


def cmd_mark_uploaded(args):
    project_dir = Path(args.project_dir)
    template_path = Path(args.template) if args.template else infer_latest_template(project_dir)
    if template_path is None:
        raise SystemExit("没有找到可标记的上传模板，请用 --template 指定成功上传的 xlsx/xlsm 文件。")

    product_name = infer_product_name(project_dir, args.product or "")
    sku_count = infer_sku_count(project_dir, args.sku_count)
    status_path = mark_uploaded_success(
        project_dir,
        product_name=product_name,
        latest_template=template_path,
        sku_count=sku_count,
        uploaded_at=args.uploaded_at,
        notes=args.note or "",
    )
    print("已标记上传成功。")
    print(f"项目：{project_dir}")
    print(f"产品：{product_name}")
    print(f"成功模板：{template_path}")
    print(f"SKU 数：{sku_count if sku_count is not None else '未识别'}")
    print(f"项目状态：{status_path}")


def cmd_learn_success_templates(args):
    rules, json_path, report_path = learn_success_templates(
        source_dir=args.source,
        json_path=args.json_output,
        report_path=args.report_output,
    )
    print(f"成功样板数：{rules['template_count']}")
    print(f"Product Type 数：{rules['product_type_count']}")
    print(f"规则 JSON：{json_path}")
    print(f"规则报告：{report_path}")


def cmd_workbench(args):
    run_workbench(host=args.host, port=args.port, open_browser=not args.no_open)


def _status_label(status):
    labels = {
        "uploaded_success": "已上传",
        "ready_for_upload": "待上传",
        "needs_manual_fix": "待修正",
        "blocked": "卡住",
        "not_started": "未开始",
    }
    return labels.get(status, status)


def _error_text(error_count):
    if error_count is None:
        return "-"
    return str(error_count)


def _print_table(headers, rows):
    widths = [len(header) for header in headers]
    for row in rows:
        for index, cell in enumerate(row):
            widths[index] = max(widths[index], _display_width(cell))

    def fmt_row(row):
        cells = []
        for index, cell in enumerate(row):
            cells.append(str(cell) + " " * (widths[index] - _display_width(cell)))
        return "  ".join(cells)

    print(fmt_row(headers))
    print(fmt_row(["-" * width for width in widths]))
    for row in rows:
        print(fmt_row(row))


def _display_width(value):
    return sum(2 if ord(char) > 127 else 1 for char in str(value))


def build_parser():
    parser = argparse.ArgumentParser(description="上品流程本地 MVP")
    subparsers = parser.add_subparsers(dest="command", required=True)

    init = subparsers.add_parser("init", help="初始化目录并生成产品资料模板")
    init.set_defaults(func=cmd_init)

    new_project = subparsers.add_parser("new-project", help="创建一个产品项目")
    new_project.add_argument("name", help="项目名，例如：黑色收纳袋")
    new_project.set_defaults(func=cmd_new_project)

    list_cmd = subparsers.add_parser("list-projects", help="列出已有项目状态")
    list_cmd.add_argument("--paths", action="store_true", help="只输出项目路径")
    list_cmd.set_defaults(func=cmd_list_projects)

    validate = subparsers.add_parser("validate", help="读取产品资料表并生成上传前自检报告")
    validate.add_argument("intake", help="产品资料 xlsx 路径")
    validate.add_argument("-o", "--output", help="自检报告输出路径")
    validate.set_defaults(func=cmd_validate)

    analyze = subparsers.add_parser("analyze-project", help="扫描项目资料并生成产品资料草稿")
    analyze.add_argument("project_dir", help="项目文件夹路径")
    analyze.add_argument("-o", "--output", help="草稿产品资料表输出路径")
    analyze.set_defaults(func=cmd_analyze_project)

    report = subparsers.add_parser("parse-report", help="解析 processing-summary 常见错误码")
    report.add_argument("report", help="processing-summary 文件路径，支持 xlsx/csv/tsv/txt")
    report.set_defaults(func=cmd_parse_report)

    inspect = subparsers.add_parser("inspect-template", help="扫描亚马逊模板字段和必填项")
    inspect.add_argument("template", help="亚马逊模板 xlsx/xlsm 路径")
    inspect.add_argument("-o", "--output", help="模板扫描报告输出路径")
    inspect.set_defaults(func=cmd_inspect_template)

    fill = subparsers.add_parser("fill-template", help="把自动提炼草稿写入亚马逊模板 Template/模板 页")
    fill.add_argument("project_dir", help="项目文件夹路径")
    fill.add_argument("--draft", help="指定自动提炼草稿 xlsx 路径")
    fill.add_argument("--template", help="指定亚马逊模板 xlsx/xlsm 路径")
    fill.add_argument("-o", "--output", help="输出填好后的模板路径")
    fill.add_argument("--write-report", action="store_true", help="额外生成写入报告；默认只输出填好的上传表格")
    fill.set_defaults(func=cmd_fill_template)

    learn = subparsers.add_parser("learn-report", help="把 processing-summary 报错沉淀到复盘库")
    learn.add_argument("reports", nargs="+", help="一个或多个 processing-summary 文件路径")
    learn.add_argument("--product", help="商品名，用于复盘归档")
    learn.add_argument("--note", help="本次修复备注")
    learn.set_defaults(func=cmd_learn_report)

    check_template = subparsers.add_parser("check-template", help="检查已填好的上传模板")
    check_template.add_argument("template", help="已填好的 xlsx/xlsm 上传模板")
    check_template.add_argument("-o", "--output", help="自检报告输出路径；指定后会生成报告")
    check_template.add_argument("--write-report", action="store_true", help="额外生成模板自检报告；默认只在终端输出结果")
    check_template.set_defaults(func=cmd_check_template)

    auto_fill = subparsers.add_parser("auto-fill", help="自动提炼草稿、写入模板并执行模板自检")
    auto_fill.add_argument("project_dir", help="项目文件夹路径")
    auto_fill.add_argument("--draft", help="指定产品资料草稿 xlsx 路径")
    auto_fill.add_argument("--template", help="指定亚马逊模板 xlsx/xlsm 路径")
    auto_fill.add_argument("-o", "--output", help="输出填好后的模板路径")
    auto_fill.add_argument("--force", action="store_true", help="即使项目已标记上传成功也重新生成")
    auto_fill.add_argument("--write-reports", action="store_true", help="额外生成写入报告和模板自检报告；默认只输出填好的上传表格")
    auto_fill.set_defaults(func=cmd_auto_fill)

    mark_uploaded = subparsers.add_parser("mark-uploaded", help="把项目标记为上传成功")
    mark_uploaded.add_argument("project_dir", help="项目文件夹路径")
    mark_uploaded.add_argument("--template", help="成功上传的 xlsx/xlsm 模板路径；不填则自动取 05_填表版本 最新文件")
    mark_uploaded.add_argument("--product", help="产品名；不填则从状态或项目目录推断")
    mark_uploaded.add_argument("--sku-count", type=int, help="SKU 数；不填则从状态或产品资料表推断")
    mark_uploaded.add_argument("--uploaded-at", help="上传成功日期，格式 YYYY-MM-DD；不填则使用今天")
    mark_uploaded.add_argument("--note", help="备注")
    mark_uploaded.set_defaults(func=cmd_mark_uploaded)

    learn_success = subparsers.add_parser("learn-success-templates", help="从成功上传模板提炼字段规则")
    learn_success.add_argument("--source", help="成功模板目录，默认 data/success_templates")
    learn_success.add_argument("--json-output", help="规则 JSON 输出路径，默认 data/success_template_rules.json")
    learn_success.add_argument("--report-output", help="规则报告输出路径，默认 outputs/成功模板规则提炼报告.md")
    learn_success.set_defaults(func=cmd_learn_success_templates)

    workbench = subparsers.add_parser("workbench", help="启动本地工作台")
    workbench.add_argument("--host", default="127.0.0.1", help="监听地址，默认 127.0.0.1")
    workbench.add_argument("--port", type=int, default=8766, help="端口，默认 8766")
    workbench.add_argument("--no-open", action="store_true", help="只启动服务，不自动打开浏览器")
    workbench.set_defaults(func=cmd_workbench)

    return parser


def main():
    parser = build_parser()
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()

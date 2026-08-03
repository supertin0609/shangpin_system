import json
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

from .paths import DATA_DIR, OUTPUTS_DIR
from .report_parser import extract_processing_summary


LEARNINGS_PATH = DATA_DIR / "error_learnings.json"
REPORT_PATH = OUTPUTS_DIR / "上传报错复盘库.md"


FIELD_RULE_HINTS = {
    "Dangerous Goods Regulations": "普通非危险品应填 Not Applicable，不要填 No；如是液体、化学品、喷雾、带电池产品需单独判断。",
    "Item Depth Unit": "模板专属条件必填：当模板包含 item_depth_width_height 字段，且该类目触发本体尺寸时，Item Depth/Height/Width 的数值和单位必须成对填写；不要套用到没有这些字段的模板。",
    "Material": "Material 在很多类目会条件必填；从采购资料或产品材质中提前填入，并确认值在 Valid Values 中。",
    "List Price": "List Price 上传前必须填数字；Haul Price / BZR Price 通常也要同步填写。",
    "Brand Name": "Haul Generic 路线 Brand 应为 Generic，并检查 Manufacturer、标题、描述和包装表达是否一致；图片字段默认不处理。",
    "Main Image URL": "图片字段默认不处理、不迁移、不写入；只有用户明确要求并提供已确认合规的公开图片 URL 时才单独处理。",
    "Item Type Keyword": "必须使用 Browse Data / BTG 中的有效值，并确认与 product_type 和实际商品匹配。",
    "Product Compliance Certificate": "TOWEL 等类目若 processing-summary 报该字段必填，通常填 Not Applicable；加入同 product_type 自检。",
    "Compliance - Towel End Use": "TOWEL 在当前 ships_globally 条件下该字段可能不允许提交；无明确要求时留空，不要主动填 Other Towel。",
    "Compliance Weave Type": "TOWEL 在当前 ships_globally 条件下该字段可能不允许提交；processing-summary 报 90248 时应清空。",
    "Compliance - Outer Surface Material": "TOWEL 在当前 ships_globally 条件下该字段可能不允许提交；材质信息填 Material/Fabric Type，不要填到该合规字段。",
}


def load_learnings():
    if not LEARNINGS_PATH.exists():
        return {"records": []}
    return json.loads(LEARNINGS_PATH.read_text(encoding="utf-8"))


def save_learnings(data):
    LEARNINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    LEARNINGS_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def learn_reports(paths, product_name=None, note=""):
    data = load_learnings()
    existing_keys = {
        (record.get("source_path"), record.get("code"), record.get("field"), record.get("sku"), record.get("message"))
        for record in data.get("records", [])
    }

    added = []
    for path in paths:
        extracted = extract_processing_summary(path)
        for item in extracted["by_sku"] or extracted["by_code"]:
            record = {
                "learned_at": datetime.now().isoformat(timespec="seconds"),
                "product_name": product_name or guess_product_name(path),
                "source_path": str(path),
                "code": item.get("code", ""),
                "category": item.get("category", ""),
                "message": item.get("message", ""),
                "field": item.get("field", ""),
                "sku": item.get("sku", ""),
                "summary": extracted.get("summary", {}),
                "rule_hint": hint_for_field(item.get("field", ""), item.get("message", "")),
                "note": note
            }
            key = (record["source_path"], record["code"], record["field"], record["sku"], record["message"])
            if key in existing_keys:
                continue
            data.setdefault("records", []).append(record)
            existing_keys.add(key)
            added.append(record)

    save_learnings(data)
    write_learning_report(data)
    return added, LEARNINGS_PATH, REPORT_PATH


def guess_product_name(path):
    name = Path(path).stem
    for token in ["-processing-summary", "_processing-summary", "processing-summary"]:
        name = name.replace(token, "")
    return name


def hint_for_field(field, message):
    text = f"{field} {message}"
    for key, hint in FIELD_RULE_HINTS.items():
        if key.lower() in text.lower():
            return hint
    if "required but missing" in message.lower():
        return "条件必填缺失：把该字段加入同 product_type / 同模板的上传前自检清单，后续同类目提前补齐；不要默认套用到所有模板。"
    if "can't accept" in message.lower():
        return "有效值错误：回到 Valid Values / Data Definitions 查询允许值，禁止自由填写。"
    return "记录为待归纳规则；下次遇到同字段同错误时补充固定修复动作。"


def write_learning_report(data):
    records = data.get("records", [])
    by_code = Counter(record["code"] for record in records)
    by_field = Counter(clean_field(record["field"]) for record in records)
    grouped = defaultdict(list)
    for record in records:
        grouped[(record["code"], clean_field(record["field"]), record["message"])].append(record)

    lines = [
        "# 上传报错复盘库",
        "",
        f"- 总记录数：{len(records)}",
        f"- 错误码种类：{len(by_code)}",
        f"- 字段种类：{len(by_field)}",
        "",
        "## 高频错误码",
        ""
    ]
    for code, count in by_code.most_common():
        lines.append(f"- {code}: {count}")

    lines.extend(["", "## 高频字段", ""])
    for field, count in by_field.most_common():
        lines.append(f"- {field}: {count}")

    lines.extend(["", "## 规则沉淀", ""])
    for (code, field, message), items in sorted(grouped.items(), key=lambda item: (-len(item[1]), item[0])):
        sample = items[0]
        products = sorted(set(item["product_name"] for item in items if item.get("product_name")))
        skus = [item["sku"] for item in items if item.get("sku")]
        lines.extend([
            f"### {code} - {field}",
            "",
            f"- 出现次数：{len(items)}",
            f"- 商品：{', '.join(products) or '未记录'}",
            f"- 报错：{message}",
            f"- 自检规则：{sample.get('rule_hint')}",
            f"- 代表 SKU：{', '.join(skus[:8]) or '汇总级错误'}",
            ""
        ])

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)
    REPORT_PATH.write_text("\n".join(lines), encoding="utf-8")


def clean_field(field):
    if not field:
        return "未识别字段"
    return field.split("(")[0].strip()

import json
import re
from pathlib import Path

from .paths import CONFIG_DIR, OUTPUTS_DIR
from .listing_rules import validate_listing_row
from .workbook_io import REQUIRED_CORE_FIELDS, read_intake_rows


TEXT_FIELDS = [
    "title",
    "bullet_1",
    "bullet_2",
    "bullet_3",
    "bullet_4",
    "bullet_5",
    "description",
    "notes"
]
COPY_FIELDS = [
    "title",
    "bullet_1",
    "bullet_2",
    "bullet_3",
    "bullet_4",
    "bullet_5",
    "description",
]
CJK_RE = re.compile(r"[\u4e00-\u9fff]")

PRICE_FIELDS_ALLOWED_DRAFT_EMPTY = {"list_price", "haul_price"}
PARENT_OPTIONAL_CORE_FIELDS = {
    "list_price",
    "package_length_in",
    "package_width_in",
    "package_height_in",
    "package_weight_lb",
    "batteries_required",
    "dangerous_goods",
}


def _load_json(name):
    return json.loads((CONFIG_DIR / name).read_text(encoding="utf-8"))


def _text(value):
    if value is None:
        return ""
    return str(value).strip()


def _lower(value):
    return _text(value).lower()


def _number(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def _has_cjk(value):
    return bool(CJK_RE.search(_text(value)))


BAD_COPY_PATTERNS = [
    ("Generic Product", "标题或描述仍是通用兜底商品名。"),
    ("Designed for everyday use", "五点仍是通用兜底场景。"),
    ("home, travel, office", "描述仍是通用兜底场景。"),
    ("Includes the product shown in the listing photos", "五点没有提炼真实包装内容。"),
    ("wide range of everyday needs", "五点过泛，没有商品特征。"),
]


def _copy_quality_findings(row, row_number):
    findings = []
    title = _text(row.get("title"))
    copy = " ".join(_text(row.get(field)) for field in COPY_FIELDS)
    lower_copy = copy.lower()
    for phrase, reason in BAD_COPY_PATTERNS:
        if phrase.lower() in lower_copy:
            findings.append({
                "severity": "error",
                "row": row_number,
                "field": "copy",
                "message": f"文案质量不合格：{reason}",
                "fix": "重新按商品类型、1688标题和竞品标题生成标题/五点/描述，不允许使用兜底模板。"
            })
    if "videos" in title.lower():
        findings.append({
            "severity": "error",
            "row": row_number,
            "field": "title",
            "message": "标题里出现 VIDEOS，疑似从网页噪音误抓。",
            "fix": "删除网页噪音，改成真实商品标题。"
        })

    product_context = " ".join(_lower(row.get(field)) for field in [
        "product_name", "title", "item_type_keyword", "accessories", "notes"
    ])
    material = _lower(row.get("material"))
    category = _lower(row.get("category"))
    item_type = _lower(row.get("item_type_keyword"))
    if any(keyword in product_context for keyword in ["cat", "猫", "q-tip", "catnip"]) and (
        "cookbook" in item_type or "home" == category or material == "metal"
    ):
        findings.append({
            "severity": "error",
            "row": row_number,
            "field": "商品理解",
            "message": "猫玩具资料疑似被网页噪音污染，类目/关键词/材质不匹配。",
            "fix": "猫玩具应优先按 pet / cat-toys / Cotton and Felt 等宠物玩具方向重写。"
        })
    return findings


def _has_brand_in_copy(row):
    brand = _text(row.get("brand"))
    if not brand or brand.lower() == "generic":
        return False
    content = " ".join(_text(row.get(field)) for field in TEXT_FIELDS)
    return brand.lower() in content.lower()


def _is_parent_row(row):
    return _lower(row.get("parentage_level")) == "parent"


def _is_child_row(row):
    return _lower(row.get("parentage_level")) == "child"


def validate_intake(path):
    rows = read_intake_rows(path)
    haul_rules = _load_json("haul_rules.json")
    risky_words = _load_json("risky_words.json")
    findings = []

    if not rows:
        findings.append({
            "severity": "error",
            "row": "-",
            "field": "产品资料",
            "message": "资料表没有可用数据行。",
            "fix": "至少保留一行产品资料。"
        })
        return findings

    for row in rows:
        row_number = row.get("_row_number")
        route = _lower(row.get("route"))
        brand = _text(row.get("brand"))
        manufacturer = _text(row.get("manufacturer"))
        category = _lower(row.get("category"))

        is_parent = _is_parent_row(row)
        is_child = _is_child_row(row)

        for field in REQUIRED_CORE_FIELDS:
            if field in PRICE_FIELDS_ALLOWED_DRAFT_EMPTY or (is_parent and field in PARENT_OPTIONAL_CORE_FIELDS):
                continue
            if _text(row.get(field)) == "":
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": field,
                    "message": f"{field} 为空。",
                    "fix": "上传前先补齐这个核心字段。"
                })

        if not is_parent and _text(row.get("list_price")) == "":
            findings.append({
                "severity": "info",
                "row": row_number,
                "field": "list_price",
                "message": "List Price 留空，按当前流程等待人工填写。",
                "fix": "定价后填写 List Price；Haul Price 后续与 List Price 保持一致。"
            })
        if not is_parent and _text(row.get("haul_price")) == "":
            findings.append({
                "severity": "info",
                "row": row_number,
                "field": "haul_price",
                "message": "Haul Price 留空，按当前流程等待 List Price 确认。",
                "fix": "定价后填写为与 List Price 一致，或按类目策略调整。"
            })

        if "generic" in route:
            if brand.lower() != "generic":
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "brand",
                    "message": "Generic/Haul Generic 路线下 Brand 不是 Generic。",
                    "fix": "如果确定走无品牌路线，把 Brand 改为 Generic；否则改走品牌路线。"
                })
            if manufacturer and manufacturer.lower() != "generic":
                findings.append({
                    "severity": "warning",
                    "row": row_number,
                    "field": "manufacturer",
                    "message": "Generic 路线下 Manufacturer 与无品牌逻辑不一致。",
                    "fix": "建议 Manufacturer 也按 Generic 逻辑处理，并确认包装、文案没有品牌痕迹；图片字段默认不处理。"
                })

            copy = " ".join(_lower(row.get(field)) for field in TEXT_FIELDS)
            for hint in risky_words["generic_forbidden_brand_hints"]:
                if hint in copy:
                    findings.append({
                        "severity": "warning",
                        "row": row_number,
                        "field": "copy",
                        "message": f"Generic 文案里出现品牌痕迹提示词：{hint}",
                        "fix": "检查标题、五点、描述是否真的适合无品牌路线；图片字段默认不处理。"
                    })

        for field in COPY_FIELDS:
            if _has_cjk(row.get(field)):
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": field,
                    "message": f"{field} 中包含中文，不能直接进入 Amazon 文案。",
                    "fix": "改成自然的跨境英语表达，避免中文、供应商原文或机器直译残留。"
                })
        for field, message, fix in validate_listing_row(row):
            findings.append({
                "severity": "error",
                "row": row_number,
                "field": field,
                "message": message,
                "fix": fix
            })
        findings.extend(_copy_quality_findings(row, row_number))

        if is_parent:
            if _text(row.get("parent_sku")):
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "parent_sku",
                    "message": "Parent 行不应填写 parent_sku。",
                    "fix": "Parent 行只填写自己的 SKU、parentage_level=Parent 和 variation_theme。"
                })
            if not _text(row.get("variation_theme")):
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "variation_theme",
                    "message": "Parent 行缺少 variation_theme。",
                    "fix": "按实际差异填写 Color、Size 或 ColorSize。"
                })
        elif is_child:
            if not _text(row.get("parent_sku")):
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "parent_sku",
                    "message": "Child 行缺少 parent_sku。",
                    "fix": "填写对应 Parent SKU。"
                })
            if not _text(row.get("variation_theme")):
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "variation_theme",
                    "message": "Child 行缺少 variation_theme。",
                    "fix": "与 Parent 行保持一致，例如 Color。"
                })
        elif _text(row.get("parent_sku")) or _text(row.get("variation_theme")):
            findings.append({
                "severity": "warning",
                "row": row_number,
                "field": "variation",
                "message": "独立 SKU 行存在父子变体字段。",
                "fix": "如果不是父子变体路线，请清空 parent_sku、parentage_level 和 variation_theme。"
            })

        if route == "brand":
            if brand.lower() == "generic":
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "brand",
                    "message": "品牌路线下 Brand 仍然是 Generic。",
                    "fix": "填写真实品牌，并确保 Manufacturer、包装、文案统一；图片字段默认不处理。"
                })
            if manufacturer and brand and manufacturer.lower() != brand.lower():
                findings.append({
                    "severity": "warning",
                    "row": row_number,
                    "field": "manufacturer",
                    "message": "品牌路线下 Manufacturer 与 Brand 不一致。",
                    "fix": "确认这是有意设置；否则统一 Brand 和 Manufacturer。"
                })
            if not _has_brand_in_copy(row):
                findings.append({
                    "severity": "info",
                    "row": row_number,
                    "field": "copy",
                    "message": "品牌路线文案里没有出现 Brand。",
                    "fix": "确认包装、标题、五点、描述的品牌表达是否统一；图片字段默认不处理。"
                })

        if "haul" in route:
            for keyword in haul_rules["blocked_category_keywords"]:
                if keyword in category:
                    findings.append({
                        "severity": "error",
                        "row": row_number,
                        "field": "category",
                        "message": f"类目疑似不适合 Haul：{keyword}",
                        "fix": "先确认该类目是否可做 Haul，必要时改普通上架路线。"
                    })

            haul_price = _number(row.get("haul_price")) or _number(row.get("target_price"))
            ceiling = haul_rules["category_price_ceiling"].get(category, haul_rules["default_price_ceiling"])
            if haul_price is not None and haul_price > ceiling:
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": "haul_price",
                    "message": f"Haul 价格 {haul_price:g} 高于当前规则上限 {ceiling:g}。",
                    "fix": "调整价格或确认该类目真实 Haul 上限后更新 config/haul_rules.json。"
                })

        list_price = _number(row.get("list_price"))
        haul_price = _number(row.get("haul_price"))
        if list_price is not None and haul_price is not None and haul_price > list_price:
            findings.append({
                "severity": "error",
                "row": row_number,
                "field": "haul_price",
                "message": "Haul Price 高于 List Price。",
                "fix": "确认 List Price、Haul Price、Min/Max 价格逻辑一致。"
            })

        for field in ["package_length_in", "package_width_in", "package_height_in", "package_weight_lb"]:
            value = _number(row.get(field))
            if value is not None and value <= 0:
                findings.append({
                    "severity": "error",
                    "row": row_number,
                    "field": field,
                    "message": f"{field} 必须大于 0。",
                    "fix": "按美国站常用单位填写：尺寸 inches，重量 pounds。"
                })

        copy = " ".join(_lower(row.get(field)) for field in TEXT_FIELDS)
        for word in risky_words["claim_words"]:
            if word in copy:
                findings.append({
                    "severity": "warning",
                    "row": row_number,
                    "field": "copy",
                    "message": f"文案里出现高风险宣称词：{word}",
                    "fix": "弱化医疗功效、情绪干预和绝对化表达，改成客观场景描述。"
                })

        for keyword in risky_words["high_risk_categories"]:
            if keyword in category:
                findings.append({
                    "severity": "warning",
                    "row": row_number,
                    "field": "category",
                    "message": f"类目含高风险关键词：{keyword}",
                    "fix": "上传前确认是否涉及食品、皮肤外用、医疗功效或受限类目。"
                })

    return findings


def write_validation_report(intake_path, findings, output_path=None):
    intake_path = Path(intake_path)
    if output_path is None:
        output_path = OUTPUTS_DIR / f"{intake_path.stem}_自检报告.md"
    else:
        output_path = Path(output_path)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    error_count = sum(1 for item in findings if item["severity"] == "error")
    warning_count = sum(1 for item in findings if item["severity"] == "warning")
    info_count = sum(1 for item in findings if item["severity"] == "info")

    lines = [
        f"# {intake_path.name} 上传前自检报告",
        "",
        f"- Error: {error_count}",
        f"- Warning: {warning_count}",
        f"- Info: {info_count}",
        ""
    ]

    if not findings:
        lines.extend(["未发现明显问题。", ""])
    else:
        lines.extend(["## 问题清单", ""])
        for item in findings:
            lines.extend([
                f"### [{item['severity'].upper()}] 行 {item['row']} - {item['field']}",
                "",
                f"- 问题：{item['message']}",
                f"- 建议：{item['fix']}",
                ""
            ])

    output_path.write_text("\n".join(lines), encoding="utf-8")
    return output_path

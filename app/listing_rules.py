import json
import re
from collections import Counter

from .paths import CONFIG_DIR


GENERIC_KEYWORD_FIELDS = {"generic_keywords", "generic_keyword", "search_terms", "search_terms_1"}


def _load_rules():
    return json.loads((CONFIG_DIR / "listing_rules.json").read_text(encoding="utf-8"))


def _text(value):
    if value is None:
        return ""
    return str(value).strip()


def _char_len(value):
    return len(_text(value))


def _contains_term(text, term):
    haystack = _text(text).lower()
    needle = _text(term).lower()
    if not needle:
        return False
    if re.search(r"[a-z0-9]", needle):
        return re.search(rf"(?<![a-z0-9]){re.escape(needle)}(?![a-z0-9])", haystack) is not None
    return needle in haystack


def _contains_any_char(text, chars):
    text = _text(text)
    return [char for char in chars if char and char in text]


def _title_words(title):
    return re.findall(r"[A-Za-z][A-Za-z0-9'-]*", _text(title))


def _is_title_cased_word(word):
    return word.isupper() or word[:1].isupper()


def _paragraphs(value):
    return [part.strip() for part in re.split(r"\n\s*\n+", _text(value)) if part.strip()]


def validate_listing_row(row):
    rules = _load_rules()
    findings = []
    title = _text(row.get("title"))
    bullets = [_text(row.get(f"bullet_{idx}")) for idx in range(1, 6)]
    description = _text(row.get("description"))
    copy_fields = ["title", "bullet_1", "bullet_2", "bullet_3", "bullet_4", "bullet_5", "description"]
    all_copy = " ".join(_text(row.get(field)) for field in copy_fields)

    if '"' in all_copy or "“" in all_copy or "”" in all_copy:
        findings.append(("copy", "Listing 文案包含双引号。", "按纯文本规则删除双引号，改用自然句表达。"))

    non_us_units = re.findall(r"\b\d+(?:\.\d+)?\s*(?:cm|centimeter|centimeters|kg|kilogram|kilograms|g|gram|grams)\b|厘米|公斤|千克|克", all_copy, re.I)
    if non_us_units:
        findings.append(("copy", f"Listing 文案疑似残留非目标单位：{', '.join(non_us_units[:6])}。", "长度统一换算为 inch，重量统一换算为 lb。"))

    if title:
        title_length = _char_len(title)
        if title_length < rules["title_min_chars"]:
            findings.append(("title", f"标题长度为 {title_length} 字符，低于 {rules['title_min_chars']} 字符。", "扩展标题，优先补充核心关键词、长尾词、用途、场景和款式信息，保持自然可读。"))
        if title_length > rules["title_max_chars"]:
            findings.append(("title", f"标题长度为 {title_length} 字符，超过 {rules['title_max_chars']} 字符。", "压缩标题，保留核心关键词、长尾词、功能和场景。"))

        forbidden_chars = _contains_any_char(title, rules.get("title_forbidden_chars", []))
        if forbidden_chars:
            findings.append(("title", f"标题包含 Amazon 禁用特殊字符：{', '.join(forbidden_chars)}。", "删除这些字符；除明确属于品牌名外，标题不得使用 ! $ ? _ { } ^ ¬ ¦。"))

        decorative_chars = _contains_any_char(title, rules.get("title_decorative_chars", []))
        if decorative_chars:
            findings.append(("title", f"标题包含装饰性或高风险字符：{', '.join(decorative_chars)}。", "标题只使用自然英文、数字、空格、英文逗号、连字符或必要括号，避免装饰符号。"))

        cjk_punctuation = _contains_any_char(title, rules.get("title_cjk_punctuation", []))
        if cjk_punctuation:
            findings.append(("title", f"标题包含中文或全角标点：{', '.join(cjk_punctuation)}。", "改为英文半角标点，避免中文标点进入 Amazon 标题。"))

        repeated_punctuation = sorted(set(re.findall(r"([,./\\-])\1+", title)))
        if repeated_punctuation:
            findings.append(("title", f"标题包含重复标点：{', '.join(repeated_punctuation)}。", "删除重复标点，保持标题像正常商品名而不是装饰或关键词堆砌。"))

        if re.search(r"<[^>]+>|&[A-Za-z]+;", title):
            findings.append(("title", "标题包含 HTML 标签或实体。", "删除 HTML 标签、实体和网页残留噪音，只保留纯文本商品标题。"))

        if re.search(r"[\r\n\t]|\s{2,}", title):
            findings.append(("title", "标题包含换行、制表符或连续多空格。", "标题使用单行纯文本，并把连续空格压缩为一个空格。"))

        promotional_hits = [term for term in rules.get("title_promotional_terms", []) if _contains_term(title, term)]
        if promotional_hits:
            findings.append(("title", f"标题包含促销、平台背书、物流或售后类表达：{', '.join(promotional_hits[:8])}。", "标题只写商品本身，不写 Free Shipping、Sale、Best Seller、Warranty、Fast Delivery 等非商品信息。"))

        high_risk_claim_hits = [term for term in rules.get("title_high_risk_claim_terms", []) if _contains_term(title, term)]
        if high_risk_claim_hits:
            findings.append(("title", f"标题包含医疗、护理、清洁、防护或结果型功效表达：{', '.join(high_risk_claim_hits[:8])}。", "标题不要写治疗、预防、牙齿清洁、杀菌、防护、健康改善等结果；改写为商品类型、款式、对象和使用场景。"))

        ingestion_hits = [term for term in rules.get("title_ingestion_and_absolute_terms", []) if _contains_term(title, term)]
        if ingestion_hits:
            findings.append(("title", f"标题包含摄入暗示或绝对化成分宣称：{', '.join(ingestion_hits[:8])}。", "非食品/非补充剂标题避免 Flavor、Edible、Food Grade、Natural、Safe 等表达；气味用 Scent，材质和成分移到字段、五点或描述中。"))

    lowercase_words = set(rules["title_lowercase_words"])
    bad_case_words = [
        word for word in _title_words(title)
        if word.lower() not in lowercase_words and not _is_title_cased_word(word)
    ]
    if bad_case_words:
        findings.append(("title", f"标题存在未按 Title Case 处理的词：{', '.join(bad_case_words[:6])}。", "除介词、冠词、连词外，标题单词首字母大写。"))

    counted_words = [
        word.lower() for word in _title_words(title)
        if word.lower() not in lowercase_words
    ]
    repeated = sorted(word for word, count in Counter(counted_words).items() if count > 2)
    if repeated:
        findings.append(("title", f"标题中非介词单词重复超过两次：{', '.join(repeated[:6])}。", "减少重复词，避免关键词堆砌。"))

    material_terms = set(rules["title_material_terms"])
    title_material_hits = [term for term in material_terms if _contains_term(title, term)]
    if title_material_hits:
        findings.append(("title", f"标题包含敏感材质或合规宣称：{', '.join(title_material_hits[:6])}。", "标题允许资料可证明的普通材质词，但不要写 Food Grade、BPA Free、Waterproof、Eco Friendly 等材质安全、等级、环保或不含某物宣称。"))

    for idx, bullet in enumerate(bullets, start=1):
        if not bullet:
            continue
        length = _char_len(bullet)
        if length < rules["bullet_min_chars"] or length > rules["bullet_max_chars"]:
            findings.append((f"bullet_{idx}", f"Bullet {idx} 长度为 {length} 字符，不在 {rules['bullet_min_chars']}-{rules['bullet_max_chars']} 字符范围。", "重写该卖点，保持关键词开头并覆盖真实产品信息。"))
        if ";" in bullet:
            findings.append((f"bullet_{idx}", f"Bullet {idx} 包含分号。", "按新规则删除分号，改用逗号或句号。"))
        if not bullet.endswith("."):
            findings.append((f"bullet_{idx}", f"Bullet {idx} 未以句号结尾。", "每条 Bullet 末尾使用英文句号。"))
        prefix = bullet.split(":", 1)[0] if ":" in bullet else ""
        prefix_words = _title_words(prefix)
        if ":" not in bullet or not (3 <= len(prefix_words) <= 5):
            findings.append((f"bullet_{idx}", f"Bullet {idx} 缺少 3-5 个词的关键词开头加冒号格式。", "格式示例：Travel Ready Design: 后接完整卖点句。"))

    prefixes = [bullet.split(":", 1)[0].strip().lower() for bullet in bullets if ":" in bullet]
    if len(prefixes) != len(set(prefixes)):
        findings.append(("bullet_points", "Bullet 开头关键词存在重复。", "每条 Bullet 使用不同的 3-5 词卖点开头。"))

    if description:
        desc_len = _char_len(description)
        if desc_len < rules["description_min_chars"] or desc_len > rules["description_max_chars"]:
            findings.append(("description", f"Description 长度为 {desc_len} 字符，不在 {rules['description_min_chars']}-{rules['description_max_chars']} 字符范围。", "按 4 段商务英语描述重写，控制总字符数。"))
        paragraphs = _paragraphs(description)
        if len(paragraphs) != 4:
            findings.append(("description", f"Description 当前为 {len(paragraphs)} 段，不是 4 段。", "按新规则改为 4 段独立描述。"))
        bad_paragraphs = []
        for index, paragraph in enumerate(paragraphs, start=1):
            match = re.search(r"[A-Za-z]", paragraph)
            if match and not paragraph[match.start()].isupper():
                bad_paragraphs.append(str(index))
        if bad_paragraphs:
            findings.append(("description", f"Description 第 {', '.join(bad_paragraphs)} 段首个英文单词未大写。", "每段首单词首字母大写。"))

    for field, value in row.items():
        if field in GENERIC_KEYWORD_FIELDS and _text(value):
            length = _char_len(value)
            if length > rules["generic_keywords_max_chars"]:
                findings.append((field, f"Generic Keywords 长度为 {length} 字符，超过 {rules['generic_keywords_max_chars']} 字符。", "压缩长尾词，删除重复标题关键词。"))

    restricted_hits = [term for term in rules["restricted_copy_terms"] if _contains_term(all_copy, term)]
    if restricted_hits:
        findings.append(("copy", f"Listing 文案包含高风险或禁用词：{', '.join(restricted_hits[:12])}。", "删除禁用词、侵权品牌词、儿童/孕妇/医疗/杀菌/防护/绝对化相关表达。"))

    return findings

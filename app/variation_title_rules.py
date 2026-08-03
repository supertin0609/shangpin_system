import json
import re

from .paths import CONFIG_DIR


PARENT_COLOR_PHRASE = "Multiple Colors Available"
PARENT_STYLE_PHRASE = "Multiple Styles Available"
PARENT_ONLY_TITLE_PHRASES = (PARENT_COLOR_PHRASE, PARENT_STYLE_PHRASE)

PACK_UNIT_RE = re.compile(
    r"\b\d+(?:\.\d+)?\s*(?:pcs?|pieces?|pack|packs|sets?|pairs?|count|ct)\b",
    re.I,
)
TITLE_PACK_PREFIX_RE = re.compile(
    r"^\s*((?:\d+(?:\.\d+)?\s*(?:pcs?|pieces?|pack|packs|sets?|pairs?|count|ct)\b[\s,+/&-]*)+)",
    re.I,
)


def _text(value):
    return str(value or "").strip()


def _title_max_chars():
    rules_path = CONFIG_DIR / "listing_rules.json"
    if not rules_path.exists():
        return 125
    try:
        return int(json.loads(rules_path.read_text(encoding="utf-8")).get("title_max_chars") or 125)
    except (OSError, ValueError, TypeError, json.JSONDecodeError):
        return 125


def _clean_spaces(value):
    return re.sub(r"\s+", " ", _text(value)).strip(" ,")


def _normalize_pack_text(value):
    text = _clean_spaces(value)
    if not text:
        return ""
    text = re.sub(r"(?<=\d)(?=[A-Za-z])", " ", text)
    text = re.sub(r"\bpc\b", "Pcs", text, flags=re.I)
    text = re.sub(r"\bpcs\b", "Pcs", text, flags=re.I)
    text = re.sub(r"\bpieces\b", "Pcs", text, flags=re.I)
    text = re.sub(r"\bpacks\b", "Pack", text, flags=re.I)
    text = re.sub(r"\bpack\b", "Pack", text, flags=re.I)
    text = re.sub(r"\bsets\b", "Set", text, flags=re.I)
    text = re.sub(r"\bset\b", "Set", text, flags=re.I)
    text = re.sub(r"\bpairs\b", "Pair", text, flags=re.I)
    text = re.sub(r"\bpair\b", "Pair", text, flags=re.I)
    text = re.sub(r"\bcounts?\b|\bct\b", "Count", text, flags=re.I)
    return _clean_spaces(text)


def package_prefix_for_row(row):
    set_count = _text(row.get("set_count"))
    if set_count:
        if PACK_UNIT_RE.search(set_count):
            return _normalize_pack_text(set_count)
        try:
            number = float(set_count)
        except ValueError:
            number = 0
        if number > 1:
            count_text = str(int(number)) if number.is_integer() else str(number).rstrip("0").rstrip(".")
            return f"{count_text} Pack"

    title = _text(row.get("title"))
    match = TITLE_PACK_PREFIX_RE.match(title)
    if match:
        return _normalize_pack_text(match.group(1))
    return ""


def strip_package_prefix(title):
    text = _clean_spaces(title)
    match = TITLE_PACK_PREFIX_RE.match(text)
    if not match:
        return text
    return _clean_spaces(text[match.end():])


def _split_title_segments(title):
    return [_clean_spaces(part) for part in re.split(r"\s*,\s*", _text(title)) if _clean_spaces(part)]


def _contains_phrase(text, phrase):
    return re.search(rf"(?<![A-Za-z0-9]){re.escape(_text(phrase))}(?![A-Za-z0-9])", _text(text), re.I) is not None


def _remove_parent_only_phrases(title):
    text = _text(title)
    for phrase in PARENT_ONLY_TITLE_PHRASES:
        text = re.sub(rf"\s*,?\s*{re.escape(phrase)}\s*,?", ", ", text, flags=re.I)
    return _clean_spaces(text)


def _attr_values(rows, field):
    values = []
    seen = set()
    for row in rows:
        value = _clean_spaces(row.get(field))
        if not value:
            continue
        key = value.lower()
        if key in seen:
            continue
        seen.add(key)
        values.append(value)
    return values


def _is_attr_segment(segment, attr_values):
    lowered = _text(segment).lower()
    return any(lowered == value.lower() for value in attr_values)


def _clean_title_base(title, attr_values):
    title = strip_package_prefix(_remove_parent_only_phrases(title))
    segments = [
        segment
        for segment in _split_title_segments(title)
        if not _is_attr_segment(segment, attr_values)
    ]
    return segments


def _best_base_segments(child_rows, attr_values):
    candidates = []
    for row in child_rows:
        segments = _clean_title_base(row.get("title"), attr_values)
        if segments:
            candidates.append(segments)
    if not candidates:
        product_name = next((_text(row.get("product_name")) for row in child_rows if _text(row.get("product_name"))), "")
        return [product_name] if product_name else ["Product"]
    return max(candidates, key=lambda item: len(", ".join(item)))


def _compose_title(prefix, base_segments, suffix_segments, max_chars=None):
    max_chars = max_chars or _title_max_chars()
    prefix = _clean_spaces(prefix)
    base_segments = [_clean_spaces(segment) for segment in base_segments if _clean_spaces(segment)]
    suffix_segments = [_clean_spaces(segment) for segment in suffix_segments if _clean_spaces(segment)]

    def render(segments):
        title = ", ".join(segments)
        if prefix:
            title = f"{prefix} {title}" if title else prefix
        return _clean_spaces(title)

    working_base = list(base_segments)
    title = render(working_base + suffix_segments)
    while len(title) > max_chars and working_base:
        words = working_base[-1].split()
        if len(words) > 3:
            working_base[-1] = " ".join(words[:-1])
        elif len(working_base) > 1:
            working_base.pop()
        else:
            break
        title = render(working_base + suffix_segments)

    if len(title) <= max_chars:
        return title

    suffix_text = ", ".join(suffix_segments)
    reserved = len(prefix) + len(suffix_text) + (1 if prefix else 0) + (2 if suffix_text else 0)
    available = max(max_chars - reserved, 20)
    base_text = ", ".join(working_base)
    words = base_text.split()
    while words and len(" ".join(words)) > available:
        words.pop()
    trimmed_base = " ".join(words) or base_text[:available].rstrip(" ,")
    return render([trimmed_base] + suffix_segments)[:max_chars].rstrip(" ,")


def _common_package_prefix(child_rows):
    prefixes = [package_prefix_for_row(row) for row in child_rows]
    prefixes = [prefix for prefix in prefixes if prefix]
    if not prefixes:
        return ""
    first = prefixes[0]
    if all(prefix.lower() == first.lower() for prefix in prefixes):
        return first
    return ""


def _parent_suffixes(child_rows):
    colors = _attr_values(child_rows, "color")
    sizes = _attr_values(child_rows, "size")
    suffixes = []
    if len(colors) > 1:
        suffixes.append(PARENT_COLOR_PHRASE)
    if len(sizes) > 1:
        suffixes.append(PARENT_STYLE_PHRASE)
    return suffixes


def build_parent_title(child_rows, fallback_row=None):
    child_rows = [row for row in child_rows if _text(row.get("title")) or _text(row.get("product_name"))]
    if not child_rows:
        return _text((fallback_row or {}).get("title"))
    attr_values = _attr_values(child_rows, "color") + _attr_values(child_rows, "size")
    prefix = _common_package_prefix(child_rows) or package_prefix_for_row(fallback_row or {})
    base_segments = _best_base_segments(child_rows, attr_values)
    return _compose_title(prefix, base_segments, _parent_suffixes(child_rows))


def build_child_title(row, child_rows):
    attr_values = _attr_values(child_rows, "color") + _attr_values(child_rows, "size")
    base_segments = _best_base_segments(child_rows, attr_values)
    row_attrs = []
    for field in ("color", "size"):
        value = _clean_spaces(row.get(field))
        if value and not _is_attr_segment(value, row_attrs):
            row_attrs.append(value)
    return _compose_title(package_prefix_for_row(row), base_segments, row_attrs)


def apply_variation_title_rules(rows):
    prepared = [dict(row) for row in rows]
    parent_rows = [row for row in prepared if _text(row.get("parentage_level")).lower() == "parent"]
    child_rows = [row for row in prepared if _text(row.get("parentage_level")).lower() == "child"]
    if not child_rows:
        return prepared

    for parent in parent_rows:
        parent["title"] = build_parent_title(child_rows, parent)
    for child in child_rows:
        child["title"] = build_child_title(child, child_rows)
    return prepared


def has_parent_only_phrase(title):
    return any(_contains_phrase(title, phrase) for phrase in PARENT_ONLY_TITLE_PHRASES)


def title_starts_with_package_prefix(title):
    return TITLE_PACK_PREFIX_RE.match(_text(title)) is not None

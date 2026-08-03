# Agent Instructions

This repository has project-specific Amazon template filling rules. Do not fill Amazon upload templates from general marketplace knowledge alone.

Before any task involving Amazon template filling, auto-fill, listing drafts, SKU route decisions, processing-summary fixes, or upload self-checks, read these files in this order:

1. `PROJECT_RULES.md`
2. `docs/amazon_template_fill_workflow.md`
3. `data/reference_docs/亚马逊上传表格_通用自检资料.md`
4. `app/template_writer.py`
5. `app/template_validator.py`
6. `app/success_rule_defaults.py`

Treat `app/template_writer.py` and `app/template_validator.py` as the source of truth for what is actually written and checked. Treat docs as workflow and policy guidance. If a task asks for current rules, inspect the code again instead of answering from memory.

When changing any persistent Amazon template filling default, validation rule, or reporting behavior, update this `AGENTS.md` summary in the same change so new Codex tasks do not start from stale defaults.

Current high-level defaults:

- Default route is `Haul Generic Variation` parent/child variation unless the user explicitly selects single-link or set-bundle.
- Fill `List Price` and Haul/BZR `our_price`; leave `minimum_seller_allowed_price` blank by default and do not default-fill `maximum_seller_allowed_price`.
- `Skip Offer` should remain blank.
- New item condition is `New`.
- Generic route uses `Brand = Generic` and `Manufacturer = Generic`.
- Dimensions use `Inches`; package weight uses `Pounds`.
- Ordinary non-battery / non-dangerous goods products use `batteries_required = No`, `batteries_included = No`, and `supplier_declared_dg_hz_regulation = Not Applicable`.
- For `PLANTER`, fill conditionally required `Model Name`, `Special Features`, `Mounting Type`, and `Product Compliance Certificate`; clear non-applicable `item_length_width_height` fields when Amazon ignores them.
- For `PET_TOY`, fill conditionally required `Subject Character`; for ordinary cat toys use the product subject such as `Cat`.
- Controlled enum fields must match the template `Valid Values` exactly, including case and slash formatting. Check fields such as `variation_theme#1.name`, `parentage_level`, and relationship type against the current template before delivery; for example, a PET_TOY template may require `COLOR`, not `Color`.
- Title self-checks now enforce Amazon-facing character and structure risks in addition to the project 100-125 character target: block prohibited/decorative characters, repeated punctuation, HTML/noise, promotional or logistics claims, medical/cleaning/protection result claims, ingestion wording such as `Flavor`, absolute ingredient claims such as `Natural`, and sensitive material/compliance claims such as `Food Grade`, `BPA Free`, or `Eco Friendly` in titles. Use `Scent` for non-food scent attributes. Ordinary verifiable materials such as `Velvet`, `Plastic`, or `Silicone` may appear in titles when useful.
- Always run the project template self-check before delivery when a filled template is produced, but do not generate or send self-check report files unless explicitly requested.

If these instructions conflict with a direct user instruction in the current conversation, follow the user instruction, but explicitly call out the deviation from the project default.

"""Write product rows into the country-specific TikTok Shop batch template.

The template lives in assets-{id|ph|th}/batch-product-source.xlsx and has the
following sheets: Template, Instruction, Image, Example, HiddenStyle,
HiddenAttr, SpecialProductListingType, TemplateConfig, Category, Brand,
ShippingInsurance, Condition, and (ID/TH only) Unavailable categories /
Kategori yang tidak tersedia.

This module only writes to the 'Template' sheet starting at row 2; the rest
of the workbook is left untouched so that TikTok's validator still accepts
the file.

ROW STRATEGY (v3 — TikTok-compatible):
    The EasyBoss source table has one row per (color × size) SKU. We expand
    each SKU into its own TikTok listing row (each row = one variant of one
    product). This matches TikTok's official Example sheet where every
    (color, size) combination is a separate listing.

    Common fields (description, size_chart, parcel, brand, category, images,
    delivery, cod, pre_order_time) are shared across all rows of the same
    product.

    `output_copies` duplicates each variant N times with a unique random
    suffix in the title and seller_sku, so duplicate listings can be
    detected and removed later.

    When `fill_sizes_enabled` is on, the writer falls back to the standard
    sizes list ONLY when the source product has no variation-2 values
    (single-SKU products). Otherwise each row keeps its own size value.

PER-COUNTRY BEHAVIOUR (v3.3 — multi-country):
    TIKTOK_COLUMNS is no longer a module-level constant. Use
    `get_columns_for(country)` which reads the first row of the Template
    sheet of the country-specific template. This handles:
      - ID: 40 cols incl. minimum_order_quantity (col 28) — Unique to ID
      - PH/TH: 39 cols
      - pre_order_time dropped for PH (legacy compat); ID/TH keep it

    Country-specific product_property (Neckline, Season, ...) is resolved
    via `presets.by_country(country).property_preferred_defaults`. Only ID
    sets these (Cowl Neck / Semua musim). PH/TH let HiddenAttr decide.

    ID-only category translation (EasyBoss Indonesian path → official
    Indonesian name) lives in presets.by_country('ID').category_translation.
"""
from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

import openpyxl

from .config import get_assets_dir, get_settings
from .presets import by_country
from .source_reader import Product, Variant


TEMPLATE_FILENAME = "batch-product-source.xlsx"


# ---------------------------------------------------------------------------
# Country-aware column resolution
# ---------------------------------------------------------------------------
_COLUMNS_CACHE: dict[str, list[str]] = {}


def get_columns_for(country: str) -> list[str]:
    """Return the Template-sheet column headers for the given country.

    Reads the first row of the country-specific template (e.g. assets-id/
    batch-product-source.xlsx for ID). Result is cached per-country.

    Note: we deliberately do NOT use openpyxl's read_only=True here — that
    mode can underestimate max_column for sheets whose last few cells
    contain only a header string (no formulas / data), which silently
    drops product_property/100403 from the ID template.
    """
    if country in _COLUMNS_CACHE:
        return _COLUMNS_CACHE[country]
    template = get_assets_dir(country) / TEMPLATE_FILENAME
    wb = openpyxl.load_workbook(template)
    try:
        if "Template" not in wb.sheetnames:
            raise ValueError(f"模板文件缺少 'Template' sheet：{template}")
        ws = wb["Template"]
        cols: list[str] = []
        max_col = ws.max_column
        for c_idx in range(1, max_col + 1):
            v = ws.cell(row=1, column=c_idx).value
            if v is None:
                continue
            s = str(v).strip()
            if s:
                cols.append(s)
    finally:
        wb.close()
    _COLUMNS_CACHE[country] = cols
    return cols


def template_path_for(country: str) -> Path:
    return get_assets_dir(country) / TEMPLATE_FILENAME


# Back-compat shim: legacy code that imports TIKTOK_COLUMNS still works
# when the GUI hasn't switched country yet (defaults to PH).
def _get_columns_legacy() -> list[str]:
    return get_columns_for("PH")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _kg_to_grams(v: Any) -> Any:
    if v is None or v == "":
        return None
    try:
        f = float(v)
        return int(round(f * 1000))
    except (TypeError, ValueError):
        return v


def _clean_str(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip()


def _to_number(v: Any) -> Any:
    """Coerce a value to int/float so Excel stores it as a number."""
    if v is None or v == "":
        return None
    if isinstance(v, bool):
        return int(v)
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip()
    if not s:
        return None
    for ch in [",", " ", "₱", "$", "￥", "¥", "PHP", "php", "kg", "KG", "g", "G", "cm", "CM"]:
        s = s.replace(ch, "")
    try:
        f = float(s)
        if f.is_integer():
            return int(f)
        return f
    except ValueError:
        return v


# Variation name translations: Chinese → English (TikTok backend expects English).
VAR_NAME_TRANSLATIONS: dict[str, str] = {
    # Chinese
    "颜色": "Color",
    "顏色": "Color",
    "尺码": "Size",
    "尺碼": "Size",
    "尺寸": "Size",
    "规格": "Specification",
    # English-already is a pass-through (we just normalize capitalization)
    "color": "Color",
    "colour": "Color",
    "size": "Size",
}


def _translate_var_name(name: str) -> str:
    if not name:
        return ""
    key = _clean_str(name)
    if key in VAR_NAME_TRANSLATIONS:
        return VAR_NAME_TRANSLATIONS[key]
    return key


# ID-only: 中文/英文颜色 → 印尼语颜色（EasyBoss 印尼模板要求印尼语）
_ID_COLOR_TRANSLATION: dict[str, str] = {
    # Chinese
    "白色": "Putih", "黑色": "Hitam", "灰色": "Abu-abu", "红色": "Merah",
    "蓝色": "Biru", "绿色": "Hijau", "黄色": "Kuning", "粉色": "Merah Muda",
    "紫色": "Ungu", "橙色": "Oranye", "棕色": "Coklat", "米色": "Krem",
    "卡其": "Khaki", "咖啡": "Coklat", "军绿": "Hijau Tua",
    "深蓝": "Biru Tua", "浅蓝": "Biru Muda",
    "花色": "Beraneka Ragam", "条纹": "Garis-garis", "格子": "Kotak-kotak",
    # English → Indonesian
    "white": "Putih", "black": "Hitam", "gray": "Abu-abu", "grey": "Abu-abu",
    "red": "Merah", "blue": "Biru", "green": "Hijau", "yellow": "Kuning",
    "pink": "Merah Muda", "purple": "Ungu", "orange": "Oranye",
    "brown": "Coklat", "beige": "Krem", "khaki": "Khaki",
    "navy": "Biru Tua", "sky blue": "Biru Muda", "army green": "Hijau Tua",
    "multicolor": "Beraneka Ragam", "striped": "Garis-garis",
    "plaid": "Kotak-kotak",
}


def _translate_color_id(value: str) -> str:
    """ID-only: 中文/英文颜色 → 印尼语颜色值。PH/TH 走英文，不调用此函数。"""
    if not value:
        return ""
    key = _clean_str(value)
    # Match case-insensitively — EasyBoss sources often use title-case
    # ("White") while our lookup table uses lowercase keys.
    if key in _ID_COLOR_TRANSLATION:
        return _ID_COLOR_TRANSLATION[key]
    lower = key.lower()
    if lower in _ID_COLOR_TRANSLATION:
        return _ID_COLOR_TRANSLATION[lower]
    # Try lowercased first word for things like "Sky Blue" → "sky blue"
    return key


# A single output row: a dict {col_name: value}
OutputRow = dict[str, Any]


# ---------------------------------------------------------------------------
# ID-only: property defaults from HiddenAttr
# ---------------------------------------------------------------------------
# 9 对 HiddenAttr pair (C1-C18) → Template C29-C37 (PH) / C30-C38 (TH) / C31-C39 (ID)
# 后 1 列 (100403) 没有 HiddenAttr pair → 留空。
# HiddenAttr 列对 → prop_id:
#   pair 0 (C1/C2)   → 100157 Material
#   pair 1 (C3/C4)   → 100198 Pattern
#   pair 2 (C5/C6)   → 100393 Neckline
#   pair 3 (C7/C8)   → 100395 Sleeve
#   pair 4 (C9/C10)  → 100397 Season
#   pair 5 (C11/C12) → 100398 Style
#   pair 6 (C13/C14) → 100399 Fit type
#   pair 7 (C15/C16) → 100400 Stretch
#   pair 8 (C17/C18) → 100401 Care
_ID_PROPERTY_PAIR_INDEX: list[tuple[int, int, str]] = [
    (1, 2, "product_property/100157"),
    (3, 4, "product_property/100198"),
    (5, 6, "product_property/100393"),
    (7, 8, "product_property/100395"),
    (9, 10, "product_property/100397"),
    (11, 12, "product_property/100398"),
    (13, 14, "product_property/100399"),
    (15, 16, "product_property/100400"),
    (17, 18, "product_property/100401"),
]


def _get_property_fallbacks(country: str, category: str) -> dict[str, str]:
    """ID-only: read HiddenAttr to find the first valid value per prop_id.

    Returns dict {prop_id: value} for all 9 product_property cols. PH/TH
    return {} — they don't use the HiddenAttr sheet.
    """
    if country != "ID":
        return {}
    try:
        # NOTE: do NOT use read_only=True here — openpyxl in read-only mode
        # can underestimate max_row for sheets with sparse rows, which
        # silently truncates the HiddenAttr table and leaves the writer
        # without fallback values for product_property.
        wb = openpyxl.load_workbook(
            get_assets_dir("ID") / TEMPLATE_FILENAME,
            data_only=True,
        )
    except (OSError, KeyError):
        return {}
    try:
        if "HiddenAttr" not in wb.sheetnames:
            return {}
        ha = wb["HiddenAttr"]
        # Build fallback: first value per prop_id whose category matches
        # `category` (substring match on either side, to tolerate minor
        # whitespace / separator differences). Preserve Excel row order so
        # the chosen value is the template's recommended default (R2-style),
        # not alphabetical-first.
        by_prop: dict[str, str] = {}
        cat_lower = (category or "").strip().lower()
        max_row = ha.max_row
        for cat_col, val_col, prop_id in _ID_PROPERTY_PAIR_INDEX:
            if prop_id in by_prop:
                continue
            for r in range(2, max_row + 1):
                cat_v = ha.cell(row=r, column=cat_col).value
                val_v = ha.cell(row=r, column=val_col).value
                if not cat_v or not val_v:
                    continue
                cat_v_l = str(cat_v).lower()
                if cat_lower in cat_v_l or cat_v_l in cat_lower:
                    by_prop[prop_id] = str(val_v).strip()
                    break
        return by_prop
    finally:
        wb.close()


def _apply_category_translation(country: str, category: str) -> str:
    """ID-only: map EasyBoss Indonesian category path to official Indonesian name."""
    if country != "ID" or not category:
        return category
    preset = by_country("ID")
    table = preset.category_translation
    if not table:
        return category
    # Exact match first
    if category in table:
        return table[category]
    # Substring fallback (works for both directions of path separators)
    cat_lower = category.lower()
    for k, v in table.items():
        if k.lower() == cat_lower:
            return v
    return category


# ---------------------------------------------------------------------------
# Common field resolution
# ---------------------------------------------------------------------------
def _resolve_common_fields(
    product: Product,
    settings: dict[str, Any],
    country: str,
) -> dict[str, Any]:
    """Resolve fields that are shared across all rows of a product."""
    # Brand: prefer settings override, else product.brand, else fallback
    brand_fallback = by_country(country).settings.get("brand_value") or "No brand"
    brand = (
        settings["brand_value"]
        if settings.get("brand_enabled") and settings.get("brand_value")
        else (product.brand or "")
    )
    if not brand:
        brand = brand_fallback

    price = (
        settings["price_value"]
        if settings.get("price_enabled")
        else None
    )
    quantity = (
        settings["quantity_value"]
        if settings.get("quantity_enabled")
        else None
    )
    cod = settings["cod_value"] if settings.get("cod_enabled") else ""
    category = (
        settings["category_value"]
        if settings.get("category_enabled")
        else product.category
    )
    category = _apply_category_translation(country, category)
    description = (
        settings["description_value"]
        if settings.get("description_enabled")
        else product.description
    )
    size_chart = (
        settings["size_chart_value"]
        if settings.get("size_chart_enabled")
        else product.size_chart
    )

    if settings.get("parcel_enabled"):
        weight = settings.get("parcel_weight_value")
        length = settings.get("parcel_length_value")
        width = settings.get("parcel_width_value")
        height = settings.get("parcel_height_value")
    else:
        weight = _kg_to_grams(product.parcel_weight_kg)
        length = product.parcel_length
        width = product.parcel_width
        height = product.parcel_height

    common: dict[str, Any] = {
        "brand": _clean_str(brand),
        "category": _clean_str(category),
        "description": _clean_str(description),
        "size_chart": _clean_str(size_chart),
        "weight": weight,
        "length": length,
        "width": width,
        "height": height,
        "cod": cod,
        "quantity": quantity,
        "price_override": price,
        "delivery": _clean_str(settings.get("delivery_value", "")),
    }

    # ID-only: minimum_order_quantity + shipping_insurance
    if country == "ID":
        common["minimum_order_quantity"] = settings.get(
            "minimum_order_quantity_value", 1
        )
        if settings.get("shipping_insurance_enabled"):
            common["shipping_insurance"] = settings.get("shipping_insurance_value", 0)
        else:
            common["shipping_insurance"] = ""

    return common


# ---------------------------------------------------------------------------
# Row builders
# ---------------------------------------------------------------------------
def _build_row_for_variant(
    product: Product,
    variant: Variant,
    settings: dict[str, Any],
    common: dict[str, Any],
    country: str,
    copy_idx: int = 0,
    copy_suffix: str = "",
    apply_suffix_to_first: bool = False,
    property_fallbacks: dict[str, str] | None = None,
) -> OutputRow:
    """Build a single TikTok row for one (variant, copy) combination."""
    use_suffix = (copy_idx > 0) or apply_suffix_to_first

    # Title
    title_prefix = settings["title_prefix"] if settings.get("title_prefix_enabled") else ""
    suffix = copy_suffix if use_suffix else ""
    title = f"{title_prefix}{product.product_name}{suffix}".strip()

    # Variant-level values
    var1_name = _translate_var_name(product.var1_name or "颜色")
    var1_value_raw = _clean_str(variant.var1)
    var1_value = _translate_color_id(var1_value_raw) if country == "ID" else var1_value_raw
    var2_name = _translate_var_name(product.var2_name or "尺码")

    if variant.var2:
        var2_value = _clean_str(variant.var2)
    elif settings.get("fill_sizes_enabled"):
        std = settings.get("standard_sizes") or "S,M,L,XL,2XL,3XL"
        var2_value = std
    else:
        var2_value = ""

    # TH-only: empty var1_value → fallback to default_color_value
    if country == "TH" and not var1_value:
        if settings.get("default_color_enabled"):
            var1_value = settings.get("default_color_value", "As Picture")

    # Images
    var1_image = _clean_str(variant.sku_image)
    images = list(product.images[:9])
    while len(images) < 9:
        images.append("")
    images = [_clean_str(i) for i in images]

    # Price / quantity
    # v3.3.4: price_enabled/quantity_enabled 勾选即覆盖每行(GUI 承诺的语义);
    # 未勾选时才透传源表值,源表也缺失则用 override 兜底(None)。
    price_value = (
        _to_number(common["price_override"])
        if common["price_override"] is not None
        else (_to_number(variant.price) if variant.price not in (None, "") else None)
    )
    quantity_value = (
        _to_number(common["quantity"])
        if common["quantity"] is not None
        else (_to_number(variant.stock) if variant.stock not in (None, "") else None)
    )

    # Seller SKU
    base_sku = _clean_str(variant.platform_sku) or product.master_sku()
    seller_sku = base_sku + copy_suffix if copy_suffix and use_suffix else base_sku

    # Build row using country-specific columns
    cols = get_columns_for(country)
    row: OutputRow = {col: "" for col in cols}
    row["category"] = common["category"]
    row["brand"] = common["brand"]
    row["product_name"] = title
    row["product_description"] = common["description"]
    row["main_image"] = var1_image if var1_image else (images[0] if images else "")
    for i, img in enumerate(images[:9]):
        col = "main_image" if i == 0 else f"image_{i + 1}"
        if i == 0 and var1_image:
            continue
        row[col] = img
    row["property_name_1"] = var1_name
    row["property_value_1"] = var1_value
    row["property_1_image"] = var1_image
    row["property_name_2"] = var2_name
    row["property_value_2"] = var2_value
    row["parcel_weight"] = _to_number(common["weight"])
    row["parcel_length"] = _to_number(common["length"])
    row["parcel_width"] = _to_number(common["width"])
    row["parcel_height"] = _to_number(common["height"])
    row["delivery"] = common["delivery"]
    row["price"] = price_value
    row["quantity"] = quantity_value
    row["seller_sku"] = seller_sku
    row["size_chart"] = common["size_chart"]
    row["cod"] = common["cod"]

    # ID-only: minimum_order_quantity
    if country == "ID" and "minimum_order_quantity" in cols:
        row["minimum_order_quantity"] = common.get("minimum_order_quantity", 1)

    # ID-only: pre_order_time / shipping_insurance (only present in ID template)
    if country == "ID":
        if "pre_order_time" in cols:
            row["pre_order_time"] = settings.get("pre_order_time_value", "")
        if "shipping_insurance" in cols:
            row["shipping_insurance"] = common.get("shipping_insurance", "")

    # product_property: ID gets HiddenAttr fallbacks + PREFERRED; PH/TH empty
    if country == "ID" and property_fallbacks:
        for prop_id, value in property_fallbacks.items():
            if prop_id in cols:
                row[prop_id] = value
        # PREFERRED defaults override fallbacks (v1.0.14 validated: Cowl Neck /
        # Semua musim are both legal values for T-shirt in HiddenAttr)
        preferred = by_country("ID").property_preferred_defaults or {}
        for prop_id, value in preferred.items():
            if prop_id in cols and prop_id in row:
                row[prop_id] = value

    return row


def build_rows_for_product(
    product: Product,
    settings: dict[str, Any],
    country: str,
    copy_suffixes: list[str] | None = None,
    apply_suffix_to_first: bool = False,
) -> list[OutputRow]:
    """Build all TikTok rows for one Product."""
    if copy_suffixes is None:
        copies = max(1, int(settings.get("output_copies", 1) or 1))
        copy_suffixes = [""] * copies
    else:
        copies = len(copy_suffixes) if copy_suffixes else 1

    common = _resolve_common_fields(product, settings, country)
    # ID-only: compute HiddenAttr fallback dict once per product (cheap, but
    # avoids re-opening the template for every variant row)
    property_fallbacks = _get_property_fallbacks(country, common["category"])

    rows: list[OutputRow] = []

    if not product.variants:
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, Variant(), settings, common, country,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
                apply_suffix_to_first=apply_suffix_to_first,
                property_fallbacks=property_fallbacks,
            ))
        return rows

    for v in product.variants:
        for c_idx in range(copies):
            rows.append(_build_row_for_variant(
                product, v, settings, common, country,
                copy_idx=c_idx, copy_suffix=copy_suffixes[c_idx],
                apply_suffix_to_first=apply_suffix_to_first,
                property_fallbacks=property_fallbacks,
            ))
    return rows


# ---------------------------------------------------------------------------
# Excel writer
# ---------------------------------------------------------------------------
def write_tiktok_xlsx(
    output_path: Path,
    rows: list[OutputRow],
    country: str,
) -> Path:
    """Copy the country template to output_path, write `rows` into Template sheet."""
    template_src = template_path_for(country)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_src, output_path)
    wb = openpyxl.load_workbook(output_path)
    if "Template" not in wb.sheetnames:
        wb.close()
        raise ValueError(f"模板文件缺少 'Template' sheet：{template_src}")
    ws = wb["Template"]

    # Build header→col index from the first row of the Template sheet
    header_row: list[str] = []
    for c in ws[1]:
        v = c.value
        header_row.append(_clean_str(v) if v is not None else "")
    cols = get_columns_for(country)

    col_idx: dict[str, int] = {}
    for i, h in enumerate(header_row, 1):
        if h and h in cols:
            col_idx[h] = i

    # Clear any pre-existing data rows in the Template sheet
    start_row = 2
    max_existing = ws.max_row
    if max_existing >= start_row:
        ws.delete_rows(start_row, max_existing - start_row + 1)

    # Write new rows
    for r_off, row_dict in enumerate(rows):
        excel_row = start_row + r_off
        for col_name, value in row_dict.items():
            ci = col_idx.get(col_name)
            if ci is None:
                continue
            ws.cell(row=excel_row, column=ci, value=value)

    wb.save(output_path)
    wb.close()
    return output_path
"""Per-country presets for the multi-country TK table converter.

Defines the deltas that make ID / PH / TH different so that the rest of the
codebase can ask `presets.by_country(code).xxx` instead of branching on
country in every function.

This file is the single source of truth for per-country differences:
- Default conversion settings (brand/price/category/cod/title/size_chart/...)
- Template assets subdirectory (each country ships its own 模板)
- Country-specific category translation table (ID only — it has the EasyBoss
  Indonesian path → official Indonesian name remap)
- Country-specific product_property PREFERRED defaults (ID only — v1.0.14
  validated Cowl Neck / Semua musim are in the legal HiddenAttr list)
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class CountryPreset:
    country_code: str        # "ID" / "PH" / "TH"
    country_name: str        # "印尼" / "菲律宾" / "泰国" (display in GUI)
    app_name: str            # GUI 标题里的本地化名
    assets_subdir: str         # "assets-id" / "assets-ph" / "assets-th"
    template_filename: str = "batch-product-source.xlsx"
    settings: dict[str, Any] = field(default_factory=dict)
    category_translation: dict[str, str] | None = None
    property_preferred_defaults: dict[str, str] | None = None


# ---------------------------------------------------------------------------
# ID settings (印尼) — 来自 tk-id-converter-repo v1.0.14
# ---------------------------------------------------------------------------
ID_SETTINGS: dict[str, Any] = {
    "download_images_enabled": True,
    "title_prefix_enabled": True,
    "title_prefix": "Kaos Unisex Oversize ",
    "brand_enabled": True,
    "brand_value": "Tidak ada merek",
    "price_enabled": True,
    "price_value": 115000,           # IDR
    "quantity_enabled": True,
    "quantity_value": 999999,
    "cod_enabled": True,
    "cod_value": "Y",                # ID 支持 COD
    "fill_sizes_enabled": True,
    "standard_sizes": "S,M,L,XL,XXL,XXXL",
    "title_random_suffix_enabled": True,
    "random_suffix_length": 6,
    "category_enabled": True,
    "category_value": "Atasan Pria/T-shirt",
    "output_copies": 1,
    "split_output_files": False,
    "description_enabled": True,
    "description_value": (
        "Kaos unisex oversize, bahan adem, nyaman dipakai sehari-hari. "
        "Lihat bagan ukuran sebelum memesan."
    ),
    "size_chart_enabled": True,
    "size_chart_value": "",          # 用户手动传 Media Center URL
    "parcel_enabled": True,
    "parcel_weight_value": 200,
    "parcel_length_value": 15,
    "parcel_width_value": 15,
    "parcel_height_value": 5,
    "pre_order_time_value": "",
    "delivery_value": "",
    # ID 独有字段
    "minimum_order_quantity_value": 1,
    "shipping_insurance_enabled": False,
    "shipping_insurance_value": 1000,
}


# ---------------------------------------------------------------------------
# PH settings (菲律宾) — 来自 tk-ph-converter-repo v3.2.3
# ---------------------------------------------------------------------------
PH_SETTINGS: dict[str, Any] = {
    "download_images_enabled": True,
    "title_prefix_enabled": True,
    "title_prefix": "COD Unisex T-shirt【S-3XL】 ",
    "brand_enabled": True,
    "brand_value": "No brand",
    "price_enabled": True,
    "price_value": 356,              # PHP
    "quantity_enabled": True,
    "quantity_value": 999,
    "cod_enabled": True,
    "cod_value": "Y",
    "fill_sizes_enabled": True,
    "standard_sizes": "S,M,L,XL,2XL,3XL",
    "title_random_suffix_enabled": True,
    "random_suffix_length": 3,
    "category_enabled": True,
    "category_value": "Men's Tops/T-shirts",
    "output_copies": 2,              # PH 默认出 2 份（防查重）
    "split_output_files": False,
    "description_enabled": True,
    "description_value": (
        "High-quality product with careful packaging. "
        "Material: premium fabric, soft and comfortable, breathable for daily wear. "
        "Care: machine washable, retains shape after washing. "
        "Size: please refer to the size chart image before ordering. "
        "Shipping: orders ship within 1-2 business days; delivery typically takes 3-8 days."
    ),
    "size_chart_enabled": True,
    "size_chart_value": (
        "https://raw.githubusercontent.com/yuqingyangxt1-wq/"
        "tk-ph-converter/main/assets/default_size_chart.png"
    ),
    "parcel_enabled": True,
    "parcel_weight_value": 200,
    "parcel_length_value": 10,
    "parcel_width_value": 10,
    "parcel_height_value": 5,
    "pre_order_time_value": "",
    "delivery_value": "",
}


# ---------------------------------------------------------------------------
# TH settings (泰国) — 来自 tk-th-converter-repo v1.0.1
# ---------------------------------------------------------------------------
TH_SETTINGS: dict[str, Any] = {
    "download_images_enabled": True,
    "title_prefix_enabled": True,
    "title_prefix": "เสื้อยืด Oversize ",
    "brand_enabled": True,
    "brand_value": "No brand",
    "price_enabled": True,
    "price_value": 199,              # THB
    "quantity_enabled": True,
    "quantity_value": 999,
    "cod_enabled": True,
    "cod_value": "N",                # TH 后台不要 COD
    "fill_sizes_enabled": True,
    "standard_sizes": "S,M,L,XL,2XL,3XL",
    "default_color_enabled": True,    # TH 独有兜底颜色
    "default_color_value": "As Picture",
    "title_random_suffix_enabled": True,
    "random_suffix_length": 3,
    "category_enabled": True,
    "category_value": "Womenswear & Underwear>Women's Tops>Women's T-shirts",
    "output_copies": 2,
    "split_output_files": False,
    "description_enabled": True,
    "description_value": (
        "เสื้อยืด Oversize คุณภาพดี ใส่สบาย เนื้อผ้านุ่ม ระบายอากาศได้ดี "
        "ดูแลรักษาง่าย ซักเครื่องได้ กรุณาดูตารางไซส์ก่อนสั่งซื้อ"
    ),
    "size_chart_enabled": True,
    "size_chart_value": "",          # TH 必传 Media Center URL
    "parcel_enabled": True,
    "parcel_weight_value": 200,
    "parcel_length_value": 15,
    "parcel_width_value": 15,
    "parcel_height_value": 5,
    "pre_order_time_value": "",
    "delivery_value": "",
}


# ---------------------------------------------------------------------------
# ID-only: EasyBoss 印尼语类目路径 → 官方模板印尼语类目名
# 来自 tk-id-converter-repo v1.0.10 (24 条，覆盖 ID HiddenStyle 25 类目里大多数)
# ---------------------------------------------------------------------------
_ID_CATEGORY_TRANSLATION: dict[str, str] = {
    # 男装上衣
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>T-shirt": "Atasan Pria/T-shirt",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kaus Polo": "Atasan Pria/Kaus Polo",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Kemeja": "Atasan Pria/Kemeja",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rajut": "Atasan Pria/Pakaian Rajut",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Hoodie & Sweatshirt": "Atasan Pria/Hoodie & Jumper",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Rompi & Gilet": "Atasan Pria/Rompi Waistcoat & Gilet",
    "Pakaian & Pakaian Dalam Pria>Atasan Pria>Jaket & Mantel": "Atasan Pria/Jaket & Mantel",
    # 男装下装
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Pendek": "Bawahan Pria/Celana pendek",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Jeans": "Bawahan Pria/Jeans",
    "Pakaian & Pakaian Dalam Pria>Celana Pria>Celana Panjang": "Bawahan Pria/Celana Pria",
    # 男士套装
    "Setelan Pria>Set Pakaian Pria": "Setelan & Overall Pria/Set Pakaian Pria",
    "Setelan Pria>Setelan Pria": "Setelan & Overall Pria/Setelan Resmi",
    "Setelan Pria>Overall": "Setelan & Overall Pria/Overall",
    # 男士内衣袜子
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Pria": "Pakaian Dalam Pria/Pakaian Dalam",
    "Pakaian Dalam & Kaus Kaki Pria>Tank Top & Pakaian Dalam": "Pakaian Dalam Pria/Pakaian Dalam",
    "Pakaian Dalam & Kaus Kaki Pria>Pakaian Dalam Hangat": "Pakaian Dalam Pria/Pakaian Dalam Termal",
    "Pakaian Dalam & Kaus Kaki Pria>Kaus Kaki": "Pakaian Dalam Pria/Kaus kaki",
    # 男士睡衣
    "Pakaian Tidur & Pakaian Santai Pria>Piyama & Loungewear": "Baju Tidur dan Baju Santai Pria/Piyama",
    "Pakaian Tidur & Pakaian Santai Pria>Robe Pria": "Baju Tidur dan Baju Santai Pria/Kimono Mandi & Rias",
    "Pakaian Tidur & Pakaian Santai Pria>Nightshirt": "Baju Tidur dan Baju Santai Pria/Piyama Midi",
    "Pakaian Tidur & Pakaian Santai Pria>Piyama Pria": "Baju Tidur dan Baju Santai Pria/Piama Terusan Pria",
    # 男士特殊场合
    "Pakaian Acara Khusus Pria>Kostum": "Pakaian Khusus Pria/Kostum & Aksesoris",
    "Pakaian Acara Khusus Pria>Pakaian Kerja": "Pakaian Khusus Pria/Pakaian Kerja & Seragam",
    "Pakaian Acara Khusus Pria>Pakaian Tradisional": "Pakaian Khusus Pria/Baju Tradisional",
}


# ---------------------------------------------------------------------------
# ID-only: product_property PREFERRED 默认值（v1.0.14 验证）
# ---------------------------------------------------------------------------
_ID_PROPERTY_PREFERRED: dict[str, str] = {
    "product_property/100393": "Cowl Neck",     # Neckline — HiddenAttr R243 合法
    "product_property/100397": "Semua musim",    # Season — HiddenAttr R45 合法
}


# ---------------------------------------------------------------------------
# Presets registry
# ---------------------------------------------------------------------------
PRESETS: dict[str, CountryPreset] = {
    "ID": CountryPreset(
        country_code="ID",
        country_name="印尼",
        app_name="TK印尼表格转化工具",
        assets_subdir="assets-id",
        settings=ID_SETTINGS,
        category_translation=_ID_CATEGORY_TRANSLATION,
        property_preferred_defaults=_ID_PROPERTY_PREFERRED,
    ),
    "PH": CountryPreset(
        country_code="PH",
        country_name="菲律宾",
        app_name="TK菲律宾表格转化工具",
        assets_subdir="assets-ph",
        settings=PH_SETTINGS,
    ),
    "TH": CountryPreset(
        country_code="TH",
        country_name="泰国",
        app_name="TK泰国表格转化工具",
        assets_subdir="assets-th",
        settings=TH_SETTINGS,
    ),
}


def by_country(code: str) -> CountryPreset:
    """Return the preset for a country code. Raises KeyError if unknown."""
    if code not in PRESETS:
        raise KeyError(
            f"Unknown country code: {code!r}. Available: {list(PRESETS)}"
        )
    return PRESETS[code]


def country_choices() -> list[tuple[str, str]]:
    """Return (code, display_name) tuples for GUI dropdown population."""
    return [(c, p.country_name) for c, p in PRESETS.items()]
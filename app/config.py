"""Configuration management - load/save config.json, provide defaults.

v3.3 (multi-country): 支持 ID / PH / TH 三国切换；settings 来自 app.presets，
country 字段默认 PH（基线仓库是 PH 版）。GUI 切换国家时写入 cfg["country"]，
load_config() 据此从 presets 取对应 settings。
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from app import presets


__version__ = "3.3.0"  # multi-country: ID/PH/TH 三国合一 (基线 PH v3.2.3 + ID v1.0.14 + TH v1.0.1)
DEFAULT_COUNTRY = "PH"   # 基线仓库是 PH 版；GUI 启动时也可改成 ID/TH


def app_name_for(country: str) -> str:
    return presets.by_country(country).app_name


def get_app_dir() -> Path:
    """Return the writable directory next to the executable.

    This is where config.json (and any future user files) live. For both
    a PyInstaller --onefile exe and a `python -m app.main` source run,
    this is the directory containing the exe / project root.
    """
    if getattr(sys, "frozen", False):  # PyInstaller bundle
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def get_assets_dir(country: str | None = None) -> Path:
    """Return the country-specific assets directory.

    Source run: <project>/assets-{id|ph|th}/
    PyInstaller: <exe-dir>/assets-{id|ph|th}/

    Falls back to assets-ph/ if country is None (shouldn't happen in normal
    use; only used by tests that don't care about country).
    """
    code = country or DEFAULT_COUNTRY
    return get_app_dir() / presets.by_country(code).assets_subdir


def default_config(country: str = DEFAULT_COUNTRY) -> dict[str, Any]:
    """The shipped defaults — sourced from presets.by_country(country)."""
    p = presets.by_country(country)
    return {
        "version": 2,
        "country": country,                      # 多国家: 'ID' / 'PH' / 'TH'
        "product_xlsx_last": "",
        "product_xlsx_output_dir": "",
        "product_xlsx_last_output_dir": "",
        "product_pool_dir": "",
        "product_xlsx_settings": dict(p.settings),  # country-specific defaults
        # Source column mapping (override which EasyBoss/源列 maps to what)
        # If a key is empty/None, the reader falls back to auto-detection.
        "source_column_mapping": {
            "product_name": "产品名",
            "brand": "品牌",
            "category": "产品类目",
            "platform_sku": "平台SKU",
            "var1_name": "规格1名称",
            "var1_value": "规格1选项",
            "var2_name": "规格2名称",
            "var2_value": "规格2选项",
            "var3_name": "规格3名称",
            "var3_value": "规格3选项",
            "price": "税前价格",
            "stock": "库存",
            "sku_image": "SKU图片",
            "image_1": "产品图片1",
            "image_2": "产品图片2",
            "image_3": "产品图片3",
            "image_4": "产品图片4",
            "image_5": "产品图片5",
            "image_6": "产品图片6",
            "image_7": "产品图片7",
            "image_8": "产品图片8",
            "image_9": "产品图片9",
            "parcel_weight_kg": "包裹重量（KG）",
            "parcel_length": "包裹长度（CM）",
            "parcel_width": "包裹宽度（CM）",
            "parcel_height": "包裹高度（CM）",
            "description": "产品描述",
            "size_chart": "尺码图",
        },
    }


def config_path() -> Path:
    return get_app_dir() / "config.json"


def load_config(path: Path | None = None) -> dict[str, Any]:
    """Load config from disk, merging with defaults so new keys appear.

    If cfg has no 'country' key (legacy PH-only files), assume PH.
    """
    p = path or config_path()
    country = DEFAULT_COUNTRY
    if p.exists():
        try:
            with p.open("r", encoding="utf-8") as f:
                on_disk = json.load(f)
            country = on_disk.get("country", DEFAULT_COUNTRY)
        except (OSError, json.JSONDecodeError):
            pass
    if country not in presets.PRESETS:
        country = DEFAULT_COUNTRY
    cfg = default_config(country)
    if p.exists():
        try:
            with p.open("r", encoding="utf-8") as f:
                on_disk = json.load(f)
            # Shallow merge at top level, deep merge for *_settings / mapping
            for k, v in on_disk.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except (OSError, json.JSONDecodeError):
            # Corrupt config → fall back to defaults but don't overwrite
            pass
    return cfg


def save_config(cfg: dict[str, Any], path: Path | None = None) -> None:
    p = path or config_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_settings(cfg: dict[str, Any]) -> dict[str, Any]:
    """Return the product_xlsx_settings sub-dict, applying defaults for missing keys."""
    country = cfg.get("country", DEFAULT_COUNTRY)
    defaults = default_config(country)["product_xlsx_settings"]
    s = dict(defaults)
    s.update(cfg.get("product_xlsx_settings", {}))
    return s


def update_settings(cfg: dict[str, Any], **kwargs: Any) -> None:
    cfg.setdefault("product_xlsx_settings", {})
    cfg["product_xlsx_settings"].update(kwargs)


def set_country(cfg: dict[str, Any], country: str) -> None:
    """Switch config to a different country.

    Resets product_xlsx_settings to that country's defaults — user has to
    re-tune their tweaks for the new country's pipeline. The rest of cfg
    (paths, source_column_mapping) is preserved.
    """
    if country not in presets.PRESETS:
        raise ValueError(
            f"Unknown country: {country!r}. Available: {list(presets.PRESETS)}"
        )
    cfg["country"] = country
    cfg["product_xlsx_settings"] = dict(presets.by_country(country).settings)


# v3.3.0: 兼容旧代码里的 APP_NAME 引用
APP_NAME = app_name_for(DEFAULT_COUNTRY)

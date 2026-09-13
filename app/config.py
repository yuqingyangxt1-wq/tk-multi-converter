"""Configuration management - load/save config.json, provide defaults.

v3.3 (multi-country): 支持 ID / PH / TH 三国切换；settings 来自 app.presets，
country 字段默认 PH（基线仓库是 PH 版）。GUI 切换国家时写入 cfg["country"]，
load_config() 据此从 presets 取对应 settings。

v3.3.2 修复 (2026-09-13): 每个国家单独保存一份 product_xlsx_settings，
切换国家不再 reset 已设置的参数。结构：
  cfg["settings_by_country"]["ID"] = {...}
  cfg["settings_by_country"]["PH"] = {...}
  cfg["settings_by_country"]["TH"] = {...}
兼容老配置：cfg["product_xlsx_settings"] 在 load 时迁移到
settings_by_country[当前 country]。
"""
from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from app import presets


__version__ = "3.3.9"  # fix: ID — restore PREFERRED only for Neckline (Cowl Neck, user's 套头圆领 T-shirt). Season/Care still go through HiddenAttr first-row defaults (Musim semi / Dicuci dengan Tangan Saja) because v3.3.7 PREFERRED Semua musim / Cuci Kering were rejected by TikTok Shop seller's T-shirt category dropdown
DEFAULT_COUNTRY = "PH"   # 基线仓库是 PH 版；GUI 启动时也可改成 ID/TH


def _settings_for(country: str) -> dict[str, Any]:
    """Return the presets-defined default settings for a country (fresh copy)."""
    return dict(presets.by_country(country).settings)


def _empty_settings_by_country() -> dict[str, dict[str, Any]]:
    """Build a fresh {ID:{}, PH:{}, TH:{}} seeded with each country's preset defaults."""
    return {code: _settings_for(code) for code in presets.PRESETS}


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
    """The shipped defaults — sourced from presets.by_country(country).

    settings_by_country holds one settings dict per country so the user's
    tweaks persist across country switches. product_xlsx_settings is kept
    as a legacy mirror of settings_by_country[country] for backward compat
    with any caller that still reads the old key.
    """
    p = presets.by_country(country)
    return {
        "version": 3,
        "country": country,                      # 多国家: 'ID' / 'PH' / 'TH'
        "product_xlsx_last": "",
        "product_xlsx_output_dir": "",
        "product_xlsx_last_output_dir": "",
        "product_pool_dir": "",
        # v3.3.2: per-country persisted settings
        "settings_by_country": _empty_settings_by_country(),
        # legacy mirror — kept so older code paths that read the old key
        # still work; load_config() rewrites this from settings_by_country.
        "product_xlsx_settings": dict(p.settings),
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

    Migration: if on-disk cfg has the v3.3 'product_xlsx_settings' key but
    no 'settings_by_country', move it into settings_by_country[<current country>]
    so the user's tweaks survive the upgrade instead of being lost the first
    time they switch country.
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
            # Migrate v3.3 (or v3.2) 'product_xlsx_settings' into
            # settings_by_country[<country>] before merging.
            legacy = on_disk.pop("product_xlsx_settings", None)
            if legacy and isinstance(legacy, dict):
                sbc = cfg.setdefault("settings_by_country", {})
                # Stash legacy values into the slot for whatever country was active.
                bucket = sbc.setdefault(country, {})
                # Don't clobber — merge: existing defaults win for missing keys,
                # but legacy values win for keys present.
                for k, v in legacy.items():
                    if v not in (None, "", []):
                        bucket[k] = v
            # Shallow merge at top level, deep merge for *_settings / mapping
            for k, v in on_disk.items():
                if isinstance(v, dict) and isinstance(cfg.get(k), dict):
                    cfg[k].update(v)
                else:
                    cfg[k] = v
        except (OSError, json.JSONDecodeError):
            # Corrupt config → fall back to defaults but don't overwrite
            pass
    # Always mirror the active country's settings into the legacy key so
    # any caller still reading the old field sees consistent data.
    cfg["product_xlsx_settings"] = dict(
        cfg.get("settings_by_country", {}).get(country, _settings_for(country))
    )
    return cfg


def save_config(cfg: dict[str, Any], path: Path | None = None) -> None:
    p = path or config_path()
    p.parent.mkdir(parents=True, exist_ok=True)
    with p.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_settings(cfg: dict[str, Any]) -> dict[str, Any]:
    """Return the active country's settings, with that country's preset
    defaults filled in for any missing keys.
    """
    country = cfg.get("country", DEFAULT_COUNTRY)
    sbc = cfg.setdefault("settings_by_country", {})
    bucket = sbc.setdefault(country, _settings_for(country))
    # Fill in any newly-added preset keys for this country.
    defaults = _settings_for(country)
    out = dict(defaults)
    out.update(bucket)
    # Mirror to legacy key for callers that still read it.
    cfg["product_xlsx_settings"] = dict(out)
    return out


def update_settings(cfg: dict[str, Any], **kwargs: Any) -> None:
    """Update the active country's settings in place (does not touch other
    countries' saved values).
    """
    country = cfg.get("country", DEFAULT_COUNTRY)
    sbc = cfg.setdefault("settings_by_country", {})
    bucket = sbc.setdefault(country, _settings_for(country))
    bucket.update(kwargs)
    cfg["product_xlsx_settings"] = dict(bucket)


def country_choices() -> list[tuple[str, str]]:
    """List of (code, display_name) for all supported countries, in display order."""
    return presets.country_choices()


def set_country(cfg: dict[str, Any], country: str) -> None:
    """Switch config to a different country.

    v3.3.2: 每个国家的 settings 独立保存 — 切换后只切换 cfg['country']，
    不重置已保存的参数。其它路径/列映射也保留。
    """
    if country not in presets.PRESETS:
        raise ValueError(
            f"Unknown country: {country!r}. Available: {list(presets.PRESETS)}"
        )
    cfg["country"] = country
    # Ensure the new country has a bucket (first visit seeds it from presets).
    sbc = cfg.setdefault("settings_by_country", _empty_settings_by_country())
    sbc.setdefault(country, _settings_for(country))
    # Mirror the new country's settings into the legacy key for any
    # downstream code that still reads the old field.
    cfg["product_xlsx_settings"] = dict(sbc[country])


# v3.3.0: 兼容旧代码里的 APP_NAME 引用
APP_NAME = app_name_for(DEFAULT_COUNTRY)

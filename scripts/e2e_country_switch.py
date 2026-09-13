"""E2E: 复现用户场景 — 三国切换后转换必须自动用对应国家参数。

用户报告: 用印尼转换完,再切泰国转换不会自动用泰国参数,必须重启才行。
根因: converter.py 单文件模式分支(默认)漏传 country + write_tiktok_xlsx 传了
未定义的 template_src + 文件名硬编码 _TKPH_;main.py pool 拆分分支同病。

本测试: mock 一张 EasyBoss 风格源表,模拟 GUI 的完整切换流程
(ID 单文件 → TH 单文件 → PH 拆分 → 回 ID),断言每次输出的文件名/列数/类目/价格。
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parent
sys.path.insert(0, str(REPO))

import openpyxl

from app import config as cfg_mod
from app.converter import convert_source

WORK = REPO / "e2e_out"
MOCK = WORK / "mock_source.xlsx"

PASS, FAIL = 0, 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  [OK]   {name}")
    else:
        FAIL += 1
        print(f"  [FAIL] {name}  {detail}")


def make_mock_source():
    """EasyBoss 导出风格的源表: 2 产品 x 2 颜色 x 3 尺码 = 12 行."""
    headers = [
        "产品名", "产品类目", "品牌", "产品描述",
        "规格1名称", "规格1选项", "规格2名称", "规格2选项",
        "税前价格", "库存", "主图",
        "产品图片1", "产品图片2", "产品图片3",
        "平台SKU", "尺码图",
    ]
    rows = []
    for p_idx, pname in enumerate(["测试圆领T恤", "测试连帽卫衣"], 1):
        for color, color_id in [("白色", "white"), ("黑色", "black")]:
            for size in ["S", "M", "L"]:
                rows.append([
                    pname,
                    "男装>T恤",                    # 源表类目(会被 settings 覆盖)
                    "No brand",
                    f"{pname} 描述文案",
                    "颜色", color, "尺码", size,
                    356, 999,                      # 价格/库存(会被覆盖)
                    f"https://img.example.com/p{p_idx}.jpg",
                    f"https://img.example.com/p{p_idx}.jpg",
                    f"https://img.example.com/p{p_idx}_2.jpg",
                    f"https://img.example.com/p{p_idx}_3.jpg",
                    f"SKU-{p_idx}-{color_id}-{size}",
                    "",
                ])
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "导出#SKU"
    ws.append(headers)
    for r in rows:
        ws.append(r)
    wb.save(MOCK)
    return MOCK


def convert_and_inspect(cfg, tag, expect_country, expect_split=False):
    """用 cfg 当前国家跑一次转换,返回 (out_xlsx, headers)."""
    settings = cfg_mod.get_settings(cfg)
    country = cfg.get("country", "PH")
    out_dir = WORK / f"run_{tag}"
    result = convert_source(
        source_xlsx=MOCK,
        output_dir=out_dir,
        settings=settings,
        country=country,
        download_imgs=False,
    )
    out_xlsx = result.output_paths[0]
    wb = openpyxl.load_workbook(out_xlsx)
    ws = wb["Template"]
    headers = [str(c.value or "") for c in ws[1]]
    return result, out_xlsx, headers, ws


def main():
    if WORK.exists():
        shutil.rmtree(WORK)
    WORK.mkdir(parents=True)
    make_mock_source()

    # ============================================================
    print("\n=== 场景 1: 印尼 ID 单文件模式(用户主路径) ===")
    cfg = cfg_mod.default_config("ID")
    id_price = 111111
    cfg_mod.update_settings(cfg, price_value=id_price, price_enabled=True)
    result, out_xlsx, headers, ws = convert_and_inspect(cfg, "id1", "ID")

    check("转换成功不崩(此前 TypeError)", result.row_count == 12,
          f"rows={result.row_count}")
    check("文件名 _TKID_", "_TKID_" in out_xlsx.name, out_xlsx.name)
    check("ID 模板 40 列", len([h for h in headers if h]) == 40,
          f"cols={len([h for h in headers if h])}")
    check("ID 有 pre_order_time / minimum_order_quantity 列(ID 独有)",
          "pre_order_time" in headers and "minimum_order_quantity" in headers)

    def cellv(ws, headers, label, row=2):
        if label not in headers:
            return None
        return ws.cell(row, headers.index(label) + 1).value

    check("ID 类目=印尼语 Atasan Pria/T-shirt",
          cellv(ws, headers, "category") == "Atasan Pria/T-shirt",
          repr(cellv(ws, headers, "category")))
    check(f"ID 价格={id_price}(IDR preset 覆盖)",
          cellv(ws, headers, "price") == id_price,
          repr(cellv(ws, headers, "price")))
    check("ID 颜色 白色→Putih",
          cellv(ws, headers, "property_value_1") == "Putih",
          repr(cellv(ws, headers, "property_value_1")))
    check("ID 品牌=Tidak ada merek",
          cellv(ws, headers, "brand") == "Tidak ada merek",
          repr(cellv(ws, headers, "brand")))

    # ============================================================
    print("\n=== 场景 2: 切泰国 TH 单文件模式(不重启!) ===")
    cfg_mod.set_country(cfg, "TH")
    th_price = 222
    cfg_mod.update_settings(cfg, price_value=th_price, price_enabled=True)
    th_copies = cfg_mod.get_settings(cfg).get("output_copies", 1) or 1
    result, out_xlsx, headers, ws = convert_and_inspect(cfg, "th1", "TH")

    check("TH 转换成功(行数=12×preset份数)",
          result.row_count == 12 * max(1, int(th_copies)),
          f"rows={result.row_count}, copies={th_copies}")
    check("TH 无 minimum_order_quantity 列(非 ID 模板)",
          "minimum_order_quantity" not in headers)
    check("文件名 _TKTH_(此前硬编码 _TKPH_)", "_TKTH_" in out_xlsx.name,
          out_xlsx.name)
    check("TH 模板 39 列(不是 ID 的 40)",
          len([h for h in headers if h]) == 39,
          f"cols={len([h for h in headers if h])}")
    check("TH 无 shipping_insurance 列", "shipping_insurance" not in headers)
    check(f"TH 价格={th_price}",
          cellv(ws, headers, "price") == th_price,
          repr(cellv(ws, headers, "price")))
    th_cat = cellv(ws, headers, "category")
    check("TH 类目≠印尼语", th_cat != "Atasan Pria/T-shirt", repr(th_cat))
    check("TH 品牌非 Tidak ada merek",
          cellv(ws, headers, "brand") != "Tidak ada merek",
          repr(cellv(ws, headers, "brand")))

    # ============================================================
    print("\n=== 场景 3: 切菲律宾 PH 拆分文件模式 ===")
    cfg_mod.set_country(cfg, "PH")
    ph_price = 333
    cfg_mod.update_settings(cfg, price_value=ph_price, price_enabled=True,
                            split_output_files=True, output_copies=2)
    result, out_xlsx, headers, ws = convert_and_inspect(cfg, "ph1", "PH")

    check("PH 拆分 2 个文件", len(result.output_paths) == 2,
          f"files={len(result.output_paths)}")
    check("文件名 _TKPH_", all("_TKPH_" in p.name for p in result.output_paths),
          str([p.name for p in result.output_paths]))
    check("PH 模板 39 列", len([h for h in headers if h]) == 39,
          f"cols={len([h for h in headers if h])}")
    check(f"PH 价格={ph_price}",
          cellv(ws, headers, "price") == ph_price,
          repr(cellv(ws, headers, "price")))
    check("PH 类目=Men's Tops/T-shirts",
          cellv(ws, headers, "category") == "Men's Tops/T-shirts",
          repr(cellv(ws, headers, "category")))

    # ============================================================
    print("\n=== 场景 4: 切回印尼,参数没被 PH/TH 污染 ===")
    cfg_mod.set_country(cfg, "ID")
    result, out_xlsx, headers, ws = convert_and_inspect(cfg, "id2", "ID")

    check("回 ID 文件名 _TKID_", "_TKID_" in out_xlsx.name, out_xlsx.name)
    check("回 ID 类目=Atasan Pria/T-shirt",
          cellv(ws, headers, "category") == "Atasan Pria/T-shirt",
          repr(cellv(ws, headers, "category")))
    check("回 ID 价格仍=111111(未被 TH/PH 污染)",
          cellv(ws, headers, "price") == id_price,
          repr(cellv(ws, headers, "price")))
    check("回 ID 40 列", len([h for h in headers if h]) == 40,
          f"cols={len([h for h in headers if h])}")

    # ============================================================
    print("\n=== 场景 5: settings_by_country 隔离持久化 ===")
    sbc = cfg["settings_by_country"]
    check("ID bucket price=111111", sbc["ID"].get("price_value") == id_price,
          repr(sbc["ID"].get("price_value")))
    check("TH bucket price=222", sbc["TH"].get("price_value") == th_price,
          repr(sbc["TH"].get("price_value")))
    check("PH bucket price=333", sbc["PH"].get("price_value") == ph_price,
          repr(sbc["PH"].get("price_value")))

    print(f"\n{'=' * 50}")
    print(f"RESULT: {PASS} passed, {FAIL} failed")
    sys.exit(1 if FAIL else 0)


if __name__ == "__main__":
    main()
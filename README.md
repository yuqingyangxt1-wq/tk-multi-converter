# TK 多国家表格转化工具 v3.3

把 EasyBoss 导出的 #SKU 表格（或 TikTok 批量上传模板本身）一键转成 TikTok Shop 印尼（ID）/ 菲律宾（PH）/ 泰国（TH）任一国家的批量上传模板。

> 这是把 TK-ID-Converter / TK-PH-Converter / TK-TH-Converter 三个独立工具合并到一个软件里。窗口顶部 目标国家 Radio 切换 ID/PH/TH，对应使用该国家的官方模板 + 默认设置（品牌/价格/类目/COD/包装）+ ID 专属的印尼语类目翻译 + HiddenAttr 兜底。

---

## 功能

- 三国切换：顶部 Radio 切换 ID/PH/TH，自动加载对应官方模板 + 国家默认设置
- 转化：源 xlsx → TikTok 批量上传模板，自动套用设置
- 入池 / 提取：产品池
- 设置：全字段可配置
- 源表识别：EasyBoss #SKU 或 TikTok 模板
- 多副本：每产品 N 个 listing（带随机后缀，PH/TH 默认 2 份防查重）
- ID 颜色翻译：白色/White → Putih 等
- ID 类目翻译：EasyBoss 印尼语路径 → 官方类目
- ID 属性兜底：HiddenAttr 模板里查合法 Neckline/Season + PREFERRED 默认 (Cowl Neck / Semua musim)
- ID 专属字段：minimum_order_quantity + pre_order_time + shipping_insurance
- 进度：底部真实进度条 + 后台线程

---

## 目录结构

- app/ - 源码包 (main.py / presets.py / config.py / source_reader.py / converter.py / tiktok_writer.py / product_pool.py)
- assets-id/ - 印尼官方模板 (40 列)
- assets-ph/ - 菲律宾官方模板 (39 列)
- assets-th/ - 泰国官方模板 (39 列)
- .github/workflows/build-windows.yml - Actions 打包
- TK-Multi-Converter.spec - PyInstaller 配置
- build_windows.bat / build_simple.bat
- BUILD.md - 打包说明

---

## 使用方法

### Windows：双击 exe
1. 解压 TK-Multi-Converter_windows.zip
2. 双击 TK-Multi-Converter.exe（macOS 需 Wine/CrossOver）
3. 顶部 目标国家 选 ID/PH/TH（默认 PH）
4. 选源 xlsx → 开始转化 → 同目录生成 *_TK<COUNTRY>_<ts>.xlsx

### macOS：跑源码
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python -m app

---

## 版本历史

- v3.3 - 多国家合版 (本次): ID/PH/TH 三国切换 + ID 专属翻译和兜底
- v3.2 - TK-TH-Converter 独立版
- v1.0.x - TK-ID-Converter 独立版
- v1.2 - TK-PH-Converter 独立版（PH v3.2.3 基线）
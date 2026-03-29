"""
loader.py — 数据加载与验证模块
负责读取上传的Excel文件，统一清洗字段，返回标准化的DataFrame字典
"""
import pandas as pd
import numpy as np
import streamlit as st


# ── Sheet名称映射（常见变体）──
SHEET_ALIASES = {
    "ar": ["应收账款表", "应收账款", "AR", "Sheet1"],
    "ap": ["应付账款表", "应付账款", "AP", "Sheet1"],
    "adv": ["预收账款表", "预收账款", "预收", "Sheet1"],
    "cf": ["经营性现金流", "经营现金流", "现金流", "Sheet1"],
    "income_cost": ["Sheet1", "收入成本", "明细"],
}

FS_SHEETS = {
    "income_stmt": ["利润表", "损益表", "合并利润表", "综合损益表"],
    "balance_sheet": ["资产负债表", "合并资产负债表", "Balance"],
    "cash_flow": ["现金流量表", "合并现金流量表", "CashFlow"],
}


def _read_sheet(file_obj, candidates):
    """尝试读取候选Sheet名称中的第一个存在的Sheet"""
    try:
        xl = pd.ExcelFile(file_obj)
        sheets = xl.sheet_names
        for c in candidates:
            if c in sheets:
                return pd.read_excel(file_obj, sheet_name=c)
        # 如果都没找到，读第一个Sheet
        return pd.read_excel(file_obj, sheet_name=0)
    except Exception as e:
        raise ValueError(f"读取文件失败：{e}")


def _clean_df(df, num_cols=None, str_cols=None, filter_col=None):
    """通用清洗：数字列转float，字符串列strip，剔除标记行"""
    df = df.copy()
    if num_cols:
        for c in num_cols:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    if str_cols:
        for c in str_cols:
            if c in df.columns:
                df[c] = df[c].astype(str).str.strip()
    if filter_col and filter_col in df.columns:
        df[filter_col] = pd.to_numeric(df[filter_col], errors="coerce").fillna(1)
        df = df[df[filter_col] == 0]
    return df


def load_and_validate(file_ic, file_fs, file_ar, file_ap, file_adv, file_cf, config):
    """
    加载全部底表，返回标准化数据字典 data{}
    """
    data = {}
    year = config["report_year"]   # e.g. "2025年"
    base = config["base_year"]     # e.g. "2024年"
    rp_tag = config["rp_ar_tag"]   # e.g. "兴源高中"
    rp_names = set(config["related_parties"])

    # ── 1. 收入成本明细表 ──
    ic = _read_sheet(file_ic, SHEET_ALIASES["income_cost"])
    ic = _clean_df(
        ic,
        num_cols=["本期借方"],
        str_cols=["摘要", "科目", "类型", "合作/自营", "区域", "标准名称",
                  "入场时间", "撤场时间", "分析是否剔除"],
        filter_col="分析是否剔除"
    )
    # 标记关联方
    ic["is_related"] = ic["标准名称"].isin(rp_names)
    data["ic"] = ic
    data["ic_rev"] = ic[ic["科目"] == "收入"]
    data["ic_cst"] = ic[ic["科目"] == "成本"]
    data["years"] = sorted(ic["摘要"].unique().tolist())

    # ── 2. 三大报表 ──
    fs_raw = {}
    try:
        xl_fs = pd.ExcelFile(file_fs)
        sheets = xl_fs.sheet_names
        for key, candidates in FS_SHEETS.items():
            for c in candidates:
                if c in sheets:
                    fs_raw[key] = pd.read_excel(file_fs, sheet_name=c, index_col=0)
                    break
            if key not in fs_raw and sheets:
                # 按顺序取
                idx = {"income_stmt": 0, "balance_sheet": 1, "cash_flow": 2}
                i = idx.get(key, 0)
                if i < len(sheets):
                    fs_raw[key] = pd.read_excel(file_fs, sheet_name=sheets[i], index_col=0)
    except Exception as e:
        raise ValueError(f"三大报表读取失败：{e}")
    data["fs"] = fs_raw

    # ── 3. 应收账款表 ──
    ar = _read_sheet(file_ar, SHEET_ALIASES["ar"])
    ar = _clean_df(
        ar,
        num_cols=["期末余额"],
        str_cols=["摘要", "内外部", "合作/自有", "区域", "业态", "标准项目", "分析是否剔除"],
        filter_col="分析是否剔除"
    )
    # 字段统一：合作/自有 → 合作/自营，业态 → 类型，标准项目 → 标准名称
    if "合作/自有" in ar.columns:
        ar = ar.rename(columns={"合作/自有": "合作/自营"})
    if "业态" in ar.columns:
        ar = ar.rename(columns={"业态": "类型"})
    if "标准项目" in ar.columns:
        ar = ar.rename(columns={"标准项目": "标准名称"})

    # 主体AR：自有 + 外部 + 非关联方
    ar_main = ar[
        (ar.get("合作/自营", ar.get("合作/自有", pd.Series(dtype=str))) == "自有") &
        (ar["内外部"] == "外部")
    ].copy()
    # 关联方AR
    ar_rp = ar[ar["内外部"] == rp_tag].copy()
    data["ar"] = ar
    data["ar_main"] = ar_main
    data["ar_rp"] = ar_rp

    # ── 4. 应付账款表 ──
    ap = _read_sheet(file_ap, SHEET_ALIASES["ap"])
    ap = _clean_df(
        ap,
        num_cols=["期末余额"],
        str_cols=["摘要", "内外部", "合作/自营", "公司", "类型", "标准项目"],
    )
    if "标准项目" in ap.columns:
        ap = ap.rename(columns={"标准项目": "标准名称"})
    if "公司" in ap.columns:
        ap = ap.rename(columns={"公司": "区域"})
    # AP：自有 + 非内部 + 非工资
    ap_main = ap[
        (ap.get("合作/自营", pd.Series("自有", index=ap.index)) == "自有") &
        (~ap["内外部"].isin(["内部", "工资"]))
    ].copy()
    data["ap"] = ap
    data["ap_main"] = ap_main

    # ── 5. 预收账款表 ──
    adv = _read_sheet(file_adv, SHEET_ALIASES["adv"])
    adv = _clean_df(
        adv,
        num_cols=["期末余额"],
        str_cols=["摘要", "内外部", "合作/自营", "公司", "类型", "标准项目"],
    )
    if "标准项目" in adv.columns:
        adv = adv.rename(columns={"标准项目": "标准名称"})
    if "公司" in adv.columns:
        adv = adv.rename(columns={"公司": "区域"})
    adv_main = adv[
        (adv.get("合作/自营", pd.Series("自有", index=adv.index)) == "自有") &
        (adv["内外部"] != "内部")
    ].copy()
    data["adv"] = adv
    data["adv_main"] = adv_main

    # ── 6. 项目级经营现金流（可选）──
    if file_cf is not None:
        try:
            cf = _read_sheet(file_cf, SHEET_ALIASES["cf"])
            cf = _clean_df(
                cf,
                num_cols=["值"],
                str_cols=["摘要", "现金流量表项", "合作/自营", "区域", "标准名称",
                          "收支", "分析是否剔除"],
                filter_col="分析是否剔除"
            )
            cf_own = cf[cf.get("合作/自营", pd.Series("自有", index=cf.index)) == "自有"]
            data["cf"] = cf
            data["cf_own"] = cf_own
            data["has_cf"] = True
        except Exception:
            data["has_cf"] = False
    else:
        data["has_cf"] = False

    data["config"] = config
    return data

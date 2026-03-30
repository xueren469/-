"""
loader.py — 数据加载与验证模块  v1.3

关联方处理原则：
  - 收入成本表、现金流表：河南区域整体分析，不拆分关联方
  - 应收账款表：单独拆出关联方，在第三章展示净敞口和周转天数差异
  - 关联方识别依据：标准项目名称匹配界面配置的关联方清单

分析是否剔除生效范围：
  - 收入成本表：1=剔除（已统一为0/1格式）
  - 现金流表：1或'剔除'=过滤（兼容两种格式）
  - 应收/应付/预收：全量读入，不过滤

字段名映射（你的字段名 → 系统标准名）：
  年份→摘要 / 标准项目名称→标准名称 / 管理区域→区域
  管理类型→合作/自营 / 业态类型→类型 / 金额→值（现金流）
  分析用收支→收支 / 分析用标准项目名称→标准名称（现金流优先）
  分析用管理区域→区域 / 分析用管理类型→合作/自营
"""

import re
import pandas as pd
import numpy as np


# ══════════════════════════════════════════
# Sheet 名称候选（按优先级排列）
# ══════════════════════════════════════════
SHEET_IC  = ["项目收入成本表", "Sheet1", "收入成本"]
SHEET_AR  = ["项目应收账款表", "应收账款表", "应收账款"]
SHEET_AP  = ["项目应付账款表", "应付账款表", "应付账款"]
SHEET_ADV = ["项目预收账款表", "预收账款表", "预收账款"]
SHEET_CF  = ["经营性现金流", "全部数据", "Sheet1"]
FS_SHEETS = {
    "income_stmt":   ["利润表", "合并利润表", "损益表"],
    "balance_sheet": ["资产负债表", "合并资产负债表"],
    "cash_flow":     ["现金流量表", "合并现金流量表"],
}


# ══════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════

def _read_sheet(file_obj, candidates):
    xl = pd.ExcelFile(file_obj)
    sheets = xl.sheet_names
    for c in candidates:
        if c in sheets:
            return pd.read_excel(file_obj, sheet_name=c)
    return pd.read_excel(file_obj, sheet_name=0)


def _clean_cols(df):
    """去掉列名里多余的空格"""
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    return df


def _rename(df, mapping):
    """只重命名存在的字段"""
    return df.rename(columns={k: v for k, v in mapping.items() if k in df.columns})


def _to_float(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0)
    return df


def _to_str(df, cols):
    for c in cols:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()
    return df


def _apply_drop_filter(df, col="分析是否剔除"):
    """
    通用剔除过滤：支持 0/1 和 0/'剔除' 两种格式
      保留：0 / '0' / '否' / 'FALSE' / '0.0'
      剔除：1 / '1' / '是' / '剔除' / 'TRUE'
    """
    if col not in df.columns:
        return df
    drop_vals = {"1", "是", "剔除", "true", "TRUE", "1.0"}
    mask = df[col].astype(str).str.strip().str.lower().isin(
        {"1", "是", "剔除", "true", "1.0"}
    )
    return df[~mask].copy()


def _read_fs_sheet(file_fs, candidates):
    """读三大报表中的一张，返回以科目为索引的DataFrame"""
    xl = pd.ExcelFile(file_fs)
    sheets = xl.sheet_names
    target = next((c for c in candidates if c in sheets), None)
    if target is None:
        return None
    df = pd.read_excel(file_fs, sheet_name=target, header=0)
    df = _clean_cols(df)
    df = df.rename(columns={df.columns[0]: "科目"})
    df["科目"] = df["科目"].astype(str).str.strip()
    for col in df.columns[1:]:
        df[col] = pd.to_numeric(df[col], errors="coerce")
    df = df.dropna(subset=["科目"])
    df = df[~df["科目"].isin(["nan", "---", "　", ""])]
    return df.set_index("科目")


# ══════════════════════════════════════════
# 主加载函数
# ══════════════════════════════════════════

def load_and_validate(file_ic, file_fs, file_ar, file_ap, file_adv, file_cf, config):
    data = {}
    year     = config["report_year"]        # "2025年"
    base     = config["base_year"]          # "2024年"
    rp_names = set(config["related_parties"])  # 关联方标准项目名称集合（仅AR表使用）


    # ══ 1. 收入成本明细表 ══
    # 关联方不拆分，整体纳入河南区域分析
    # 分析是否剔除：1=剔除，0=保留（支持旧格式 是/否）
    ic = _read_sheet(file_ic, SHEET_IC)
    ic = _clean_cols(ic)
    ic = _rename(ic, {
        "年份":       "摘要",
        "金额":       "本期借方",
        "标准项目名称": "标准名称",
        "业态类型":   "类型",
        "管理类型":   "合作/自营",
        "管理区域":   "区域",
        "分析用科目": "科目",
    })
    ic = _apply_drop_filter(ic, "分析是否剔除")
    ic = _to_float(ic, ["本期借方"])
    ic = _to_str(ic, ["摘要", "科目", "类型", "合作/自营", "区域",
                       "标准名称", "入场时间", "撤场时间"])

    # 标记关联方（收入成本表也标记，供第二章使用）
    # 兼容列名映射前后两种状态
    _name_col = "标准名称" if "标准名称" in ic.columns else (
                "标准项目名称" if "标准项目名称" in ic.columns else None)
    ic["is_related"] = ic[_name_col].isin(rp_names) if _name_col else False

    data["ic"]     = ic
    data["ic_rev"] = ic[ic["科目"] == "收入"]
    data["ic_cst"] = ic[ic["科目"] == "成本"]
    data["years"]  = sorted(ic["摘要"].unique().tolist())


    # ══ 2. 三大报表 ══
    fs_raw = {}
    try:
        # 利润表
        pl = _read_fs_sheet(file_fs, FS_SHEETS["income_stmt"])
        if pl is not None:
            cols = pl.columns.tolist()
            if len(cols) >= 2 and year not in cols:
                pl = pl.rename(columns={cols[0]: year, cols[1]: base})
        fs_raw["income_stmt"] = pl

        # 资产负债表（第1列=期末/当年，第2列=年初/上年）
        bs = _read_fs_sheet(file_fs, FS_SHEETS["balance_sheet"])
        if bs is not None:
            cols = bs.columns.tolist()
            if len(cols) >= 2:
                bs = bs.rename(columns={cols[0]: year, cols[1]: base})
        fs_raw["balance_sheet"] = bs

        # 现金流量表
        cf_stmt = _read_fs_sheet(file_fs, FS_SHEETS["cash_flow"])
        if cf_stmt is not None:
            cols = cf_stmt.columns.tolist()
            if len(cols) >= 2 and year not in cols:
                cf_stmt = cf_stmt.rename(columns={cols[0]: year, cols[1]: base})
        fs_raw["cash_flow"] = cf_stmt

    except Exception as e:
        raise ValueError(f"三大报表读取失败：{e}")
    data["fs"] = fs_raw


    # ══ 3. 应收账款表 ══
    # 关联方通过标准名称匹配识别，单独拆出供第三章使用
    # 不使用分析是否剔除（全量读入）
    ar = _read_sheet(file_ar, SHEET_AR)
    ar = _clean_cols(ar)
    ar = _rename(ar, {
        "年份":       "摘要",
        "标准项目名称": "标准名称",
        "业态类型":   "类型",
        "管理类型":   "合作/自营",
        "管理区域":   "区域",
    })
    ar = _to_float(ar, ["期末余额"])
    ar = _to_str(ar, ["摘要", "标准名称", "类型", "合作/自营", "区域"])

    # 关联方识别：匹配配置的项目名称清单
    ar["is_related"] = ar["标准名称"].isin(rp_names)

    # 主体AR：自有 + 非关联方
    zy = "合作/自营"
    ar_main = ar[
        (~ar["is_related"]) &
        (ar[zy] == "自有" if zy in ar.columns else True)
    ].copy()

    # 关联方AR：供第三章单独列示
    ar_rp = ar[ar["is_related"]].copy()

    data["ar"]      = ar
    data["ar_main"] = ar_main
    data["ar_rp"]   = ar_rp


    # ══ 4. 应付账款表 ══
    # 全量读入，自有口径
    ap = _read_sheet(file_ap, SHEET_AP)
    ap = _clean_cols(ap)
    ap = _rename(ap, {
        "年份":       "摘要",
        "标准项目名称": "标准名称",
        "业态类型":   "类型",
        "管理类型":   "合作/自营",
        "管理区域":   "区域",
        "省市":       "省份",
    })
    ap = _to_float(ap, ["期末余额"])
    ap = _to_str(ap, ["摘要", "标准名称", "合作/自营", "区域"])

    zy_ap = "合作/自营" if "合作/自营" in ap.columns else None
    ap_main = ap[(ap[zy_ap] == "自有") if zy_ap else True].copy()

    data["ap"]      = ap
    data["ap_main"] = ap_main


    # ══ 5. 预收账款表 ══
    # 全量读入，自有口径
    adv = _read_sheet(file_adv, SHEET_ADV)
    adv = _clean_cols(adv)
    adv = _rename(adv, {
        "年份":       "摘要",
        "标准项目名称": "标准名称",
        "业态类型":   "类型",
        "管理类型":   "合作/自营",
        "管理区域":   "区域",
        "省市":       "省份",
    })
    adv = _to_float(adv, ["期末余额"])
    adv = _to_str(adv, ["摘要", "标准名称", "合作/自营", "区域"])

    zy_adv = "合作/自营" if "合作/自营" in adv.columns else None
    adv_main = adv[(adv[zy_adv] == "自有") if zy_adv else True].copy()

    data["adv"]      = adv
    data["adv_main"] = adv_main


    # ══ 6. 项目级经营现金流（可选）══
    # 关联方不拆分，整体纳入河南区域
    # 分析是否剔除：1/'剔除'=过滤（排除异常小额数据）
    if file_cf is not None:
        try:
            cf = _read_sheet(file_cf, SHEET_CF)
            cf = _clean_cols(cf)

            # 分析用字段优先（比原始字段更准确）
            if "分析用标准项目名称" in cf.columns:
                cf["标准名称"] = cf["分析用标准项目名称"]
            if "分析用管理区域" in cf.columns:
                cf["区域"] = cf["分析用管理区域"]
            if "分析用管理类型" in cf.columns:
                cf["合作/自营"] = cf["分析用管理类型"]
            if "分析用收支" in cf.columns:
                cf["收支"] = cf["分析用收支"]

            cf = _rename(cf, {
                "年份": "摘要",
                "金额": "值",
            })

            # 过滤异常数据（芜湖佳景等非物业主营业务的小额数据）
            cf = _apply_drop_filter(cf, "分析是否剔除")

            cf = _to_float(cf, ["值"])
            cf = _to_str(cf, ["摘要", "现金流量表项", "收支",
                               "标准名称", "区域", "合作/自营"])

            # 标记关联方（供第四章拆分河南区域用）
            cf["is_related"] = cf["标准名称"].isin(rp_names)

            zy_cf = "合作/自营" if "合作/自营" in cf.columns else None
            cf_own     = cf[(cf[zy_cf] == "自有") if zy_cf else True].copy()
            # 自有·非关联方（纯市场化口径）
            cf_mkt     = cf_own[~cf_own["is_related"]].copy()
            # 自有·关联方（供单独展示）
            cf_rp      = cf_own[cf_own["is_related"]].copy()

            data["cf"]     = cf
            data["cf_own"] = cf_own
            data["cf_mkt"] = cf_mkt   # 纯市场化
            data["cf_rp"]  = cf_rp    # 关联方
            data["has_cf"] = True

        except Exception as e:
            data["has_cf"] = False
    else:
        data["has_cf"] = False

    data["config"] = config
    return data

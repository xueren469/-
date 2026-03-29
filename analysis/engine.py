"""
engine.py — 五章分析引擎
接收标准化数据字典，输出各章分析结果字典
"""
import pandas as pd
import numpy as np


def _piv(df, grp, val="本期借方"):
    """按区域/类型分组做年份透视"""
    if grp not in df.columns:
        return pd.DataFrame()
    p = df.groupby([grp, "摘要"])[val].sum().unstack("摘要").fillna(0)
    return p


def _piv2(df, grp1, grp2, val="本期借方"):
    p = df.groupby([grp1, grp2, "摘要"])[val].sum().unstack("摘要").fillna(0)
    return p


def _gm(rev, cst):
    """毛利率"""
    return (rev - cst) / rev.replace(0, np.nan)


def _get_fs_value(fs_dict, sheet_key, row_keywords, col_keyword=None, default=0.0):
    """从三大报表中按关键字提取数值"""
    df = fs_dict.get(sheet_key)
    if df is None:
        return default
    try:
        for kw in row_keywords:
            matched = [i for i in df.index if kw in str(i)]
            if matched:
                row = df.loc[matched[0]]
                if col_keyword:
                    col_matched = [c for c in df.columns if col_keyword in str(c)]
                    if col_matched:
                        return float(row[col_matched[0]])
                else:
                    # 取最后一列（通常是当年）
                    return float(row.iloc[-1])
    except Exception:
        pass
    return default


# ══════════════════════════════════════════════════════════
# 第一章
# ══════════════════════════════════════════════════════════
def chapter1(data, config):
    fs = data.get("fs", {})
    ic = data["ic"]
    year = config["report_year"]
    base = config["base_year"]
    rev_target = config["revenue_target"]

    # 从收入成本表计算收入/成本（主口径）
    ic_mkt = ic[~ic["is_related"]]
    rev_25 = ic_mkt[(ic_mkt["摘要"] == year) & (ic_mkt["科目"] == "收入")]["本期借方"].sum()
    rev_24 = ic_mkt[(ic_mkt["摘要"] == base) & (ic_mkt["科目"] == "收入")]["本期借方"].sum()
    cst_25 = ic_mkt[(ic_mkt["摘要"] == year) & (ic_mkt["科目"] == "成本")]["本期借方"].sum()
    cst_24 = ic_mkt[(ic_mkt["摘要"] == base) & (ic_mkt["科目"] == "成本")]["本期借方"].sum()

    gm_25 = rev_25 - cst_25
    gm_24 = rev_24 - cst_24
    gm_rate_25 = gm_25 / rev_25 if rev_25 > 0 else 0
    gm_rate_24 = gm_24 / rev_24 if rev_24 > 0 else 0
    rev_growth = (rev_25 - rev_24) / rev_24 if rev_24 > 0 else 0
    cst_growth = (cst_25 - cst_24) / cst_24 if cst_24 > 0 else 0

    # 从三大报表提取利润表科目
    def fv(keywords, sheet="income_stmt", col=None):
        return _get_fs_value(fs, sheet, keywords, col)

    # 尝试从利润表提取费用数据
    sell_25 = fv(["销售费用"])
    sell_24 = fv(["销售费用"])
    mgmt_25 = fv(["管理费用"])
    mgmt_24 = fv(["管理费用"])
    fin_25  = fv(["财务费用"])
    fin_24  = fv(["财务费用"])
    op_25   = fv(["营业利润"])
    net_25  = fv(["净利润", "归属于母公司"])
    net_24  = fv(["净利润", "归属于母公司"])
    other_income_25 = fv(["其他收益", "资产处置"])
    impair_25 = fv(["减值", "信用减值"])
    rd_25   = fv(["研发费用", "研发"])

    # 如果利润表读取失败（全为0），用收入成本表估算
    if net_25 == 0:
        net_25 = gm_25 * 0.25  # 估算
    if net_24 == 0:
        net_24 = gm_24 * 0.25

    # 资产负债表
    total_asset_25 = fv(["资产合计", "总资产"], "balance_sheet")
    total_liab_25  = fv(["负债合计", "总负债"], "balance_sheet")
    cash_25        = fv(["货币资金"], "balance_sheet")
    st_loan_25     = fv(["短期借款"], "balance_sheet")
    lt_loan_25     = fv(["长期借款"], "balance_sheet")
    goodwill_25    = fv(["商誉"], "balance_sheet")
    cur_asset_25   = fv(["流动资产合计"], "balance_sheet")
    cur_liab_25    = fv(["流动负债合计"], "balance_sheet")

    da_ratio_25 = total_liab_25 / total_asset_25 if total_asset_25 > 0 else 0
    current_ratio = cur_asset_25 / cur_liab_25 if cur_liab_25 > 0 else 0
    total_debt = st_loan_25 + lt_loan_25
    net_debt = total_debt - cash_25

    # 现金流量表
    ocf = fv(["经营活动产生的现金流量净额", "经营活动净额"], "cash_flow")
    capex = fv(["购建固定资产", "资本支出"], "cash_flow")
    fcf = ocf - abs(capex)
    cash_from_sales = fv(["销售商品", "提供劳务"], "cash_flow")
    sales_collection_rate = cash_from_sales / rev_25 if rev_25 > 0 else 0

    # EBITDA估算
    fa_25 = fv(["固定资产", "使用权资产"], "balance_sheet")
    fa_24_est = fa_25 * 1.05
    est_da = max(fa_24_est - fa_25, 0) + max(goodwill_25 * 0.02, 0)
    interest_25 = abs(fin_25)
    ebitda = op_25 + est_da + interest_25
    debt_ebitda = total_debt / ebitda if ebitda > 0 else 0
    interest_cover = op_25 / interest_25 if interest_25 > 0 else 99

    # 一次性收益识别（取其他收益中较大的部分）
    onetime_gain = other_income_25  # 简化处理
    adj_net_25 = net_25 - onetime_gain * 0.75  # 税后
    ocf_quality = ocf / adj_net_25 if adj_net_25 > 0 else 0

    res = {
        # 收入
        "rev_25": rev_25, "rev_24": rev_24,
        "rev_growth": rev_growth, "rev_target": rev_target / 100,
        "cst_25": cst_25, "cst_24": cst_24, "cst_growth": cst_growth,
        # 毛利
        "gm_25": gm_25, "gm_24": gm_24,
        "gm_rate_25": gm_rate_25, "gm_rate_24": gm_rate_24,
        # 费用
        "sell_25": sell_25, "sell_24": sell_24,
        "mgmt_25": mgmt_25, "mgmt_24": mgmt_24,
        "fin_25": fin_25, "fin_24": fin_24,
        "rd_25": rd_25, "impair_25": impair_25,
        "other_income_25": other_income_25, "onetime_gain": onetime_gain,
        # 利润
        "op_25": op_25, "net_25": net_25, "net_24": net_24,
        "adj_net_25": adj_net_25,
        # 资产负债
        "da_ratio_25": da_ratio_25, "current_ratio": current_ratio,
        "total_debt": total_debt, "net_debt": net_debt,
        "goodwill_25": goodwill_25, "total_asset_25": total_asset_25,
        "cash_25": cash_25, "st_loan_25": st_loan_25,
        # 偿债
        "debt_ebitda": debt_ebitda, "interest_cover": interest_cover,
        "ebitda": ebitda,
        # 现金流
        "ocf": ocf, "fcf": fcf, "ocf_quality": ocf_quality,
        "sales_collection_rate": sales_collection_rate,
    }
    return res


# ══════════════════════════════════════════════════════════
# 第二章
# ══════════════════════════════════════════════════════════
def chapter2(data, config):
    ic = data["ic"]
    year = config["report_year"]
    base = config["base_year"]
    rp_names = set(config["related_parties"])

    rev = data["ic_rev"]
    cst = data["ic_cst"]

    # 主体（剔关联方）
    rev_m = rev[~rev["is_related"]]
    cst_m = cst[~cst["is_related"]]

    # ── 四象限 ──
    def quadrant(r, c, cat_col, cat_val, mode_col, mode_val):
        rf = r[(r[cat_col] == cat_val) & (r[mode_col] == mode_val)]
        cf = c[(c[cat_col] == cat_val) & (c[mode_col] == mode_val)]
        r25 = rf[rf["摘要"] == year]["本期借方"].sum()
        r24 = rf[rf["摘要"] == base]["本期借方"].sum()
        c25 = cf[cf["摘要"] == year]["本期借方"].sum()
        c24 = cf[cf["摘要"] == base]["本期借方"].sum()
        return {
            "rev_25": r25, "rev_24": r24,
            "rev_chg": (r25 - r24) / r24 if r24 > 0 else 0,
            "gm_rate_25": (r25 - c25) / r25 if r25 > 0 else 0,
            "gm_rate_24": (r24 - c24) / r24 if r24 > 0 else 0,
            "gm_25": r25 - c25, "gm_24": r24 - c24,
        }

    quadrants = {
        "住宅·自有": quadrant(rev_m, cst_m, "类型", "居民生活服务", "合作/自营", "自有"),
        "住宅·合作": quadrant(rev_m, cst_m, "类型", "居民生活服务", "合作/自营", "合作"),
        "后勤·自有": quadrant(rev_m, cst_m, "类型", "企事业后勤综合服务", "合作/自营", "自有"),
        "后勤·合作": quadrant(rev_m, cst_m, "类型", "企事业后勤综合服务", "合作/自营", "合作"),
    }

    # ── 破桶效应：按项目类型归因 ──
    def classify_project(row):
        entry = row.get("入场时间", "")
        exit_ = row.get("撤场时间", "0")
        if str(exit_) not in ["0", "nan", "", "0.0"]:
            return "撤场项目"
        if str(entry) == year:
            return f"{year}新入场"
        if str(entry) == base:
            return f"{base}新入场"
        if str(entry) in ["2022年", "2023年"]:
            return "成长期(22-23年入场)"
        if str(entry) == "存续":
            return "存续老项目"
        return "其他"

    rev_m2 = rev_m.copy()
    rev_m2["项目类型"] = rev_m2.apply(classify_project, axis=1)
    cst_m2 = cst_m.copy()
    cst_m2["项目类型"] = cst_m2.apply(classify_project, axis=1)

    rev_by_type = rev_m2.groupby(["项目类型", "摘要"])["本期借方"].sum().unstack("摘要").fillna(0)
    cst_by_type = cst_m2.groupby(["项目类型", "摘要"])["本期借方"].sum().unstack("摘要").fillna(0)

    type_analysis = {}
    for t in rev_by_type.index:
        r25 = rev_by_type.loc[t].get(year, 0)
        r24 = rev_by_type.loc[t].get(base, 0)
        c25 = cst_by_type.loc[t].get(year, 0) if t in cst_by_type.index else 0
        c24 = cst_by_type.loc[t].get(base, 0) if t in cst_by_type.index else 0
        type_analysis[t] = {
            "rev_25": r25, "rev_24": r24,
            "rev_chg": r25 - r24,
            "gm_rate_25": (r25 - c25) / r25 if r25 > 0 else 0,
            "gm_rate_24": (r24 - c24) / r24 if r24 > 0 else 0,
        }

    # ── 区域毛利排行 ──
    if "区域" in rev_m.columns:
        r_rgn = rev_m.groupby(["区域", "摘要"])["本期借方"].sum().unstack("摘要").fillna(0)
        c_rgn = cst_m.groupby(["区域", "摘要"])["本期借方"].sum().unstack("摘要").fillna(0)
        region_analysis = {}
        for rgn in r_rgn.index:
            r25 = r_rgn.loc[rgn].get(year, 0)
            r24 = r_rgn.loc[rgn].get(base, 0)
            c25 = c_rgn.loc[rgn].get(year, 0) if rgn in c_rgn.index else 0
            c24 = c_rgn.loc[rgn].get(base, 0) if rgn in c_rgn.index else 0
            if r25 > 0:
                region_analysis[rgn] = {
                    "rev_25": r25, "rev_24": r24,
                    "gm_rate_25": (r25 - c25) / r25,
                    "gm_rate_24": (r24 - c24) / r24 if r24 > 0 else 0,
                    "gm_ppt": ((r25-c25)/r25 - (r24-c24)/r24) if r24 > 0 else 0,
                    "gm_25": r25 - c25, "gm_24": r24 - c24,
                    "gm_chg": (r25-c25) - (r24-c24),
                }
    else:
        region_analysis = {}

    # ── 项目毛利率分布 ──
    r_proj = rev_m[rev_m["摘要"] == year].groupby("标准名称")["本期借方"].sum()
    c_proj = cst_m[cst_m["摘要"] == year].groupby("标准名称")["本期借方"].sum()
    proj = pd.DataFrame({"rev": r_proj, "cst": c_proj}).fillna(0)
    proj = proj[proj["rev"] > 50000]  # 收入>5万
    proj["gm_rate"] = (proj["rev"] - proj["cst"]) / proj["rev"]
    proj["loss"] = proj["cst"] - proj["rev"]

    bins = [-np.inf, 0, 0.05, 0.10, 0.20, np.inf]
    labels = ["亏损(<0%)", "微利(0-5%)", "正常(5-10%)", "良好(10-20%)", "优质(>20%)"]
    proj["band"] = pd.cut(proj["gm_rate"], bins=bins, labels=labels)
    proj_dist = proj.groupby("band", observed=True).agg(
        count=("gm_rate", "count"),
        rev_sum=("rev", "sum")
    ).reset_index()

    loss_proj = proj[proj["gm_rate"] < 0].sort_values("loss", ascending=False).head(10)

    return {
        "quadrants": quadrants,
        "type_analysis": type_analysis,
        "region_analysis": region_analysis,
        "proj_dist": proj_dist.to_dict("records"),
        "loss_top10": loss_proj.reset_index()[["标准名称", "rev", "gm_rate", "loss"]].to_dict("records"),
    }


# ══════════════════════════════════════════════════════════
# 第三章
# ══════════════════════════════════════════════════════════
def chapter3(data, config):
    year = config["report_year"]
    base = config["base_year"]
    rp_tag = config["rp_ar_tag"]

    ar_m = data["ar_main"]
    ap_m = data["ap_main"]
    adv_m = data["adv_main"]
    ar_rp = data["ar_rp"]
    ic_rev = data["ic_rev"]

    # ── AR 5年趋势 ──
    ar_trend = ar_m.groupby("摘要")["期末余额"].sum().sort_index()

    # ── 关联方 AR 历年趋势 ──
    rp_trend = ar_rp.groupby("摘要")["期末余额"].sum().sort_index()

    # ── 按区域计算净资金敞口 ──
    def get_region_balance(df, region_col="区域"):
        if region_col not in df.columns:
            return pd.DataFrame()
        return df[df["摘要"] == year].groupby(region_col)["期末余额"].sum()

    def get_region_balance_base(df, region_col="区域"):
        if region_col not in df.columns:
            return pd.Series(dtype=float)
        return df[df["摘要"] == base].groupby(region_col)["期末余额"].sum()

    # AR
    ar_r25 = get_region_balance(ar_m, "区域")
    ar_r24 = get_region_balance_base(ar_m, "区域")
    ar_avg = (ar_r25 + ar_r24.reindex(ar_r25.index, fill_value=0)) / 2

    # AP
    ap_col = "区域" if "区域" in ap_m.columns else "公司"
    ap_r25 = get_region_balance(ap_m, ap_col)

    # 预收
    adv_col = "区域" if "区域" in adv_m.columns else "公司"
    adv_r25 = get_region_balance(adv_m, adv_col)

    # 收入（用于周转天数）
    rev_r25 = ic_rev[
        (ic_rev["摘要"] == year) & (ic_rev["合作/自营"] == "自有") & (~ic_rev["is_related"])
    ].groupby("区域")["本期借方"].sum() if "区域" in ic_rev.columns else pd.Series(dtype=float)

    # 合并
    region_df = pd.DataFrame({
        "ar_25": ar_r25,
        "ar_24": ar_r24,
        "ar_avg": ar_avg,
        "ap_25": ap_r25.reindex(ar_r25.index, fill_value=0),
        "adv_25": adv_r25.reindex(ar_r25.index, fill_value=0),
        "rev_25": rev_r25.reindex(ar_r25.index, fill_value=0),
    }).fillna(0)

    region_df["gap"] = region_df["ar_25"] - region_df["ap_25"] - region_df["adv_25"]
    region_df["ar_days"] = np.where(
        region_df["rev_25"] > 0,
        360 * region_df["ar_avg"] / region_df["rev_25"], 0
    )
    region_df = region_df[region_df["rev_25"] > 0].sort_values("gap", ascending=False)

    # ── 分业态周转天数 ──
    def turnover_by_type(biz_type):
        ar_t = ar_m[ar_m["类型"] == biz_type] if "类型" in ar_m.columns else ar_m
        rev_t = ic_rev[
            (ic_rev["合作/自营"] == "自有") &
            (~ic_rev["is_related"]) &
            (ic_rev.get("类型", pd.Series("居民生活服务", index=ic_rev.index)) == biz_type)
        ]
        ar25 = ar_t[ar_t["摘要"] == year]["期末余额"].sum()
        ar24 = ar_t[ar_t["摘要"] == base]["期末余额"].sum()
        rev25 = rev_t[rev_t["摘要"] == year]["本期借方"].sum()
        avg = (ar25 + ar24) / 2
        days = 360 * avg / rev25 if rev25 > 0 else 0
        return {"ar_25": ar25, "ar_24": ar24, "rev_25": rev25, "days": days}

    # 全公司整体
    ar_total_25 = ar_m[ar_m["摘要"] == year]["期末余额"].sum()
    ar_total_24 = ar_m[ar_m["摘要"] == base]["期末余额"].sum()
    rev_total_25 = ic_rev[
        (ic_rev["摘要"] == year) & (ic_rev["合作/自营"] == "自有") & (~ic_rev["is_related"])
    ]["本期借方"].sum()
    overall_days = 360 * (ar_total_25 + ar_total_24) / 2 / rev_total_25 if rev_total_25 > 0 else 0

    # 关联方整体
    rp_25 = ar_rp[ar_rp["摘要"] == year]["期末余额"].sum()
    rp_24 = ar_rp[ar_rp["摘要"] == base]["期末余额"].sum()

    return {
        "ar_trend": ar_trend.to_dict(),
        "rp_trend": rp_trend.to_dict(),
        "region_df": region_df.round(0).to_dict("index"),
        "overall_days": overall_days,
        "residential_days": turnover_by_type("居民生活服务")["days"],
        "commercial_days": turnover_by_type("企事业后勤综合服务")["days"],
        "rp_ar_25": rp_25,
        "rp_ar_24": rp_24,
        "ar_total_25": ar_total_25,
    }


# ══════════════════════════════════════════════════════════
# 第四章
# ══════════════════════════════════════════════════════════
def chapter4(data, config):
    year = config["report_year"]
    base = config["base_year"]
    ch1 = data.get("_ch1", {})

    result = {
        "debt_ebitda": ch1.get("debt_ebitda", 0),
        "interest_cover": ch1.get("interest_cover", 0),
        "da_ratio_25": ch1.get("da_ratio_25", 0),
        "current_ratio": ch1.get("current_ratio", 0),
        "net_debt": ch1.get("net_debt", 0),
        "cash_25": ch1.get("cash_25", 0),
        "st_loan_25": ch1.get("st_loan_25", 0),
        "cash_cover": ch1.get("cash_25", 0) / ch1.get("st_loan_25", 1) if ch1.get("st_loan_25", 0) > 0 else 0,
        "ocf": ch1.get("ocf", 0),
        "fcf": ch1.get("fcf", 0),
        "ocf_quality": ch1.get("ocf_quality", 0),
        "net_25": ch1.get("net_25", 0),
        "adj_net_25": ch1.get("adj_net_25", 0),
        "onetime_gain": ch1.get("onetime_gain", 0),
        "sales_collection_rate": ch1.get("sales_collection_rate", 0),
        "region_cf": {},
        "has_cf": data.get("has_cf", False),
    }

    # ── 区域现金净额（需项目级现金流）──
    if data.get("has_cf"):
        cf_own = data["cf_own"]
        rp_names = set(config["related_parties"])

        # 纯市场化（剔关联方）
        cf_mkt = cf_own[~cf_own["标准名称"].isin(rp_names)] if "标准名称" in cf_own.columns else cf_own

        if "区域" in cf_mkt.columns and "现金流量表项" in cf_mkt.columns:
            cf_y = cf_mkt[cf_mkt["摘要"] == year]
            cash_in = cf_y[cf_y["现金流量表项"] == "销售商品、提供劳务收到的现金"].groupby("区域")["值"].sum()
            cash_out = cf_y[cf_y["现金流量表项"].isin([
                "购买商品、接受劳务支付的现金",
                "支付给职工以及为职工支付的现金"
            ])].groupby("区域")["值"].sum()
            cf_df = pd.DataFrame({"in": cash_in, "out": cash_out}).fillna(0)
            cf_df["net"] = cf_df["in"] - cf_df["out"]
            result["region_cf"] = cf_df.sort_values("net").round(0).to_dict("index")

    return result


# ══════════════════════════════════════════════════════════
# 第五章 — 综合结论
# ══════════════════════════════════════════════════════════
def chapter5(ch1, ch2, ch3, ch4, config):
    year = config["report_year"]
    rev_target = config["revenue_target"] / 100
    rev_growth = ch1.get("rev_growth", 0)
    gm_chg = ch1.get("gm_rate_25", 0) - ch1.get("gm_rate_24", 0)
    da_ratio = ch1.get("da_ratio_25", 0)
    net_debt = ch1.get("net_debt", 0)
    fcf = ch1.get("fcf", 0)
    ar_total = ch3.get("ar_total_25", 0)
    rp_ar = ch3.get("rp_ar_25", 0)

    # 评分逻辑
    def score_growth():
        gap = rev_growth - rev_target
        if gap >= 0: return "A"
        elif gap >= -0.05: return "B+"
        elif gap >= -0.10: return "B"
        elif gap >= -0.15: return "C+"
        else: return "C"

    def score_profit():
        if gm_chg >= 0.01: return "B+"
        elif gm_chg >= 0: return "B"
        elif gm_chg >= -0.005: return "C+"
        else: return "D+"

    def score_safety():
        if da_ratio < 0.5 and net_debt < 0: return "A"
        elif da_ratio < 0.6: return "B+"
        elif da_ratio < 0.7: return "B"
        else: return "C"

    def score_ar():
        ar_rev_ratio = ar_total / ch1.get("rev_25", 1)
        if ar_rev_ratio < 0.2: return "A"
        elif ar_rev_ratio < 0.3: return "B"
        elif ar_rev_ratio < 0.4: return "C+"
        else: return "D"

    def score_cf():
        ocf_q = ch1.get("ocf_quality", 0)
        if ocf_q > 1.2 and fcf > 0: return "A"
        elif ocf_q > 1.0: return "B+"
        elif ocf_q > 0.8: return "B"
        else: return "C"

    scores = {
        "收入增长质量": score_growth(),
        "盈利能力": score_profit(),
        "财务安全性": score_safety(),
        "应收账款管理": score_ar(),
        "现金流质量": score_cf(),
    }

    # 生成关键问题列表
    issues = []
    region_analysis = ch2.get("region_analysis", {})

    # 问题1：破桶效应
    surviving = ch2.get("type_analysis", {}).get("存续老项目", {})
    exit_proj = ch2.get("type_analysis", {}).get("撤场项目", {})
    if surviving.get("rev_chg", 0) < 0 or exit_proj.get("rev_24", 0) > 0:
        issues.append({
            "priority": "P1", "level": "red",
            "title": "破桶效应：存续老项目收入萎缩，增长依赖新项目输血",
            "detail": f"存续老项目收入变动{surviving.get('rev_chg', 0)/1e4:.0f}万，一旦拓新放缓将立即暴露",
            "chapter": "第二章"
        })

    # 问题2：毛利下滑最严重区域
    worst_regions = sorted(
        [(k, v) for k, v in region_analysis.items() if v.get("gm_ppt", 0) < -0.03],
        key=lambda x: x[1].get("gm_chg", 0)
    )[:2]
    for rgn, v in worst_regions:
        issues.append({
            "priority": "P2", "level": "red",
            "title": f"{rgn}毛利率大幅下滑（{v.get('gm_ppt',0)*100:.1f}ppt）",
            "detail": f"毛利额减少{abs(v.get('gm_chg',0))/1e4:.0f}万，需立即专项整治",
            "chapter": "第二章"
        })

    # 问题3：超长账期区域
    region_df = ch3.get("region_df", {})
    danger_ar = sorted(
        [(k, v) for k, v in region_df.items() if v.get("ar_days", 0) > 365],
        key=lambda x: -x[1].get("ar_days", 0)
    )[:2]
    for rgn, v in danger_ar:
        issues.append({
            "priority": "P3", "level": "red",
            "title": f"{rgn}应收账款周转{v.get('ar_days',0):.0f}天，形同准坏账",
            "detail": f"应收余额{v.get('ar_25',0)/1e4:.0f}万，需立即启动法律催收",
            "chapter": "第三章"
        })

    # 问题4：现金漏斗
    region_cf = ch4.get("region_cf", {})
    drain_regions = sorted(
        [(k, v) for k, v in region_cf.items() if v.get("net", 0) < -500000],
        key=lambda x: x[1].get("net", 0)
    )[:1]
    for rgn, v in drain_regions:
        issues.append({
            "priority": "P4", "level": "yellow",
            "title": f"{rgn}净消耗现金{abs(v.get('net',0))/1e4:.0f}万，是主要现金漏斗",
            "detail": "成本结构需系统性优化",
            "chapter": "第四章"
        })

    return {
        "scores": scores,
        "issues": issues[:9],  # 最多9项
    }


# ══════════════════════════════════════════════════════════
# 总入口
# ══════════════════════════════════════════════════════════
def run_all_chapters(data, config):
    year = config["report_year"]
    base = config["base_year"]

    ch1_res = chapter1(data, config)
    data["_ch1"] = ch1_res

    ch2_res = chapter2(data, config)
    ch3_res = chapter3(data, config)
    ch4_res = chapter4(data, config)
    ch5_res = chapter5(ch1_res, ch2_res, ch3_res, ch4_res, config)

    # KPI摘要（供Streamlit界面显示）
    rev = ch1_res.get("rev_25", 0)
    kpi_summary = {
        "revenue_25": f"{rev/1e8:.2f}亿元" if rev > 1e7 else f"{rev/1e4:.0f}万元",
        "revenue_chg": f"{ch1_res.get('rev_growth',0)*100:+.1f}%",
        "gm_rate_25": f"{ch1_res.get('gm_rate_25',0)*100:.1f}%",
        "gm_rate_chg": f"{(ch1_res.get('gm_rate_25',0)-ch1_res.get('gm_rate_24',0))*100:+.2f}ppt",
        "debt_ratio_25": f"{ch1_res.get('da_ratio_25',0)*100:.1f}%",
        "debt_ratio_chg": "▼ 降杠杆" if ch1_res.get("net_debt", 0) < 0 else "",
        "fcf": f"{ch1_res.get('fcf',0)/1e4:.0f}万元",
    }

    return {
        "ch1": ch1_res,
        "ch2": ch2_res,
        "ch3": ch3_res,
        "ch4": ch4_res,
        "ch5": ch5_res,
        "kpi_summary": kpi_summary,
        "config": config,
    }

# -*- coding: utf-8 -*-
"""按论文表1、表2口径构建企业面板：数据要素水平(DEL)与产业链韧性(RES)"""
import os
import numpy as np
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
PROV_XLSX = r"D:\软件库\微信聊天\xwechat_files\wxid_9xgpzj79zcod22_5ef8\temp\RWTemp\2026-11\9e20f478899dc29eb19741386f9343c8\省级面板数据2014-2024(1).xlsx"

SOE_KW = ["国有资产监督管理委员会", "国资委", "人民政府", "财政厅", "财政局", "国务院",
          "交通厅", "管理委员会", "国家铁路集团", "招商局集团", "兵器工业集团",
          "中信集团", "上海电气控股", "东浩兰生", "中央汇金", "全国社会保障基金"]
NONSOE_EXTRA = {"上海大众企业管理有限公司"}
EAST = ["北京", "天津", "河北", "辽宁", "上海", "江苏", "浙江", "福建", "山东", "广东", "海南"]


def load_wide():
    df = pd.read_csv(os.path.join(BASE, "raw_statements.csv"), dtype={"code": str})
    df = df[df.field != "ERROR"].copy()
    df["code"] = df["code"].str[-6:]
    w = df.pivot_table(index=["code", "year"], columns="field", values="value", aggfunc="last").reset_index()
    w["year"] = w["year"].astype(int)
    return w


def entropy(mat: pd.DataFrame, positive: dict):
    n = len(mat)
    norm = pd.DataFrame(index=mat.index)
    for c, pos in positive.items():
        x = mat[c].astype(float)
        lo, hi = x.min(), x.max()
        norm[c] = ((x - lo) / (hi - lo)) if pos else ((hi - x) / (hi - lo))
        norm[c] = norm[c].clip(0, 1).fillna(0)
    p = norm.div(norm.sum(axis=0).replace(0, np.nan), axis=1).fillna(0)
    k = 1 / np.log(n)
    with np.errstate(divide="ignore", invalid="ignore"):
        e = -k * (p * np.log(p.where(p > 0))).sum(axis=0)
    d = (1 - e).replace(0, np.nan)
    w = (d / d.sum()).fillna(1 / len(d))
    return (norm * w).sum(axis=1), w


def winsorize(s, lo=0.01, hi=0.99):
    return s.clip(s.quantile(lo), s.quantile(hi))


def main():
    w = load_wide()
    info = pd.read_csv(os.path.join(BASE, "firm_info.csv"), dtype={"code": str})
    df = w.merge(info[["code", "name", "sector", "province_raw", "actual_holder", "listing_date"]],
                 on="code", how="left")
    df = df.sort_values(["code", "year"]).reset_index(drop=True)

    # ---- 需跨年信息：先在全样本(2012-2023)上计算，再截取2015-2023 ----
    df["roa_calc"] = df.NETPROFIT / df.TOTASSET
    df["roa_std3"] = df.groupby("code")["roa_calc"].transform(
        lambda s: s.rolling(3, min_periods=3).std())
    df["growth"] = df.groupby("code")["BIZINCO"].pct_change()
    med = df.groupby(["sector", "year"])["growth"].transform("median")
    df["ind_adj_growth"] = df.growth - med

    df = df[(df.year >= 2015) & (df.year <= 2023)].copy()
    df = df[(df.TOTASSET > 0) & (df.BIZINCO > 0)].copy()

    # 平衡面板：保留2015—2023年连续9年均有观测的公司（剔除退市、被吸收合并等年度不全样本）
    cnt = df.groupby("code")["year"].nunique()
    keep = cnt[cnt == 9].index
    df = df[df.code.isin(keep)].copy()

    # 研发费用缺失按 0 处理（2018 年前多数公司未单列）
    df["rd_missing"] = df.DEVEEXPE.isna().astype(int)
    df["DEVEEXPE"] = df.DEVEEXPE.fillna(0)

    # 表1
    df["intan_ratio"] = df.INTAASSET / df.TOTASSET
    df["rd_intensity"] = df.DEVEEXPE / df.BIZINCO
    df["dcapex_intensity"] = df.ACQUASSETCASH / df.BIZINCO
    # 表2
    df["cash_ratio"] = df.MANANETR / df.TOTASSET
    df["cost_ratio"] = df.BIZCOST / df.BIZINCO
    df["asset_turn"] = df.BIZINCO / df.TOTASSET
    df["current_ratio"] = df.TOTCURRASSET / df.TOTALCURRLIAB
    # 控制/中介
    df["Size"] = np.log(df.TOTASSET)
    df["Lev"] = df.TOTLIAB / df.TOTASSET
    df["Roa"] = df.roa_calc
    df["Growth"] = df.growth
    df["Info"] = np.log1p(df.BIZINCO / df.ACCORECE)
    df["Eff"] = np.log1p(df.BIZINCO / df.FIXEDASSENET)
    df["Fin"] = df.FINEXPE / df.TOTLIAB * 100

    win_vars = ["intan_ratio", "rd_intensity", "dcapex_intensity", "roa_std3", "cash_ratio",
                "cost_ratio", "ind_adj_growth", "asset_turn", "current_ratio",
                "Size", "Lev", "Growth", "Roa", "Info", "Eff", "Fin"]
    for c in win_vars:
        df[c] = winsorize(df[c])

    def soe_flag(row):
        h = str(row.actual_holder)
        if h in ("nan", "无", ""):
            return np.nan
        if h in NONSOE_EXTRA:
            return 0
        return 1 if any(k in h for k in SOE_KW) else 0

    df["soe"] = df.apply(soe_flag, axis=1)
    df["east"] = df.province_raw.isin(EAST).astype(int)
    df["infra"] = (df.sector == "公路桥梁").astype(int)

    # 熵值法：三维度分别赋权后再合成 DEL
    df["DE1"], _ = entropy(df[["intan_ratio"]], {"intan_ratio": True})
    df["DE2"], _ = entropy(df[["rd_intensity"]], {"rd_intensity": True})
    df["DE3"], _ = entropy(df[["dcapex_intensity"]], {"dcapex_intensity": True})
    parts = df[["DE1", "DE2", "DE3"]]
    p = parts.div(parts.sum(axis=0).replace(0, np.nan), axis=1).fillna(0)
    e = -1 / np.log(len(df)) * (p * np.log(p.where(p > 0))).sum(axis=0)
    dw = (1 - e) / (1 - e).sum()
    df["DEL"] = (parts * dw).sum(axis=1)

    df["RESIST"], wr = entropy(df[["roa_std3", "cash_ratio", "cost_ratio"]],
                               {"roa_std3": False, "cash_ratio": True, "cost_ratio": False})
    df["RECOVER"], wc = entropy(df[["ind_adj_growth", "asset_turn", "current_ratio"]],
                                {"ind_adj_growth": True, "asset_turn": True, "current_ratio": True})
    df["RES"] = (df.RESIST + df.RECOVER) / 2
    df["RES_geo"] = np.sqrt(df.RESIST.clip(lower=0) * df.RECOVER.clip(lower=0))
    # 六指标统一熵权（稳健性替代测度）
    df["RES_alt"], _ = entropy(df[["roa_std3", "cash_ratio", "cost_ratio",
                                   "ind_adj_growth", "asset_turn", "current_ratio"]],
                               {"roa_std3": False, "cash_ratio": True, "cost_ratio": False,
                                "ind_adj_growth": True, "asset_turn": True, "current_ratio": True})

    prov = pd.read_excel(PROV_XLSX, sheet_name="面板数据").rename(columns={"province": "province_raw"})
    df = df.merge(prov, on=["province_raw", "year"], how="left")
    for c in ["roaddensity", "freight", "pgdp", "urban"]:
        df["z_" + c] = (df[c] - df[c].mean()) / df[c].std()
    df["IV"] = df[["z_roaddensity", "z_freight", "z_pgdp", "z_urban"]].mean(axis=1)

    df = df.rename(columns={"code": "stkcd", "name": "firm", "province_raw": "province"})
    df["year"] = df.year.astype(int)
    df["stkcd"] = df.stkcd.astype(str)
    df.to_csv(os.path.join(BASE, "firm_panel_raw.csv"), index=False, encoding="utf-8-sig")

    print("样本量", len(df), "公司数", df.stkcd.nunique(), "年份", df.year.min(), df.year.max())
    print("DEL权重", dw.round(4).to_dict())
    print("抵抗维度", wr.round(4).to_dict())
    print("恢复维度", wc.round(4).to_dict())
    show = ["RES", "RESIST", "RECOVER", "DEL", "DE1", "DE2", "DE3", "Size", "Lev", "Growth", "Roa"]
    print(df[show].describe().T[["count", "mean", "std", "min", "50%", "max"]].round(3).to_string())
    print("soe", df.soe.value_counts(dropna=False).to_dict(), "east", df.east.mean().round(3))


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""外部交叉核验：以东方财富独立报表数据比对新浪口径抓取结果"""
import json, os, time
import numpy as np
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = os.path.dirname(os.path.abspath(__file__))
URL = "https://datacenter.eastmoney.com/securities/api/data/v1/get"
HDR = {"User-Agent": "Mozilla/5.0"}

MAP = {
    "TOTASSET": ("GBALANCE", "TOTAL_ASSETS"),
    "TOTLIAB": ("GBALANCE", "TOTAL_LIABILITIES"),
    "TOTCURRASSET": ("GBALANCE", "TOTAL_CURRENT_ASSETS"),
    "TOTALCURRLIAB": ("GBALANCE", "TOTAL_CURRENT_LIAB"),
    "INTAASSET": ("GBALANCE", "INTANGIBLE_ASSET"),
    "ACCORECE": ("GBALANCE", "ACCOUNTS_RECE"),
    "FIXEDASSENET": ("GBALANCE", "FIXED_ASSET"),
    "BIZINCO": ("GINCOME", "OPERATE_INCOME"),
    "BIZTOTINCO": ("GINCOME", "TOTAL_OPERATE_INCOME"),
    "BIZCOST": ("GINCOME", "OPERATE_COST"),
    "DEVEEXPE": ("GINCOME", "RESEARCH_EXPENSE"),
    "FINEXPE": ("GINCOME", "FINANCE_EXPENSE"),
    "NETPROFIT": ("GINCOME", "NETPROFIT"),
    "PARENETP": ("GINCOME", "PARENT_NETPROFIT"),
    "MANANETR": ("GCASHFLOW", "NETCASH_OPERATE"),
    "ACQUASSETCASH": ("GCASHFLOW", "CONSTRUCT_LONG_ASSET"),
}
REPORTS = ["GBALANCE", "GINCOME", "GCASHFLOW"]


def fetch(secucode, report):
    p = {"reportName": f"RPT_F10_FINANCE_{report}", "columns": "ALL",
         "filter": f'(SECUCODE="{secucode}")', "pageNumber": 1, "pageSize": 60,
         "sortColumns": "REPORT_DATE", "sortTypes": "-1"}
    for attempt in range(3):
        try:
            d = requests.get(URL, params=p, headers=HDR, timeout=40).json()["result"]["data"]
            out = {}
            for row in d:
                rd = str(row.get("REPORT_DATE", ""))[:10]
                if rd.endswith("12-31"):
                    out[rd[:4]] = row
            return report, out
        except Exception:
            if attempt == 2:
                return report, {}
            time.sleep(2)


def main():
    panel = pd.read_csv(os.path.join(BASE, "firm_panel_raw.csv"), dtype={"stkcd": str})
    firms = panel[["stkcd"]].drop_duplicates().stkcd.tolist()
    tasks = []
    with ThreadPoolExecutor(max_workers=10) as ex:
        futs = {}
        for c in firms:
            sc = f"{c}.{'SH' if c.startswith('6') else 'SZ'}"
            for r in REPORTS:
                futs[ex.submit(fetch, sc, r)] = c
        store = {c: {} for c in firms}
        done = 0
        for fu in as_completed(futs):
            c = futs[fu]
            report, out = fu.result()
            store[c][report] = out
            done += 1
            if done % 60 == 0:
                print(f"{done}/{len(futs)}", flush=True)

    rows = []
    for r in panel[["stkcd", "year"]].itertuples(index=False):
        y = str(r.year)
        for mine, (rep, fld) in MAP.items():
            em = None
            blk = store.get(r.stkcd, {}).get(rep, {}).get(y)
            if blk:
                em = blk.get(fld)
                if em in ("", None):
                    em = None
            rows.append({"stkcd": r.stkcd, "year": r.year, "item": mine, "em_value": em})
    emdf = pd.DataFrame(rows)
    emdf = emdf.pivot_table(index=["stkcd", "year"], columns="item",
                            values="em_value", aggfunc="last").reset_index()
    emdf.to_csv(os.path.join(BASE, "em_crosscheck_raw.csv"), index=False, encoding="utf-8-sig")

    cmp = panel[["stkcd", "year"] + list(MAP)].merge(emdf, on=["stkcd", "year"], how="left",
                                                     suffixes=("", "_em"))
    rec = []
    for item in MAP:
        a = pd.to_numeric(cmp[item], errors="coerce")
        b = pd.to_numeric(cmp[item + "_em"], errors="coerce")
        both = a.notna() & b.notna() & (b != 0)
        rel = (a - b).abs() / b.abs()
        rec.append({
            "科目": item,
            "可比观测": int(both.sum()),
            "未获取": int((b.isna()).sum()),
            "相对误差<0.1%": int((rel[both] < 0.001).sum()),
            "0.1%—1%": int(((rel[both] >= 0.001) & (rel[both] < 0.01)).sum()),
            "1%—5%": int(((rel[both] >= 0.01) & (rel[both] < 0.05)).sum()),
            ">5%": int((rel[both] >= 0.05).sum()),
            "中位相对误差": round(float(rel[both].median()), 6) if both.any() else None,
        })
    rep = pd.DataFrame(rec)
    rep["一致率(<1%)"] = ((rep["相对误差<0.1%"] + rep["0.1%—1%"]) / rep["可比观测"]).round(4)
    rep.to_csv(os.path.join(BASE, "em_crosscheck_summary.csv"), index=False, encoding="utf-8-sig")
    cmp.to_csv(os.path.join(BASE, "em_crosscheck_detail.csv"), index=False, encoding="utf-8-sig")
    print(rep.to_string(index=False))


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""抓取东方财富公司概况：注册省份、实际控制人、上市日期、所属行业"""
import os
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor

BASE = os.path.dirname(os.path.abspath(__file__))
URL = "https://datacenter.eastmoney.com/securities/api/data/v1/get"
HDR = {"User-Agent": "Mozilla/5.0"}


def fetch(row):
    sec = f"{row.code}.{'SH' if row.symbol.startswith('sh') else 'SZ'}"
    p = {"reportName": "RPT_F10_BASIC_ORGINFO", "columns": "ALL",
         "filter": f'(SECUCODE="{sec}")', "pageNumber": 1, "pageSize": 5}
    try:
        d = requests.get(URL, params=p, headers=HDR, timeout=30).json()["result"]["data"][0]
        return {"code": row.code, "name": row.name, "sector": row.sector,
                "org_name": d.get("ORG_NAME"), "province_raw": d.get("PROVINCE"),
                "reg_address": d.get("REG_ADDRESS"), "actual_holder": d.get("ACTUAL_HOLDER"),
                "listing_date": (d.get("LISTING_DATE") or "")[:10],
                "em_industry": d.get("EM2016"), "csrc_industry": d.get("INDUSTRYCSRC1")}
    except Exception as e:
        return {"code": row.code, "name": row.name, "sector": row.sector, "error": str(e)[:100]}


if __name__ == "__main__":
    firms = pd.read_csv(os.path.join(BASE, "firm_list.csv"), dtype={"code": str})
    with ThreadPoolExecutor(max_workers=8) as ex:
        rows = list(ex.map(fetch, firms.itertuples()))
    df = pd.DataFrame(rows)
    df.to_csv(os.path.join(BASE, "firm_info.csv"), index=False, encoding="utf-8-sig")
    print(df.shape, df.error.notna().sum() if "error" in df else 0)
    print(df[["code", "name", "province_raw", "actual_holder", "listing_date"]].to_string())

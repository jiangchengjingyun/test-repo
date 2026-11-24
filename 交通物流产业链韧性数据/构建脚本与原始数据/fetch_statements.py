# -*- coding: utf-8 -*-
"""抓取新浪财经合并报表年度数据（与论文数据来源一致）"""
import json, os, time, random
import pandas as pd
import requests
from concurrent.futures import ThreadPoolExecutor, as_completed

BASE = os.path.dirname(os.path.abspath(__file__))
LIST = os.path.join(BASE, "firm_list.csv")
OUT = os.path.join(BASE, "raw_statements.csv")
URL = "https://quotes.sina.cn/cn/api/openapi.php/CompanyFinanceService.getFinanceReport2022"
HDR = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)", "Referer": "https://finance.sina.com.cn"}

FIELDS = {
    "fzb": ["TOTASSET", "TOTCURRASSET", "TOTALCURRLIAB", "INTAASSET", "ACCORECE",
            "NOTESACCORECE", "FIXEDASSENET", "TOTLIAB", "TOTLIABSHAREQUI"],
    "lrb": ["BIZINCO", "BIZTOTINCO", "BIZCOST", "DEVEEXPE", "FINEXPE", "NETPROFIT", "PARENETP", "TOTPROFIT"],
    "llb": ["MANANETR", "ACQUASSETCASH"],
}
YEARS = [str(y) for y in range(2012, 2024)]


def fetch_one(code, source, session):
    params = {"paperCode": code, "source": source, "type": "0", "page": "1", "num": "60"}
    for attempt in range(4):
        try:
            r = session.get(URL, params=params, timeout=45, headers=HDR)
            js = r.json()["result"]["data"]["report_list"]
            rows = []
            for date, block in js.items():
                if not date.endswith("1231"):
                    continue
                year = date[:4]
                if year not in YEARS:
                    continue
                vals = {}
                for it in block.get("data", []):
                    f = it.get("item_field")
                    v = it.get("item_value")
                    if f in FIELDS[source]:
                        try:
                            vals[f] = float(v) if v not in (None, "", "--") else None
                        except (TypeError, ValueError):
                            vals[f] = None
                for f in FIELDS[source]:
                    rows.append({"code": code, "year": year, "source": source,
                                 "field": f, "value": vals.get(f)})
            return rows
        except Exception as e:
            if attempt == 3:
                return [{"code": code, "year": "", "source": source, "field": "ERROR", "value": str(e)[:120]}]
            time.sleep(2 + attempt * 2)


def main():
    firms = pd.read_csv(LIST, dtype={"code": str})
    tasks = [(r.symbol, r.code, s) for r in firms.itertuples() for s in FIELDS]
    out, errors = [], []
    with requests.Session() as session:
        with ThreadPoolExecutor(max_workers=8) as ex:
            futs = {ex.submit(fetch_one, sym, s, session): (sym, code, s) for sym, code, s in tasks}
            done = 0
            for fu in as_completed(futs):
                sym, code, s = futs[fu]
                rows = fu.result()
                done += 1
                if any(r["field"] == "ERROR" for r in rows):
                    errors.append((code, s))
                out.extend(rows)
                if done % 30 == 0:
                    print(f"{done}/{len(tasks)} 完成", flush=True)
    df = pd.DataFrame(out)
    df.to_csv(OUT, index=False, encoding="utf-8-sig")
    print("总行数", len(df), "失败任务", errors)
    print(df.groupby("source")["value"].apply(lambda x: x.notna().sum()))


if __name__ == "__main__":
    main()

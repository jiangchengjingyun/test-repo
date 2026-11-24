# -*- coding: utf-8 -*-
"""导出 Stata 可读的 Excel 数据集、dta 文件与 do 文件"""
import os
import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
OUTDIR = os.path.abspath(os.path.join(BASE, "..", "output"))
os.makedirs(OUTDIR, exist_ok=True)

PROV_XLSX = r"D:\软件库\微信聊天\xwechat_files\wxid_9xgpzj79zcod22_5ef8\temp\RWTemp\2026-11\9e20f478899dc29eb19741386f9343c8\省级面板数据2014-2024(1).xlsx"
XLSX = os.path.join(OUTDIR, "交通物流产业链韧性_企业面板数据_2015-2023.xlsx")

RAW_COLS = ["TOTASSET", "TOTLIAB", "TOTCURRASSET", "TOTALCURRLIAB", "INTAASSET", "ACCORECE",
            "FIXEDASSENET", "BIZINCO", "BIZCOST", "DEVEEXPE", "FINEXPE", "NETPROFIT",
            "MANANETR", "ACQUASSETCASH"]
RATIO_COLS = ["intan_ratio", "rd_intensity", "dcapex_intensity", "roa_std3", "cash_ratio",
              "cost_ratio", "ind_adj_growth", "asset_turn", "current_ratio"]
SCORE_COLS = ["DE1", "DE2", "DE3", "DEL", "RESIST", "RECOVER", "RES", "RES_geo", "RES_alt"]
CTRL_COLS = ["Size", "Lev", "Growth", "Roa"]
MED_COLS = ["Info", "Eff", "Fin"]
PROV_COLS = ["freight", "pgdp", "industry2", "roaddensity", "open", "urban", "IV"]

PAPER_T3 = {
    "RES": (0.390, 0.071, 0.075, 0.384, 0.644),
    "RESIST": (0.604, 0.110, 0.141, 0.607, 0.920),
    "RECOVER": (0.176, 0.106, 0.010, 0.151, 0.642),
    "DEL": (0.098, 0.115, 0.000, 0.046, 0.523),
    "DE1": (0.125, 0.209, 0.000, 0.046, 1.000),
    "DE2": (0.095, 0.217, 0.000, 0.000, 1.000),
    "DE3": (0.076, 0.139, 0.000, 0.034, 1.000),
    "Size": (23.273, 1.441, 20.249, 23.163, 26.651),
    "Lev": (0.468, 0.209, 0.083, 0.454, 1.204),
    "Growth": (0.125, 0.437, -0.639, 0.059, 2.694),
    "Roa": (0.026, 0.071, -0.447, 0.033, 0.163),
}

DICT = [
    ("stkcd", "证券代码", "6位代码，保留前导零", "新浪财经行业成分股", "标识"),
    ("firm", "证券简称", "抓取时的简称，含ST/退市标记", "新浪财经", "标识"),
    ("year", "年份", "2015—2023", "—", "标识"),
    ("province", "注册省份", "公司注册地所在省级行政区", "东方财富公司概况", "分组/合并省级面板"),
    ("sector", "板块", "交通运输/公路桥梁（新浪行业分类）", "新浪财经", "分组"),
    ("soe", "产权性质", "1=国有，0=非国有，缺失=未披露实际控制人", "东方财富实际控制人字段判定", "表8分组"),
    ("east", "东部地区", "1=东部（京沪津冀辽苏浙闽鲁粤琼）", "按注册省份划分", "表8分组"),
    ("infra", "基础设施型", "1=公路桥梁，0=交通运输", "新浪行业分类", "表8分组"),
    ("TOTASSET", "资产总计", "合并资产负债表，元", "新浪财经合并报表", "分母"),
    ("TOTLIAB", "负债合计", "元", "同上", "Lev/Fin"),
    ("TOTCURRASSET", "流动资产合计", "元", "同上", "冗余资源"),
    ("TOTALCURRLIAB", "流动负债合计", "元", "同上", "冗余资源"),
    ("INTAASSET", "无形资产净额", "元", "同上", "表1 无形资产占比"),
    ("ACCORECE", "应收账款", "元", "同上", "信息路径"),
    ("FIXEDASSENET", "固定资产净额", "元", "同上", "效率路径"),
    ("BIZINCO", "营业收入", "元", "合并利润表", "分母/周转"),
    ("BIZCOST", "营业成本", "元", "同上", "成本控制"),
    ("DEVEEXPE", "研发费用", "元；2018年前多数公司未单列，缺失按0处理", "同上", "表1 研发投入强度"),
    ("FINEXPE", "财务费用", "元", "同上", "金融路径"),
    ("NETPROFIT", "净利润", "元", "同上", "ROA"),
    ("MANANETR", "经营活动现金流量净额", "元", "合并现金流量表", "现金流保障"),
    ("ACQUASSETCASH", "购建固定资产、无形资产和其他长期资产支付的现金", "元", "同上", "表1 数字资本开支强度"),
    ("intan_ratio", "无形资产占比", "INTAASSET/TOTASSET", "计算", "表1 DE1"),
    ("rd_intensity", "研发投入强度", "DEVEEXPE/BIZINCO", "计算", "表1 DE2"),
    ("dcapex_intensity", "数字资本开支强度", "ACQUASSETCASH/BIZINCO", "计算", "表1 DE3"),
    ("roa_std3", "盈利稳定性", "近三年ROA标准差（反向指标）", "计算", "表2 抵抗能力"),
    ("cash_ratio", "现金流保障", "MANANETR/TOTASSET", "计算", "表2 抵抗能力"),
    ("cost_ratio", "成本控制", "BIZCOST/BIZINCO（反向指标）", "计算", "表2 抵抗能力"),
    ("ind_adj_growth", "行业调整增长", "营业收入增长率-同板块同年度中位数", "计算", "表2 恢复能力"),
    ("asset_turn", "资产周转", "BIZINCO/TOTASSET", "计算", "表2 恢复能力"),
    ("current_ratio", "冗余资源", "TOTCURRASSET/TOTALCURRLIAB", "计算", "表2 恢复能力"),
    ("DE1", "数据基础投入得分", "无形资产占比极差标准化×熵权", "计算", "表1/表7"),
    ("DE2", "数据资源化得分", "研发投入强度极差标准化×熵权", "计算", "表1/表7"),
    ("DE3", "数据场景应用得分", "数字资本开支强度极差标准化×熵权", "计算", "表1/表7"),
    ("DEL", "数据要素水平", "DE1—DE3按熵权0.288/0.470/0.241合成", "计算", "核心解释变量"),
    ("RESIST", "抵抗能力", "盈利稳定性、现金流保障、成本控制熵权合成", "计算", "表2/表4列4"),
    ("RECOVER", "恢复能力", "行业调整增长、资产周转、冗余资源熵权合成", "计算", "表2/表4列5"),
    ("RES", "产业链韧性", "(RESIST+RECOVER)/2", "计算", "被解释变量"),
    ("RES_geo", "韧性（几何平均）", "sqrt(RESIST×RECOVER)", "计算", "表5列1"),
    ("RES_alt", "韧性（六指标统一熵权）", "六指标统一熵值法综合得分", "计算", "表5列2"),
    ("Size", "企业规模", "ln(TOTASSET)", "计算", "控制变量"),
    ("Lev", "资产负债率", "TOTLIAB/TOTASSET", "计算", "控制变量"),
    ("Growth", "营业收入增长率", "营业收入同比增速", "计算", "控制变量"),
    ("Roa", "总资产收益率", "NETPROFIT/TOTASSET", "计算", "控制变量"),
    ("Info", "信息效率", "ln(1+BIZINCO/ACCORECE)", "计算", "表6信息路径"),
    ("Eff", "运营效率", "ln(1+BIZINCO/FIXEDASSENET)", "计算", "表6效率路径"),
    ("Fin", "融资成本", "FINEXPE/TOTLIAB×100", "计算", "表6金融路径"),
    ("freight", "货运量", "亿吨公里，省级", "省级面板（用户提供）", "工具变量"),
    ("pgdp", "人均地区生产总值", "元，省级", "省级面板（用户提供）", "工具变量"),
    ("industry2", "第二产业占比", "%，省级", "省级面板（用户提供）", "控制/描述"),
    ("roaddensity", "公路密度", "公里/平方公里，省级", "省级面板（用户提供）", "工具变量"),
    ("open", "对外开放度", "进出口总额/GDP，%，省级", "省级面板（用户提供）", "控制/描述"),
    ("urban", "城镇化率", "% ，省级", "省级面板（用户提供）", "工具变量"),
    ("IV", "省级发展环境工具变量", "公路密度、货运量、人均GDP、城镇化率Z值等权平均", "计算", "表5列6"),
]


def build_analysis(df):
    cols = (["stkcd", "firm", "year", "province", "sector", "soe", "east", "infra"]
            + RATIO_COLS + SCORE_COLS + CTRL_COLS + MED_COLS + PROV_COLS)
    return df[cols].copy()


def style_sheet(ws, freeze="A2", ncol=None):
    header_fill = PatternFill("solid", fgColor="1F3864")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF", size=10)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = freeze
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 32
    for j in range(1, (ncol or ws.max_column) + 1):
        letter = get_column_letter(j)
        width = max(len(str(ws.cell(1, j).value or "")) * 1.8, 10)
        ws.column_dimensions[letter].width = min(width, 34)


def main():
    df = pd.read_csv(os.path.join(BASE, "firm_panel_raw.csv"), dtype={"stkcd": str})
    df["stkcd"] = df.stkcd.str.zfill(6)
    analysis = build_analysis(df)

    raw = pd.read_csv(os.path.join(BASE, "raw_statements.csv"), dtype={"code": str})
    raw["code"] = raw.code.str[-6:]
    raw_w = raw[raw.field != "ERROR"].pivot_table(index=["code", "year"], columns="field",
                                                  values="value", aggfunc="last").reset_index()
    raw_w = raw_w.rename(columns={"code": "stkcd", "year": "年度"})
    raw_w["stkcd"] = raw_w.stkcd.str.zfill(6)
    for c in raw_w.columns:
        if c not in ("stkcd", "年度"):
            raw_w[c] = (raw_w[c] / 1e8).round(4)

    prov = pd.read_excel(PROV_XLSX, sheet_name="面板数据")

    info = pd.read_csv(os.path.join(BASE, "firm_info.csv"), dtype={"code": str})
    info["in_sample"] = info.code.isin(df.stkcd.astype(str)).astype(int)
    info = info[["code", "name", "sector", "province_raw", "actual_holder", "listing_date",
                 "em_industry", "in_sample"]].rename(columns={
        "code": "证券代码", "name": "证券简称", "sector": "板块", "province_raw": "注册省份",
        "actual_holder": "实际控制人", "listing_date": "上市日期",
        "em_industry": "东财行业", "in_sample": "进入最终样本"})

    # 描述统计对照
    rows = []
    desc = df[list(PAPER_T3)].describe().T
    for v, p in PAPER_T3.items():
        d = desc.loc[v]
        rows.append({"变量": v, "观测值": int(d["count"]), "均值_复现": round(d["mean"], 3),
                     "均值_论文": p[0], "标准差_复现": round(d["std"], 3), "标准差_论文": p[1],
                     "最小值_复现": round(d["min"], 3), "最小值_论文": p[2],
                     "中位数_复现": round(d["50%"], 3), "中位数_论文": p[3],
                     "最大值_复现": round(d["max"], 3), "最大值_论文": p[4]})
    desc_df = pd.DataFrame(rows)

    note = pd.DataFrame({"项目": [
        "数据集名称", "样本区间", "样本单位", "样本量", "公司数", "板块构成", "覆盖省份",
        "企业财务数据来源", "公司属性来源", "省级面板来源", "抓取时点", "抓取方式",
        "样本筛选", "缩尾处理", "缺失值处理", "熵值法口径", "复现情况",
        "与论文的差异", "配套文件", "使用方式"],
        "说明": [
        "交通物流产业链韧性企业面板数据（论文《数据要素驱动交通物流产业链韧性提升》配套数据）",
        "2015—2023年（原始报表回溯至2012年，用于计算近三年ROA标准差与增长率）",
        "上市公司—年度",
        "891个观测",
        "99家（新浪行业分类原始名单107家：交通运输87家、公路桥梁20家）",
        "交通运输（物流服务型）726个观测，公路桥梁（基础设施型）165个观测",
        "29个省级行政区",
        "新浪财经公开合并资产负债表、利润表、现金流量表（年报口径），按证券代码与会计年度逐条抓取",
        "东方财富公司概况：注册省份、实际控制人、上市日期",
        "用户提供的《省级面板数据2014-2024》，31省×11年=341个观测，无缺失",
        "2026年11月",
        "公开接口逐公司抓取，原始科目单位统一换算为亿元（保留4位小数），可复核",
        "剔除2015—2023年年度不全的8家公司（退市长油、ST东北高、国恒退、华北高速、外运发展、营口港、退市中昌、退市海医），保留99家平衡面板",
        "连续变量在1%和99%分位处缩尾（论文表3口径），熵值法在缩尾后的891个观测上计算",
        "研发费用2018年前多数公司未单列，按0处理并保留rd_missing标记；实际控制人为“无”的4家公司产权性质留缺失",
        "极差标准化（负向指标反向处理）+熵值法；抵抗、恢复两维度内部分别赋权，再等权合成RES",
        "熵权与论文表1、表2高度一致：DEL三维度0.288/0.470/0.241（论文0.268/0.477/0.255），抵抗0.091/0.372/0.538（论文0.126/0.480/0.394），恢复0.144/0.533/0.323（论文0.134/0.534/0.332）",
        "产权性质改用实际控制人口径，覆盖855个观测（论文原仅405个）；研发费用缺失处理、ROA标准差窗口等细节差异使个别统计量与论文表3存在小幅差异，详见“05_描述统计”",
        "同目录下：企业面板dta（Stata直接打开）、复现do文件",
        "Stata：use 交通物流_企业面板.dta, clear 后直接运行do文件；Excel：首行为变量名，可直接import excel"]})

    with pd.ExcelWriter(XLSX, engine="openpyxl") as xw:
        note.to_excel(xw, sheet_name="00_使用说明", index=False)
        analysis.to_excel(xw, sheet_name="01_分析面板", index=False)
        raw_w.to_excel(xw, sheet_name="02_原始财务数据", index=False)
        prov.to_excel(xw, sheet_name="03_省级面板", index=False)
        pd.DataFrame(DICT, columns=["变量名", "变量名称", "构造口径", "数据来源", "对应论文位置"]).to_excel(
            xw, sheet_name="04_指标定义", index=False)
        desc_df.to_excel(xw, sheet_name="05_描述统计", index=False)
        info.to_excel(xw, sheet_name="06_公司名单", index=False)

    wb = load_workbook(XLSX)
    for name in wb.sheetnames:
        ws = wb[name]
        style_sheet(ws)
        if name in ("01_分析面板", "02_原始财务数据"):
            for row in ws.iter_rows(min_row=2, min_col=9 if name == "01_分析面板" else 3):
                for cell in row:
                    cell.number_format = "0.0000"
    ws = wb["00_使用说明"]
    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 110
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    wb.save(XLSX)

    # Stata dta
    dta = analysis.copy()
    for c in dta.columns:
        if dta[c].dtype == object and c not in ("stkcd", "firm", "province", "sector"):
            dta[c] = pd.to_numeric(dta[c], errors="coerce")
    dta.to_stata(os.path.join(OUTDIR, "交通物流_企业面板.dta"), write_index=False, version=118)

    print("已导出", XLSX)
    print(analysis.shape, "dta ok")
    print(desc_df.to_string(index=False))


if __name__ == "__main__":
    main()

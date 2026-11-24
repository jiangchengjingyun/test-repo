# -*- coding: utf-8 -*-
"""追加“核验样本回归”工作表（结果取自 Stata 运行日志）"""
import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUT = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "output"))
XLSX = os.path.join(OUT, "交通物流产业链韧性_企业面板数据_2015-2023.xlsx")

ROWS = [
    ("基准（论文口径）", "全部观测", 891, "-0.0536 (0.0391)", "-1.37", "0.174",
     "与论文表4列(2)的-0.0500基本一致"),
    ("剔除实体错配", "data_valid=1", 873, "-0.0667 (0.0392)", "-1.70", "0.092",
     "剔除601313、600087后负向略强，10%水平边缘显著"),
    ("仅运输主业", "data_valid=1 & transport_core=1", 621, "-0.0067 (0.0720)", "-0.09", "0.926",
     "系数接近零，说明基准负向效应主要由板块内非运输企业贡献"),
    ("剔除跨源差异", "data_valid=1 & xcheck_flag=0", 850, "-0.0669 (0.0396)", "-1.69", "0.094",
     "剔除24个跨源差异观测后结果稳定"),
]

NOTE = ("说明：均控制企业规模、资产负债率、营业收入增长率、总资产收益率，并含企业固定效应与年份固定效应，"
        "标准误聚类到企业层面。第3行是重要提示——按证监会行业口径仅保留交通运输、仓储和邮政业企业后，"
        "数据要素水平的系数由-0.0536变为-0.0067且完全不显著，表明论文的负向结果在一定程度上"
        "来自新浪“交通运输”板块中混入的交运设备制造、房地产、公用事业等非运输企业；"
        "若论文强调“交通物流企业”的结论，建议以第3行作为主回归或补充说明。")

df = pd.DataFrame(ROWS, columns=["样本设定", "筛选条件", "观测数", "DEL系数(标准误)", "t值", "p值", "解读"])

wb = load_workbook(XLSX)
if "10_核验样本回归" in wb.sheetnames:
    del wb["10_核验样本回归"]
ws = wb.create_sheet("10_核验样本回归")
ws.append(list(df.columns))
for r in df.itertuples(index=False):
    ws.append(list(r))
ws.append([])
ws.append(["注", NOTE])

for c in ws[1]:
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws.freeze_panes = "A2"
for i, w in enumerate([16, 34, 9, 20, 8, 8, 60], start=1):
    ws.column_dimensions[get_column_letter(i)].width = w
last = ws.max_row
ws.cell(last, 2).alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=last, start_column=2, end_row=last, end_column=7)
ws.row_dimensions[last].height = 60

wb.save(XLSX)
print("sheet列表：", wb.sheetnames)

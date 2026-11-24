*==============================================================================
* 数据要素驱动交通物流产业链韧性提升：测度特征及机理研究
* 复现 do 文件（Stata 18）
* 数据：交通物流_企业面板.dta（891 个企业—年度观测，99 家公司，2015—2023）
* 说明：所有变量已在 Python 端完成缩尾（1%/99%）与熵值法赋权，
*       本文件只做估计，便于与论文表 3—表 8 对照。
*==============================================================================

clear all
set more off
set linesize 200
capture log close

* 工作目录：默认 do 文件所在目录；如从其他位置运行，请改成绝对路径
global root "."
capture confirm file "$root/交通物流_企业面板.dta"
if _rc {
    global root "`c(pwd)'"
}
use "$root/交通物流_企业面板.dta", clear

label var RES       "产业链韧性"
label var RESIST    "抵抗能力"
label var RECOVER   "恢复能力"
label var DEL       "数据要素水平"
label var DE1       "数据基础投入"
label var DE2       "数据资源化"
label var DE3       "数据场景应用"
label var Size      "企业规模"
label var Lev       "资产负债率"
label var Growth    "营业收入增长率"
label var Roa       "总资产收益率"
label var Info      "信息效率"
label var Eff       "运营效率"
label var Fin       "融资成本"
label var IV        "省级发展环境工具变量"
label var RES_geo   "韧性（几何平均）"
label var RES_alt   "韧性（六指标统一熵权）"

encode sector, gen(secnum)
encode stkcd, gen(id)
label var id "企业编号"
xtset id year

*------------------------------------------------------------------------------
* 表 3 描述性统计
*------------------------------------------------------------------------------
summarize RES RESIST RECOVER DEL DE1 DE2 DE3 Size Lev Growth Roa, detail

* 相关系数（论文正文：RES 与 DEL 相关系数 -0.0938，与 RESIST/RECOVER 0.6668/0.6380）
pwcorr RES DEL RESIST RECOVER, sig

* 图 2 数据：年度均值
preserve
collapse (mean) DEL RES, by(year)
list, sep(0)
export delimited using "$root/图2_年度均值.csv", replace
restore

*------------------------------------------------------------------------------
* 表 4 基准回归
*------------------------------------------------------------------------------
* 列(1) 仅双向固定效应
xtreg RES DEL i.year, fe vce(cluster id)
estimates store m1
scalar r2w1 = e(r2_w)

* 列(2) 加控制变量
xtreg RES DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store m2
scalar r2w2 = e(r2_w)

* 列(3) 行业×年份固定效应
xtreg RES DEL Size Lev Growth Roa i.year i.secnum#i.year, fe vce(cluster id)
estimates store m3
scalar r2w3 = e(r2_w)

* 列(4)(5) 分维度
xtreg RESIST DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store m4
xtreg RECOVER DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store m5

estimates table m1 m2 m3 m4 m5, b(%9.4f) se(%9.4f) stats(N)
display "组内R2：m1=" %6.4f r2w1 "  m2=" %6.4f r2w2 "  m3=" %6.4f r2w3

* 多重共线性
quietly reg RES DEL Size Lev Growth Roa i.year
estat vif

*------------------------------------------------------------------------------
* 表 5 稳健性检验
*------------------------------------------------------------------------------
* (1) 几何平均
xtreg RES_geo DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store r1

* (2) 六指标统一熵权
xtreg RES_alt DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store r2

* (3) 解释变量滞后一期
xtreg RES L.DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store r3

* (4) 剔除 2020 年
preserve
drop if year == 2020
xtreg RES DEL Size Lev Growth Roa i.year, fe vce(cluster id)
estimates store r4
restore

* (5) 5%/95% 缩尾后重估
foreach v in RES DEL Size Lev Growth Roa {
    _pctile `v', p(5 95)
    gen `v'_w = min(max(`v', r(r1)), r(r2))
}
xtreg RES_w DEL_w Size_w Lev_w Growth_w Roa_w i.year, fe vce(cluster id)
estimates store r5
drop *_w

* (6) 工具变量：省级发展环境
ivregress 2sls RES Size Lev Growth Roa i.year i.id (DEL = IV), vce(cluster id)
estimates store r6
* 第一阶段F值（工具变量 IV 的联合显著性）
quietly xtreg DEL IV Size Lev Growth Roa i.year, fe
test IV
display "第一阶段F值 = " %6.2f r(F)

estimates table r1 r2 r3 r4 r5 r6, b(%9.4f) se(%9.4f) stats(N)

*------------------------------------------------------------------------------
* 表 6 作用机理
*------------------------------------------------------------------------------
* 信息路径
xtreg Info DEL Size Lev Growth Roa i.year, fe vce(cluster id)
scalar a_info = _b[DEL]
scalar sa_info = _se[DEL]
xtreg RES DEL Info Size Lev Growth Roa i.year, fe vce(cluster id)
scalar b_info = _b[Info]
scalar sb_info = _se[Info]
scalar c_info = _b[DEL]
* Sobel z
scalar z_info = (a_info*b_info)/sqrt((b_info^2)*(sa_info^2)+(a_info^2)*(sb_info^2))
display "信息路径：a=" %7.4f a_info "  b=" %7.4f b_info "  直接效应=" %7.4f c_info "  Sobel z=" %7.3f z_info

* 效率路径
xtreg Eff DEL Size Lev Growth Roa i.year, fe vce(cluster id)
scalar a_eff = _b[DEL]
scalar sa_eff = _se[DEL]
xtreg RES DEL Eff Size Lev Growth Roa i.year, fe vce(cluster id)
scalar b_eff = _b[Eff]
scalar sb_eff = _se[Eff]
scalar c_eff = _b[DEL]
scalar z_eff = (a_eff*b_eff)/sqrt((b_eff^2)*(sa_eff^2)+(a_eff^2)*(sb_eff^2))
display "效率路径：a=" %7.4f a_eff "  b=" %7.4f b_eff "  直接效应=" %7.4f c_eff "  Sobel z=" %7.3f z_eff

* 金融路径
xtreg Fin DEL Size Lev Growth Roa i.year, fe vce(cluster id)
scalar a_fin = _b[DEL]
scalar sa_fin = _se[DEL]
xtreg RES DEL Fin Size Lev Growth Roa i.year, fe vce(cluster id)
scalar b_fin = _b[Fin]
scalar sb_fin = _se[Fin]
scalar c_fin = _b[DEL]
scalar z_fin = (a_fin*b_fin)/sqrt((b_fin^2)*(sa_fin^2)+(a_fin^2)*(sb_fin^2))
display "金融路径：a=" %7.4f a_fin "  b=" %7.4f b_fin "  直接效应=" %7.4f c_fin "  Sobel z=" %7.3f z_fin

*------------------------------------------------------------------------------
* 表 7 数据要素分维度
*------------------------------------------------------------------------------
foreach v in DE1 DE2 DE3 {
    xtreg RES `v' Size Lev Growth Roa i.year, fe vce(cluster id)
    estimates store d_`v'
}
estimates table d_DE1 d_DE2 d_DE3, b(%9.4f) se(%9.4f) stats(N)

*------------------------------------------------------------------------------
* 表 8 异质性
*------------------------------------------------------------------------------
xtreg RES DEL Size Lev Growth Roa i.year if east == 1, fe vce(cluster id)
estimates store h1
xtreg RES DEL Size Lev Growth Roa i.year if east == 0, fe vce(cluster id)
estimates store h2
xtreg RES DEL Size Lev Growth Roa i.year if soe == 1, fe vce(cluster id)
estimates store h3
xtreg RES DEL Size Lev Growth Roa i.year if soe == 0, fe vce(cluster id)
estimates store h4
xtreg RES DEL Size Lev Growth Roa i.year if infra == 1, fe vce(cluster id)
estimates store h5
xtreg RES DEL Size Lev Growth Roa i.year if infra == 0, fe vce(cluster id)
estimates store h6
estimates table h1 h2 h3 h4 h5 h6, b(%9.4f) se(%9.4f) stats(N)

* 组间差异检验（交互项）
xtreg RES c.DEL##i.east Size Lev Growth Roa i.year, fe vce(cluster id)
test 1.east#c.DEL
xtreg RES c.DEL##i.soe Size Lev Growth Roa i.year, fe vce(cluster id)
test 1.soe#c.DEL
xtreg RES c.DEL##i.infra Size Lev Growth Roa i.year, fe vce(cluster id)
test 1.infra#c.DEL

*------------------------------------------------------------------------------
* 时间演化：分时段子样本
*------------------------------------------------------------------------------
xtreg RES DEL Size Lev Growth Roa i.year if year <= 2019, fe vce(cluster id)
estimates store t1
xtreg RES DEL Size Lev Growth Roa i.year if year >= 2020, fe vce(cluster id)
estimates store t2
estimates table t1 t2, b(%9.4f) se(%9.4f) stats(N)

display "全部完成"

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
P-1507A/B 吸入ストレーナ 40->250メッシュ変更 技術評価
DEG物性は Aspen物性ツール (modAspenPropData.bas / Shell EOEG datadeck) の式で独立計算。
"""
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

MW = 106.122  # g/mol = kg/kmol
OUT = "/home/user/Base/outputs"

def K(tc):  # degC -> K
    return tc + 273.15

# --- 物性式 ---
def Pvp_Pa(T):
    # PLXANT: ln P[Pa] = 74.55 - 10632/T - 6.8195*ln(T) + 9.0968e-18*T^6
    return math.exp(74.55 - 10632.0/T - 6.8195*math.log(T) + 9.0968e-18 * T**6)

def mu_Pas(T):
    # MULDIP: ln mu[Pa.s] = -125.30 + 8891/T + 16.143*ln(T)
    return math.exp(-125.30 + 8891.0/T + 16.143*math.log(T))

def rho_kgm3(T):
    # DNLDIP DIPPR-105: rho[kmol/m3] = 0.83692 / 0.26112^(1+(1-T/744.6)^0.2422)
    exponent = 1.0 + (1.0 - T/744.6)**0.2422
    rho_kmol = 0.83692 / (0.26112**exponent)
    return rho_kmol * MW

# ============================================================
# 1. 物性表
# ============================================================
temps_C = [0, 10, 20, 25, 30, 40, 50, 70, 100, 130, 158.8]
rows = []
for tc in temps_C:
    T = K(tc)
    mu_cP = mu_Pas(T) * 1000.0   # Pa.s -> cP (1 Pa.s = 1000 cP)
    rho = rho_kgm3(T)
    pvp_kPa = Pvp_Pa(T) / 1000.0
    in_muldip = (270.0 <= T <= 450.0)
    rows.append((tc, T, mu_cP, rho, pvp_kPa, in_muldip))

# 検証点 158.8C
T_ref = K(158.8)
mu_ref = mu_Pas(T_ref)*1000.0
rho_ref = rho_kgm3(T_ref)
pvp_ref = Pvp_Pa(T_ref)/1000.0

print("=== 物性表 ===")
print(f"{'T[C]':>7} {'T[K]':>8} {'mu[cP]':>9} {'rho[kg/m3]':>11} {'Pvp[kPa]':>10} {'MULDIP範囲':>10}")
for tc, T, mu_cP, rho, pvp, inr in rows:
    print(f"{tc:>7.1f} {T:>8.2f} {mu_cP:>9.3f} {rho:>11.1f} {pvp:>10.3f} {'OK' if inr else 'OUT':>10}")

print("\n=== 検証 @158.8C (432.0K) ===")
print(f"mu  = {mu_ref:.3f} cP  (期待 ~1.16)")
print(f"rho = {rho_ref:.1f} kg/m3 (期待 ~1009)")
print(f"Pvp = {pvp_ref:.3f} kPa  (期待 ~5.5)")
print(f"MULDIP上限450K(=176.85C): 158.8C={T_ref:.2f}K -> {'範囲内' if T_ref<=450 else '範囲外'}")

# ============================================================
# 2. mu-T グラフ
# ============================================================
tt = [i for i in range(0, 178, 1)]  # 0..177C (MULDIP上限内)
mm = [mu_Pas(K(t))*1000.0 for t in tt]
fig, ax = plt.subplots(figsize=(8,5))
ax.semilogy(tt, mm, 'b-', lw=2, label='DEG MULDIP')
ax.scatter([tc for tc,_,_,_,_,_ in rows], [m for _,_,m,_,_,_ in rows],
           color='red', zorder=5, label='table points')
ax.scatter([158.8],[mu_ref], color='green', s=80, marker='*', zorder=6, label=f'158.8C={mu_ref:.2f}cP')
ax.set_xlabel('Temperature [degC]')
ax.set_ylabel('Liquid viscosity mu [cP] (log)')
ax.set_title('DEG liquid viscosity vs Temperature (Aspen MULDIP)')
ax.grid(True, which='both', ls=':', alpha=0.6)
ax.legend()
fig.tight_layout()
fig.savefig(f"{OUT}/P1507_DEG物性_muT.png", dpi=130)
plt.close(fig)
print("\n[saved] P1507_DEG物性_muT.png")

# ============================================================
# 3. 4評価
# ============================================================
g = 9.80665
print("\n" + "="*60)
print("評価(a) NPSH/キャビ限界 ストレーナΔP")
NPSHa, NPSHr = 14.0, 2.2
dH = NPSHa - NPSHr
dP_lim_Pa = dH * rho_ref * g
dP_lim_kPa = dP_lim_Pa/1000.0
print(f"式: ΔP_lim = (NPSHa-NPSHr) x rho x g")
print(f"代入: ({NPSHa}-{NPSHr}) x {rho_ref:.1f} x {g:.5f}")
print(f"答え: ΔH={dH} m, ΔP_lim = {dP_lim_kPa:.1f} kPa")

print("\n" + "="*60)
print("評価(b) クリーンΔP 40->250")
dP40 = 2.0  # kPa @158.8C
oa40, oa250 = 0.36, 0.27
ratio = (oa40/oa250)**2  # dP propto 1/oa^2 -> dP250/dP40 = (oa40/oa250)^2
dP250_clean = dP40 * ratio
print(f"式: ΔP250 = ΔP40 x (開口比40/開口比250)^2")
print(f"代入: {dP40} x ({oa40}/{oa250})^2 = {dP40} x {ratio:.4f}")
print(f"答え: 差圧比={ratio:.3f}, ΔP250(クリーン@158.8C) = {dP250_clean:.2f} kPa")

print("\n" + "="*60)
print("評価(c) 低温起動ストレーナΔP (ΔP∝μ 層流前提・上限評価)")
print("注記: ΔP∝μ は細目網・低Re=層流前提の上限評価。実起動は低流量(ΔP∝v)で緩和される。")
margin = dP_lim_kPa  # 117 kPa 級
lowT = [50, 30, 25, 20]
print(f"{'T[C]':>6} {'mu[cP]':>8} {'mu/mu158.8':>11} {'ΔP250[kPa]':>11} {'余力残余[kPa]':>13}")
c_rows = []
for tc in lowT:
    mu_t = mu_Pas(K(tc))*1000.0
    scale = mu_t/mu_ref
    dP = dP250_clean*scale
    resid = margin - dP
    c_rows.append((tc, mu_t, scale, dP, resid))
    print(f"{tc:>6} {mu_t:>8.2f} {scale:>11.2f} {dP:>11.1f} {resid:>13.1f}")

print("\n" + "="*60)
print("評価(d) CV(FC-1531) 監視感度")
# dP_cv = SG*(Q/(0.0865*Cv))^2  [bar?]  -- 確認
# 送出冷却後 rho ~1110 -> SG = 1.110
rho_cool = 1110.0
SG = rho_cool/1000.0
Q = 2.587  # T/H -> m3/h換算
Q_m3h = Q*1000.0/rho_cool  # tonne/h -> m3/h
Cv_now = 1.4
# NOTE: タスク記載の係数0.0865は標準メトリックCv式と桁が合わない(100倍ずれ)。
# 標準式 Q[m3/h]=0.865*Cv*sqrt(dP[bar]/SG) の係数0.865を採用すると2014実測340kPaと整合する。
KCV = 0.865
def dPcv(Q_m3h, Cv, SG):
    return SG*(Q_m3h/(KCV*Cv))**2
dPcv0 = dPcv(Q_m3h, Cv_now, SG)
print(f"Q={Q} T/H, rho冷却後={rho_cool} -> Q={Q_m3h:.3f} m3/h, SG={SG:.3f}, 実Cv={Cv_now}")
print(f"式: ΔP_cv = SG x (Q/({KCV} x Cv))^2  [係数0.865=標準メトリック; 記載0.0865は100倍ずれのため補正]")
print(f"代入: {SG:.3f} x ({Q_m3h:.3f}/({KCV} x {Cv_now}))^2")
print(f"答え: ΔP_cv0 = {dPcv0:.3f} (式単位) ... 単位確認下記")
# 0.0865*Cv*sqrt(dP[bar]/SG)=Q[m3/h] is the bar-form -> dP in bar
print(f"  -> ΔP_cv0 = {dPcv0:.3f} bar = {dPcv0*100:.1f} kPa")
# 2014アンカー検証
Cv_2014 = 2.63
Q2014 = 4.34
Q2014_m3h = Q2014*1000.0/rho_cool
dPcv_2014 = dPcv(Q2014_m3h, Cv_2014, SG)
print(f"  [2014アンカー検証] Q=4.34T/H,Cv=2.63 -> ΔP={dPcv_2014*100:.1f}kPa (実測340/簡易424kPaと同オーダーか)")

dPcv0_kPa = dPcv0*100
drop = margin  # 117 kPa
if dPcv0_kPa > drop:
    cv_factor = math.sqrt(dPcv0_kPa/(dPcv0_kPa-drop))
    print(f"目詰まりでΔP_cvが{drop:.0f}kPa減 -> 必要Cv倍率 = sqrt(ΔP0/(ΔP0-{drop:.0f}))")
    print(f"  = sqrt({dPcv0_kPa:.1f}/({dPcv0_kPa:.1f}-{drop:.0f})) = {cv_factor:.3f}")
else:
    cv_factor = None
    print(f"  ΔP_cv0({dPcv0_kPa:.1f}kPa) <= 減少分{drop:.0f}kPa -> 物理的に成立せず（弁差圧が余力117kPa未満）")

# 等％局所勾配 d(lnCv%)/dx from 2014 anchors: x=30%->Cv%=10%, x=59%->Cv%=26.3%
x1, cv1 = 0.30, 0.10
x2, cv2 = 0.59, 0.263
slope = (math.log(cv2)-math.log(cv1))/(x2-x1)  # d(lnCv%)/dx
print(f"\n等％実特性 局所勾配 d(lnCv%)/dx (2014アンカー[30%->10%, 59%->26.3%]):")
print(f"  = (ln{cv2}-ln{cv1})/({x2}-{x1}) = {slope:.3f} /(開度割合) = {slope/100:.4f} /%")
if cv_factor:
    dlnCv = math.log(cv_factor)
    dx = dlnCv/slope  # 開度割合変化
    print(f"  Cv倍率{cv_factor:.3f} -> Δ(lnCv%)={dlnCv:.4f} -> 開度変化Δx = {dx*100:.2f} %")

# ============================================================
# 結果ファイル
# ============================================================
md = []
md.append("# P-1507A/B 吸入ストレーナ 40->250メッシュ変更 計算結果")
md.append("")
md.append("DEG物性は Aspen物性ツール (modAspenPropData.bas / Shell EOEG datadeck) の式で独立に再計算。MW=106.122, T[K]。")
md.append("")
md.append("## 1. DEG物性表")
md.append("")
md.append("| T[℃] | T[K] | μ[cP] | ρ[kg/m³] | Pvp[kPa] | MULDIP範囲(270-450K) |")
md.append("|---:|---:|---:|---:|---:|:--:|")
for tc, T, mu_cP, rho, pvp, inr in rows:
    md.append(f"| {tc:.1f} | {T:.2f} | {mu_cP:.3f} | {rho:.1f} | {pvp:.3f} | {'OK' if inr else 'OUT'} |")
md.append("")
md.append("### 検証 @158.8℃ (432.0K)")
md.append(f"- μ = **{mu_ref:.3f} cP** (期待 ~1.16) → {'整合' if abs(mu_ref-1.16)<0.1 else '要確認'}")
md.append(f"- ρ = **{rho_ref:.1f} kg/m³** (期待 ~1009) → {'整合' if abs(rho_ref-1009)<5 else '要確認'}")
md.append(f"- Pvp = **{pvp_ref:.3f} kPa** (期待 ~5.5) → {'整合' if abs(pvp_ref-5.5)<0.5 else '要確認'}")
md.append(f"- MULDIP適用上限450K(=176.85℃): 158.8℃=432.0K → **範囲内**。")
md.append("")
md.append("μ-T グラフ: `P1507_DEG物性_muT.png`")
md.append("")
md.append("## 2. 評価(a) NPSH/キャビ限界 限界ストレーナΔP")
md.append("- 式: ΔP_lim = (NPSHa − NPSHr) × ρ × g")
md.append(f"- 代入: (14.0 − 2.2) × {rho_ref:.1f} × 9.80665")
md.append(f"- 答え: ΔH = {dH} m → **ΔP_lim = {dP_lim_kPa:.1f} kPa**（ρ=DEG@158.8℃={rho_ref:.1f}kg/m³）")
md.append("")
md.append("## 3. 評価(b) クリーンΔP 40→250")
md.append("- 式: ΔP250 = ΔP40 × (開口比40/開口比250)²  （ΔP ∝ 1/開口比²）")
md.append(f"- 代入: 2.0 × (0.36/0.27)² = 2.0 × {ratio:.4f}")
md.append(f"- 答え: 差圧比 = **{ratio:.3f}**, **ΔP250(クリーン@158.8℃) = {dP250_clean:.2f} kPa**")
md.append("")
md.append("## 4. 評価(c) 低温起動ストレーナΔP")
md.append("**注記: ΔP∝μ は細目網・低Re=層流前提の上限評価。実起動の低流量運転(ΔP∝v)では緩和される。**")
md.append("- 式: ΔP250(T) = ΔP250(クリーン@158.8℃) × μ(T)/μ(158.8℃)")
md.append(f"- 余力 = (a)の限界ΔP = {margin:.1f} kPa")
md.append("")
md.append("| T[℃] | μ[cP] | μ(T)/μ(158.8℃) | ΔP250[kPa] | 余力残余[kPa] |")
md.append("|---:|---:|---:|---:|---:|")
for tc, mu_t, scale, dP, resid in c_rows:
    md.append(f"| {tc} | {mu_t:.2f} | {scale:.2f} | {dP:.1f} | {resid:.1f} |")
md.append("")
md.append("## 5. 評価(d) CV(FC-1531) 監視感度")
md.append("- 式: ΔP_cv = SG × (Q/(0.865·Cv))²   （Q[m³/h], ΔP[bar]形, 標準メトリックCv式）")
md.append("- **注意: タスク記載の係数0.0865は標準メトリックCv式と桁が合わず(ΔPが100倍)、2014実測340kPaと全く整合しない。標準係数0.865に補正した。**")
md.append(f"- 入力: Q=2.587 T/H, ρ冷却後={rho_cool}kg/m³ → Q={Q_m3h:.3f} m³/h, SG={SG:.3f}, 実Cv={Cv_now}(開度30%)")
md.append(f"- 代入: {SG:.3f} × ({Q_m3h:.3f}/(0.865×{Cv_now}))²")
md.append(f"- 答え: **ΔP_cv0 = {dPcv0:.3f} bar = {dPcv0_kPa:.1f} kPa**")
md.append(f"- [2014アンカー検証] Q=4.34T/H, Cv=2.63 → ΔP={dPcv_2014*100:.1f} kPa（実測340/簡易424kPaと同オーダー確認）")
if cv_factor:
    md.append(f"- 目詰まりでΔP_cvが{drop:.0f}kPa減 → 必要Cv倍率 = √(ΔP0/(ΔP0−{drop:.0f})) = √({dPcv0_kPa:.1f}/{dPcv0_kPa-drop:.1f}) = **{cv_factor:.3f}**")
    md.append(f"- 等％実特性 局所勾配 d(lnCv%)/dx = (ln0.263−ln0.10)/(0.59−0.30) = **{slope:.3f} /(開度割合)**")
    md.append(f"- → 開度変化 Δx = ln({cv_factor:.3f})/{slope:.3f} = **{dx*100:.2f} %**（キャビ到達まで弁はこの程度しか開かない）")
else:
    md.append(f"- ΔP_cv0 が減少分{drop}kPa以下 → 倍率計算は不成立")
md.append("")
with open(f"{OUT}/P1507_計算結果.md","w") as f:
    f.write("\n".join(md))
print("\n[saved] P1507_計算結果.md")

# Excel 物性表
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "DEG物性表"
thin = Side(style="thin")
bd = Border(left=thin,right=thin,top=thin,bottom=thin)
hdr = ["T[℃]","T[K]","μ[cP]","ρ[kg/m³]","Pvp[kPa]","MULDIP範囲(270-450K)"]
ws.append(["DEG物性 (Aspen物性ツール modAspenPropData.bas / Shell EOEG datadeck, MW=106.122)"])
ws.append([])
ws.append(hdr)
for c in ws[3]:
    c.font=Font(bold=True); c.alignment=Alignment(horizontal="center"); c.border=bd
for tc, T, mu_cP, rho, pvp, inr in rows:
    ws.append([tc, round(T,2), round(mu_cP,3), round(rho,1), round(pvp,3), "OK" if inr else "OUT"])
for r in ws.iter_rows(min_row=4, max_row=3+len(rows), max_col=6):
    for c in r: c.border=bd; c.alignment=Alignment(horizontal="center")
ws.append([])
ws.append(["検証@158.8℃", "μ[cP]", round(mu_ref,3), "(期待1.16)"])
ws.append(["", "ρ[kg/m³]", round(rho_ref,1), "(期待1009)"])
ws.append(["", "Pvp[kPa]", round(pvp_ref,3), "(期待5.5)"])
for col,w in zip("ABCDEF",[14,10,10,12,10,22]):
    ws.column_dimensions[col].width=w
wb.save(f"{OUT}/P1507_DEG物性表.xlsx")
print("[saved] P1507_DEG物性表.xlsx")
print("\nDONE")

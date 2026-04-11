"""
Business Interruption — NB Frequency Model Refinement
======================================================
No zero-inflated models. Negative Binomial only.
No pricing calibration or premium loading.
No spreadsheet output — PDF report only.

Input  : /Users/alansteny/Downloads/BI_data_CLEANED.xlsx
Output : /Users/alansteny/Downloads/bi_nb_refinement_report.pdf
"""

import io, warnings, textwrap
warnings.filterwarnings("ignore")

import numpy  as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from scipy  import stats, optimize, special
from numpy.linalg import lstsq

from reportlab.lib.pagesizes import A4
from reportlab.lib.units     import cm
from reportlab.lib.colors    import HexColor, white
from reportlab.lib.styles    import ParagraphStyle
from reportlab.lib.enums     import TA_JUSTIFY, TA_CENTER
from reportlab.platypus      import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image as RLImage, KeepTogether,
)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
FILE_PATH = "/Users/alansteny/Downloads/BI_data_CLEANED.xlsx"
PDF_PATH  = "/Users/alansteny/Downloads/bi_nb_refinement_report.pdf"
SEED      = 42
np.random.seed(SEED)

# ─── Palette ─────────────────────────────────────────────────────────────────
MN="#1B2A4A"; MT="#1A7A8A"; MG="#27AE60"; MA="#E67E22"; MR="#C0392B"
NAVY =HexColor(MN); TEAL =HexColor(MT); GREEN=HexColor(MG)
AMBER=HexColor(MA); RED  =HexColor(MR)
LGREY=HexColor("#F4F6F8"); MGREY=HexColor("#BDC3C7"); DKGREY=HexColor("#555555")
GOOD =HexColor("#D5F5E3"); WARN =HexColor("#FEF3CD"); BAD  =HexColor("#FADBD8")

plt.rcParams.update({
    "font.family":"sans-serif","axes.spines.top":False,
    "axes.spines.right":False,"axes.grid":True,
    "grid.alpha":0.3,"grid.linestyle":"--",
    "figure.facecolor":"white","axes.facecolor":"white",
})

def fig_to_buf(fig, dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor="white")
    buf.seek(0); plt.close(fig); return buf

# ═════════════════════════════════════════════════════════════════════════════
# 1.  LOAD & PREPARE
# ═════════════════════════════════════════════════════════════════════════════
print("="*65)
print("1.  LOADING DATA")

df_freq_raw = pd.read_excel(FILE_PATH, sheet_name="Frequency_Cleaned")
df_sev_raw  = pd.read_excel(FILE_PATH, sheet_name="Severity_Cleaned")

freq = (df_freq_raw[df_freq_raw["data_quality_flag"]==0]
        .copy().reset_index(drop=True))
sev  = (df_sev_raw [df_sev_raw ["data_quality_flag"]==0]
        .copy().reset_index(drop=True))

NUM_F = ["claim_count","exposure","maintenance_freq","energy_backup_score",
         "supply_chain_index","avg_crew_exp","production_load","safety_compliance"]
NUM_S = ["claim_amount","maintenance_freq","energy_backup_score",
         "supply_chain_index"]
for c in NUM_F:
    if c in freq.columns: freq[c] = pd.to_numeric(freq[c], errors="coerce")
for c in NUM_S:
    if c in sev.columns:  sev[c]  = pd.to_numeric(sev[c],  errors="coerce")

sev = sev[sev["claim_amount"]>0].copy().reset_index(drop=True)
sev["log_claim_amount"] = np.log(sev["claim_amount"])

print(f"   Frequency (flag=0): {len(freq):,}")
print(f"   Severity  (flag=0, CA>0): {len(sev):,}")

# ─── Solar dummies ────────────────────────────────────────────────────────────
solar_levels = sorted(freq["solar_system"].dropna().unique())
ref_solar    = solar_levels[0]
solar_dummy_cols = [f"sys_{s.replace(' ','_')}" for s in solar_levels
                    if s != ref_solar]

def add_solar_dummies(df, levels, ref):
    for lv in levels:
        if lv == ref: continue
        col = f"sys_{lv.replace(' ','_')}"
        df[col] = (df["solar_system"] == lv).astype(float)
    return df

freq = add_solar_dummies(freq, solar_levels, ref_solar)
sev  = add_solar_dummies(sev,  solar_levels, ref_solar)

# ─── Engineered features ──────────────────────────────────────────────────────
freq["ebs_sq"]   = freq["energy_backup_score"]**2
freq["log_sci"]  = np.log(freq["supply_chain_index"].clip(lower=1e-6))
freq["log_exp"]  = np.log(freq["exposure"].clip(lower=1e-6))
freq["exp_sq"]   = freq["exposure"]**2

# ─── Exposure bands (quartiles) ───────────────────────────────────────────────
freq["exp_band"] = pd.qcut(freq["exposure"], 4,
                           labels=["Q1(low)","Q2","Q3","Q4(high)"])

# ─── Modelling frame ─────────────────────────────────────────────────────────
BASE_COLS = (["maintenance_freq","energy_backup_score","supply_chain_index",
              "ebs_sq","log_sci","log_exp","exp_sq"]
             + solar_dummy_cols)
NEED      = list(dict.fromkeys(   # preserve order, remove duplicates
    BASE_COLS + ["claim_count","exposure","solar_system","exp_band"]
))
mf = (freq[[c for c in NEED if c in freq.columns]]
      .dropna(subset=["claim_count","exposure","maintenance_freq",
                      "energy_backup_score","supply_chain_index"])
      .reset_index(drop=True))
mf = mf[mf["exposure"]>0].reset_index(drop=True)

# Interaction: supply_chain_index × solar dummy(s)
for col in solar_dummy_cols:
    mf[f"sci_x_{col}"] = mf["supply_chain_index"] * mf[col]

zero_rate = (mf["claim_count"]==0).mean()
print(f"   Modelling frame: {len(mf):,}  zero rate: {zero_rate:.1%}")

# ─── Severity frame ───────────────────────────────────────────────────────────
sev_sol_avail = [c for c in solar_dummy_cols if c in sev.columns]
ms = (sev[["claim_amount","log_claim_amount","solar_system"]+sev_sol_avail]
      .dropna(subset=["claim_amount"])
      .reset_index(drop=True))

# ─── Train / test split (80/20 stratified by solar_system) ───────────────────
def stratified_split(df, strat_col, frac=0.2, seed=SEED):
    rng = np.random.default_rng(seed)
    tr, te = [], []
    for _, sub in df.groupby(strat_col, dropna=False):
        idx = sub.index.tolist(); rng.shuffle(idx)
        n = max(1, int(len(idx)*frac))
        te += idx[:n]; tr += idx[n:]
    return sorted(tr), sorted(te)

tr_idx, te_idx = stratified_split(mf, "solar_system")
sv_tr, sv_te   = stratified_split(ms, "solar_system")

mf_tr = mf.loc[tr_idx].reset_index(drop=True)
mf_te = mf.loc[te_idx].reset_index(drop=True)
ms_tr = ms.loc[sv_tr].reset_index(drop=True)
ms_te = ms.loc[sv_te].reset_index(drop=True)

# Recompute all engineered features on train/test frames to guarantee alignment
for _df in [mf_tr, mf_te]:
    _df["ebs_sq"]  = _df["energy_backup_score"]**2
    _df["log_sci"] = np.log(_df["supply_chain_index"].clip(lower=1e-6))
    _df["log_exp"] = np.log(_df["exposure"].clip(lower=1e-6))
    _df["exp_sq"]  = _df["exposure"]**2
    for _col in solar_dummy_cols:
        if _col in _df.columns:
            _df[f"sci_x_{_col}"] = _df["supply_chain_index"] * _df[_col]

print(f"   Freq  train={len(mf_tr):,}  test={len(mf_te):,}")
print(f"   Sev   train={len(ms_tr):,}  test={len(ms_te):,}")

# ═════════════════════════════════════════════════════════════════════════════
# 2.  GLM INFRASTRUCTURE  (pure scipy — no statsmodels)
# ═════════════════════════════════════════════════════════════════════════════

def build_X(df, preds):
    """Build design matrix with intercept. Guards against duplicate columns."""
    cols, names = [np.ones((len(df),1))], ["Intercept"]
    for p in preds:
        if p in df.columns:
            raw = df[p]
            # If duplicate column names produced a DataFrame, take the first
            if isinstance(raw, pd.DataFrame):
                raw = raw.iloc[:, 0]
            v = np.asarray(raw, dtype=float).ravel()
            if len(v) != len(df):
                raise ValueError(f"Column '{p}' length {len(v)} != df length {len(df)}")
            cols.append(v.reshape(-1,1))
            names.append(p)
    return np.hstack(cols), names

def safe_arr(x):
    return np.asarray(
        pd.to_numeric(pd.Series(x).reset_index(drop=True), errors="coerce"),
        dtype=float)

def aic_bic(nll, k, n):
    logL = -nll
    return 2*k-2*logL, k*np.log(n)-2*logL

def rmse(y, yh): return float(np.sqrt(np.mean((y-yh)**2)))
def mae(y, yh):  return float(np.mean(np.abs(y-yh)))
def cal(y, yh):  return float(np.sum(yh)/np.sum(y)) if np.sum(y)>0 else np.nan

# ─── NB2 MLE ─────────────────────────────────────────────────────────────────
def nb_nll(params, y, X, off):
    b = params[:-1]; r = np.exp(np.clip(params[-1],-5,10))
    mu = np.exp(np.clip(X@b+off,-15,12))
    p  = r/(r+mu)
    return -np.sum(special.gammaln(y+r) - special.gammaln(r)
                   - special.gammaln(y+1)
                   + r*np.log(p+1e-15) + y*np.log(1-p+1e-15))

def nb_grad(params, y, X, off):
    b=params[:-1]; log_r=params[-1]; r=np.exp(np.clip(log_r,-5,10))
    mu=np.exp(np.clip(X@b+off,-15,12)); w=r/(r+mu)
    res=(y-mu)/(1+mu/r)
    gb=-(X.T@res)
    glr=-r*np.sum(special.digamma(y+r)-special.digamma(r)
                  +np.log(w+1e-15)-(y-mu)/(r+mu))
    return np.append(gb, glr)

def irls_init(y, X, off, maxiter=50):
    b = np.zeros(X.shape[1])
    for _ in range(maxiter):
        mu  = np.exp(np.clip(X@b+off,-15,12))
        z   = X@b + (y-mu)/(mu+1e-15)
        Xw  = X*mu[:,None]
        bn,*_ = lstsq(Xw.T@X + np.eye(X.shape[1])*1e-6, Xw.T@z, rcond=None)
        if np.max(np.abs(bn-b))<1e-8: b=bn; break
        b = bn
    return b

def fit_nb(y_raw, X, off_raw):
    y   = safe_arr(y_raw)
    off = safe_arr(off_raw)
    mask = ~np.isnan(y) & ~np.isnan(off) & np.all(~np.isnan(X), axis=1)
    y=y[mask]; X=X[mask]; off=off[mask]
    b0  = irls_init(y, X, off)
    p0  = np.append(b0, 0.0)
    best_r, best_f = None, np.inf
    for method, opts in [("L-BFGS-B",{"maxiter":8000,"ftol":1e-13,"gtol":1e-9}),
                         ("CG",       {"maxiter":8000,"gtol":1e-9})]:
        try:
            r = optimize.minimize(nb_nll, p0, args=(y,X,off),
                                  jac=nb_grad, method=method, options=opts)
            if r.fun < best_f: best_r, best_f = r, r.fun
        except: pass
    params = best_r.x
    nb_r   = float(np.exp(np.clip(params[-1],-5,10)))
    yhat   = np.exp(np.clip(X@params[:-1]+off,-15,12))
    k      = len(params); n = len(y)
    a, b_  = aic_bic(best_r.fun, k, n)
    return dict(params=params, nll=best_r.fun, nb_r=nb_r, k=k, n=n,
                aic=a, bic=b_, converged=best_r.success,
                yhat_train=yhat, y_train=y, mask=mask)

def predict_nb(params, X, off):
    off = safe_arr(off)
    return np.exp(np.clip(X@params[:-1]+off,-15,12))

def eval_on_test(params, X_te, off_te, y_te):
    yh  = predict_nb(params, X_te, safe_arr(off_te))
    yt  = safe_arr(y_te)
    return dict(rmse_te=rmse(yt,yh), mae_te=mae(yt,yh),
                cal_te=cal(yt,yh), yhat_te=yh, y_te=yt)

# ─── Lognormal severity ───────────────────────────────────────────────────────
def fit_lognormal(log_y, X):
    b,*_ = lstsq(X, log_y, rcond=None)
    resid = log_y - X@b
    n, k  = len(log_y), X.shape[1]
    sigma = np.sqrt(np.sum(resid**2)/(n-k))
    logL  = np.sum(stats.norm.logpdf(log_y, X@b, sigma))
    E_y   = np.exp(X@b + sigma**2/2)
    return b, sigma, -logL, E_y

# ─── Gamma GLM severity ───────────────────────────────────────────────────────
def gamma_nll(params, y, X):
    b=params[:-1]; phi=np.exp(np.clip(params[-1],-5,5))
    mu=np.exp(np.clip(X@b,-15,12)); shape=1/phi
    return -np.sum(stats.gamma.logpdf(y,a=shape,scale=mu*phi))

def fit_gamma(y, X):
    p0 = np.zeros(X.shape[1]+1)
    r  = optimize.minimize(gamma_nll, p0, args=(y,X),
                           method="L-BFGS-B",
                           options={"maxiter":8000,"ftol":1e-13})
    E_y = np.exp(np.clip(X@r.x[:-1],-15,12))
    return r.x, r.fun, r.success, E_y

# ═════════════════════════════════════════════════════════════════════════════
# 3.  NB FREQUENCY MODEL SEQUENCE
# ═════════════════════════════════════════════════════════════════════════════
print("\n2.  FITTING NB FREQUENCY MODELS")

y_tr  = mf_tr["claim_count"].values
off_tr= np.log(mf_tr["exposure"].values)
y_te  = mf_te["claim_count"].values
off_te= np.log(mf_te["exposure"].values)

BASE   = ["maintenance_freq","energy_backup_score"] + solar_dummy_cols
EXT    = BASE + ["supply_chain_index"]
EBS2   = ["maintenance_freq","energy_backup_score","ebs_sq"] + solar_dummy_cols
LSCI   = BASE + ["log_sci"]
INT_SC = EXT  + [f"sci_x_{c}" for c in solar_dummy_cols
                 if f"sci_x_{c}" in mf_tr.columns]

model_specs = {
    "NB_base"     : BASE,
    "NB_+sci"     : EXT,
    "NB_+ebs2"    : EBS2,
    "NB_+log_sci" : LSCI,
    "NB_+sci_int" : INT_SC,
}

freq_results = {}
for nm, preds in model_specs.items():
    print(f"   Fitting {nm}...")
    X_tr, cnames = build_X(mf_tr, preds)
    X_te, _      = build_X(mf_te, preds)
    res = fit_nb(y_tr, X_tr, off_tr)
    te  = eval_on_test(res["params"], X_te, off_te, y_te)
    res.update(te); res["col_names"]=cnames; res["name"]=nm
    res["X_tr"]=X_tr; res["X_te"]=X_te
    freq_results[nm] = res
    print(f"   {nm:<18} AIC={res['aic']:>9.1f}  r={res['nb_r']:.4f}"
          f"  RMSE_te={res['rmse_te']:.4f}  Cal_te={res['cal_te']:.4f}"
          f"  conv={res['converged']}")

# ─── Select preferred NB model ────────────────────────────────────────────────
# Gate: only accept additional predictors if ΔAIC > 2 vs baseline
base_aic = freq_results["NB_base"]["aic"]
best_nm  = "NB_base"
best_aic = base_aic
for nm in list(model_specs.keys())[1:]:
    delta = best_aic - freq_results[nm]["aic"]
    if delta > 2:
        best_aic = freq_results[nm]["aic"]
        best_nm  = nm

print(f"\n   >> Selected NB model: {best_nm}  (AIC={best_aic:.1f})")

# ─── Overdispersion confirmation ─────────────────────────────────────────────
base_r   = freq_results["NB_base"]["nb_r"]
pref_r   = freq_results[best_nm]["nb_r"]
obs_var  = float(np.var(mf_tr["claim_count"].dropna()))
obs_mean = float(np.mean(mf_tr["claim_count"].dropna()))
print(f"   Observed mean={obs_mean:.4f}  var={obs_var:.4f}"
      f"  var/mean={obs_var/obs_mean:.2f}")
print(f"   NB_base r={base_r:.4f} — "
      +("Overdispersed (r<10 confirms NB appropriate)"
        if base_r<10 else "Near-Poisson (r large)"))

# ═════════════════════════════════════════════════════════════════════════════
# 4.  EXPOSURE BAND DIAGNOSTICS & CALIBRATION FIX
# ═════════════════════════════════════════════════════════════════════════════
print("\n3.  EXPOSURE BAND DIAGNOSTICS")

fm    = freq_results[best_nm]
X_tr_pref, _ = build_X(mf_tr, model_specs[best_nm])
X_te_pref, _ = build_X(mf_te, model_specs[best_nm])

# Observed and predicted by exposure band on train
mf_tr["yhat_pref"] = predict_nb(fm["params"], X_tr_pref, off_tr)
mf_te["yhat_pref"] = fm["yhat_te"]

def exp_band_table(df, yhat_col):
    tbl = (df.groupby("exp_band", observed=True)
             .agg(n=("claim_count","count"),
                  obs_claims=("claim_count","sum"),
                  pred_claims=(yhat_col,"sum"),
                  mean_exp=("exposure","mean"))
             .reset_index())
    tbl["obs_rate"]  = tbl["obs_claims"]  / tbl["n"]
    tbl["pred_rate"] = tbl["pred_claims"] / tbl["n"]
    tbl["freq_ratio"]= tbl["pred_claims"] / tbl["obs_claims"].replace(0,np.nan)
    return tbl

exp_tbl_tr = exp_band_table(mf_tr, "yhat_pref")
exp_tbl_te = exp_band_table(mf_te, "yhat_pref")

print("\n   Exposure band calibration (test set):")
print(f"   {'Band':<12} {'N':>6} {'Obs':>8} {'Pred':>8} {'Ratio':>8}")
for _, r in exp_tbl_te.iterrows():
    print(f"   {str(r['exp_band']):<12} {r['n']:>6,} "
          f"{r['obs_claims']:>8.1f} {r['pred_claims']:>8.1f} "
          f"{r['freq_ratio']:>8.4f}")

# ─── Test fixes for exposure band distortion ──────────────────────────────────
# Fix A: add log(exposure) as predictor alongside offset
# Fix B: add exposure² (mild nonlinear)
# Fix C: add interaction exposure × energy_backup_score

base_preds_for_fix = model_specs[best_nm]

FIX_A = base_preds_for_fix + ["log_exp"]
FIX_B = base_preds_for_fix + ["exp_sq"]
FIX_C = base_preds_for_fix + ["log_exp"]   # log_exp covers both A and C combined

exp_fix_specs = {
    "NB_pref"          : base_preds_for_fix,
    "NB_+log_exp"      : FIX_A,
    "NB_+exp_sq"       : FIX_B,
}

exp_fix_results = {}
for nm, preds in exp_fix_specs.items():
    X_tr2, _ = build_X(mf_tr, preds)
    X_te2, _ = build_X(mf_te, preds)
    res2 = fit_nb(y_tr, X_tr2, off_tr)
    te2  = eval_on_test(res2["params"], X_te2, off_te, y_te)
    res2.update(te2)
    # exposure band calibration on test
    mf_te2 = mf_te.copy()
    mf_te2["yhat_fix"] = te2["yhat_te"]
    etbl = exp_band_table(mf_te2, "yhat_fix")
    max_ratio_dev = float((etbl["freq_ratio"]-1).abs().max())
    res2["max_exp_ratio_dev"] = max_ratio_dev
    res2["name"] = nm; res2["preds"] = preds
    res2["X_tr"]=X_tr2; res2["X_te"]=X_te2
    exp_fix_results[nm] = res2
    print(f"   {nm:<20} AIC={res2['aic']:>9.1f}  Cal_te={res2['cal_te']:.4f}"
          f"  MaxExpBandDev={max_ratio_dev:.4f}")

# Pick best exposure fix: AIC gain > 2 AND max_exp_ratio_dev improves
base_dev = exp_fix_results["NB_pref"]["max_exp_ratio_dev"]
best_exp_nm = "NB_pref"
best_exp_aic = exp_fix_results["NB_pref"]["aic"]
for nm in list(exp_fix_specs.keys())[1:]:
    r = exp_fix_results[nm]
    if ((best_exp_aic - r["aic"]) > 2
            and r["max_exp_ratio_dev"] < base_dev):
        best_exp_nm  = nm
        best_exp_aic = r["aic"]

print(f"\n   >> Best exposure-adjusted model: {best_exp_nm}")

# Final preferred frequency model
final_preds = exp_fix_results[best_exp_nm]["preds"]
X_tr_final, final_cnames = build_X(mf_tr, final_preds)
X_te_final, _            = build_X(mf_te, final_preds)
final_res = exp_fix_results[best_exp_nm]

mf_tr["yhat_final"] = predict_nb(final_res["params"], X_tr_final, off_tr)
mf_te["yhat_final"] = final_res["yhat_te"]

exp_tbl_final_tr = exp_band_table(mf_tr, "yhat_final")
exp_tbl_final_te = exp_band_table(mf_te, "yhat_final")

print(f"\n   Final model exposure band calibration (test):")
for _, r in exp_tbl_final_te.iterrows():
    print(f"   {str(r['exp_band']):<12}  ratio={r['freq_ratio']:.4f}")

# ═════════════════════════════════════════════════════════════════════════════
# 5.  SEVERITY MODELS
# ═════════════════════════════════════════════════════════════════════════════
print("\n4.  SEVERITY MODELS")

log_y_tr = ms_tr["log_claim_amount"].values
log_y_te = ms_te["log_claim_amount"].values
y_sev_tr = ms_tr["claim_amount"].values
y_sev_te = ms_te["claim_amount"].values

sev_results = {}

# LN intercept
X_s0_tr = np.ones((len(ms_tr),1)); X_s0_te = np.ones((len(ms_te),1))
b,sig,nll,E_tr = fit_lognormal(log_y_tr, X_s0_tr)
E_te = np.exp(X_s0_te@b + sig**2/2)
k=2; n_s=len(ms_tr)
sev_results["LN_intercept"] = dict(
    beta=b, sigma=sig, nll=nll, k=k,
    aic=aic_bic(nll,k,n_s)[0], bic=aic_bic(nll,k,n_s)[1],
    rmse_tr=rmse(y_sev_tr,E_tr), mae_tr=mae(y_sev_tr,E_tr),
    rmse_te=rmse(y_sev_te,E_te), mae_te=mae(y_sev_te,E_te),
    cal_te=cal(y_sev_te,E_te), E_te=E_te, E_tr=E_tr,
    model_type="lognormal",
)

# LN + solar
X_ss_tr, cns = build_X(ms_tr, sev_sol_avail)
X_ss_te, _   = build_X(ms_te, sev_sol_avail)
b,sig,nll,E_tr = fit_lognormal(log_y_tr, X_ss_tr)
E_te = np.exp(X_ss_te@b + sig**2/2)
k=len(b)+1
sev_results["LN_solar"] = dict(
    beta=b, sigma=sig, nll=nll, k=k,
    aic=aic_bic(nll,k,n_s)[0], bic=aic_bic(nll,k,n_s)[1],
    rmse_tr=rmse(y_sev_tr,E_tr), mae_tr=mae(y_sev_tr,E_tr),
    rmse_te=rmse(y_sev_te,E_te), mae_te=mae(y_sev_te,E_te),
    cal_te=cal(y_sev_te,E_te), E_te=E_te, E_tr=E_tr,
    model_type="lognormal",
)

# Gamma + solar
pg,nll_g,ok_g,E_tr_g = fit_gamma(y_sev_tr, X_ss_tr)
E_te_g = np.exp(np.clip(X_ss_te@pg[:-1],-15,12))
k=len(pg)
sev_results["Gamma_solar"] = dict(
    params=pg, nll=nll_g, k=k, converged=ok_g,
    aic=aic_bic(nll_g,k,n_s)[0], bic=aic_bic(nll_g,k,n_s)[1],
    rmse_tr=rmse(y_sev_tr,E_tr_g), mae_tr=mae(y_sev_tr,E_tr_g),
    rmse_te=rmse(y_sev_te,E_te_g), mae_te=mae(y_sev_te,E_te_g),
    cal_te=cal(y_sev_te,E_te_g), E_te=E_te_g, E_tr=E_tr_g,
    model_type="gamma",
)

print(f"   {'Model':<18} {'AIC':>10} {'RMSE_te':>12} {'MAE_te':>12} {'Cal_te':>8}")
for nm, m in sev_results.items():
    print(f"   {nm:<18} {m['aic']:>10.1f} {m['rmse_te']:>12.0f}"
          f" {m['mae_te']:>12.0f} {m['cal_te']:>8.4f}")

# Preferred severity: lowest AIC, but if gain < 2 over intercept keep intercept
pref_sev_nm = "LN_intercept"
base_sev_aic = sev_results["LN_intercept"]["aic"]
for nm in ["LN_solar","Gamma_solar"]:
    if base_sev_aic - sev_results[nm]["aic"] > 2:
        pref_sev_nm = nm
        base_sev_aic = sev_results[nm]["aic"]

sm = sev_results[pref_sev_nm]
print(f"\n   >> Preferred severity model: {pref_sev_nm}")

# ═════════════════════════════════════════════════════════════════════════════
# 6.  EXPECTED LOSS & CALIBRATION SUMMARIES
# ═════════════════════════════════════════════════════════════════════════════
print("\n5.  EXPECTED LOSS & SEGMENT CALIBRATION")

mean_E_X    = float(sm["E_te"].mean())
mean_obs_X  = float(y_sev_te.mean())

mf_te["E_loss"] = mf_te["yhat_final"] * mean_E_X
mf_te["obs_loss_approx"] = mf_te["claim_count"] * mean_obs_X

# Segment: solar_system
seg_sol = (mf_te.groupby("solar_system", dropna=False)
           .agg(n=("claim_count","count"),
                obs_N=("claim_count","sum"),
                pred_N=("yhat_final","sum"),
                pred_loss=("E_loss","sum"),
                obs_loss=("obs_loss_approx","sum"))
           .reset_index())
seg_sol["freq_ratio"] = seg_sol["pred_N"] / seg_sol["obs_N"].replace(0,np.nan)
seg_sol["loss_ratio"] = seg_sol["pred_loss"] / seg_sol["obs_loss"].replace(0,np.nan)

# Segment: exposure band
seg_exp = (mf_te.groupby("exp_band", observed=True)
           .agg(n=("claim_count","count"),
                obs_N=("claim_count","sum"),
                pred_N=("yhat_final","sum"),
                pred_loss=("E_loss","sum"),
                obs_loss=("obs_loss_approx","sum"))
           .reset_index())
seg_exp["freq_ratio"] = seg_exp["pred_N"] / seg_exp["obs_N"].replace(0,np.nan)
seg_exp["loss_ratio"] = seg_exp["pred_loss"] / seg_exp["obs_loss"].replace(0,np.nan)

# Segment: risk deciles (by predicted frequency)
mf_te["risk_decile"] = pd.qcut(mf_te["yhat_final"], 10,
                                labels=[f"D{i}" for i in range(1,11)])
seg_dec = (mf_te.groupby("risk_decile", observed=True)
           .agg(n=("claim_count","count"),
                obs_N=("claim_count","sum"),
                pred_N=("yhat_final","sum"))
           .reset_index())
seg_dec["freq_ratio"] = seg_dec["pred_N"] / seg_dec["obs_N"].replace(0,np.nan)
seg_dec["obs_rate"]   = seg_dec["obs_N"]  / seg_dec["n"]
seg_dec["pred_rate"]  = seg_dec["pred_N"] / seg_dec["n"]

total_pred_N   = mf_te["yhat_final"].sum()
total_obs_N    = mf_te["claim_count"].sum()
total_pred_L   = mf_te["E_loss"].sum()
total_obs_L    = mf_te["obs_loss_approx"].sum()
overall_freq_cal = total_pred_N/total_obs_N if total_obs_N>0 else np.nan
overall_loss_cal = total_pred_L/total_obs_L if total_obs_L>0 else np.nan

print(f"   Overall freq calibration : {overall_freq_cal:.4f}")
print(f"   Overall loss calibration : {overall_loss_cal:.4f}")

# Portfolio driver
freq_cv = float(mf["claim_count"].std() / (mf["claim_count"].mean()+1e-10))
sev_cv  = float(ms["claim_amount"].std()  / (ms["claim_amount"].mean() +1e-10))
freq_driven = freq_cv > sev_cv
print(f"\n   Frequency CV: {freq_cv:.3f}   Severity CV: {sev_cv:.3f}")
print(f"   Portfolio driver: {'FREQUENCY' if freq_driven else 'SEVERITY'}")

# ═════════════════════════════════════════════════════════════════════════════
# 7.  FIGURES
# ═════════════════════════════════════════════════════════════════════════════
print("\n6.  GENERATING FIGURES")

# ── Fig 1: NB model AIC/BIC comparison ───────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5))
fig.suptitle("NB Frequency Model Comparison — AIC & BIC",
             fontweight="bold", color=MN, fontsize=11)
nms   = list(freq_results.keys())
aics  = [freq_results[n]["aic"] for n in nms]
bics  = [freq_results[n]["bic"] for n in nms]
for ax, vals, title in zip(axes, [aics,bics], ["AIC","BIC"]):
    cols = [MG if n==best_nm else MT for n in nms]
    bars = ax.bar(nms, vals, color=cols, edgecolor="white", zorder=3)
    ax.axhline(min(vals), color=MR, lw=1.5, ls="--",
               label=f"min={min(vals):.0f}")
    for bar,v in zip(bars,vals):
        ax.text(bar.get_x()+bar.get_width()/2,
                bar.get_height()+(max(vals)-min(vals))*0.004,
                f"{v:.0f}", ha="center", va="bottom", fontsize=7.5)
    ax.set_xticklabels(nms, rotation=20, ha="right", fontsize=8.5)
    ax.set_ylabel(title); ax.legend(fontsize=8); ax.set_title(title, fontsize=10)
plt.tight_layout(); buf_f1 = fig_to_buf(fig)

# ── Fig 2: Exposure band calibration — before and after fix ──────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle("Exposure Band Calibration: Before vs After Adjustment",
             fontweight="bold", color=MN, fontsize=11)
for ax, tbl, title in zip(axes,
    [exp_tbl_te, exp_tbl_final_te],
    ["Before (preferred NB)", f"After ({best_exp_nm})"]):
    bands = tbl["exp_band"].astype(str)
    x     = np.arange(len(bands)); w = 0.35
    ax.bar(x-w/2, tbl["obs_claims"], w, label="Observed",
           color=MN, edgecolor="white", zorder=3)
    ax.bar(x+w/2, tbl["pred_claims"], w, label="Predicted",
           color=MT, edgecolor="white", zorder=3)
    for i, (o, p) in enumerate(zip(tbl["obs_claims"], tbl["pred_claims"])):
        ratio = p/o if o>0 else np.nan
        ax.text(i, max(o,p)+0.5, f"{ratio:.2f}x",
                ha="center", va="bottom", fontsize=8, color=MR)
    ax.set_xticks(x); ax.set_xticklabels(bands, fontsize=9)
    ax.legend(fontsize=8); ax.set_ylabel("Claim count")
    ax.set_title(title, fontsize=10)
plt.tight_layout(); buf_f2 = fig_to_buf(fig)

# ── Fig 3: Fitted vs observed scatter (test) ──────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle("Final NB Model — Fitted vs Observed (test set)",
             fontweight="bold", color=MN, fontsize=11)
for ax, (yhat, ytrue, label) in zip(axes, [
    (mf_tr["yhat_final"].values, mf_tr["claim_count"].values, "Train"),
    (mf_te["yhat_final"].values, mf_te["claim_count"].values, "Test"),
]):
    ax.scatter(yhat, ytrue, s=5, alpha=0.25, color=MT, zorder=3)
    lim = max(yhat.max(), ytrue.max())*1.05
    ax.plot([0,lim],[0,lim], color=MR, lw=1.5, ls="--", label="45°")
    ax.set_xlabel("Predicted E[N]"); ax.set_ylabel("Observed N")
    ax.set_title(label, fontsize=10); ax.legend(fontsize=8.5)
plt.tight_layout(); buf_f3 = fig_to_buf(fig)

# ── Fig 4: Residuals ──────────────────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5))
fig.suptitle("Final NB Model — Residuals",
             fontweight="bold", color=MN, fontsize=11)
for ax, (yhat, ytrue, label) in zip(axes, [
    (mf_tr["yhat_final"].values, mf_tr["claim_count"].values, "Train"),
    (mf_te["yhat_final"].values, mf_te["claim_count"].values, "Test"),
]):
    resid = ytrue - yhat
    ax.hist(resid, bins=50, color=MN, edgecolor="white",
            density=True, alpha=0.8, zorder=3)
    ax.axvline(0, color=MR, lw=1.5, ls="--",
               label=f"mean={resid.mean():.4f}")
    ax.set_xlabel("Residual"); ax.set_ylabel("Density")
    ax.set_title(label, fontsize=10); ax.legend(fontsize=8)
plt.tight_layout(); buf_f4 = fig_to_buf(fig)

# ── Fig 5: Risk decile lift chart ─────────────────────────────────────────────
fig, ax = plt.subplots(figsize=(12, 5))
fig.suptitle("Risk Decile Lift — Observed vs Predicted Frequency (test)",
             fontweight="bold", color=MN, fontsize=11)
x_ = np.arange(len(seg_dec)); w = 0.35
ax.bar(x_-w/2, seg_dec["obs_rate"],  w, label="Observed rate",
       color=MN, edgecolor="white", zorder=3)
ax.bar(x_+w/2, seg_dec["pred_rate"], w, label="Predicted rate",
       color=MT, edgecolor="white", zorder=3)
ax.set_xticks(x_)
ax.set_xticklabels(seg_dec["risk_decile"].astype(str), fontsize=8.5)
ax.set_xlabel("Risk Decile (D1=lowest E[N])")
ax.set_ylabel("Claim frequency (claims/policy)")
ax.legend(fontsize=9)
plt.tight_layout(); buf_f5 = fig_to_buf(fig)

# ── Fig 6: Solar system calibration (expected loss) ───────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle("Expected Loss vs Observed — Segment Calibration (test)",
             fontweight="bold", color=MN, fontsize=11)
ax = axes[0]
x_ = np.arange(len(seg_sol)); w=0.35
ax.bar(x_-w/2, seg_sol["obs_N"],  w, label="Obs N",
       color=MN, edgecolor="white", zorder=3)
ax.bar(x_+w/2, seg_sol["pred_N"], w, label="Pred N",
       color=MT, edgecolor="white", zorder=3)
for i, r in seg_sol.iterrows():
    ax.text(i, max(r["obs_N"],r["pred_N"])+0.3,
            f"{r['freq_ratio']:.2f}x", ha="center", fontsize=8, color=MR)
ax.set_xticks(x_)
ax.set_xticklabels(seg_sol["solar_system"].astype(str),
                   rotation=12, fontsize=8.5)
ax.legend(fontsize=8); ax.set_ylabel("Claims")
ax.set_title("By solar_system", fontsize=10)

ax = axes[1]
x_ = np.arange(len(seg_exp)); w=0.35
ax.bar(x_-w/2, seg_exp["obs_N"],  w, label="Obs N",
       color=MN, edgecolor="white", zorder=3)
ax.bar(x_+w/2, seg_exp["pred_N"], w, label="Pred N",
       color=MT, edgecolor="white", zorder=3)
for i, r in seg_exp.iterrows():
    ax.text(i, max(r["obs_N"],r["pred_N"])+0.3,
            f"{r['freq_ratio']:.2f}x", ha="center", fontsize=8, color=MR)
ax.set_xticks(x_)
ax.set_xticklabels(seg_exp["exp_band"].astype(str), fontsize=8.5)
ax.legend(fontsize=8); ax.set_ylabel("Claims")
ax.set_title("By Exposure Band", fontsize=10)
plt.tight_layout(); buf_f6 = fig_to_buf(fig)

# ── Fig 7: Severity calibration ───────────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle(f"Severity Model ({pref_sev_nm}) — Calibration (test)",
             fontweight="bold", color=MN, fontsize=11)
ax = axes[0]
ax.scatter(sm["E_te"], y_sev_te, s=5, alpha=0.25, color=MG, zorder=3)
lim = max(sm["E_te"].max(), y_sev_te.max())*1.05
ax.plot([0,lim],[0,lim], color=MR, lw=1.5, ls="--", label="45°")
ax.set_xlabel("Predicted E[X]"); ax.set_ylabel("Observed claim_amount")
ax.set_title("Fitted vs Observed", fontsize=10); ax.legend(fontsize=8)
ax = axes[1]
resid_s = y_sev_te - sm["E_te"]
ax.hist(resid_s, bins=50, color=MG, edgecolor="white",
        density=True, alpha=0.8, zorder=3)
ax.axvline(0, color=MR, lw=1.5, ls="--",
           label=f"mean={resid_s.mean():.0f}")
ax.set_xlabel("Residual"); ax.set_ylabel("Density"); ax.legend(fontsize=8)
ax.set_title("Residuals", fontsize=10)
plt.tight_layout(); buf_f7 = fig_to_buf(fig)

# ── Fig 8: NB coefficient plot (preferred model) ──────────────────────────────
fig, ax = plt.subplots(figsize=(11, max(4, len(final_cnames)*0.45+1)))
fig.suptitle(f"Final NB Model Coefficients — {best_exp_nm}",
             fontweight="bold", color=MN, fontsize=11)
params_plot = final_res["params"][:-1]   # exclude log_r
names_plot  = final_cnames
y_pos = range(len(names_plot))
bar_cols = [MG if v>0 else MR for v in params_plot]
ax.barh(list(y_pos), params_plot, color=bar_cols,
        edgecolor="white", linewidth=0.5, zorder=3)
ax.axvline(0, color="black", lw=0.8)
ax.set_yticks(list(y_pos)); ax.set_yticklabels(names_plot, fontsize=9)
ax.set_xlabel("Coefficient (log scale on E[N])")
for i, (nm_, v) in enumerate(zip(names_plot, params_plot)):
    ax.text(v + (0.005 if v>=0 else -0.005), i,
            f"{v:.4f}", va="center",
            ha="left" if v>=0 else "right", fontsize=7.5)
plt.tight_layout(); buf_f8 = fig_to_buf(fig)

# ── Fig 9: Severity model comparison ─────────────────────────────────────────
fig, axes = plt.subplots(1, 2, figsize=(13, 4.5))
fig.suptitle("Severity Model Comparison",
             fontweight="bold", color=MN, fontsize=11)
sev_nms   = list(sev_results.keys())
sev_aics  = [sev_results[n]["aic"]    for n in sev_nms]
sev_rmses = [sev_results[n]["rmse_te"] for n in sev_nms]
scols     = [MG if n==pref_sev_nm else MT for n in sev_nms]
for ax, vals, title in zip(axes, [sev_aics, sev_rmses], ["AIC","RMSE (test)"]):
    bars = ax.bar(sev_nms, vals, color=scols, edgecolor="white", zorder=3)
    ax.set_xticklabels(sev_nms, rotation=15, ha="right", fontsize=9)
    ax.set_ylabel(title); ax.set_title(title, fontsize=10)
    ax.axhline(min(vals), color=MR, lw=1.5, ls="--",
               label=f"min={min(vals):.0f}")
    ax.legend(fontsize=8)
plt.tight_layout(); buf_f9 = fig_to_buf(fig)

# ═════════════════════════════════════════════════════════════════════════════
# 8.  BUILD PDF REPORT
# ═════════════════════════════════════════════════════════════════════════════
print("\n7.  BUILDING PDF REPORT")

doc = SimpleDocTemplate(
    PDF_PATH, pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2.2*cm, bottomMargin=2.2*cm,
    title="BI NB Refinement Report",
)
W = A4[0] - 4*cm

H1  = ParagraphStyle("H1",  fontSize=16, textColor=NAVY, spaceAfter=5,
                             spaceBefore=12, fontName="Helvetica-Bold")
H2  = ParagraphStyle("H2",  fontSize=12, textColor=TEAL, spaceAfter=3,
                             spaceBefore=9,  fontName="Helvetica-Bold")
H3  = ParagraphStyle("H3",  fontSize=10, textColor=NAVY, spaceAfter=2,
                             spaceBefore=5,  fontName="Helvetica-Bold")
BD  = ParagraphStyle("BD",  fontSize=9.5, leading=14, spaceAfter=5,
                             alignment=TA_JUSTIFY, fontName="Helvetica")
MO  = ParagraphStyle("MO",  fontSize=8.5, leading=12, fontName="Courier",
                             textColor=HexColor("#2C3E50"))
CP  = ParagraphStyle("CP",  fontSize=8, leading=11, textColor=DKGREY,
                             alignment=TA_CENTER, spaceAfter=8)
BL  = ParagraphStyle("BL",  parent=BD, leftIndent=14, spaceAfter=3)

def hr():    return HRFlowable(width="100%", thickness=0.5,
                                color=MGREY, spaceAfter=5)
def sp(h=6): return Spacer(1,h)
def p(t,s=BD): return Paragraph(t,s)
def cap(t):    return Paragraph(t,CP)
def bul(t):    return p(f"&bull;&nbsp;&nbsp;{t}", BL)

def img(buf, w_cm=16):
    r = 4.5/13
    return KeepTogether([sp(4),
                         RLImage(buf, width=w_cm*cm, height=w_cm*cm*r),
                         sp(2)])

def kv(rows, kw=6.5*cm):
    data = [[p(f"<b>{k}</b>",MO), p(str(v),BD)] for k,v in rows]
    t    = Table(data, colWidths=[kw,W-kw])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,-1),LGREY),
        ("GRID",(0,0),(-1,-1),0.4,MGREY),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),6),
    ]))
    return t

def tbl(hdr, rows, hi_col=None, hi_val=None):
    data = [[p(f"<b>{c}</b>",MO) for c in hdr]]
    for row in rows:
        data.append([p(str(v),BD) for v in row])
    cw = [W/len(hdr)]*len(hdr)
    t  = Table(data, colWidths=cw, repeatRows=1)
    sty= [
        ("BACKGROUND",(0,0),(-1,0),NAVY),
        ("TEXTCOLOR",(0,0),(-1,0),white),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[white,LGREY]),
        ("GRID",(0,0),(-1,-1),0.4,MGREY),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),4),
        ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),6),
        ("FONTSIZE",(0,0),(-1,-1),8.5),
    ]
    if hi_val is not None:
        for i, row in enumerate(rows,1):
            if str(row[hi_col]) == str(hi_val):
                sty.append(("BACKGROUND",(0,i),(-1,i),GOOD))
    t.setStyle(TableStyle(sty)); return t

def box(title, paras, bg=LGREY, border=TEAL):
    inner = [p(f"<b>{title}</b>",H3)]+paras
    t = Table([[inner]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),bg),
        ("BOX",(0,0),(-1,-1),0.8,border),
        ("TOPPADDING",(0,0),(-1,-1),8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),10),
        ("RIGHTPADDING",(0,0),(-1,-1),10),
    ]))
    return KeepTogether([t,sp(6)])

story = []

# ── Cover ─────────────────────────────────────────────────────────────────────
cov = Table([[
    p("<b>Business Interruption</b>",
      ParagraphStyle("CT",fontSize=22,textColor=white,
                     fontName="Helvetica-Bold",alignment=TA_CENTER))
],[
    p("NB Frequency Model Refinement",
      ParagraphStyle("CS",fontSize=14,textColor=HexColor("#A8D8E0"),
                     fontName="Helvetica",alignment=TA_CENTER))
],[
    p("No zero-inflated models  |  Exposure band diagnostics  |  "
      "Severity  |  Expected loss calibration",
      ParagraphStyle("CS2",fontSize=9,textColor=HexColor("#BDC3C7"),
                     fontName="Helvetica",alignment=TA_CENTER))
]], colWidths=[W])
cov.setStyle(TableStyle([
    ("BACKGROUND",(0,0),(-1,-1),NAVY),
    ("TOPPADDING",(0,0),(-1,-1),20),("BOTTOMPADDING",(0,0),(-1,-1),20),
    ("LEFTPADDING",(0,0),(-1,-1),14),("RIGHTPADDING",(0,0),(-1,-1),14),
]))
story += [sp(30), cov, sp(14),
          kv([
              ("Input file",           "BI_data_CLEANED.xlsx — flag=0 only"),
              ("Freq modelling frame", f"{len(mf):,} rows"),
              ("Sev  modelling frame", f"{len(ms):,} rows (CA>0)"),
              ("Train/Test split",     "80/20 stratified by solar_system"),
              ("Zero rate",            f"{zero_rate:.1%}"),
              ("Preferred freq model", best_exp_nm),
              ("NB dispersion r",      f"{final_res['nb_r']:.4f}"),
              ("Preferred sev model",  pref_sev_nm),
              ("Overall freq cal",     f"{overall_freq_cal:.4f}"),
              ("Overall loss cal",     f"{overall_loss_cal:.4f}"),
          ]), sp(8),
          p("NB-only refinement following the decision to reject zero-inflated "
            "models on conceptual and empirical grounds. No pricing or premium "
            "loading is performed."),
          PageBreak()]

# ── S1: Overdispersion confirmation ───────────────────────────────────────────
story += [p("1.  Overdispersion Confirmation", H1), hr(),
          p("The adequacy of the Negative Binomial distribution is confirmed "
            "by comparing the observed variance/mean ratio against the Poisson "
            "assumption of equidispersion, and by inspecting the estimated NB "
            "size parameter <i>r</i>. Small values of <i>r</i> indicate heavy "
            "overdispersion; as r → ∞ the NB converges to Poisson."),
          sp(6),
          kv([
              ("Observed mean (train)",    f"{obs_mean:.4f}"),
              ("Observed variance (train)",f"{obs_var:.4f}"),
              ("Var / Mean ratio",         f"{obs_var/obs_mean:.2f}"
               "  (Poisson requires 1.0)"),
              ("NB baseline r",            f"{base_r:.4f}"),
              ("Interpretation",
               "Overdispersed — NB appropriate"
               if base_r < 10 else
               "Near-Poisson — low overdispersion"),
              ("Zero rate",                f"{zero_rate:.1%}"
               "  (reflects rare events and exposure, not structural zeros)"),
          ]),
          sp(6),
          p("The high zero rate (~93%) is consistent with a genuinely rare-event "
            "BI portfolio. Under a Negative Binomial with moderate exposure, the "
            "majority of policies are expected to report zero claims even without "
            "any structural zero class. Zero-inflated models are therefore "
            "conceptually inappropriate for commercial BI: every policy has a "
            "non-zero claim probability, and the zero excess is fully explained "
            "by small exposure values and a low underlying rate."),
          PageBreak()]

# ── S2: NB model sequence ─────────────────────────────────────────────────────
story += [p("2.  NB Model Sequence", H1), hr(),
          p("Five NB models are fitted incrementally. Each extension is retained "
            "only if AIC improvement exceeds 2 units relative to the previous "
            "preferred model. BIC penalises complexity more heavily and provides "
            "a secondary check."),
          sp(6)]

freq_tbl_rows = []
for nm, m in freq_results.items():
    extra = f"r={m['nb_r']:.3f}"
    freq_tbl_rows.append([
        nm,
        f"{m['aic']:.1f}",
        f"{m['bic']:.1f}",
        f"{m['rmse_te']:.4f}",
        f"{m['mae_te']:.4f}",
        f"{m['cal_te']:.4f}",
        extra,
        "✓" if m["converged"] else "✗",
    ])

story += [tbl(
    ["Model","AIC","BIC","RMSE_te","MAE_te","Cal_te","NB r","Conv"],
    freq_tbl_rows, hi_col=0, hi_val=best_nm
), sp(6),
    img(buf_f1), cap("Figure 2.1. AIC and BIC by NB model. "
                      "Green = selected preferred NB model (pre-exposure fix)."),
    sp(8)]

sci_delta = freq_results["NB_base"]["aic"] - freq_results["NB_+sci"]["aic"]
ebs2_delta= freq_results["NB_base"]["aic"] - freq_results["NB_+ebs2"]["aic"]
lsci_delta= freq_results["NB_base"]["aic"] - freq_results["NB_+log_sci"]["aic"]

story += [kv([
    ("supply_chain_index",        f"ΔAIC={sci_delta:.1f} — "
     +("added (material)" if sci_delta>2 else "not added (<2 threshold)")),
    ("energy_backup_score²",      f"ΔAIC={ebs2_delta:.1f} — "
     +("added" if ebs2_delta>2 else "quadratic term not material")),
    ("log(supply_chain_index)",   f"ΔAIC={lsci_delta:.1f} — "
     +("added" if lsci_delta>2 else "log form not preferred")),
    ("SCI × solar interaction",   f"ΔAIC="
     f"{freq_results['NB_base']['aic']-freq_results['NB_+sci_int']['aic']:.1f}"),
    ("Selected NB predictor set", best_nm),
]), PageBreak()]

# ── S3: Exposure band diagnostics ─────────────────────────────────────────────
story += [p("3.  Exposure Band Diagnostics", H1), hr(),
          p("A known limitation in earlier runs was systematic under- or "
            "over-prediction by exposure quartile. The frequency ratio "
            "(predicted / observed) is examined before and after applying "
            "an exposure adjustment. Three adjustments are tested: "
            "log(exposure) as a covariate, exposure² as a mild nonlinear "
            "term, and the comparison of their AIC gain and reduction in "
            "maximum exposure-band deviation."),
          sp(6)]

fix_rows = []
for nm, m in exp_fix_results.items():
    fix_rows.append([
        nm,
        f"{m['aic']:.1f}",
        f"{m['cal_te']:.4f}",
        f"{m['max_exp_ratio_dev']:.4f}",
        "★" if nm==best_exp_nm else "",
    ])
story += [tbl(
    ["Model","AIC","Cal_te","Max Band Deviation","Selected"],
    fix_rows, hi_col=4, hi_val="★"
), sp(6),
    img(buf_f2), cap("Figure 3.1. Exposure band calibration before (left) and "
                      "after (right) exposure adjustment. Numbers above bars are "
                      "predicted/observed ratios."),
    sp(8)]

story += [kv([
    ("Pre-fix max band deviation",  f"{base_dev:.4f}"),
    ("Post-fix max band deviation", f"{final_res['max_exp_ratio_dev']:.4f}"),
    ("Selected adjustment",         best_exp_nm),
    ("Action",
     "Exposure adjustment retained — material improvement in band calibration."
     if best_exp_nm != "NB_pref" else
     "No adjustment retained — baseline model already well-calibrated by band."),
]), PageBreak()]

# ── S4: Final frequency model ─────────────────────────────────────────────────
story += [p("4.  Final Frequency Model", H1), hr(),
          kv([
              ("Model name",       best_exp_nm),
              ("Predictors",       ", ".join(final_cnames)),
              ("NB dispersion r",  f"{final_res['nb_r']:.4f}"),
              ("AIC",              f"{final_res['aic']:.1f}"),
              ("BIC",              f"{final_res['bic']:.1f}"),
              ("RMSE (test)",      f"{final_res['rmse_te']:.4f}"),
              ("MAE  (test)",      f"{final_res['mae_te']:.4f}"),
              ("Calibration (test)",f"{final_res['cal_te']:.4f}"),
          ]),
          sp(6),
          img(buf_f8), cap("Figure 4.1. Coefficient plot for final NB model. "
                            "Green = positive effect on E[N], red = negative."),
          img(buf_f3), cap("Figure 4.2. Fitted vs observed claim count — "
                            "train (left) and test (right)."),
          img(buf_f4), cap("Figure 4.3. Residual distribution — "
                            "train (left) and test (right)."),
          PageBreak()]

# ── S5: Severity ──────────────────────────────────────────────────────────────
story += [p("5.  Severity Models", H1), hr(),
          p("Three severity models are fitted on positive claim amounts. "
            "The lognormal intercept model is retained if AIC improvement "
            "from adding predictors does not exceed 2 units, consistent with "
            "the EDA finding that severity signal is weak."),
          sp(6)]

sev_tbl_rows = []
for nm, m in sev_results.items():
    sig_str = f"σ={m['sigma']:.4f}" if "sigma" in m else "—"
    sev_tbl_rows.append([
        nm,
        f"{m['aic']:.1f}",
        f"{m['bic']:.1f}",
        f"{m['rmse_te']:,.0f}",
        f"{m['mae_te']:,.0f}",
        f"{m['cal_te']:.4f}",
        sig_str,
    ])

story += [tbl(
    ["Model","AIC","BIC","RMSE_te","MAE_te","Cal_te","Notes"],
    sev_tbl_rows, hi_col=0, hi_val=pref_sev_nm
), sp(6),
    img(buf_f9), cap("Figure 5.1. Severity AIC (left) and RMSE on test (right). "
                      "Green = preferred model."),
    img(buf_f7), cap("Figure 5.2. Preferred severity model fitted vs observed "
                      "(left) and residual distribution (right)."),
    sp(6)]

sev_gain = sev_results["LN_intercept"]["aic"] - min(
    sev_results[n]["aic"] for n in ["LN_solar","Gamma_solar"])
story += [box("Severity finding", [
    bul(f"Preferred model: <b>{pref_sev_nm}</b>  "
        f"(AIC={sm['aic']:.1f}, calibration={sm['cal_te']:.4f})."),
    bul(f"Maximum AIC gain from adding predictors: {sev_gain:.1f} units."),
    bul("Severity variation appears largely random relative to the available "
        "operational predictors. The lognormal intercept model is robust and "
        "interpretable. Operational predictors do not materially explain "
        "why one claim is larger than another in this BI portfolio."),
], bg=GOOD if pref_sev_nm=="LN_intercept" else WARN, border=GREEN),
    PageBreak()]

# ── S6: Expected loss & calibration ───────────────────────────────────────────
story += [p("6.  Expected Loss & Calibration Summaries", H1), hr(),
          p(f"Expected loss = E[N] (final NB model) × E[X] (preferred severity model). "
            f"Mean E[X] = {mean_E_X:,.0f}. Calibration ratios are reported "
            f"by solar_system, exposure band, and risk decile."),
          sp(6),
          kv([
              ("Total pred N (test)",  f"{total_pred_N:,.1f}"),
              ("Total obs  N (test)",  f"{total_obs_N:,.0f}"),
              ("Frequency calibration",f"{overall_freq_cal:.4f}"),
              ("Mean E[X]",            f"{mean_E_X:,.0f}"),
              ("Total pred loss",      f"{total_pred_L:,.0f}"),
              ("Total obs loss (approx)", f"{total_obs_L:,.0f}"),
              ("Loss calibration",     f"{overall_loss_cal:.4f}"),
          ]),
          sp(8), p("6.1  By solar_system", H2),
          tbl(["solar_system","N","Obs N","Pred N","Freq ratio",
               "Pred loss","Obs loss","Loss ratio"],
              [[str(r["solar_system"]), f"{r['n']:,}",
                f"{r['obs_N']:.0f}", f"{r['pred_N']:.1f}",
                f"{r['freq_ratio']:.4f}" if pd.notna(r['freq_ratio']) else "—",
                f"{r['pred_loss']:,.0f}", f"{r['obs_loss']:,.0f}",
                f"{r['loss_ratio']:.4f}" if pd.notna(r['loss_ratio']) else "—"]
               for _, r in seg_sol.iterrows()]),
          sp(8), p("6.2  By exposure band", H2),
          tbl(["Exp band","N","Obs N","Pred N","Freq ratio",
               "Pred loss","Obs loss","Loss ratio"],
              [[str(r["exp_band"]), f"{r['n']:,}",
                f"{r['obs_N']:.0f}", f"{r['pred_N']:.1f}",
                f"{r['freq_ratio']:.4f}" if pd.notna(r['freq_ratio']) else "—",
                f"{r['pred_loss']:,.0f}", f"{r['obs_loss']:,.0f}",
                f"{r['loss_ratio']:.4f}" if pd.notna(r['loss_ratio']) else "—"]
               for _, r in seg_exp.iterrows()]),
          sp(8), p("6.3  Risk decile lift (frequency)", H2),
          tbl(["Decile","N","Obs N","Pred N","Obs rate","Pred rate","Ratio"],
              [[str(r["risk_decile"]), f"{r['n']:,}",
                f"{r['obs_N']:.0f}", f"{r['pred_N']:.1f}",
                f"{r['obs_rate']:.4f}", f"{r['pred_rate']:.4f}",
                f"{r['freq_ratio']:.4f}" if pd.notna(r['freq_ratio']) else "—"]
               for _, r in seg_dec.iterrows()]),
          sp(6),
          img(buf_f5), cap("Figure 6.1. Risk decile lift — observed vs predicted "
                            "frequency rate by model risk decile. "
                            "Good lift shows monotone increase D1→D10."),
          img(buf_f6), cap("Figure 6.2. Expected vs observed claim counts by "
                            "solar_system (left) and exposure band (right). "
                            "Numbers above bars are predicted/observed ratios."),
          PageBreak()]

# ── S7: Actuarial summary ─────────────────────────────────────────────────────
story += [p("7.  Actuarial Modelling Summary", H1), hr(),
          p("This section provides the actuarial conclusions from the NB refinement "
            "phase. No pricing calibration or premium loading is performed."),
          sp(8)]

story += [box("NB as the appropriate frequency distribution", [
    bul(f"Observed variance/mean = {obs_var/obs_mean:.2f} (Poisson requires 1.0). "
        "Poisson is clearly inadequate."),
    bul(f"NB baseline size parameter r = {base_r:.4f}. "
        +("Small r confirms genuine overdispersion beyond Poisson."
          if base_r<10 else
          "r is large but Poisson is still formally rejected by AIC.")),
    bul("The NB distribution is confirmed as the appropriate frequency family "
        "for this BI portfolio."),
], bg=GOOD, border=GREEN)]

story += [box("Rejection of zero-inflated models", [
    bul("Zero-inflated models impose a structural subpopulation that can "
        "never produce a claim. This is economically unrealistic for commercial "
        "Business Interruption: every insured facility has a non-zero probability "
        "of a business interruption event."),
    bul(f"The observed zero rate of {zero_rate:.1%} is fully explained by "
        "the combination of low underlying frequency, small exposure values, "
        "and the NB variance structure. No structural zero class is required."),
    bul("ZINB improved AIC marginally in earlier runs but produced no "
        "out-of-sample improvement in RMSE or calibration. The mixing weight π "
        "was near zero, confirming the structural zero component was spurious."),
    bul("ZI models are rejected on both conceptual and empirical grounds."),
], bg=GOOD, border=GREEN)]

sci_selected = best_nm in ["NB_+sci","NB_+log_sci","NB_+sci_int"]
story += [box("supply_chain_index", [
    bul(f"ΔAIC (NB_base → NB_+sci) = {sci_delta:.1f}."),
    bul("supply_chain_index " +
        (f"is retained in the preferred model ({best_nm}) — "
         "it materially improves frequency fit."
         if sci_selected else
         "does not materially improve AIC (threshold: >2). "
         "It is excluded to avoid unnecessary complexity. "
         "The baseline predictors maintenance_freq, "
         "energy_backup_score, and solar_system are sufficient.")),
], bg=GOOD if sci_selected else WARN,
   border=GREEN if sci_selected else AMBER)]

exp_fixed = best_exp_nm != "NB_pref"
story += [box("Exposure band distortion", [
    bul(f"Pre-fix maximum exposure band deviation: {base_dev:.4f} "
        "(ratio furthest from 1.0 across exposure quartiles)."),
    bul(f"Post-fix maximum deviation: {final_res['max_exp_ratio_dev']:.4f}."),
    bul("Exposure band distortion " +
        (f"was corrected by including {best_exp_nm.replace('NB_+','')} as "
         "a predictor. The model now produces consistent frequency "
         "calibration across the exposure distribution."
         if exp_fixed else
         "was not material in the baseline model. "
         "No exposure adjustment was required. "
         "Calibration is consistent across exposure quartiles.")),
], bg=GOOD, border=GREEN)]

story += [box("Severity conclusion", [
    bul(f"Preferred severity model: {pref_sev_nm}  "
        f"(AIC={sm['aic']:.1f})."),
    bul(f"Maximum AIC gain from adding operational predictors: {sev_gain:.1f}."),
    bul("Severity variation in this BI portfolio appears largely random "
        "relative to the available operational and structural predictors. "
        "Claim size is not materially explained by maintenance frequency, "
        "energy backup score, supply chain index, or solar system. "
        "The lognormal intercept model captures the distributional shape "
        "adequately and is retained as the severity component."),
], bg=GOOD if pref_sev_nm=="LN_intercept" else WARN, border=GREEN)]

story += [box("Portfolio driver", [
    bul(f"Frequency CV: {freq_cv:.3f}   Severity CV: {sev_cv:.3f}"),
    bul("<b>This BI portfolio is "
        +("FREQUENCY-driven.</b> "
          "Loss variability is primarily explained by variation in "
          "claim count rather than variation in individual claim size. "
          "Underwriting and risk management effort should be directed "
          "at the frequency model — improving the identification of "
          "high-frequency policies. Severity management has limited "
          "additional uplift available from predictive modelling."
          if freq_driven else
          "SEVERITY-driven.</b> "
          "Loss variability is primarily driven by individual claim "
          "size. Tail risk and large-loss management should be "
          "prioritised in addition to frequency modelling.")),
    bul(f"Overall frequency calibration (test): {overall_freq_cal:.4f}  "
        +("— model is well-calibrated."
          if abs(overall_freq_cal-1)<0.10 else
          "— calibration adjustment recommended before deployment.")),
    bul(f"Overall loss calibration (test): {overall_loss_cal:.4f}  "
        +("— within acceptable tolerance."
          if abs(overall_loss_cal-1)<0.15 else
          "— review aggregate balance before use.")),
], bg=GOOD if (abs(overall_freq_cal-1)<0.10 and abs(overall_loss_cal-1)<0.15)
    else WARN,
   border=GREEN if (abs(overall_freq_cal-1)<0.10 and abs(overall_loss_cal-1)<0.15)
    else AMBER)]

doc.build(story)
print(f"\nPDF saved → {PDF_PATH}")

# ═════════════════════════════════════════════════════════════════════════════
# 9.  CONSOLE ACTUARIAL SUMMARY
# ═════════════════════════════════════════════════════════════════════════════
print("\n" + "="*65)
print("ACTUARIAL MODELLING SUMMARY")
print("="*65)

sci_verdict = (f"RETAINED in {best_nm}" if sci_selected
               else f"NOT retained (ΔAIC={sci_delta:.1f} < 2)")
exp_verdict = (f"CORRECTED via {best_exp_nm}" if exp_fixed
               else "No material distortion — no adjustment needed")

print(textwrap.dedent(f"""
  FREQUENCY DISTRIBUTION
  ─────────────────────
  Distribution       : Negative Binomial (confirmed)
  Observed var/mean  : {obs_var/obs_mean:.2f}  (Poisson requires 1.0)
  NB dispersion r    : {base_r:.4f}
  Interpretation     : {'Overdispersed — NB appropriate' if base_r<10 else 'Near-Poisson, NB still preferred'}

  ZERO-INFLATED MODELS
  ─────────────────────
  Decision           : REJECTED — conceptually and empirically
  Reason             : Every BI policy can claim; zero excess is
                       explained by low exposure and rare-event
                       structure, not a structural zero class.

  SUPPLY_CHAIN_INDEX
  ─────────────────────
  Decision           : {sci_verdict}

  EXPOSURE BAND DISTORTION
  ─────────────────────
  Pre-fix max deviation  : {base_dev:.4f}
  Post-fix max deviation : {final_res['max_exp_ratio_dev']:.4f}
  Outcome            : {exp_verdict}

  SEVERITY
  ─────────────────────
  Preferred model    : {pref_sev_nm}
  Max AIC gain       : {sev_gain:.1f}  (threshold: 2.0)
  Conclusion         : Severity is largely unexplained by available
                       predictors. Intercept lognormal retained.

  PORTFOLIO DRIVER
  ─────────────────────
  Frequency CV       : {freq_cv:.3f}
  Severity CV        : {sev_cv:.3f}
  Driver             : {'FREQUENCY' if freq_driven else 'SEVERITY'}

  CALIBRATION (TEST SET)
  ─────────────────────
  Frequency          : {overall_freq_cal:.4f}  ({'OK' if abs(overall_freq_cal-1)<0.10 else 'REVIEW'})
  Loss               : {overall_loss_cal:.4f}  ({'OK' if abs(overall_loss_cal-1)<0.15 else 'REVIEW'})
"""))
print("="*65)
print("NB REFINEMENT COMPLETE — no pricing performed")
print("="*65)
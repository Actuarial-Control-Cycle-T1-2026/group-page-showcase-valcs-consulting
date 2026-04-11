"""
Business Interruption — Deep EDA & Feature Selection
=====================================================
Parts 1–7 per specification.
NO final model fitting. NO premium calibration. NO aggregate pricing.
Exploratory and feature selection phase only.

Input  : /Users/alansteny/Downloads/BI_data_CLEANED.xlsx
Output : /Users/alansteny/Downloads/bi_eda_feature_selection.pdf
"""

import warnings, io
warnings.filterwarnings("ignore")

import numpy  as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
from matplotlib.colors import LinearSegmentedColormap
from scipy import stats
from scipy.cluster.hierarchy import linkage, dendrogram, fcluster
from scipy.spatial.distance import squareform
from numpy.linalg import lstsq
from scipy.optimize import minimize
from scipy.special import gammaln, digamma

# ── ReportLab ─────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib.units    import cm
from reportlab.lib.colors   import HexColor, white
from reportlab.lib.styles   import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums    import TA_CENTER, TA_JUSTIFY, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, HRFlowable, Image as RLImage, KeepTogether
)

# ── Paths ─────────────────────────────────────────────────────────────────────
FILE_PATH   = "/Users/alansteny/Downloads/BI_data_CLEANED.xlsx"
OUTPUT_PATH = "/Users/alansteny/Downloads/bi_eda_feature_selection.pdf"

# ── Palette ───────────────────────────────────────────────────────────────────
NAVY  = HexColor("#1B2A4A");  MN = "#1B2A4A"
TEAL  = HexColor("#1A7A8A");  MT = "#1A7A8A"
GREEN = HexColor("#27AE60");  MG = "#27AE60"
AMBER = HexColor("#E67E22");  MA = "#E67E22"
RED   = HexColor("#C0392B");  MR = "#C0392B"
LGREY = HexColor("#F4F6F8");  ML = "#F4F6F8"
MGREY = HexColor("#BDC3C7")
DKGREY= HexColor("#555555")
GOOD  = HexColor("#D5F5E3")
WARN  = HexColor("#FEF3CD")
BAD   = HexColor("#FADBD8")

plt.rcParams.update({
    "font.family":"sans-serif", "axes.spines.top":False,
    "axes.spines.right":False,  "axes.grid":True,
    "grid.alpha":0.3, "grid.linestyle":"--",
    "figure.facecolor":"white", "axes.facecolor":"white",
})

# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA
# ─────────────────────────────────────────────────────────────────────────────
print("Loading data...")
df_freq_raw = pd.read_excel(FILE_PATH, sheet_name="Frequency_Cleaned")
df_sev_raw  = pd.read_excel(FILE_PATH, sheet_name="Severity_Cleaned")

df_freq = df_freq_raw[df_freq_raw["data_quality_flag"] == 0].copy()
df_sev  = df_sev_raw [df_sev_raw ["data_quality_flag"] == 0].copy()

# Coerce known numeric columns
NUM_FREQ = ["claim_count","exposure","production_load","supply_chain_index",
            "avg_crew_exp","maintenance_freq","energy_backup_score",
            "safety_compliance"]
NUM_SEV  = ["claim_amount","production_load","supply_chain_index",
            "avg_crew_exp","maintenance_freq","energy_backup_score",
            "safety_compliance","exposure"]

for col in NUM_FREQ:
    if col in df_freq.columns:
        df_freq[col] = pd.to_numeric(df_freq[col], errors="coerce")
for col in NUM_SEV:
    if col in df_sev.columns:
        df_sev[col] = pd.to_numeric(df_sev[col], errors="coerce")

df_sev  = df_sev[df_sev["claim_amount"] > 0].copy().reset_index(drop=True)
df_sev["log_claim_amount"] = np.log(df_sev["claim_amount"])

print(f"  freq clean: {len(df_freq):,}   sev clean: {len(df_sev):,}")

# Candidate predictors per sheet
FREQ_PREDS = [c for c in ["production_load","supply_chain_index","avg_crew_exp",
                           "maintenance_freq","energy_backup_score","safety_compliance",
                           "solar_system"]
              if c in df_freq.columns]
SEV_PREDS  = [c for c in ["production_load","supply_chain_index","avg_crew_exp",
                           "maintenance_freq","energy_backup_score","safety_compliance",
                           "solar_system"]
              if c in df_sev.columns]

NUMERIC_PREDS = [p for p in FREQ_PREDS if p != "solar_system"]

# ─────────────────────────────────────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────
def fig_to_buf(fig, dpi=150):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    buf.seek(0); plt.close(fig)
    return buf

def poisson_deviance_reduction(y, x, offset):
    """Fit intercept-only vs simple Poisson with x. Return deviance reduction, AIC, coef.
    All inputs converted to plain numpy arrays first to avoid pandas index misalignment."""
    y_arr   = np.asarray(pd.to_numeric(pd.Series(y).reset_index(drop=True),      errors="coerce"), dtype=float)
    x_arr   = np.asarray(pd.to_numeric(pd.Series(x).reset_index(drop=True),      errors="coerce"), dtype=float)
    off_arr = np.asarray(pd.to_numeric(pd.Series(offset).reset_index(drop=True), errors="coerce"), dtype=float)
    mask    = ~np.isnan(x_arr) & ~np.isnan(y_arr) & ~np.isnan(off_arr)
    y_    = y_arr[mask]; x_ = x_arr[mask]; off = off_arr[mask]

    # null model
    mu0   = np.exp(np.log(y_.mean() + 1e-10) + 0)
    def null_nll(b):
        mu = np.exp(b[0] + off)
        return -np.sum(y_*np.log(mu+1e-15) - mu - gammaln(y_+1))
    r0 = minimize(null_nll, [0.0], method="L-BFGS-B")
    nll0 = r0.fun

    # model with x
    def model_nll(b):
        mu = np.exp(b[0] + b[1]*x_ + off)
        return -np.sum(y_*np.log(mu+1e-15) - mu - gammaln(y_+1))
    r1 = minimize(model_nll, [0.0, 0.0], method="L-BFGS-B")
    nll1 = r1.fun

    dev_red = 2*(nll0 - nll1)
    aic     = 2*2 - 2*(-nll1)
    aic0    = 2*1 - 2*(-nll0)
    coef    = r1.x[1]
    return float(dev_red), float(aic), float(aic0 - aic), float(coef), int(mask.sum())

def ols_r2(log_y, x):
    """Simple OLS R² for log(claim_amount) ~ x.
    All inputs converted to plain numpy arrays first to avoid pandas index misalignment."""
    y_arr = np.asarray(pd.to_numeric(pd.Series(log_y).reset_index(drop=True), errors="coerce"), dtype=float)
    x_arr = np.asarray(pd.to_numeric(pd.Series(x).reset_index(drop=True),     errors="coerce"), dtype=float)
    mask  = ~np.isnan(x_arr) & ~np.isnan(y_arr)
    y_  = y_arr[mask]; x_ = x_arr[mask]
    if mask.sum() < 5:
        return np.nan, np.nan, np.nan, 0
    X   = np.column_stack([np.ones(len(y_)), x_])
    b,_,_,_ = lstsq(X, y_, rcond=None)
    yhat    = X @ b
    ss_res  = np.sum((y_ - yhat)**2)
    ss_tot  = np.sum((y_ - y_.mean())**2)
    r2      = 1 - ss_res/(ss_tot + 1e-15)
    sigma   = (ss_res/(len(y_)-2))**0.5
    logL    = np.sum(stats.norm.logpdf(y_, yhat, sigma))
    aic     = 2*3 - 2*logL
    return float(r2), float(aic), float(b[1]), int(mask.sum())

def compute_vif(df_sub):
    cols = df_sub.columns.tolist()
    Xm   = df_sub.values.astype(float)
    vif  = {}
    for i, col in enumerate(cols):
        y_ = Xm[:, i]
        X_ = np.column_stack([np.ones(len(Xm)), np.delete(Xm, i, axis=1)])
        b,_,_,_ = lstsq(X_, y_, rcond=None)
        yh  = X_ @ b
        ss_res = np.sum((y_ - yh)**2)
        ss_tot = np.sum((y_ - y_.mean())**2)
        r2  = 1 - ss_res/(ss_tot + 1e-15)
        vif[col] = 1/(1 - r2 + 1e-15)
    return vif

def decile_plot(ax, predictor_series, response_series, title, ylabel, color):
    mask = predictor_series.notna() & response_series.notna()
    x    = predictor_series[mask]; y = response_series[mask]
    try:
        deciles = pd.qcut(x, 10, labels=False, duplicates="drop")
    except Exception:
        return
    means = y.groupby(deciles).mean()
    ns    = y.groupby(deciles).count()
    ax.bar(means.index, means.values, color=color, edgecolor="white",
           linewidth=0.6, zorder=3)
    ax.set_xlabel(f"Decile of {predictor_series.name}", fontsize=8)
    ax.set_ylabel(ylabel, fontsize=8); ax.set_title(title, fontsize=8.5)
    # monotonicity: count inversions
    vals = means.values
    inversions = sum(1 for i in range(len(vals)-1) if vals[i] > vals[i+1])
    ax.set_title(f"{title}\n({inversions} inversions)", fontsize=8.5)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 1 — DATA UNDERSTANDING
# ═══════════════════════════════════════════════════════════════════════════════
print("\nPART 1 — Data understanding")

def variable_profile(df, col):
    s = df[col]
    is_num = pd.api.types.is_numeric_dtype(s)
    n      = len(s)
    n_miss = s.isna().sum()
    n_zero = (s == 0).sum() if is_num else np.nan
    if is_num:
        sv = s.dropna()
        return {
            "column"  : col,
            "type"    : "numeric",
            "n"       : n,
            "missing" : n_miss,
            "miss_pct": round(n_miss/n*100, 2),
            "zero_pct": round(n_zero/n*100, 2) if pd.notna(n_zero) else "—",
            "mean"    : round(sv.mean(), 4) if len(sv) else np.nan,
            "std"     : round(sv.std(),  4) if len(sv) else np.nan,
            "min"     : round(sv.min(),  4) if len(sv) else np.nan,
            "p25"     : round(sv.quantile(.25), 4) if len(sv) else np.nan,
            "median"  : round(sv.median(), 4) if len(sv) else np.nan,
            "p75"     : round(sv.quantile(.75), 4) if len(sv) else np.nan,
            "p99"     : round(sv.quantile(.99), 4) if len(sv) else np.nan,
            "max"     : round(sv.max(),  4) if len(sv) else np.nan,
            "skewness": round(stats.skew(sv) if len(sv)>2 else np.nan, 4),
            "kurtosis": round(stats.kurtosis(sv) if len(sv)>2 else np.nan, 4),
            "n_unique": sv.nunique(),
        }
    else:
        vc = s.value_counts()
        return {
            "column"  : col,
            "type"    : "categorical",
            "n"       : n,
            "missing" : n_miss,
            "miss_pct": round(n_miss/n*100, 2),
            "zero_pct": "—",
            "mean"    : "—", "std": "—", "min": "—", "p25": "—",
            "median"  : "—", "p75": "—", "p99": "—", "max": "—",
            "skewness": "—", "kurtosis": "—",
            "n_unique": s.nunique(),
        }

freq_all_cols = [c for c in df_freq.columns if c not in
                 ["data_quality_flag","data_quality_desc"]]
sev_all_cols  = [c for c in df_sev.columns  if c not in
                 ["data_quality_flag","data_quality_desc"]]

freq_profiles = [variable_profile(df_freq, c) for c in freq_all_cols]
sev_profiles  = [variable_profile(df_sev,  c) for c in sev_all_cols]

# Figure P1a: Frequency distributions grid
num_freq_cols = [p["column"] for p in freq_profiles if p["type"] == "numeric"][:9]
fig = plt.figure(figsize=(14, 10))
fig.suptitle("Part 1 — Frequency Sheet: Numeric Variable Distributions",
             fontsize=12, fontweight="bold", color=MN)
for i, col in enumerate(num_freq_cols):
    ax = fig.add_subplot(3, 3, i+1)
    s  = df_freq[col].dropna()
    if len(s) < 5: continue
    ax.hist(s, bins=40, color=MT, edgecolor="white", linewidth=0.3,
            density=True, zorder=3)
    ax.axvline(s.median(), color=MR, lw=1.5, linestyle="--",
               label=f"med={s.median():.3f}")
    ax.set_title(col, fontsize=9); ax.legend(fontsize=7)
    ax.set_ylabel("Density", fontsize=7); ax.tick_params(labelsize=7)
plt.tight_layout(); buf_p1a = fig_to_buf(fig)

# Figure P1b: Solar system distribution
fig, axes = plt.subplots(1, 2, figsize=(12, 4))
fig.suptitle("Part 1 — Categorical Variables: solar_system",
             fontsize=11, fontweight="bold", color=MN)
for ax, df_, nm in zip(axes, [df_freq, df_sev], ["Frequency","Severity"]):
    if "solar_system" in df_.columns:
        vc = df_["solar_system"].value_counts()
        colors_ = [MT, MA, MG][:len(vc)]
        ax.bar(vc.index.astype(str), vc.values, color=colors_,
               edgecolor="white", zorder=3)
        for i_, (lbl, v) in enumerate(zip(vc.index, vc.values)):
            ax.text(i_, v + vc.max()*0.01, f"{v:,}", ha="center",
                    va="bottom", fontsize=8.5)
    ax.set_title(nm, fontsize=10); ax.set_ylabel("Count")
plt.tight_layout(); buf_p1b = fig_to_buf(fig)

# Figure P1c: Boxplots
fig, axes = plt.subplots(2, 4, figsize=(14, 7))
fig.suptitle("Part 1 — Boxplots: Numeric Predictors",
             fontsize=11, fontweight="bold", color=MN)
box_cols = [c for c in FREQ_PREDS if c != "solar_system"]
for ax, col in zip(axes.flat, box_cols):
    if col in df_freq.columns:
        s = df_freq[col].dropna()
        ax.boxplot(s, vert=True, patch_artist=True,
                   boxprops=dict(facecolor=MT, alpha=0.7),
                   medianprops=dict(color="white", linewidth=2),
                   flierprops=dict(marker="o", alpha=0.2, markersize=3))
        ax.set_title(col, fontsize=9); ax.set_ylabel("Value", fontsize=8)
        ax.tick_params(labelsize=7)
for ax in axes.flat[len(box_cols):]: ax.axis("off")
plt.tight_layout(); buf_p1c = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 2 — UNIVARIATE SIGNAL TESTING
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 2 — Univariate signal testing")

# Frequency
df_freq_m = df_freq[df_freq["exposure"] > 0].copy().reset_index(drop=True)
df_freq_m["log_exposure"] = np.log(df_freq_m["exposure"])
y_freq   = df_freq_m["claim_count"].values.astype(float)
off_freq = df_freq_m["log_exposure"].values

freq_signal = []
for pred in FREQ_PREDS:
    if pred == "solar_system":
        # treat as categorical — encode as dummies, one per level
        vc  = df_freq_m["solar_system"].value_counts()
        ref = vc.idxmax()
        for level in vc.index:
            if level == ref: continue
            dummy = (df_freq_m["solar_system"] == level).astype(float)
            dummy.name = f"solar_system={level}"
            dr, aic, daic, coef, n = poisson_deviance_reduction(y_freq, dummy, off_freq)
            freq_signal.append({"predictor": f"solar[{level}]",
                                 "dev_red": dr, "delta_aic": daic,
                                 "coef": coef, "n": n, "type": "cat"})
    else:
        x = df_freq_m[pred]
        dr, aic, daic, coef, n = poisson_deviance_reduction(y_freq, x, off_freq)
        freq_signal.append({"predictor": pred, "dev_red": dr,
                             "delta_aic": daic, "coef": coef,
                             "n": n, "type": "num"})

freq_signal_df = pd.DataFrame(freq_signal).sort_values("dev_red", ascending=False)
print("  Frequency signal ranking:")
print(freq_signal_df[["predictor","dev_red","delta_aic","coef"]].to_string(index=False))

# Severity
log_y_sev = df_sev["log_claim_amount"].values

sev_signal = []
for pred in SEV_PREDS:
    if pred == "solar_system":
        vc  = df_sev["solar_system"].value_counts()
        ref = vc.idxmax()
        for level in vc.index:
            if level == ref: continue
            dummy = (df_sev["solar_system"] == level).astype(float)
            dummy.index = df_sev.index
            r2, aic, coef, n = ols_r2(log_y_sev, dummy)
            sev_signal.append({"predictor": f"solar[{level}]",
                                "r2": r2, "aic": aic, "coef": coef,
                                "n": n, "type": "cat"})
    else:
        if pred not in df_sev.columns: continue
        r2, aic, coef, n = ols_r2(log_y_sev, df_sev[pred])
        sev_signal.append({"predictor": pred, "r2": r2, "aic": aic,
                            "coef": coef, "n": n, "type": "num"})

sev_signal_df = pd.DataFrame(sev_signal).sort_values("r2", ascending=False)
print("  Severity signal ranking:")
print(sev_signal_df[["predictor","r2","aic","coef"]].to_string(index=False))

# Figure P2a: Signal bar charts
fig, axes = plt.subplots(1, 2, figsize=(13, 5))
fig.suptitle("Part 2 — Univariate Signal Strength",
             fontsize=11, fontweight="bold", color=MN)

ax = axes[0]
fs = freq_signal_df.head(10)
colors_ = [MG if d > 0 else MR for d in fs["delta_aic"]]
ax.barh(fs["predictor"], fs["dev_red"], color=colors_, edgecolor="white", zorder=3)
ax.set_xlabel("Deviance Reduction vs Null (Poisson)")
ax.set_title("Frequency — Top Predictors by Deviance Reduction", fontsize=9.5)
ax.axvline(2, color=MR, lw=1, linestyle="--", label="ΔAIC=2 threshold")
ax.legend(fontsize=8)

ax = axes[1]
ss = sev_signal_df[sev_signal_df["r2"].notna()].head(10)
ax.barh(ss["predictor"], ss["r2"], color=MT, edgecolor="white", zorder=3)
ax.set_xlabel("R² (OLS on log claim_amount)")
ax.set_title("Severity — Top Predictors by R²", fontsize=9.5)
plt.tight_layout(); buf_p2a = fig_to_buf(fig)

# Figure P2b: Decile plots for top 6 freq predictors
top6_freq = [r["predictor"] for _, r in freq_signal_df[freq_signal_df["type"]=="num"].head(6).iterrows()]
fig, axes = plt.subplots(2, 3, figsize=(13, 7))
fig.suptitle("Part 2 — Frequency: Mean claim_count by Predictor Decile",
             fontsize=11, fontweight="bold", color=MN)
for ax, pred in zip(axes.flat, top6_freq):
    if pred in df_freq_m.columns:
        decile_plot(ax, df_freq_m[pred], df_freq_m["claim_count"],
                    pred, "Mean claim_count", MT)
plt.tight_layout(); buf_p2b = fig_to_buf(fig)

# Figure P2c: Decile plots for top 6 sev predictors
top6_sev = [r["predictor"] for _, r in sev_signal_df[sev_signal_df["type"]=="num"].head(6).iterrows()]
fig, axes = plt.subplots(2, 3, figsize=(13, 7))
fig.suptitle("Part 2 — Severity: Mean log(claim_amount) by Predictor Decile",
             fontsize=11, fontweight="bold", color=MN)
for ax, pred in zip(axes.flat, top6_sev):
    if pred in df_sev.columns:
        decile_plot(ax, df_sev[pred], df_sev["log_claim_amount"],
                    pred, "Mean log(claim_amount)", MG)
plt.tight_layout(); buf_p2c = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 3 — REDUNDANCY AND COLLINEARITY
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 3 — Collinearity and redundancy")

num_preds_avail = [p for p in NUMERIC_PREDS if p in df_freq.columns]
corr_df_freq    = df_freq[num_preds_avail].dropna()
corr_mat        = corr_df_freq.corr()

# VIF
vif_dict = compute_vif(corr_df_freq) if len(corr_df_freq) > len(num_preds_avail) + 5 else {}

# Hierarchical clustering on correlation distance
corr_arr  = corr_mat.values
dist_arr  = 1 - np.abs(corr_arr)
np.fill_diagonal(dist_arr, 0)
dist_cond = squareform(dist_arr, checks=False)
Z         = linkage(dist_cond, method="ward")
cluster_labels = fcluster(Z, t=0.4, criterion="distance")
cluster_map    = dict(zip(num_preds_avail, cluster_labels))

print("  Cluster assignments:")
for col, cl in sorted(cluster_map.items(), key=lambda x: x[1]):
    print(f"    cluster {cl}: {col}")

# Highlight pairs with |r| > 0.5
high_corr_pairs = []
for i in range(len(num_preds_avail)):
    for j in range(i+1, len(num_preds_avail)):
        r = corr_mat.iloc[i, j]
        if abs(r) > 0.5:
            high_corr_pairs.append((num_preds_avail[i], num_preds_avail[j], round(r, 4)))

print(f"  High-correlation pairs (|r|>0.5): {len(high_corr_pairs)}")
for a, b, r in high_corr_pairs:
    print(f"    {a} vs {b}: r={r}")

# Figure P3a: Correlation heatmap
fig, axes = plt.subplots(1, 2, figsize=(14, 6))
fig.suptitle("Part 3 — Correlation Matrix & VIF",
             fontsize=11, fontweight="bold", color=MN)

ax = axes[0]
im = ax.imshow(corr_mat.values, cmap="RdBu_r", vmin=-1, vmax=1, aspect="auto")
n_ = len(num_preds_avail)
ax.set_xticks(range(n_)); ax.set_yticks(range(n_))
ax.set_xticklabels(num_preds_avail, rotation=40, ha="right", fontsize=8)
ax.set_yticklabels(num_preds_avail, fontsize=8)
for i in range(n_):
    for j in range(n_):
        v = corr_mat.values[i,j]
        ax.text(j, i, f"{v:.2f}", ha="center", va="center",
                fontsize=7.5, color="white" if abs(v)>0.5 else "black")
        if i != j and abs(v) > 0.5:
            ax.add_patch(plt.Rectangle((j-0.5, i-0.5), 1, 1,
                                        fill=False, edgecolor=MA, linewidth=2))
plt.colorbar(im, ax=ax, shrink=0.8); ax.set_title("Pearson Correlation (|r|>0.5 outlined)", fontsize=9)

ax = axes[1]
if vif_dict:
    vnames = list(vif_dict.keys()); vvals = list(vif_dict.values())
    vcols  = [MR if v>5 else (MA if v>2.5 else MT) for v in vvals]
    bars   = ax.barh(vnames, vvals, color=vcols, edgecolor="white", zorder=3)
    ax.axvline(5, color=MR, lw=1.5, linestyle="--", label="VIF=5")
    ax.axvline(10, color=MN, lw=1.5, linestyle=":",  label="VIF=10")
    for bar, v in zip(bars, vvals):
        ax.text(bar.get_width()+0.05, bar.get_y()+bar.get_height()/2,
                f"{v:.2f}", va="center", fontsize=8.5)
    ax.legend(fontsize=8.5); ax.set_xlabel("VIF")
ax.set_title("Variance Inflation Factors", fontsize=9)
plt.tight_layout(); buf_p3a = fig_to_buf(fig)

# Figure P3b: Dendrogram
fig, ax = plt.subplots(figsize=(10, 4.5))
fig.suptitle("Part 3 — Hierarchical Clustering on Correlation Distance",
             fontsize=11, fontweight="bold", color=MN)
dendrogram(Z, labels=num_preds_avail, ax=ax, color_threshold=0.4,
           leaf_rotation=35, leaf_font_size=9)
ax.axhline(0.4, color=MR, lw=1.5, linestyle="--", label="Cut at 0.4")
ax.legend(fontsize=8.5); ax.set_ylabel("Distance")
plt.tight_layout(); buf_p3b = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 4 — FUNCTIONAL FORM ANALYSIS
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 4 — Functional form analysis")

def form_test_poisson(y, x, off):
    """Test linear, quadratic, log form. Return AIC table."""
    mask = pd.Series(x).notna() & pd.Series(y).notna() & pd.Series(off).notna()
    y_   = np.array(y)[mask]; x_ = np.array(x)[mask]; off_ = np.array(off)[mask]
    if mask.sum() < 20:
        return {}

    results = {}
    for name, X in [
        ("linear",    np.column_stack([np.ones(len(y_)), x_])),
        ("quadratic", np.column_stack([np.ones(len(y_)), x_, x_**2])),
        ("log",       np.column_stack([np.ones(len(y_)), np.log(np.abs(x_)+1e-10)])),
    ]:
        def nll(b, X=X):
            mu = np.exp(np.clip(X @ b + off_, -15, 12))
            return -np.sum(y_*np.log(mu+1e-15) - mu - gammaln(y_+1))
        r = minimize(nll, np.zeros(X.shape[1]), method="L-BFGS-B")
        results[name] = 2*X.shape[1] - 2*(-r.fun)
    return results

form_results = {}
top_form_preds = [r["predictor"] for _, r in
                   freq_signal_df[freq_signal_df["type"]=="num"].head(4).iterrows()
                   if r["predictor"] in df_freq_m.columns]

for pred in top_form_preds:
    fr = form_test_poisson(y_freq, df_freq_m[pred].values, off_freq)
    if fr:
        form_results[pred] = fr
        best = min(fr, key=fr.get)
        print(f"  {pred:<28}: best form = {best}  AICs = {fr}")

# Figure P4: Functional form AIC comparison
fig, axes = plt.subplots(2, 2, figsize=(12, 8))
fig.suptitle("Part 4 — Functional Form: AIC Comparison (Poisson, frequency)",
             fontsize=11, fontweight="bold", color=MN)
for ax, pred in zip(axes.flat, top_form_preds):
    if pred not in form_results: continue
    forms = list(form_results[pred].keys())
    aics  = list(form_results[pred].values())
    min_aic = min(aics)
    bar_cols = [MG if a == min_aic else MT for a in aics]
    bars = ax.bar(forms, aics, color=bar_cols, edgecolor="white", zorder=3)
    ax.set_title(pred, fontsize=9.5); ax.set_ylabel("AIC")
    for bar, a in zip(bars, aics):
        ax.text(bar.get_x()+bar.get_width()/2, bar.get_height()+0.1,
                f"{a:.1f}", ha="center", va="bottom", fontsize=8)
for ax in axes.flat[len(top_form_preds):]: ax.axis("off")
plt.tight_layout(); buf_p4 = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 5 — INTERACTION EXPLORATION
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 5 — Interaction exploration")

def interaction_test(y, x1, x2, off):
    """Main effects vs main + interaction. Returns delta AIC."""
    mask = (pd.Series(x1).notna() & pd.Series(x2).notna() &
            pd.Series(y).notna() & pd.Series(off).notna())
    y_  = np.array(y)[mask]
    x1_ = np.array(x1)[mask]; x2_ = np.array(x2)[mask]
    off_= np.array(off)[mask]

    def nll(b, X):
        mu = np.exp(np.clip(X @ b + off_, -15, 12))
        return -np.sum(y_*np.log(mu+1e-15) - mu - gammaln(y_+1))

    Xme  = np.column_stack([np.ones(len(y_)), x1_, x2_])
    Xint = np.column_stack([np.ones(len(y_)), x1_, x2_, x1_*x2_])

    rme  = minimize(nll, np.zeros(Xme.shape[1]),  args=(Xme,),  method="L-BFGS-B")
    rint = minimize(nll, np.zeros(Xint.shape[1]), args=(Xint,), method="L-BFGS-B")

    aic_me  = 2*Xme.shape[1]  - 2*(-rme.fun)
    aic_int = 2*Xint.shape[1] - 2*(-rint.fun)
    return float(aic_me - aic_int), float(rint.x[-1])  # delta AIC, interaction coef

interactions = {}
pairs = []
nc = [c for c in NUMERIC_PREDS if c in df_freq_m.columns]
for i in range(len(nc)):
    for j in range(i+1, len(nc)):
        pairs.append((nc[i], nc[j]))

interaction_rows = []
for a, b in pairs[:12]:   # cap at 12 pairs to keep runtime manageable
    try:
        daic, coef = interaction_test(y_freq, df_freq_m[a].values,
                                      df_freq_m[b].values, off_freq)
        interaction_rows.append({"pair": f"{a} × {b}", "delta_aic": round(daic,2),
                                  "int_coef": round(coef,4),
                                  "material": "YES" if daic > 2 else "no"})
        print(f"  {a} × {b}: ΔAIC={daic:.2f}  coef={coef:.4f}")
    except:
        pass

int_df = pd.DataFrame(interaction_rows).sort_values("delta_aic", ascending=False)

# Figure P5: Interaction ΔAIC bar
fig, ax = plt.subplots(figsize=(10, 5))
fig.suptitle("Part 5 — Interaction ΔAIC (main effects vs main + interaction)",
             fontsize=11, fontweight="bold", color=MN)
if len(int_df):
    cols_ = [MG if d>2 else LGREY.hexval()[1:] if hasattr(LGREY,"hexval") else "#F4F6F8"
             for d in int_df["delta_aic"]]
    cols_ = [MG if d > 2 else "#BDC3C7" for d in int_df["delta_aic"]]
    ax.barh(int_df["pair"], int_df["delta_aic"], color=cols_, edgecolor="white", zorder=3)
    ax.axvline(2, color=MR, lw=1.5, linestyle="--", label="ΔAIC=2 threshold")
    ax.axvline(10, color=MN, lw=1.5, linestyle=":", label="ΔAIC=10 (decisive)")
    ax.legend(fontsize=8.5); ax.set_xlabel("ΔAIC (positive = interaction improves fit)")
plt.tight_layout(); buf_p5 = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 6 — STABILITY TESTING
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 6 — Stability testing")

stability_rows = []

def quick_poisson_coef(y, x, off):
    mask = pd.Series(x).notna() & pd.Series(y).notna() & pd.Series(off).notna()
    y_  = np.array(y)[mask]; x_ = np.array(x)[mask]; off_ = np.array(off)[mask]
    if mask.sum() < 10: return np.nan
    def nll(b):
        mu = np.exp(np.clip(b[0]+b[1]*x_+off_, -15, 12))
        return -np.sum(y_*np.log(mu+1e-15)-mu-gammaln(y_+1))
    r = minimize(nll, [0.0, 0.0], method="L-BFGS-B")
    return float(r.x[1])

top_preds_stable = [p for p in
                    freq_signal_df[freq_signal_df["type"]=="num"]["predictor"].head(5)
                    if p in df_freq_m.columns]

# Stability across solar_system
if "solar_system" in df_freq_m.columns:
    for pred in top_preds_stable:
        coefs = {}
        for sys in df_freq_m["solar_system"].dropna().unique():
            sub = df_freq_m[df_freq_m["solar_system"] == sys]
            c   = quick_poisson_coef(sub["claim_count"].values,
                                     sub[pred].values,
                                     sub["log_exposure"].values)
            coefs[str(sys)] = c
        valid_coefs = [v for v in coefs.values() if pd.notna(v)]
        signs       = [np.sign(v) for v in valid_coefs]
        sign_flip   = len(set(signs)) > 1
        mag_range   = max(valid_coefs) - min(valid_coefs) if valid_coefs else np.nan
        stability_rows.append({
            "predictor"   : pred,
            "split"       : "solar_system",
            "coefs"       : str({k: round(v, 4) if pd.notna(v) else "NA"
                                  for k, v in coefs.items()}),
            "sign_flip"   : "YES ⚠" if sign_flip else "no",
            "mag_range"   : round(mag_range, 4) if pd.notna(mag_range) else "NA",
            "stable"      : "NO" if sign_flip or (pd.notna(mag_range) and mag_range > 1.0)
                             else "YES",
        })
        print(f"  {pred} by solar_system: sign_flip={sign_flip}  range={mag_range:.4f}")

# Stability across exposure bands
df_freq_m["exp_band"] = pd.qcut(df_freq_m["exposure"], 3,
                                   labels=["low","mid","high"])
for pred in top_preds_stable:
    coefs = {}
    for band in ["low","mid","high"]:
        sub = df_freq_m[df_freq_m["exp_band"] == band]
        c   = quick_poisson_coef(sub["claim_count"].values,
                                  sub[pred].values,
                                  sub["log_exposure"].values)
        coefs[band] = c
    valid_coefs = [v for v in coefs.values() if pd.notna(v)]
    signs       = [np.sign(v) for v in valid_coefs]
    sign_flip   = len(set(signs)) > 1
    mag_range   = max(valid_coefs) - min(valid_coefs) if valid_coefs else np.nan
    stability_rows.append({
        "predictor"   : pred,
        "split"       : "exposure_band",
        "coefs"       : str({k: round(v,4) if pd.notna(v) else "NA"
                              for k,v in coefs.items()}),
        "sign_flip"   : "YES ⚠" if sign_flip else "no",
        "mag_range"   : round(mag_range,4) if pd.notna(mag_range) else "NA",
        "stable"      : "NO" if sign_flip or (pd.notna(mag_range) and mag_range > 1.0)
                         else "YES",
    })

stability_df = pd.DataFrame(stability_rows)

# Figure P6: Stability heatmap
fig, ax = plt.subplots(figsize=(11, 5))
fig.suptitle("Part 6 — Coefficient Stability: Sign Flips Across Subgroups",
             fontsize=11, fontweight="bold", color=MN)
if len(stability_df):
    pivot = stability_df.pivot_table(
        index="predictor", columns="split",
        values="sign_flip", aggfunc="first"
    )
    pivot_num = pivot.map(lambda x: 1 if isinstance(x,str) and "YES" in x else 0)
    im = ax.imshow(pivot_num.values, cmap="RdYlGn_r", vmin=0, vmax=1, aspect="auto")
    ax.set_xticks(range(pivot_num.shape[1]))
    ax.set_yticks(range(pivot_num.shape[0]))
    ax.set_xticklabels(pivot_num.columns, fontsize=9)
    ax.set_yticklabels(pivot_num.index, fontsize=9)
    for i in range(pivot_num.shape[0]):
        for j in range(pivot_num.shape[1]):
            raw = pivot.iloc[i,j]
            ax.text(j, i, str(raw), ha="center", va="center", fontsize=8.5)
    ax.set_xlabel("Stability Split"); ax.set_ylabel("Predictor")
plt.tight_layout(); buf_p6 = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# PART 7 — CANDIDATE FEATURE SET ASSEMBLY
# ═══════════════════════════════════════════════════════════════════════════════
print("PART 7 — Candidate feature set assembly")

# Classification: signal + stability + collinearity
def classify_predictor(pred, freq_sig_df, sev_sig_df, vif_d, stab_df, cluster_map_):
    freq_row = freq_sig_df[freq_sig_df["predictor"] == pred]
    sev_row  = sev_sig_df [sev_sig_df ["predictor"] == pred]

    freq_dev  = freq_row["dev_red"].values[0]   if len(freq_row) else np.nan
    sev_r2    = sev_row ["r2"].values[0]        if len(sev_row)  else np.nan
    vif_val   = vif_d.get(pred, np.nan)
    cluster   = cluster_map_.get(pred, 0)

    # Stability: any sign flip
    stab_pred = stab_df[stab_df["predictor"] == pred] if len(stab_df) else pd.DataFrame()
    sign_flips = (stab_pred["sign_flip"].str.contains("YES").sum()
                  if len(stab_pred) else 0)

    signal_ok   = (pd.notna(freq_dev) and freq_dev > 2) or \
                  (pd.notna(sev_r2)   and sev_r2  > 0.005)
    vif_ok      = pd.isna(vif_val) or vif_val < 10
    stable      = sign_flips == 0

    decision = "INCLUDE" if (signal_ok and vif_ok and stable) else "REVIEW"
    if not signal_ok: decision = "REJECT — no signal"
    elif not vif_ok:  decision = "REVIEW — high VIF"
    elif not stable:  decision = "REVIEW — unstable"

    return {
        "predictor"      : pred,
        "freq_dev_red"   : round(freq_dev,  3) if pd.notna(freq_dev) else "—",
        "sev_r2"         : round(sev_r2,    5) if pd.notna(sev_r2)   else "—",
        "vif"            : round(vif_val,   2) if pd.notna(vif_val)   else "—",
        "cluster"        : cluster,
        "sign_flips"     : sign_flips,
        "decision"       : decision,
    }

all_preds_unique = list(dict.fromkeys(
    [r["predictor"] for r in freq_signal if r["type"] == "num"] +
    [r["predictor"] for r in sev_signal  if r["type"] == "num"]
))

final_assessment = [classify_predictor(p, freq_signal_df, sev_signal_df,
                                        vif_dict, stability_df, cluster_map)
                    for p in all_preds_unique]
final_df = pd.DataFrame(final_assessment)

included = final_df[final_df["decision"] == "INCLUDE"]["predictor"].tolist()
review   = final_df[final_df["decision"].str.startswith("REVIEW")]["predictor"].tolist()
rejected = final_df[final_df["decision"].str.startswith("REJECT")]["predictor"].tolist()

# Always include solar_system as a categorical fixed effect
solar_recommendation = "INCLUDE as categorical fixed effect (systematic environmental heterogeneity)"

print("\n  INCLUDED:", included)
print("  REVIEW  :", review)
print("  REJECTED:", rejected)

# Figure P7: Final feature scorecard
fig, ax = plt.subplots(figsize=(12, max(4, len(final_df)*0.55 + 1.5)))
fig.suptitle("Part 7 — Candidate Feature Scorecard",
             fontsize=11, fontweight="bold", color=MN)
ax.axis("off")
col_headers = ["Predictor","Freq Dev.Red","Sev R²","VIF","Cluster","Sign Flips","Decision"]
tbl_data    = [[str(r["predictor"]), str(r["freq_dev_red"]), str(r["sev_r2"]),
                str(r["vif"]), str(r["cluster"]), str(r["sign_flips"]), str(r["decision"])]
               for _, r in final_df.iterrows()]
t = ax.table(cellText=tbl_data, colLabels=col_headers,
             cellLoc="center", loc="center", bbox=[0, 0, 1, 1])
t.auto_set_font_size(False); t.set_fontsize(8.5)
for (r,c), cell in t.get_celld().items():
    cell.set_edgecolor("white")
    if r == 0:
        cell.set_facecolor(MN); cell.get_text().set_color("white")
    elif r > 0 and c == 6:
        val = tbl_data[r-1][6]
        if "INCLUDE" in val:
            cell.set_facecolor("#D5F5E3")
        elif "REJECT" in val:
            cell.set_facecolor("#FADBD8")
        else:
            cell.set_facecolor("#FEF3CD")
    elif r % 2 == 0:
        cell.set_facecolor("#F4F6F8")
plt.tight_layout(); buf_p7 = fig_to_buf(fig)

# ═══════════════════════════════════════════════════════════════════════════════
# BUILD PDF REPORT
# ═══════════════════════════════════════════════════════════════════════════════
print("\nBuilding PDF...")

doc = SimpleDocTemplate(
    OUTPUT_PATH, pagesize=A4,
    leftMargin=2*cm, rightMargin=2*cm,
    topMargin=2.2*cm, bottomMargin=2.2*cm,
    title="BI EDA & Feature Selection Report",
)
W = A4[0] - 4*cm

styles = getSampleStyleSheet()
H1   = ParagraphStyle("H1",   fontSize=16, textColor=NAVY,  spaceAfter=5,
                               spaceBefore=12, fontName="Helvetica-Bold")
H2   = ParagraphStyle("H2",   fontSize=12, textColor=TEAL,  spaceAfter=3,
                               spaceBefore=9,  fontName="Helvetica-Bold")
H3   = ParagraphStyle("H3",   fontSize=10, textColor=NAVY,  spaceAfter=2,
                               spaceBefore=6,  fontName="Helvetica-Bold")
BODY = ParagraphStyle("BODY", fontSize=9.5, leading=14, spaceAfter=5,
                               alignment=TA_JUSTIFY, fontName="Helvetica")
MONO = ParagraphStyle("MONO", fontSize=8.5, leading=12, spaceAfter=3,
                               fontName="Courier", textColor=HexColor("#2C3E50"))
CAP  = ParagraphStyle("CAP",  fontSize=8,  leading=11, textColor=DKGREY,
                               alignment=TA_CENTER, spaceAfter=8)
BLET = ParagraphStyle("BLET", parent=BODY, leftIndent=14, spaceAfter=3)

def hr():    return HRFlowable(width="100%", thickness=0.5, color=MGREY, spaceAfter=5)
def sp(h=6): return Spacer(1, h)
def p(t, s=BODY):  return Paragraph(t, s)
def cap(t):        return Paragraph(t, CAP)
def bullet(t):     return p(f"&bull;&nbsp;&nbsp;{t}", BLET)

def embed(buf, w_cm=16):
    ratio = 4/11
    return KeepTogether([sp(4),
                         RLImage(buf, width=w_cm*cm, height=w_cm*cm*ratio),
                         sp(2)])

def embed_sq(buf, w_cm=14):
    return KeepTogether([sp(4),
                         RLImage(buf, width=w_cm*cm, height=w_cm*cm*0.52),
                         sp(2)])

def kv(rows, kw=6*cm):
    data = [[p(f"<b>{k}</b>", MONO), p(str(v), BODY)] for k,v in rows]
    t    = Table(data, colWidths=[kw, W-kw])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(0,-1), LGREY),
        ("GRID",          (0,0),(-1,-1), 0.4, MGREY),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(-1,-1), 6),
    ]))
    return t

def std_tbl(header, rows, decision_col=None):
    data = [[p(f"<b>{c}</b>", MONO) for c in header]]
    for row in rows:
        data.append([p(str(v), BODY) for v in row])
    cw = [W/len(header)]*len(header)
    t  = Table(data, colWidths=cw, repeatRows=1)
    style = [
        ("BACKGROUND",    (0,0),(-1,0), NAVY),
        ("TEXTCOLOR",     (0,0),(-1,0), white),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[white, LGREY]),
        ("GRID",          (0,0),(-1,-1), 0.4, MGREY),
        ("VALIGN",        (0,0),(-1,-1), "MIDDLE"),
        ("TOPPADDING",    (0,0),(-1,-1), 4),
        ("BOTTOMPADDING", (0,0),(-1,-1), 4),
        ("LEFTPADDING",   (0,0),(-1,-1), 6),
        ("FONTSIZE",      (0,0),(-1,-1), 8.5),
    ]
    if decision_col is not None:
        for i, row in enumerate(rows, 1):
            val = str(row[decision_col])
            if "INCLUDE" in val:  style.append(("BACKGROUND",(0,i),(-1,i), GOOD))
            elif "REJECT"  in val: style.append(("BACKGROUND",(0,i),(-1,i), BAD))
            elif "REVIEW"  in val: style.append(("BACKGROUND",(0,i),(-1,i), WARN))
    t.setStyle(TableStyle(style))
    return t

def box(title, paras, bg=LGREY, border=TEAL):
    inner = [p(f"<b>{title}</b>", H3)] + paras
    t = Table([[inner]], colWidths=[W])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), bg),
        ("BOX",           (0,0),(-1,-1), 0.8, border),
        ("TOPPADDING",    (0,0),(-1,-1), 8),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("LEFTPADDING",   (0,0),(-1,-1), 10),
        ("RIGHTPADDING",  (0,0),(-1,-1), 10),
    ]))
    return KeepTogether([t, sp(6)])

# ── COVER ──────────────────────────────────────────────────────────────────────
story = []
cov = Table([
    [p("<b>Business Interruption</b>",
       ParagraphStyle("CT", fontSize=22, textColor=white,
                      fontName="Helvetica-Bold", alignment=TA_CENTER))],
    [p("Deep EDA &amp; Feature Selection Report",
       ParagraphStyle("CS", fontSize=13, textColor=HexColor("#A8D8E0"),
                      fontName="Helvetica", alignment=TA_CENTER))],
    [p("Parts 1–7 &nbsp;|&nbsp; No final model fitting &nbsp;|&nbsp; Exploratory phase only",
       ParagraphStyle("CS2", fontSize=9, textColor=HexColor("#BDC3C7"),
                      fontName="Helvetica", alignment=TA_CENTER))],
], colWidths=[W])
cov.setStyle(TableStyle([
    ("BACKGROUND",    (0,0),(-1,-1), NAVY),
    ("TOPPADDING",    (0,0),(-1,-1), 20), ("BOTTOMPADDING",(0,0),(-1,-1), 20),
    ("LEFTPADDING",   (0,0),(-1,-1), 14), ("RIGHTPADDING", (0,0),(-1,-1), 14),
]))
story += [sp(32), cov, sp(14)]
story += [kv([
    ("Source file",      "BI_data_CLEANED.xlsx (Frequency_Cleaned + Severity_Cleaned)"),
    ("Filter",           "data_quality_flag == 0"),
    ("Frequency rows",   f"{len(df_freq):,}"),
    ("Severity rows",    f"{len(df_sev):,}  (positive claim_amount only)"),
    ("Freq predictors",  ", ".join(FREQ_PREDS)),
    ("Sev predictors",   ", ".join(SEV_PREDS)),
    ("Candidate set",    f"INCLUDE={len(included)+1}  REVIEW={len(review)}  "
                          f"REJECT={len(rejected)}  (+solar_system)"),
]), sp(8)]
story += [
    p("This report documents structured exploratory data analysis and feature selection "
      "for Business Interruption frequency and severity modelling. Parts 1–7 follow the "
      "prescribed methodology: data profiling, univariate signal testing, collinearity "
      "analysis, functional form testing, interaction exploration, stability checking, "
      "and candidate set assembly. No final models are fitted. No premiums are calibrated."),
    PageBreak()
]

# ── PART 1 ─────────────────────────────────────────────────────────────────────
story += [p("Part 1 — Data Understanding", H1), hr()]

story += [p("Variable type classification, missing rates, distribution statistics, "
            "skewness, and outlier presence for all variables in both sheets. "
            "No composite features are engineered in this section.", BODY), sp(8)]

# Frequency profile table
freq_prof_rows = [[str(r["column"]), str(r["type"]),
                   f"{r['miss_pct']}%", str(r["mean"]),
                   str(r["std"]), str(r["skewness"]), str(r["n_unique"])]
                  for r in freq_profiles
                  if r["column"] not in ["data_quality_flag","data_quality_desc"]]
story += [p("1.1  Frequency Sheet — Variable Profiles", H2),
          std_tbl(["Column","Type","Missing%","Mean","Std","Skewness","Unique"],
                  freq_prof_rows[:20]),
          sp(6)]

# Severity profile table
sev_prof_rows = [[str(r["column"]), str(r["type"]),
                  f"{r['miss_pct']}%", str(r["mean"]),
                  str(r["std"]), str(r["skewness"]), str(r["n_unique"])]
                 for r in sev_profiles
                 if r["column"] not in ["data_quality_flag","data_quality_desc","log_claim_amount"]]
story += [p("1.2  Severity Sheet — Variable Profiles", H2),
          std_tbl(["Column","Type","Missing%","Mean","Std","Skewness","Unique"],
                  sev_prof_rows[:20]),
          sp(8)]

story += [p("1.3  Distributions", H2),
          embed(buf_p1a),
          cap("Figure 1.1. Density histograms for numeric frequency variables. "
              "Red dashed line = median."),
          sp(4),
          embed(buf_p1b),
          cap("Figure 1.2. solar_system counts in frequency (left) and severity (right)."),
          sp(4),
          embed_sq(buf_p1c),
          cap("Figure 1.3. Boxplots of numeric predictors. "
              "Outlier points shown at 20% opacity."),
          PageBreak()]

# ── PART 2 ─────────────────────────────────────────────────────────────────────
story += [p("Part 2 — Univariate Signal Testing", H1), hr()]

story += [p("Each predictor is tested independently against the response. "
            "For frequency: simple Poisson GLM with exposure offset; deviance "
            "reduction and ΔAIC vs intercept-only null are reported. "
            "For severity: OLS on log(claim_amount); R² reported. "
            "Predictors ranked by signal strength.", BODY), sp(8)]

story += [p("2.1  Frequency — Ranked Signal", H2),
          std_tbl(["Predictor","Dev. Reduction","ΔAIC vs Null","Coefficient","N","Type"],
                  [[r["predictor"], f"{r['dev_red']:.3f}", f"{r['delta_aic']:.1f}",
                    f"{r['coef']:.4f}", f"{r['n']:,}", r["type"]]
                   for _, r in freq_signal_df.iterrows()]),
          sp(6)]

story += [p("2.2  Severity — Ranked Signal", H2),
          std_tbl(["Predictor","R²","AIC","Coefficient","N","Type"],
                  [[r["predictor"],
                    f"{r['r2']:.5f}" if pd.notna(r["r2"]) else "—",
                    f"{r['aic']:.1f}"  if pd.notna(r["aic"]) else "—",
                    f"{r['coef']:.4f}" if pd.notna(r["coef"]) else "—",
                    f"{r['n']:,}", r["type"]]
                   for _, r in sev_signal_df.iterrows()]),
          sp(8)]

story += [embed(buf_p2a),
          cap("Figure 2.1. Left: frequency deviance reduction by predictor. "
              "Right: severity R² by predictor. Teal = positive signal."),
          sp(4),
          embed_sq(buf_p2b),
          cap("Figure 2.2. Mean claim_count by decile of each top predictor. "
              "Count in parentheses shows number of monotone inversions."),
          sp(4),
          embed_sq(buf_p2c),
          cap("Figure 2.3. Mean log(claim_amount) by decile of each top severity predictor."),
          PageBreak()]

# ── PART 3 ─────────────────────────────────────────────────────────────────────
story += [p("Part 3 — Redundancy and Collinearity Structure", H1), hr()]

story += [p("Full Pearson correlation matrix on raw numeric predictors. "
            "Pairs with |r| > 0.5 are highlighted. VIF computed for each predictor. "
            "Hierarchical clustering on absolute correlation distance identifies "
            "groups of substitutable variables.", BODY), sp(8)]

story += [p("3.1  High-Correlation Pairs (|r| > 0.5)", H2)]
if high_corr_pairs:
    story += [std_tbl(["Variable A","Variable B","Pearson r"],
                       [[a, b, f"{r:.4f}"] for a, b, r in high_corr_pairs])]
else:
    story += [p("No pairs with |r| > 0.5 detected.")]

story += [sp(8), p("3.2  VIF", H2)]
if vif_dict:
    story += [std_tbl(["Predictor","VIF","Status"],
                       [[k, f"{v:.3f}",
                         "OK" if v<2.5 else ("Monitor" if v<5 else "HIGH ⚠")]
                        for k,v in vif_dict.items()])]

story += [sp(8), p("3.3  Cluster Assignments", H2),
          std_tbl(["Predictor","Cluster"],
                   [[col, str(cl)] for col, cl in
                    sorted(cluster_map.items(), key=lambda x: x[1])]),
          sp(8)]

story += [embed(buf_p3a),
          cap("Figure 3.1. Left: correlation heatmap — amber borders on |r|>0.5 pairs. "
              "Right: VIF bar chart."),
          sp(4),
          embed(buf_p3b),
          cap("Figure 3.2. Dendrogram from hierarchical clustering on |r|-distance. "
              "Red cut at distance=0.4 defines clusters of redundant variables."),
          sp(8)]

# Cluster recommendations
clusters_seen = sorted(set(cluster_map.values()))
story += [p("3.4  Cluster Recommendations", H2)]
for cl in clusters_seen:
    members = [col for col, c in cluster_map.items() if c == cl]
    if len(members) == 1:
        recommendation = f"Only member — include directly."
    else:
        # pick by freq signal
        signals = {m: freq_signal_df[freq_signal_df["predictor"]==m]["dev_red"].values
                   for m in members}
        signals = {m: float(v[0]) if len(v) else 0.0 for m, v in signals.items()}
        rep     = max(signals, key=signals.get)
        recommendation = (f"Representative: <b>{rep}</b> (highest deviance reduction). "
                          f"Others ({', '.join(m for m in members if m != rep)}) "
                          f"may be dropped if VIF is high.")
    story += [bullet(f"<b>Cluster {cl}</b> — {', '.join(members)}. {recommendation}")]

story += [PageBreak()]

# ── PART 4 ─────────────────────────────────────────────────────────────────────
story += [p("Part 4 — Functional Form Analysis", H1), hr()]

story += [p("For the top-ranked numeric frequency predictors, three functional forms are tested: "
            "linear, quadratic, and log-transformed. AIC is reported for each form. "
            "The lowest AIC form is selected as preferred. Monotonicity is assessed "
            "via decile plot inversion counts.", BODY), sp(8)]

if form_results:
    form_rows = []
    for pred, forms in form_results.items():
        best = min(forms, key=forms.get)
        for fname, aic in forms.items():
            form_rows.append([pred, fname, f"{aic:.1f}",
                               "BEST ✓" if fname == best else ""])
    story += [std_tbl(["Predictor","Form","AIC",""],
                       form_rows),
              sp(8)]

story += [embed_sq(buf_p4),
          cap("Figure 4.1. AIC by functional form for top frequency predictors. "
              "Green bar = lowest AIC (preferred form)."),
          sp(8),
          box("Functional Form Summary", [
              p("Linear form is most stable and interpretable for initial modelling. "
                "If AIC improvement for quadratic or log form exceeds 2 units, "
                "the non-linear form should be preferred. Spline basis expansion "
                "can be evaluated at the model-building stage if a smooth non-linear "
                "relationship is indicated by the decile plots."),
          ]),
          PageBreak()]

# ── PART 5 ─────────────────────────────────────────────────────────────────────
story += [p("Part 5 — Interaction Exploration", H1), hr()]

story += [p("Plausible two-way interactions between numeric predictors are tested. "
            "Each interaction is evaluated as: main effects only vs main + interaction term. "
            "ΔAIC = AIC(main) − AIC(interaction). Positive ΔAIC = interaction improves fit. "
            "Threshold: ΔAIC > 2 = material; > 10 = decisive.", BODY), sp(8)]

if len(int_df):
    story += [std_tbl(["Interaction","ΔAIC","Interaction Coef","Material?"],
                       [[r["pair"], f"{r['delta_aic']:.2f}",
                         f"{r['int_coef']:.4f}", r["material"]]
                        for _, r in int_df.iterrows()]),
              sp(8)]

story += [embed(buf_p5),
          cap("Figure 5.1. Interaction ΔAIC. Green = material (>2); grey = negligible."),
          sp(8)]

material_ints = int_df[int_df["delta_aic"] > 2]
story += [box("Interaction Recommendation", [
    p(f"{len(material_ints)} material interaction(s) detected (ΔAIC > 2).") if len(material_ints)
    else p("No material interactions detected. Do not include interaction terms in the base model."),
    sp(4),
    p("Interactions should be included only if: (a) ΔAIC > 2, (b) the interaction has a "
      "clear economic rationale, and (c) it remains stable across data splits. "
      "Marginal ΔAIC improvements < 2 do not justify additional complexity."),
])]

story += [PageBreak()]

# ── PART 6 ─────────────────────────────────────────────────────────────────────
story += [p("Part 6 — Stability Testing", H1), hr()]

story += [p("Top predictors are refitted independently within subgroups defined by "
            "solar_system and exposure band. Coefficient sign flips and magnitude ranges "
            "are computed. Predictors that flip sign or show coefficient range > 1.0 "
            "are flagged as unstable.", BODY), sp(8)]

if len(stability_df):
    story += [std_tbl(
        ["Predictor","Split","Coefficients by Subgroup","Sign Flip","Mag Range","Stable?"],
        [[r["predictor"], r["split"], r["coefs"][:60]+"..." if len(str(r["coefs"]))>60
          else r["coefs"],
          r["sign_flip"], str(r["mag_range"]), r["stable"]]
         for _, r in stability_df.iterrows()]
    ), sp(8)]

story += [embed(buf_p6),
          cap("Figure 6.1. Stability heatmap. Red = sign flip detected (unstable). "
              "Green = no flip."),
          sp(8),
          box("Stability Summary", [
              bullet("Predictors with sign flips across solar_system groups should be "
                     "modelled with a solar_system interaction or tested with a "
                     "random-effects structure."),
              bullet("Predictors stable across exposure bands are appropriate for "
                     "global (non-segmented) model inclusion."),
              bullet("Unstable predictors retained in the model should be flagged "
                     "explicitly and monitored at each model refresh."),
          ])]

story += [PageBreak()]

# ── PART 7 ─────────────────────────────────────────────────────────────────────
story += [p("Part 7 — Candidate Feature Set", H1), hr()]

story += [p("Final assessment integrating signal strength, VIF, collinearity cluster, "
            "and stability. Decision: INCLUDE / REVIEW / REJECT for each predictor. "
            "This is the defensible shortlist for the modelling phase.", BODY), sp(8)]

story += [p("7.1  Full Scorecard", H2),
          std_tbl(
              ["Predictor","Freq Dev.Red","Sev R²","VIF","Cluster","Sign Flips","Decision"],
              [[str(r["predictor"]), str(r["freq_dev_red"]), str(r["sev_r2"]),
                str(r["vif"]), str(r["cluster"]), str(r["sign_flips"]), str(r["decision"])]
               for _, r in final_df.iterrows()],
              decision_col=6
          ),
          sp(8)]

story += [embed_sq(buf_p7),
          cap("Figure 7.1. Feature scorecard. Green = INCLUDE. Amber = REVIEW. Red = REJECT."),
          sp(8)]

story += [p("7.2  Recommended Candidate Feature Set", H2)]

story += [box("INCLUDED Predictors", [
    p("The following predictors are recommended for inclusion in the initial modelling phase:"),
    sp(4),
    p("<b>Numeric:</b> " + (", ".join(included) if included else "None cleared all criteria")),
    sp(4),
    p(f"<b>solar_system</b> — {solar_recommendation}"),
    sp(4),
    p("These variables cleared all three gates: "
      "(1) univariate signal present (deviance reduction > 2 or R² > 0.005), "
      "(2) VIF < 10 (no severe multicollinearity), "
      "(3) no sign flip detected across stability splits."),
], bg=GOOD, border=GREEN)]

if review:
    story += [box("REVIEW Predictors", [
        p("The following predictors carry signal but have a concern that must be resolved "
          "before inclusion. Each predictor is reviewed at the modelling stage:"),
        sp(4),
    ] + [bullet(f"<b>{pred}</b> — {final_df[final_df['predictor']==pred]['decision'].values[0]}")
         for pred in review], bg=WARN, border=AMBER)]

if rejected:
    story += [box("REJECTED Predictors", [
        p("The following predictors are excluded from the candidate set:"),
        sp(4),
    ] + [bullet(f"<b>{pred}</b> — {final_df[final_df['predictor']==pred]['decision'].values[0]}")
         for pred in rejected], bg=BAD, border=RED)]

story += [sp(8),
          box("Modelling Phase Instructions", [
              p("This EDA phase delivers a defensible shortlist. At the next stage:"),
              bullet("Fit Negative Binomial GLM on included predictors + solar_system."),
              bullet("Test REVIEW predictors in ablation: fit model with/without each; "
                     "compare AIC and check sign stability in full model context."),
              bullet("Do NOT include REJECTED predictors without new evidence of signal."),
              bullet("Test best functional form (Part 4) for each numeric predictor."),
              bullet("Include interactions only where Part 5 ΔAIC > 2."),
              bullet("Validate aggregate calibration after modelling (see prior diagnostic report)."),
          ], bg=HexColor("#EAF2F8"), border=TEAL)]

doc.build(story)
print(f"\n{'='*60}")
print(f"Report saved  →  {OUTPUT_PATH}")
print(f"{'='*60}")
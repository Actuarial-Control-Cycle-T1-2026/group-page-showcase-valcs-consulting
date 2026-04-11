import os, warnings, math
warnings.filterwarnings("ignore")

import numpy as np
import pandas as pd

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from scipy import stats
from scipy.cluster.hierarchy import linkage, fcluster
from scipy.spatial.distance import squareform

import statsmodels.api as sm
from statsmodels.stats.outliers_influence import variance_inflation_factor
from patsy import dmatrix

# ─────────────────────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────────────────────
FILE_PATH   = "BI_data_CLEANED.xlsx"   # change to your path
OUT_DIR     = "bi_eda_outputs"
PDF_PATH    = os.path.join(OUT_DIR, "bi_eda_feature_investigation.pdf")

FREQ_SHEET  = "Frequency_Cleaned"
SEV_SHEET   = "Severity_Cleaned"

FREQ_TARGET = "claim_count"
EXPOSURE    = "exposure"

SEV_TARGET  = "claim_amount"

GROUP_COLS  = ["solar_system", "data_quality_flag"]  # used in stability checks
EXPOSURE_BANDS = [0, 0.25, 0.5, 0.75, 1, 2, 5, np.inf]  # tune to your scale

TOP_K_FUNCTIONAL = 12
TOP_K_STABILITY  = 15
TOP_K_INTERACTIONS = 10

CORR_HIGHLIGHT = 0.5
CLUSTER_CORR_THRESHOLD = 0.5  # clusters built from 1 - |corr|

os.makedirs(OUT_DIR, exist_ok=True)

# ─────────────────────────────────────────────────────────────────────────────
# REPORTLAB
# ─────────────────────────────────────────────────────────────────────────────
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, black
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    PageBreak, Image as RLImage
)
from reportlab.lib.enums import TA_LEFT

NAVY = HexColor("#1B2A4A")
TEAL = HexColor("#1A7A8A")
LGREY = HexColor("#F4F6F8")

styles = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=styles["Heading1"], textColor=NAVY, spaceAfter=8)
H2 = ParagraphStyle("H2", parent=styles["Heading2"], textColor=TEAL, spaceAfter=6)
B  = ParagraphStyle("B",  parent=styles["BodyText"], leading=12, spaceAfter=6)

def rl_table(df, col_widths=None, font_size=8):
    data = [df.columns.tolist()] + df.values.tolist()
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), LGREY),
        ("TEXTCOLOR", (0,0), (-1,0), black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), font_size),
        ("GRID", (0,0), (-1,-1), 0.25, HexColor("#BFC5CC")),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), ["#FFFFFF", "#FAFBFC"]),
    ]))
    return t

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def classify_series(s: pd.Series) -> str:
    if pd.api.types.is_numeric_dtype(s):
        nunq = s.dropna().nunique()
        if nunq <= 10:
            return "ordinal_or_discrete_numeric"
        return "numeric"
    return "categorical"

def basic_stats_numeric(s: pd.Series) -> dict:
    x = pd.to_numeric(s, errors="coerce")
    x = x.dropna()
    if len(x) == 0:
        return dict(n=0)
    return dict(
        n=int(len(x)),
        mean=float(x.mean()),
        std=float(x.std(ddof=1)) if len(x) > 1 else 0.0,
        p01=float(np.quantile(x, 0.01)),
        p05=float(np.quantile(x, 0.05)),
        p50=float(np.quantile(x, 0.50)),
        p95=float(np.quantile(x, 0.95)),
        p99=float(np.quantile(x, 0.99)),
        min=float(x.min()),
        max=float(x.max()),
        skew=float(stats.skew(x, bias=False)) if len(x) > 2 else np.nan,
        kurtosis_excess=float(stats.kurtosis(x, fisher=True, bias=False)) if len(x) > 3 else np.nan
    )

def missing_rate(s: pd.Series) -> float:
    return float(s.isna().mean())

def zero_rate(s: pd.Series) -> float:
    if not pd.api.types.is_numeric_dtype(s):
        return np.nan
    x = pd.to_numeric(s, errors="coerce")
    return float((x == 0).mean())

def outlier_rate_iqr(s: pd.Series) -> float:
    if not pd.api.types.is_numeric_dtype(s):
        return np.nan
    x = pd.to_numeric(s, errors="coerce").dropna()
    if len(x) < 10:
        return np.nan
    q1, q3 = np.quantile(x, 0.25), np.quantile(x, 0.75)
    iqr = q3 - q1
    lo, hi = q1 - 1.5*iqr, q3 + 1.5*iqr
    return float(((x < lo) | (x > hi)).mean())

def safe_log_exposure(df):
    e = pd.to_numeric(df[EXPOSURE], errors="coerce")
    return np.log(e.clip(lower=1e-12))

def save_hist_box(series, varname, tag):
    s = series.copy()
    path_hist = os.path.join(OUT_DIR, f"{tag}_{varname}_hist.png")
    path_box  = os.path.join(OUT_DIR, f"{tag}_{varname}_box.png")

    plt.figure()
    if pd.api.types.is_numeric_dtype(s):
        x = pd.to_numeric(s, errors="coerce").dropna()
        if len(x) > 0:
            plt.hist(x, bins=30)
            plt.title(f"{tag} {varname} histogram")
    else:
        vc = s.astype("string").fillna("NA").value_counts().head(25)
        plt.bar(vc.index.astype(str), vc.values)
        plt.xticks(rotation=90)
        plt.title(f"{tag} {varname} top categories")
    plt.tight_layout()
    plt.savefig(path_hist, dpi=160)
    plt.close()

    plt.figure()
    if pd.api.types.is_numeric_dtype(s):
        x = pd.to_numeric(s, errors="coerce").dropna()
        if len(x) > 0:
            plt.boxplot(x, vert=True)
            plt.title(f"{tag} {varname} boxplot")
            plt.tight_layout()
            plt.savefig(path_box, dpi=160)
    plt.close()
    return path_hist, path_box

def decile_plot_data(x, y, w=None, n=10):
    df = pd.DataFrame({"x": x, "y": y})
    if w is not None:
        df["w"] = w
    df = df.replace([np.inf, -np.inf], np.nan).dropna()
    if len(df) < 50:
        return None

    try:
        df["bin"] = pd.qcut(df["x"], q=n, duplicates="drop")
    except Exception:
        return None

    if w is None:
        g = df.groupby("bin")["y"].mean()
        c = df.groupby("bin")["x"].mean()
    else:
        g = df.groupby("bin").apply(lambda d: np.average(d["y"], weights=d["w"]))
        c = df.groupby("bin").apply(lambda d: np.average(d["x"], weights=d["w"]))
    out = pd.DataFrame({"x_mean": c.values, "y_mean": g.values})
    return out

def save_decile_plot(dec_df, title, filename):
    path = os.path.join(OUT_DIR, filename)
    plt.figure()
    plt.plot(dec_df["x_mean"], dec_df["y_mean"], marker="o")
    plt.title(title)
    plt.tight_layout()
    plt.savefig(path, dpi=160)
    plt.close()
    return path

def poisson_univariate(df, predictor):
    d = df[[FREQ_TARGET, EXPOSURE, predictor]].copy()
    d[FREQ_TARGET] = pd.to_numeric(d[FREQ_TARGET], errors="coerce")
    d[EXPOSURE] = pd.to_numeric(d[EXPOSURE], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[EXPOSURE] > 0]
    if len(d) < 200:
        return None

    y = d[FREQ_TARGET].values
    offset = np.log(d[EXPOSURE].values)

    if pd.api.types.is_numeric_dtype(d[predictor]):
        x = pd.to_numeric(d[predictor], errors="coerce").values
        X1 = sm.add_constant(x)
        model1 = sm.GLM(y, X1, family=sm.families.Poisson(), offset=offset).fit()
    else:
        cat = d[predictor].astype("string")
        X1 = pd.get_dummies(cat, drop_first=True)
        if X1.shape[1] == 0:
            return None
        X1 = sm.add_constant(X1)
        model1 = sm.GLM(y, X1, family=sm.families.Poisson(), offset=offset).fit()

    X0 = np.ones((len(d), 1))
    model0 = sm.GLM(y, X0, family=sm.families.Poisson(), offset=offset).fit()

    dev_red = float(model0.deviance - model1.deviance)
    aic = float(model1.aic)

    coef = model1.params.iloc[1] if hasattr(model1.params, "iloc") and len(model1.params) > 1 else np.nan
    sign = float(np.sign(coef)) if pd.notna(coef) else np.nan

    # decile plot for numeric only
    dec_path = None
    mono_flag = None
    if pd.api.types.is_numeric_dtype(d[predictor]):
        dec = decile_plot_data(pd.to_numeric(d[predictor], errors="coerce"), y / d[EXPOSURE].values)
        if dec is not None and len(dec) >= 5:
            dec_path = save_decile_plot(dec, f"Freq rate vs {predictor} (deciles)", f"freq_dec_{predictor}.png")
            diffs = np.diff(dec["y_mean"].values)
            mono_flag = bool(np.all(diffs >= 0) or np.all(diffs <= 0))

    return dict(
        predictor=predictor,
        n=int(len(d)),
        deviance_reduction=dev_red,
        aic=aic,
        coef=float(coef) if pd.notna(coef) else np.nan,
        sign=sign,
        decile_plot=dec_path,
        monotone=mono_flag
    )

def severity_univariate(df, predictor):
    d = df[[SEV_TARGET, predictor]].copy()
    d[SEV_TARGET] = pd.to_numeric(d[SEV_TARGET], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[SEV_TARGET] > 0]
    if len(d) < 200:
        return None

    y = np.log(d[SEV_TARGET].values)

    if pd.api.types.is_numeric_dtype(d[predictor]):
        x = pd.to_numeric(d[predictor], errors="coerce").values
        X1 = sm.add_constant(x)
        model1 = sm.OLS(y, X1).fit()
        coef = model1.params[1] if len(model1.params) > 1 else np.nan
    else:
        cat = d[predictor].astype("string")
        X1 = pd.get_dummies(cat, drop_first=True)
        if X1.shape[1] == 0:
            return None
        X1 = sm.add_constant(X1)
        model1 = sm.OLS(y, X1).fit()
        coef = model1.params.iloc[1] if len(model1.params) > 1 else np.nan

    r2 = float(model1.rsquared)
    aic = float(model1.aic)
    sign = float(np.sign(coef)) if pd.notna(coef) else np.nan

    dec_path = None
    mono_flag = None
    if pd.api.types.is_numeric_dtype(d[predictor]):
        dec = decile_plot_data(pd.to_numeric(d[predictor], errors="coerce"), y)
        if dec is not None and len(dec) >= 5:
            dec_path = save_decile_plot(dec, f"log(Severity) vs {predictor} (deciles)", f"sev_dec_{predictor}.png")
            diffs = np.diff(dec["y_mean"].values)
            mono_flag = bool(np.all(diffs >= 0) or np.all(diffs <= 0))

    return dict(
        predictor=predictor,
        n=int(len(d)),
        r2=r2,
        aic=aic,
        coef=float(coef) if pd.notna(coef) else np.nan,
        sign=sign,
        decile_plot=dec_path,
        monotone=mono_flag
    )

def compute_vif(df_num: pd.DataFrame) -> pd.DataFrame:
    d = df_num.replace([np.inf, -np.inf], np.nan).dropna()
    if d.shape[0] < 200 or d.shape[1] < 2:
        return pd.DataFrame({"variable": d.columns.tolist(), "vif": [np.nan]*d.shape[1]})
    X = sm.add_constant(d.values)
    out = []
    for i, col in enumerate(d.columns, start=1):
        try:
            vif = variance_inflation_factor(X, i)
        except Exception:
            vif = np.nan
        out.append((col, float(vif)))
    return pd.DataFrame(out, columns=["variable", "vif"]).sort_values("vif", ascending=False)

def corr_clusters(df_num: pd.DataFrame):
    d = df_num.replace([np.inf, -np.inf], np.nan).dropna()
    if d.shape[1] < 2:
        return None, None, None
    corr = d.corr().fillna(0.0)
    dist = 1 - np.abs(corr.values)
    np.fill_diagonal(dist, 0.0)
    condensed = squareform(dist, checks=False)
    Z = linkage(condensed, method="average")
    # cluster by threshold in distance space
    t = 1 - CLUSTER_CORR_THRESHOLD
    labels = fcluster(Z, t=t, criterion="distance")
    cl = pd.DataFrame({"variable": corr.columns, "cluster": labels})
    return corr, Z, cl

def functional_forms_frequency(df, predictor):
    d = df[[FREQ_TARGET, EXPOSURE, predictor]].copy()
    d[FREQ_TARGET] = pd.to_numeric(d[FREQ_TARGET], errors="coerce")
    d[EXPOSURE] = pd.to_numeric(d[EXPOSURE], errors="coerce")
    d[predictor] = pd.to_numeric(d[predictor], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[EXPOSURE] > 0]
    if len(d) < 500:
        return None

    y = d[FREQ_TARGET].values
    offset = np.log(d[EXPOSURE].values)
    x = d[predictor].values

    X_lin = sm.add_constant(x)
    m_lin = sm.GLM(y, X_lin, family=sm.families.Poisson(), offset=offset).fit()

    X_quad = sm.add_constant(np.column_stack([x, x**2]))
    m_quad = sm.GLM(y, X_quad, family=sm.families.Poisson(), offset=offset).fit()

    # spline with 4 df
    X_spl = dmatrix("bs(x, df=4, include_intercept=False)", {"x": x}, return_type="dataframe")
    X_spl = sm.add_constant(X_spl)
    m_spl = sm.GLM(y, X_spl, family=sm.families.Poisson(), offset=offset).fit()

    return dict(
        predictor=predictor,
        aic_linear=float(m_lin.aic),
        aic_quadratic=float(m_quad.aic),
        aic_spline=float(m_spl.aic)
    )

def functional_forms_severity(df, predictor):
    d = df[[SEV_TARGET, predictor]].copy()
    d[SEV_TARGET] = pd.to_numeric(d[SEV_TARGET], errors="coerce")
    d[predictor] = pd.to_numeric(d[predictor], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[SEV_TARGET] > 0]
    if len(d) < 500:
        return None

    y = np.log(d[SEV_TARGET].values)
    x = d[predictor].values

    X_lin = sm.add_constant(x)
    m_lin = sm.OLS(y, X_lin).fit()

    X_quad = sm.add_constant(np.column_stack([x, x**2]))
    m_quad = sm.OLS(y, X_quad).fit()

    X_spl = dmatrix("bs(x, df=4, include_intercept=False)", {"x": x}, return_type="dataframe")
    X_spl = sm.add_constant(X_spl)
    m_spl = sm.OLS(y, X_spl).fit()

    return dict(
        predictor=predictor,
        aic_linear=float(m_lin.aic),
        aic_quadratic=float(m_quad.aic),
        aic_spline=float(m_spl.aic)
    )

def interaction_probe_frequency(df, a, b):
    cols = [FREQ_TARGET, EXPOSURE, a, b]
    d = df[cols].copy()
    d[FREQ_TARGET] = pd.to_numeric(d[FREQ_TARGET], errors="coerce")
    d[EXPOSURE] = pd.to_numeric(d[EXPOSURE], errors="coerce")
    d[a] = pd.to_numeric(d[a], errors="coerce")
    d[b] = pd.to_numeric(d[b], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[EXPOSURE] > 0]
    if len(d) < 800:
        return None

    y = d[FREQ_TARGET].values
    offset = np.log(d[EXPOSURE].values)
    x1 = d[a].values
    x2 = d[b].values
    inter = x1 * x2

    X0 = np.ones((len(d), 1))
    m0 = sm.GLM(y, X0, family=sm.families.Poisson(), offset=offset).fit()

    X1 = sm.add_constant(np.column_stack([x1, x2]))
    m1 = sm.GLM(y, X1, family=sm.families.Poisson(), offset=offset).fit()

    X2 = sm.add_constant(np.column_stack([x1, x2, inter]))
    m2 = sm.GLM(y, X2, family=sm.families.Poisson(), offset=offset).fit()

    dev_add = float(m1.deviance - m2.deviance)
    return dict(a=a, b=b, n=int(len(d)), deviance_reduction_from_interaction=dev_add, aic_no_inter=float(m1.aic), aic_with_inter=float(m2.aic))

def interaction_probe_severity(df, a, b):
    cols = [SEV_TARGET, a, b]
    d = df[cols].copy()
    d[SEV_TARGET] = pd.to_numeric(d[SEV_TARGET], errors="coerce")
    d[a] = pd.to_numeric(d[a], errors="coerce")
    d[b] = pd.to_numeric(d[b], errors="coerce")
    d = d.replace([np.inf, -np.inf], np.nan).dropna()
    d = d[d[SEV_TARGET] > 0]
    if len(d) < 800:
        return None

    y = np.log(d[SEV_TARGET].values)
    x1 = d[a].values
    x2 = d[b].values
    inter = x1 * x2

    X1 = sm.add_constant(np.column_stack([x1, x2]))
    m1 = sm.OLS(y, X1).fit()

    X2 = sm.add_constant(np.column_stack([x1, x2, inter]))
    m2 = sm.OLS(y, X2).fit()

    # compare AIC and delta R2
    return dict(a=a, b=b, n=int(len(d)), delta_r2=float(m2.rsquared - m1.rsquared), aic_no_inter=float(m1.aic), aic_with_inter=float(m2.aic))

def stability_univariate_poisson(df, predictor, group_col, group_value):
    d = df[df[group_col] == group_value].copy()
    res = poisson_univariate(d, predictor)
    if res is None:
        return None
    # store pvalue if numeric and single coef
    cols = [FREQ_TARGET, EXPOSURE, predictor]
    d2 = d[cols].copy().replace([np.inf, -np.inf], np.nan).dropna()
    d2 = d2[pd.to_numeric(d2[EXPOSURE], errors="coerce") > 0]
    if len(d2) < 200:
        res["pvalue"] = np.nan
        return res
    y = pd.to_numeric(d2[FREQ_TARGET], errors="coerce").values
    offset = np.log(pd.to_numeric(d2[EXPOSURE], errors="coerce").values)
    if pd.api.types.is_numeric_dtype(d2[predictor]):
        x = pd.to_numeric(d2[predictor], errors="coerce").values
        m = sm.GLM(y, sm.add_constant(x), family=sm.families.Poisson(), offset=offset).fit()
        res["pvalue"] = float(m.pvalues[1]) if len(m.pvalues) > 1 else np.nan
    else:
        res["pvalue"] = np.nan
    return res

def stability_univariate_ols(df, predictor, group_col, group_value):
    d = df[df[group_col] == group_value].copy()
    res = severity_univariate(d, predictor)
    if res is None:
        return None
    cols = [SEV_TARGET, predictor]
    d2 = d[cols].copy().replace([np.inf, -np.inf], np.nan).dropna()
    d2[SEV_TARGET] = pd.to_numeric(d2[SEV_TARGET], errors="coerce")
    d2 = d2[d2[SEV_TARGET] > 0]
    if len(d2) < 200:
        res["pvalue"] = np.nan
        return res
    y = np.log(d2[SEV_TARGET].values)
    if pd.api.types.is_numeric_dtype(d2[predictor]):
        x = pd.to_numeric(d2[predictor], errors="coerce").values
        m = sm.OLS(y, sm.add_constant(x)).fit()
        res["pvalue"] = float(m.pvalues[1]) if len(m.pvalues) > 1 else np.nan
    else:
        res["pvalue"] = np.nan
    return res

# ─────────────────────────────────────────────────────────────────────────────
# LOAD
# ─────────────────────────────────────────────────────────────────────────────
df_freq = pd.read_excel(FILE_PATH, sheet_name=FREQ_SHEET)
df_sev  = pd.read_excel(FILE_PATH, sheet_name=SEV_SHEET)

# raw variables: exclude targets and exposure from predictors
freq_predictors = [c for c in df_freq.columns if c not in [FREQ_TARGET, EXPOSURE]]
sev_predictors  = [c for c in df_sev.columns  if c not in [SEV_TARGET]]

# ─────────────────────────────────────────────────────────────────────────────
# PART 1: DATA UNDERSTANDING
# ─────────────────────────────────────────────────────────────────────────────
freq_profiles = []
freq_plot_paths = []

for c in [FREQ_TARGET, EXPOSURE] + freq_predictors:
    s = df_freq[c]
    vtype = classify_series(s)
    miss = missing_rate(s)
    zero = zero_rate(s)
    out = outlier_rate_iqr(s)
    if vtype in ["numeric", "ordinal_or_discrete_numeric"] and pd.api.types.is_numeric_dtype(s):
        st = basic_stats_numeric(s)
    else:
        st = {"n": int(s.notna().sum())}
    freq_profiles.append({
        "variable": c,
        "type": vtype,
        "missing_rate": miss,
        "zero_rate": zero,
        "outlier_rate_iqr": out,
        **{k: st.get(k, np.nan) for k in ["n","mean","std","p01","p05","p50","p95","p99","min","max","skew","kurtosis_excess"]}
    })
    ph, pb = save_hist_box(s, c, tag="freq")
    freq_plot_paths.append((c, ph, pb))

sev_profiles = []
sev_plot_paths = []

for c in [SEV_TARGET] + sev_predictors:
    s = df_sev[c]
    vtype = classify_series(s)
    miss = missing_rate(s)
    zero = zero_rate(s)
    out = outlier_rate_iqr(s)
    if vtype in ["numeric", "ordinal_or_discrete_numeric"] and pd.api.types.is_numeric_dtype(s):
        st = basic_stats_numeric(s)
    else:
        st = {"n": int(s.notna().sum())}
    sev_profiles.append({
        "variable": c,
        "type": vtype,
        "missing_rate": miss,
        "zero_rate": zero,
        "outlier_rate_iqr": out,
        **{k: st.get(k, np.nan) for k in ["n","mean","std","p01","p05","p50","p95","p99","min","max","skew","kurtosis_excess"]}
    })
    ph, pb = save_hist_box(s, c, tag="sev")
    sev_plot_paths.append((c, ph, pb))

df_freq_profile = pd.DataFrame(freq_profiles).sort_values(["type","missing_rate"], ascending=[True, False])
df_sev_profile  = pd.DataFrame(sev_profiles).sort_values(["type","missing_rate"], ascending=[True, False])

df_freq_profile.to_csv(os.path.join(OUT_DIR, "freq_variable_profile.csv"), index=False)
df_sev_profile.to_csv(os.path.join(OUT_DIR, "sev_variable_profile.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 2: UNIVARIATE SIGNAL TESTING
# ─────────────────────────────────────────────────────────────────────────────
freq_uni = []
for p in freq_predictors:
    res = poisson_univariate(df_freq, p)
    if res is not None:
        freq_uni.append(res)
df_freq_uni = pd.DataFrame(freq_uni)
if len(df_freq_uni) > 0:
    df_freq_uni = df_freq_uni.sort_values(["deviance_reduction","aic"], ascending=[False, True])
df_freq_uni.to_csv(os.path.join(OUT_DIR, "freq_univariate_rank.csv"), index=False)

sev_uni = []
for p in sev_predictors:
    res = severity_univariate(df_sev, p)
    if res is not None:
        sev_uni.append(res)
df_sev_uni = pd.DataFrame(sev_uni)
if len(df_sev_uni) > 0:
    df_sev_uni = df_sev_uni.sort_values(["r2","aic"], ascending=[False, True])
df_sev_uni.to_csv(os.path.join(OUT_DIR, "sev_univariate_rank.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 3: REDUNDANCY AND COLLINEARITY
# ─────────────────────────────────────────────────────────────────────────────
freq_num = df_freq[[c for c in freq_predictors if pd.api.types.is_numeric_dtype(df_freq[c])]].copy()
sev_num  = df_sev [[c for c in sev_predictors  if pd.api.types.is_numeric_dtype(df_sev[c])]].copy()

freq_vif = compute_vif(freq_num)
sev_vif  = compute_vif(sev_num)
freq_vif.to_csv(os.path.join(OUT_DIR, "freq_vif.csv"), index=False)
sev_vif.to_csv(os.path.join(OUT_DIR, "sev_vif.csv"), index=False)

freq_corr, freq_Z, freq_clusters = corr_clusters(freq_num)
sev_corr,  sev_Z,  sev_clusters  = corr_clusters(sev_num)

if freq_corr is not None:
    freq_corr.to_csv(os.path.join(OUT_DIR, "freq_corr.csv"))
    freq_clusters.to_csv(os.path.join(OUT_DIR, "freq_corr_clusters.csv"), index=False)
if sev_corr is not None:
    sev_corr.to_csv(os.path.join(OUT_DIR, "sev_corr.csv"))
    sev_clusters.to_csv(os.path.join(OUT_DIR, "sev_corr_clusters.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 4: FUNCTIONAL FORMS
# ─────────────────────────────────────────────────────────────────────────────
freq_top = df_freq_uni.head(TOP_K_FUNCTIONAL)["predictor"].tolist() if len(df_freq_uni) else []
sev_top  = df_sev_uni.head(TOP_K_FUNCTIONAL)["predictor"].tolist() if len(df_sev_uni) else []

freq_ff = []
for p in freq_top:
    if pd.api.types.is_numeric_dtype(df_freq[p]):
        out = functional_forms_frequency(df_freq, p)
        if out is not None:
            freq_ff.append(out)
df_freq_ff = pd.DataFrame(freq_ff)
df_freq_ff.to_csv(os.path.join(OUT_DIR, "freq_functional_forms.csv"), index=False)

sev_ff = []
for p in sev_top:
    if pd.api.types.is_numeric_dtype(df_sev[p]):
        out = functional_forms_severity(df_sev, p)
        if out is not None:
            sev_ff.append(out)
df_sev_ff = pd.DataFrame(sev_ff)
df_sev_ff.to_csv(os.path.join(OUT_DIR, "sev_functional_forms.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 5: INTERACTION EXPLORATION
# ─────────────────────────────────────────────────────────────────────────────
# Default interaction candidates. Adjust to your actual column names.
interaction_candidates = [
    ("stress", "resilience"),
    ("experience", "stress"),
]

# Keep only those present and numeric in each sheet
freq_inter = []
for a, b in interaction_candidates:
    if a in df_freq.columns and b in df_freq.columns:
        if pd.api.types.is_numeric_dtype(df_freq[a]) and pd.api.types.is_numeric_dtype(df_freq[b]):
            r = interaction_probe_frequency(df_freq, a, b)
            if r is not None:
                freq_inter.append(r)
df_freq_inter = pd.DataFrame(freq_inter).sort_values("deviance_reduction_from_interaction", ascending=False) if len(freq_inter) else pd.DataFrame()
df_freq_inter.to_csv(os.path.join(OUT_DIR, "freq_interactions.csv"), index=False)

sev_inter = []
for a, b in interaction_candidates:
    if a in df_sev.columns and b in df_sev.columns:
        if pd.api.types.is_numeric_dtype(df_sev[a]) and pd.api.types.is_numeric_dtype(df_sev[b]):
            r = interaction_probe_severity(df_sev, a, b)
            if r is not None:
                sev_inter.append(r)
df_sev_inter = pd.DataFrame(sev_inter).sort_values("delta_r2", ascending=False) if len(sev_inter) else pd.DataFrame()
df_sev_inter.to_csv(os.path.join(OUT_DIR, "sev_interactions.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 6: STABILITY TESTING
# ─────────────────────────────────────────────────────────────────────────────
def exposure_band(s):
    x = pd.to_numeric(s, errors="coerce")
    return pd.cut(x, bins=EXPOSURE_BANDS, include_lowest=True)

df_freq["_exposure_band"] = exposure_band(df_freq[EXPOSURE])

freq_stab_preds = df_freq_uni.head(TOP_K_STABILITY)["predictor"].tolist() if len(df_freq_uni) else []
sev_stab_preds  = df_sev_uni.head(TOP_K_STABILITY)["predictor"].tolist() if len(df_sev_uni) else []

freq_stability_rows = []
for p in freq_stab_preds:
    # solar_system
    if "solar_system" in df_freq.columns:
        for g in df_freq["solar_system"].dropna().unique().tolist():
            r = stability_univariate_poisson(df_freq, p, "solar_system", g)
            if r is not None:
                r["group_col"] = "solar_system"
                r["group_val"] = str(g)
                freq_stability_rows.append(r)
    # exposure band
    for g in df_freq["_exposure_band"].dropna().unique().tolist():
        r = stability_univariate_poisson(df_freq, p, "_exposure_band", g)
        if r is not None:
            r["group_col"] = "exposure_band"
            r["group_val"] = str(g)
            freq_stability_rows.append(r)
    # data_quality_flag
    if "data_quality_flag" in df_freq.columns:
        for g in df_freq["data_quality_flag"].dropna().unique().tolist():
            r = stability_univariate_poisson(df_freq, p, "data_quality_flag", g)
            if r is not None:
                r["group_col"] = "data_quality_flag"
                r["group_val"] = str(g)
                freq_stability_rows.append(r)

df_freq_stab = pd.DataFrame(freq_stability_rows)
df_freq_stab.to_csv(os.path.join(OUT_DIR, "freq_stability.csv"), index=False)

sev_stability_rows = []
for p in sev_stab_preds:
    if "solar_system" in df_sev.columns:
        for g in df_sev["solar_system"].dropna().unique().tolist():
            r = stability_univariate_ols(df_sev, p, "solar_system", g)
            if r is not None:
                r["group_col"] = "solar_system"
                r["group_val"] = str(g)
                sev_stability_rows.append(r)
    if "data_quality_flag" in df_sev.columns:
        for g in df_sev["data_quality_flag"].dropna().unique().tolist():
            r = stability_univariate_ols(df_sev, p, "data_quality_flag", g)
            if r is not None:
                r["group_col"] = "data_quality_flag"
                r["group_val"] = str(g)
                sev_stability_rows.append(r)

df_sev_stab = pd.DataFrame(sev_stability_rows)
df_sev_stab.to_csv(os.path.join(OUT_DIR, "sev_stability.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# PART 7: DECISION LOGIC FOR CANDIDATE SETS
# ─────────────────────────────────────────────────────────────────────────────
def build_reject_accept(unirank_df, stab_df, corr_clusters_df, vif_df, metric_col, top_n=25):
    accepted = []
    rejected = []

    if len(unirank_df) == 0:
        return accepted, rejected, pd.DataFrame()

    top = unirank_df.head(top_n).copy()

    # stability: flag sign flips within each predictor
    stab_flags = {}
    if len(stab_df):
        for p, g in stab_df.groupby("predictor"):
            signs = g["sign"].dropna().values
            if len(signs) >= 2 and (np.any(signs > 0) and np.any(signs < 0)):
                stab_flags[p] = "sign_flip"
            else:
                # magnitude instability: coefficient CV above 1.0 on numeric cases
                coefs = g["coef"].dropna().values
                if len(coefs) >= 3:
                    mu = np.mean(coefs)
                    sd = np.std(coefs, ddof=1)
                    if abs(mu) > 1e-12 and (sd / abs(mu)) > 1.0:
                        stab_flags[p] = "magnitude_instability"

    # redundancy: choose representative per corr cluster using univariate rank then interpretability proxy
    reps = {}
    if corr_clusters_df is not None and len(corr_clusters_df):
        joined = corr_clusters_df.merge(unirank_df[["predictor", metric_col]], left_on="variable", right_on="predictor", how="left")
        for cl, g in joined.groupby("cluster"):
            g2 = g.dropna(subset=[metric_col]).sort_values(metric_col, ascending=False)
            if len(g2):
                reps[int(cl)] = g2.iloc[0]["variable"]

    vif_map = {}
    if len(vif_df):
        vif_map = dict(zip(vif_df["variable"], vif_df["vif"]))

    # accept logic
    for _, r in top.iterrows():
        p = r["predictor"]

        if p in stab_flags:
            rejected.append((p, f"unstable:{stab_flags[p]}"))
            continue

        # if numeric and in a cluster where it is not rep, reject as redundant
        if corr_clusters_df is not None and len(corr_clusters_df):
            row = corr_clusters_df[corr_clusters_df["variable"] == p]
            if len(row):
                cl = int(row.iloc[0]["cluster"])
                rep = reps.get(cl, None)
                if rep is not None and rep != p:
                    rejected.append((p, f"redundant_in_cluster_{cl}, keep:{rep}"))
                    continue

        # VIF screen: soft reject if extreme
        if p in vif_map and pd.notna(vif_map[p]) and vif_map[p] >= 10:
            rejected.append((p, f"high_vif:{vif_map[p]:.1f}"))
            continue

        accepted.append(p)

    rec_clusters = []
    if corr_clusters_df is not None and len(corr_clusters_df):
        for cl, rep in sorted(reps.items(), key=lambda x: x[0]):
            members = corr_clusters_df[corr_clusters_df["cluster"] == cl]["variable"].tolist()
            rec_clusters.append((cl, rep, ", ".join(members)))

    rec_clusters_df = pd.DataFrame(rec_clusters, columns=["cluster", "recommended_rep", "members"])
    return accepted, rejected, rec_clusters_df

freq_accept, freq_reject, freq_cluster_recs = build_reject_accept(
    df_freq_uni, df_freq_stab, freq_clusters, freq_vif,
    metric_col="deviance_reduction", top_n=30
)
sev_accept, sev_reject, sev_cluster_recs = build_reject_accept(
    df_sev_uni, df_sev_stab, sev_clusters, sev_vif,
    metric_col="r2", top_n=30
)

pd.DataFrame({"predictor": freq_accept}).to_csv(os.path.join(OUT_DIR, "freq_candidate_features.csv"), index=False)
pd.DataFrame({"predictor": sev_accept}).to_csv(os.path.join(OUT_DIR, "sev_candidate_features.csv"), index=False)
pd.DataFrame(freq_reject, columns=["predictor","reason"]).to_csv(os.path.join(OUT_DIR, "freq_rejected.csv"), index=False)
pd.DataFrame(sev_reject, columns=["predictor","reason"]).to_csv(os.path.join(OUT_DIR, "sev_rejected.csv"), index=False)
freq_cluster_recs.to_csv(os.path.join(OUT_DIR, "freq_cluster_recommendations.csv"), index=False)
sev_cluster_recs.to_csv(os.path.join(OUT_DIR, "sev_cluster_recommendations.csv"), index=False)

# ─────────────────────────────────────────────────────────────────────────────
# BUILD PDF
# ─────────────────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(PDF_PATH, pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=1.5*cm, bottomMargin=1.5*cm)
story = []

story.append(Paragraph("Business Interruption EDA and Feature Investigation", H1))
story.append(Paragraph("Scope: EDA, univariate signal, redundancy, functional form, interaction probes, stability checks. No final pricing models.", B))
story.append(Spacer(1, 10))

# Part 1 summaries
story.append(Paragraph("Part 1 Data Understanding", H2))
story.append(Paragraph("Frequency variable profile summary (sorted by type and missingness).", B))
story.append(rl_table(df_freq_profile.head(25).round(4)))
story.append(Spacer(1, 8))
story.append(Paragraph("Severity variable profile summary (sorted by type and missingness).", B))
story.append(rl_table(df_sev_profile.head(25).round(4)))
story.append(PageBreak())

# Part 2 ranks
story.append(Paragraph("Part 2 Univariate Signal Testing", H2))
if len(df_freq_uni):
    story.append(Paragraph("Frequency ranked predictors by deviance reduction then AIC.", B))
    story.append(rl_table(df_freq_uni.head(30).round(6)))
else:
    story.append(Paragraph("Frequency univariate: insufficient results. Check data volume and exposure integrity.", B))
story.append(Spacer(1, 10))

if len(df_sev_uni):
    story.append(Paragraph("Severity ranked predictors by R2 then AIC.", B))
    story.append(rl_table(df_sev_uni.head(30).round(6)))
else:
    story.append(Paragraph("Severity univariate: insufficient results. Check claim_amount positivity and missingness.", B))
story.append(PageBreak())

# Part 3 redundancy
story.append(Paragraph("Part 3 Redundancy and Collinearity", H2))
story.append(Paragraph(f"Correlation highlight threshold: |r| > {CORR_HIGHLIGHT}. VIF shown for numeric predictors.", B))
story.append(Paragraph("Frequency VIF (top 25).", B))
story.append(rl_table(freq_vif.head(25).round(3)))
story.append(Spacer(1, 8))
story.append(Paragraph("Severity VIF (top 25).", B))
story.append(rl_table(sev_vif.head(25).round(3)))

story.append(Spacer(1, 10))
story.append(Paragraph("Frequency correlation cluster representatives.", B))
story.append(rl_table(freq_cluster_recs.head(25)))
story.append(Spacer(1, 8))
story.append(Paragraph("Severity correlation cluster representatives.", B))
story.append(rl_table(sev_cluster_recs.head(25)))
story.append(PageBreak())

# Part 4 functional forms
story.append(Paragraph("Part 4 Functional Form Analysis", H2))
if len(df_freq_ff):
    story.append(Paragraph("Frequency AIC comparison for linear vs quadratic vs spline on top predictors.", B))
    story.append(rl_table(df_freq_ff.sort_values("aic_spline").head(25).round(3)))
else:
    story.append(Paragraph("Frequency functional form: insufficient numeric top predictors or sample size too small.", B))
story.append(Spacer(1, 8))
if len(df_sev_ff):
    story.append(Paragraph("Severity AIC comparison for linear vs quadratic vs spline on top predictors.", B))
    story.append(rl_table(df_sev_ff.sort_values("aic_spline").head(25).round(3)))
else:
    story.append(Paragraph("Severity functional form: insufficient numeric top predictors or sample size too small.", B))
story.append(PageBreak())

# Part 5 interactions
story.append(Paragraph("Part 5 Interaction Exploration", H2))
story.append(Paragraph("Only pre-specified domain interactions tested if both numeric and present.", B))
if len(df_freq_inter):
    story.append(Paragraph("Frequency interaction results.", B))
    story.append(rl_table(df_freq_inter.head(25).round(6)))
else:
    story.append(Paragraph("Frequency interactions: none run or insufficient sample.", B))
story.append(Spacer(1, 8))
if len(df_sev_inter):
    story.append(Paragraph("Severity interaction results.", B))
    story.append(rl_table(df_sev_inter.head(25).round(6)))
else:
    story.append(Paragraph("Severity interactions: none run or insufficient sample.", B))
story.append(PageBreak())

# Part 6 stability
story.append(Paragraph("Part 6 Stability Testing", H2))
story.append(Paragraph("Top predictors refit within solar_system, exposure bands and data_quality_flag where available.", B))
if len(df_freq_stab):
    story.append(Paragraph("Frequency stability sample (top 40 rows).", B))
    story.append(rl_table(df_freq_stab.head(40).round(6)))
else:
    story.append(Paragraph("Frequency stability: no results.", B))
story.append(Spacer(1, 8))
if len(df_sev_stab):
    story.append(Paragraph("Severity stability sample (top 40 rows).", B))
    story.append(rl_table(df_sev_stab.head(40).round(6)))
else:
    story.append(Paragraph("Severity stability: no results.", B))
story.append(PageBreak())

# Part 7 outputs
story.append(Paragraph("Part 7 Candidate Feature Sets and Reject Log", H2))
story.append(Paragraph("Frequency candidate features (post redundancy, VIF, stability screens).", B))
story.append(rl_table(pd.DataFrame({"predictor": freq_accept}).head(50)))
story.append(Spacer(1, 8))
story.append(Paragraph("Frequency rejected predictors with reason.", B))
story.append(rl_table(pd.DataFrame(freq_reject, columns=["predictor","reason"]).head(50)))

story.append(Spacer(1, 10))
story.append(Paragraph("Severity candidate features (post redundancy, VIF, stability screens).", B))
story.append(rl_table(pd.DataFrame({"predictor": sev_accept}).head(50)))
story.append(Spacer(1, 8))
story.append(Paragraph("Severity rejected predictors with reason.", B))
story.append(rl_table(pd.DataFrame(sev_reject, columns=["predictor","reason"]).head(50)))

doc.build(story)

print("Done.")
print("Outputs in:", OUT_DIR)
print("PDF:", PDF_PATH)
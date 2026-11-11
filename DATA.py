"""
HR Workforce Retention - Power BI-style analytics dashboard (Streamlit).      Run:  streamlit run dashboard.py
Data: put 'Workforce_Retention_Dataset_-_HR.xlsx' next to this file (or in data/source/), or upload it in the sidebar.
Every visual has an "Explain this visual" panel: what it shows, how to read it, and an auto-written insight from the CURRENT filters.
"""
from __future__ import annotations
import glob
import io
import os
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from scipy import stats

st.set_page_config(page_title="HR Workforce Retention Dashboard", page_icon="📊", layout="wide")
NAVY, TEAL, CORAL, AMBER = "#1E2761", "#00798C", "#D1495B", "#E8A33D"
PAL = [TEAL, CORAL, NAVY, AMBER, "#7A8B99", "#8D6A9F", "#4C9F70"]
BANDS = {"L1": (4e5, 8e5), "L2": (8e5, 1.5e6), "L3": (1.5e6, 2.5e6), "L4": (2.5e6, 4e6)}
REQUIRED = ["Employees_Master", "Departments_Master", "Managers_Master", "Attendance", "Leave_Records", "Performance_Ratings",
            "Engagement_Surveys", "Promotion_History", "Training_History", "Risk_Scores"]
SURVEY = ["EngagementScore", "JobSatisfaction", "WorkLifeBalance", "ManagerSatisfaction", "CareerGrowthScore"]
st.markdown(f"""<style>.block-container{{padding-top:1.2rem}} .card{{background:#fff;border:1px solid #d9dfe9;border-left:5px solid {TEAL};border-radius:10px;padding:12px 16px;margin-bottom:8px}}
.card b{{font-size:27px;color:{NAVY};display:block;line-height:1.2}} .card span{{color:#5b6577;font-size:13px}} .card small{{color:{CORAL};font-size:12px}}
h1,h2,h3{{color:{NAVY}}}</style>""", unsafe_allow_html=True)
_n = [0]


# ------------------------------------------------------------------------------------------------ helpers
def card(col, label: str, value, sub: str = "") -> None:
    """Power BI-style KPI card."""
    col.markdown(f"<div class='card'><span>{label}</span><b>{value}</b><small>{sub}</small></div>", unsafe_allow_html=True)


def show(fig, what: str, read: str, insight: str = "", height: int = 380) -> None:
    """Draw a chart followed by its explanation panel."""
    _n[0] += 1
    fig.update_layout(template="plotly_white", height=height, margin=dict(l=10, r=10, t=50, b=10), colorway=PAL)
    st.plotly_chart(fig, key=f"chart{_n[0]}")
    with st.expander("ℹ️ Explain this visual"):
        st.markdown(f"**What it shows:** {what}\n\n**How to read it:** {read}" + (f"\n\n**Insight from your current filters:** {insight}" if insight else ""))


def pct(a, b) -> float:
    return round(100.0 * a / b, 1) if b else 0.0


def find_default():
    here = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()
    for base in (here, Path.cwd()):
        for pat in (os.getenv("HR_XLSX", ""), "Workforce_Retention_Dataset*.xlsx", "data/source/*.xlsx", "*.xlsx"):
            if pat:
                hits = sorted(glob.glob(str(base / pat)))
                if hits:
                    return Path(hits[0]).name, Path(hits[0]).read_bytes()
    return None


@st.cache_data(show_spinner="Loading workbook...")
def load(raw: bytes) -> dict:
    xl = pd.ExcelFile(io.BytesIO(raw))
    return {s: xl.parse(s) for s in xl.sheet_names}


@st.cache_data(show_spinner="Building the employee 360 table...")
def build(T: dict):
    """One row per employee with every measure joined; plus the attendance table with real dates."""
    e = T["Employees_Master"].copy()
    for c in ("JoiningDate", "ExitDate"):
        e[c] = pd.to_datetime(e[c], errors="coerce")
    a = T["Attendance"].copy()
    a["Month"] = pd.to_datetime(a["MonthYear"], format="%b-%Y", errors="coerce")
    end = (a["Month"].max() + pd.offsets.MonthEnd(0)) if a["Month"].notna().any() else pd.Timestamp.today().normalize()
    e["Left"] = (e["AttritionFlag"].astype(str).str.lower() == "yes").astype(int)
    e["Status"] = np.where(e["Left"] == 1, "Left", "Active")
    e["Tenure"] = ((e["ExitDate"].fillna(end) - e["JoiningDate"]).dt.days / 365.25).round(2)
    e = e.merge(T["Departments_Master"][["DepartmentID", "DepartmentName"]], on="DepartmentID", how="left")
    e = e.merge(T["Managers_Master"][["ManagerID", "ManagerName"]], left_on="ReportingManagerID", right_on="ManagerID", how="left")
    lo = e["JobLevel"].map(lambda v: BANDS.get(v, (np.nan, np.nan))[0])
    hi = e["JobLevel"].map(lambda v: BANDS.get(v, (np.nan, np.nan))[1])
    e["Compa"] = (e["CurrentSalary"] / ((lo + hi) / 2)).round(3)
    att = a.groupby("EmployeeID").agg(Overtime=("OvertimeHours", "mean"), WD=("WorkingDays", "sum"), PD=("PresentDays", "sum"), WFH=("WFHPercent", "mean")).reset_index()
    att["Attendance"] = 100 * att["PD"] / att["WD"]
    sv = T["Engagement_Surveys"][["EmployeeID"] + SURVEY].drop_duplicates("EmployeeID")
    pf = T["Performance_Ratings"].groupby("EmployeeID").agg(Rating=("PerformanceRating", "mean"), Goals=("GoalsMetPercent", "mean")).reset_index()
    lv = T["Leave_Records"][T["Leave_Records"]["ApprovalStatus"] == "Approved"].groupby("EmployeeID").agg(LeaveDays=("DaysTaken", "sum")).reset_index()
    pr = T["Promotion_History"].groupby("EmployeeID").agg(Promotions=("PromotionID", "count"), AvgHike=("SalaryHikePercent", "mean")).reset_index()
    tr = T["Training_History"].assign(Cert=lambda d: (d["CertificationEarned"].astype(str).str.lower() == "yes").astype(int))
    tr = tr.groupby("EmployeeID").agg(TrainingHours=("DurationHours", "sum"), Certs=("Cert", "sum")).reset_index()
    rk = T["Risk_Scores"][["EmployeeID", "RiskScore", "RiskCategory", "ResignationProbability", "RiskTrend"]]
    for part in (att.drop(columns=["WD", "PD"]), sv, pf, lv, pr, tr, rk):
        e = e.merge(part, on="EmployeeID", how="left")
    for c in ("LeaveDays", "Promotions", "TrainingHours", "Certs"):
        e[c] = e[c].fillna(0)
    cut = lambda s, bins, labels: pd.cut(s, bins=bins, labels=labels, right=False).astype(str)
    e["AgeBand"] = cut(e["Age"], [0, 30, 40, 50, 200], ["<30", "30-39", "40-49", "50+"])
    e["TenureBand"] = cut(e["Tenure"], [-1, 1, 3, 5, 10, 100], ["1) <1y", "2) 1-3y", "3) 3-5y", "4) 5-10y", "5) 10y+"])
    e["PayPosition"] = cut(e["Compa"], [0, .75, 1, 1.25, 10], ["1) bottom of band", "2) below midpoint", "3) above midpoint", "4) top of band"])
    e["OvertimeBand"] = cut(e["Overtime"], [0, 10, 20, 30, 1000], ["1) <10 h", "2) 10-20 h", "3) 20-30 h", "4) 30+ h"])
    e["EngagementBand"] = cut(e["EngagementScore"], [0, 4, 7, 11], ["1) low (1-3)", "2) mid (4-6)", "3) high (7-10)"])
    e["TrainingBand"] = cut(e["TrainingHours"], [-1, 1, 41, 81, 10000], ["1) none", "2) 1-40 h", "3) 41-80 h", "4) 80+ h"])
    e["Promoted"] = np.where(e["Promotions"] > 0, "Promoted", "Not promoted")
    e["SalaryBand"] = pd.cut(e["CurrentSalary"], [0, 5e5, 8e5, 1.2e6, 2e6, 1e8], labels=["<5L", "5-8L", "8-12L", "12-20L", "20L+"], right=False).astype(str)
    return e, a, end


def rate_by(df: pd.DataFrame, col: str) -> pd.DataFrame:
    g = df.groupby(col).agg(Employees=("EmployeeID", "count"), Left=("Left", "sum"), Overtime=("Overtime", "mean"), Engagement=("EngagementScore", "mean")).reset_index()
    g["Rate"] = (100 * g["Left"] / g["Employees"]).round(1)
    return g


def monthly(df: pd.DataFrame, months) -> pd.DataFrame:
    rows = []
    for m in months:
        eom = m + pd.offsets.MonthEnd(0)
        op = int(((df["JoiningDate"] < m) & (df["ExitDate"].isna() | (df["ExitDate"] >= m))).sum())
        jn = int(((df["JoiningDate"] >= m) & (df["JoiningDate"] <= eom)).sum())
        lv = int(((df["ExitDate"] >= m) & (df["ExitDate"] <= eom)).sum())
        rows.append({"Month": m.strftime("%Y-%m"), "Opening": op, "Joiners": jn, "Leavers": lv, "Closing": op + jn - lv})
    return pd.DataFrame(rows)


def mwu(a, b):
    a, b = pd.Series(a).dropna(), pd.Series(b).dropna()
    if len(a) < 3 or len(b) < 3:
        return np.nan, np.nan
    r = stats.mannwhitneyu(a, b, alternative="two-sided")
    return float(r.pvalue), float(2 * r.statistic / (len(a) * len(b)) - 1)


def km(d, ev):
    d, ev = np.asarray(d, float), np.asarray(ev, int)
    xs, ys, s = [0.0], [1.0], 1.0
    for t in np.unique(d[ev == 1]):
        n, k = (d >= t).sum(), ((d == t) & (ev == 1)).sum()
        s *= 1 - k / n
        xs.append(float(t)); ys.append(float(s))
    return xs, ys


# ------------------------------------------------------------------------------------------------ data + sidebar
st.title("📊 HR Workforce Retention Dashboard")
st.sidebar.header("Data")
up = st.sidebar.file_uploader("Upload the HR workbook (.xlsx)", type=["xlsx"])
src = (up.name, up.getvalue()) if up is not None else find_default()
if src is None:
    st.info("No workbook found. Upload 'Workforce_Retention_Dataset_-_HR.xlsx' in the sidebar, or place it next to dashboard.py.")
    st.stop()
T = load(src[1])
missing = [s for s in REQUIRED if s not in T]
if missing:
    st.error(f"The workbook is missing these sheets: {', '.join(missing)}")
    st.stop()
E, A, END = build(T)
st.caption(f"Source: {src[0]}  |  {len(E)} employees  |  data period ends {END.date()}  |  Privacy mode hides names and exact salaries.")

st.sidebar.header("Filters")
mask = st.sidebar.toggle("Privacy mode (mask names & salaries)", value=True)
pick = lambda label, col: st.sidebar.multiselect(label, sorted(E[col].dropna().astype(str).unique()))
f_dep, f_loc, f_lvl, f_sta = pick("Department", "DepartmentName"), pick("Work location", "WorkLocation"), pick("Job level", "JobLevel"), pick("Status", "Status")
F = E.copy()
for col, vals in (("DepartmentName", f_dep), ("WorkLocation", f_loc), ("JobLevel", f_lvl), ("Status", f_sta)):
    if vals:
        F = F[F[col].astype(str).isin(vals)]
if F.empty:
    st.warning("No employees match the current filters.")
    st.stop()
st.sidebar.success(f"{len(F)} of {len(E)} employees selected")
st.sidebar.caption("⚠️ The dataset looks synthetic: leavers and stayers separate almost perfectly. Treat results as a method demonstration.")

tabs = st.tabs(["1 Overview", "2 Attrition drivers", "3 People & risk", "4 Workload & engagement", "5 Career & pay", "6 Statistics & advanced", "7 Data quality", "8 Explorer", "9 How to read"])

# ------------------------------------------------------------------------------------------------ 1 overview
with tabs[0]:
    st.subheader("Executive overview")
    st.markdown("*KPI cards give the headline numbers; the charts below show where attrition is concentrated, when it happened and why people say they left.*")
    L, Act = int(F["Left"].sum()), int((F["Left"] == 0).sum())
    c = st.columns(4)
    card(c[0], "Employees", len(F), f"{Act} active")
    card(c[1], "Leavers", L, "in the data period")
    card(c[2], "Attrition rate", f"{pct(L, len(F))}%", "leavers / employees")
    card(c[3], "Active high-risk", int(((F["Left"] == 0) & (F["RiskCategory"] == "High")).sum()), "supplied risk score")
    c = st.columns(4)
    card(c[0], "Avg tenure (yrs)", round(F["Tenure"].mean(), 1))
    card(c[1], "Avg engagement", f"{F['EngagementScore'].mean():.1f} / 10")
    card(c[2], "Avg overtime (h/month)", f"{F['Overtime'].mean():.1f}")
    card(c[3], "Avg attendance", f"{F['Attendance'].mean():.1f}%")
    d = rate_by(F, "DepartmentName").sort_values("Rate")
    top = d.iloc[-1]
    show(px.bar(d, x="Rate", y="DepartmentName", orientation="h", text="Rate", title="Attrition rate by department (%)", color_discrete_sequence=[CORAL]),
         "The share of employees who left, for each department.", "Longer bar = higher attrition. Hover to see headcount and leavers.",
         f"{top['DepartmentName']} has the highest rate at {top['Rate']}% ({int(top['Left'])} of {int(top['Employees'])}). Small departments move a lot with one leaver.")
    m = monthly(F, sorted(A["Month"].dropna().unique()))
    fig = go.Figure()
    fig.add_bar(x=m["Month"], y=m["Leavers"], name="Leavers", marker_color=CORAL)
    fig.add_scatter(x=m["Month"], y=m["Closing"], name="Closing headcount", mode="lines+markers", yaxis="y2", line=dict(color=NAVY))
    fig.update_layout(title="Leavers and headcount by month", yaxis=dict(title="Leavers"), yaxis2=dict(title="Headcount", overlaying="y", side="right"))
    worst = m.loc[m["Leavers"].idxmax()]
    show(fig, "Monthly leavers (bars) against month-end headcount (line).", "Bars show how many people left that month; the line shows how many remained.",
         f"Peak month: {worst['Month']} with {int(worst['Leavers'])} leavers. Joiners in the period: {int(m['Joiners'].sum())}, so headcount only falls.")
    lv = F[F["Left"] == 1]["ExitReason"].value_counts().reset_index()
    lv.columns = ["Reason", "Leavers"]
    if len(lv):
        show(px.pie(lv, names="Reason", values="Leavers", hole=0.5, title="Exit reasons"), "Why leavers say they left.", "Bigger slice = more leavers. Personal reasons (studies, relocation) are hard to influence; compensation, manager, work-life balance and growth are controllable.",
             f"Most common reason: {lv.iloc[0]['Reason']} ({int(lv.iloc[0]['Leavers'])} people).")
    rk = F.groupby(["RiskCategory", "Status"]).size().reset_index(name="Employees")
    show(px.bar(rk, x="RiskCategory", y="Employees", color="Status", barmode="stack", title="Supplied risk category vs actual outcome", category_orders={"RiskCategory": ["Low", "Medium", "High"]}),
         "How the supplied risk score lines up with who actually left.", "A good risk score puts most leavers in 'High'.",
         "If every leaver sits in 'High', the score probably used the outcome itself (leakage) - do not treat it as independent proof.")

# ------------------------------------------------------------------------------------------------ 2 drivers
with tabs[1]:
    st.subheader("Attrition drivers")
    st.markdown("*Pick any factor to see attrition rate for each group. This is the Power BI 'slice and dice' view.*")
    DIMS = {"Department": "DepartmentName", "Manager": "ManagerName", "Work location": "WorkLocation", "Job level": "JobLevel", "Gender": "Gender", "Age band": "AgeBand", "Education": "EducationLevel",
            "Marital status": "MaritalStatus", "Tenure": "TenureBand", "Pay position in band": "PayPosition", "Overtime": "OvertimeBand", "Engagement": "EngagementBand", "Training hours": "TrainingBand", "Promotion": "Promoted"}
    dim = st.selectbox("Break attrition down by", list(DIMS))
    g = rate_by(F, DIMS[dim]).sort_values(DIMS[dim])
    hot = g.sort_values("Rate").iloc[-1]
    show(px.bar(g, x=DIMS[dim], y="Rate", text="Rate", hover_data=["Employees", "Left"], title=f"Attrition rate by {dim.lower()} (%)", color_discrete_sequence=[CORAL]),
         f"Attrition rate for each {dim.lower()} group.", "Taller bar = higher attrition. Always check the group size in the table below before drawing conclusions.",
         f"Highest: {hot[DIMS[dim]]} at {hot['Rate']}% (n={int(hot['Employees'])}).")
    st.dataframe(g.rename(columns={DIMS[dim]: dim}).round(2), hide_index=True)
    lv = F[F["Left"] == 1]
    if len(lv):
        show(px.treemap(lv, path=["DepartmentName", "ExitReason"], title="Leavers: department, then exit reason"), "Every leaver, grouped by department and stated reason.", "Box size = number of leavers. Click a department to zoom in.",
             "Shows whether one department has a single dominant reason.")

# ------------------------------------------------------------------------------------------------ 3 people & risk
with tabs[2]:
    st.subheader("People and risk")
    st.markdown("*Each dot is one employee. The watch-list lists active employees currently flagged High risk, with the reasons behind the flag.*")
    show(px.scatter(F, x="Overtime", y="EngagementScore", color="Status", hover_data=["EmployeeID", "DepartmentName"], title="Overtime vs engagement", color_discrete_map={"Left": CORAL, "Active": TEAL}),
         "Average monthly overtime against engagement score for every employee.", "Bottom-right (high overtime, low engagement) is the danger zone.",
         f"Leavers average {F[F['Left'] == 1]['Overtime'].mean():.1f} overtime hours vs {F[F['Left'] == 0]['Overtime'].mean():.1f} for active staff." if L else "No leavers in the current selection.")
    show(px.histogram(F, x="RiskScore", color="Status", nbins=20, barmode="overlay", opacity=0.7, title="Risk score distribution", color_discrete_map={"Left": CORAL, "Active": TEAL}),
         "How the supplied risk scores are spread for leavers and stayers.", "Two clearly separated humps mean the score separates the groups.", "A perfect split again points to a score built with knowledge of the outcome.")
    w = F[(F["Left"] == 0) & (F["RiskCategory"] == "High")].copy()
    w["Flags"] = w.apply(lambda r: ", ".join(x for x, ok in (("low engagement", r["EngagementScore"] <= 5), ("high overtime", r["Overtime"] >= 20), ("low attendance", r["Attendance"] < 90), ("weak rating", r["Rating"] <= 3)) if ok) or "-", axis=1)
    w["Employee"] = w["EmployeeID"] if mask else w["EmployeeID"] + " " + w["FullName"].astype(str)
    cols = ["Employee", "DepartmentName", "JobRole", "ManagerName", "RiskScore", "RiskTrend", "EngagementScore", "Overtime", "Flags"]
    st.markdown("**HR watch-list (active, High risk)** - a flag starts a supportive conversation, never a decision about pay, promotion or dismissal.")
    st.dataframe(w.sort_values("RiskScore", ascending=False)[cols].round(1), hide_index=True)

# ------------------------------------------------------------------------------------------------ 4 workload & engagement
with tabs[3]:
    st.subheader("Workload and engagement")
    st.markdown("*Heat-maps and box plots compare departments; the line chart shows how overtime or attendance moved month by month.*")
    h = F.groupby("DepartmentName")[SURVEY].mean().round(2)
    show(px.imshow(h, text_auto=".1f", aspect="auto", color_continuous_scale="RdYlGn", title="Engagement heat-map (department x survey dimension)"),
         "Average survey score (1-10) for five dimensions in each department.", "Green = better, red = worse. Look for red cells.",
         f"Lowest cell: {h.stack().idxmin()[0]} on {h.stack().idxmin()[1]} ({h.stack().min():.1f}).")
    show(px.box(F, x="DepartmentName", y="Overtime", color="Status", title="Overtime by department"), "Spread of average monthly overtime per department.", "The box is the middle 50% of employees; dots are outliers.", "")
    met = st.selectbox("Monthly trend metric", ["OvertimeHours", "AttendancePercent", "WFHPercent"])
    ta = A[A["EmployeeID"].isin(F["EmployeeID"])].merge(F[["EmployeeID", "DepartmentName"]], on="EmployeeID").groupby(["Month", "DepartmentName"])[met].mean().reset_index()
    show(px.line(ta, x="Month", y=met, color="DepartmentName", markers=True, title=f"{met} by month"), "The chosen workload metric for each department over time.", "A rising line means the metric is worsening (overtime) or improving (attendance).", "")
    num = F[["Age", "Tenure", "CurrentSalary", "Overtime", "Attendance", "EngagementScore", "Rating", "TrainingHours", "Promotions", "RiskScore", "Left"]].corr().round(2)
    show(px.imshow(num, text_auto=".1f", aspect="auto", color_continuous_scale="RdBu_r", zmin=-1, zmax=1, title="Correlation matrix"), "How strongly pairs of numbers move together.", "+1 move together, -1 move opposite, 0 unrelated. Look at the 'Left' row.",
         "Correlation is not causation; use it to choose what to investigate.", height=480)

# ------------------------------------------------------------------------------------------------ 5 career & pay
with tabs[4]:
    st.subheader("Career growth and pay")
    st.markdown("*Pay is compared with the midpoint of each job level's salary band (compa-ratio: 1.0 = paid at the midpoint).*")
    show(px.box(F, x="Status", y="Compa", color="Status", points="all", title="Pay position (compa-ratio) by status", color_discrete_map={"Left": CORAL, "Active": TEAL}),
         "Salary divided by the midpoint of the job level's band.", "Below 1.0 = paid under the band midpoint.", f"Median compa-ratio: leavers {F[F['Left'] == 1]['Compa'].median():.2f} vs active {F[F['Left'] == 0]['Compa'].median():.2f}." if L else "")
    p = rate_by(F, "DepartmentName")
    pr = F.groupby("DepartmentName").agg(Promoted=("Promotions", lambda s: 100 * (s > 0).mean()), Training=("TrainingHours", "mean")).reset_index().round(1)
    show(px.bar(pr, x="DepartmentName", y="Promoted", title="% of employees promoted, by department", color_discrete_sequence=[TEAL]), "Share of each department that has had at least one promotion.", "Low bars suggest limited visible career progression.", "")
    show(px.bar(pr, x="DepartmentName", y="Training", title="Average training hours per employee", color_discrete_sequence=[NAVY]), "Learning hours completed per employee.", "Compare with attrition: low training plus high attrition may signal weak development.", "Leavers stop accumulating training, so this chart can exaggerate the link.")
    sal = F.groupby(["JobLevel", "Gender"]).agg(Employees=("EmployeeID", "count"), AvgSalary=("CurrentSalary", "mean")).reset_index()
    if mask:
        sal["AvgSalary"] = "hidden (privacy mode)"
    st.markdown("**Pay by level and gender (aggregate check for pay gaps)**")
    st.dataframe(sal, hide_index=True)

# ------------------------------------------------------------------------------------------------ 6 statistics & advanced
with tabs[5]:
    st.subheader("Statistics and advanced analytics")
    st.markdown("*Statistical tests check whether differences between leavers and stayers are real or luck. Cliff's delta is the effect size (+1/-1 = the groups do not overlap at all).*")
    mets = ["EngagementScore", "JobSatisfaction", "WorkLifeBalance", "ManagerSatisfaction", "CareerGrowthScore", "Overtime", "Attendance", "Rating", "Goals", "Tenure", "Compa", "TrainingHours", "Age"]
    rows = []
    for mt in mets:
        pv, dl = mwu(F[F["Left"] == 1][mt], F[F["Left"] == 0][mt])
        rows.append({"Metric": mt, "Median (left)": F[F["Left"] == 1][mt].median(), "Median (stayed)": F[F["Left"] == 0][mt].median(), "p-value": pv, "Cliff's delta": dl})
    S = pd.DataFrame(rows).sort_values("p-value").round(4)
    st.markdown("**Mann-Whitney U test (leavers vs stayers)**")
    st.dataframe(S, hide_index=True)
    st.caption("p-value < 0.05 = unlikely to be chance. A delta of exactly +/-1 means every leaver is above (or below) every stayer: real data almost never does that.")
    if (S["Cliff's delta"].abs() > 0.99).any():
        st.warning("Perfect separation detected - the data is probably synthetic or recorded after people decided to leave.")
    if F["Left"].nunique() == 2:
        fl = go.Figure()
        for name, grp, colr in (("Paid below band midpoint", F[F["Compa"] < 1], CORAL), ("Paid at/above midpoint", F[F["Compa"] >= 1], TEAL)):
            if len(grp) > 2:
                xs, ys = km(grp["Tenure"], grp["Left"])
                fl.add_scatter(x=xs, y=ys, mode="lines", line_shape="hv", name=f"{name} (n={len(grp)})", line=dict(color=colr))
        fl.update_layout(title="Kaplan-Meier retention curve", xaxis_title="Tenure (years)", yaxis_title="Share still employed", yaxis_range=[0, 1.02])
        show(fl, "Survival analysis: the share of employees still employed after each number of years.", "A curve that drops earlier means people in that group leave sooner. It correctly handles people who have not left yet.",
             "Compare the two curves: a gap means pay position is linked to how long people stay.")
    try:
        from sklearn.cluster import KMeans
        from sklearn.metrics import silhouette_score
        from sklearn.preprocessing import StandardScaler
        feat = ["Overtime", "Attendance", "EngagementScore", "Compa", "Tenure", "Rating"]
        Z = F.dropna(subset=feat).copy()
        if len(Z) >= 12:
            X = StandardScaler().fit_transform(Z[feat])
            best = max(((silhouette_score(X, KMeans(n_clusters=k, n_init=10, random_state=42).fit_predict(X)), k) for k in (2, 3, 4)))
            Z["Persona"] = ("Persona " + pd.Series(KMeans(n_clusters=best[1], n_init=10, random_state=42).fit_predict(X) + 1, index=Z.index).astype(str))
            prof = Z.groupby("Persona").agg(Employees=("EmployeeID", "count"), Left=("Left", "sum"), Overtime=("Overtime", "mean"), Engagement=("EngagementScore", "mean"), Compa=("Compa", "mean")).round(2).reset_index()
            prof["Attrition %"] = (100 * prof["Left"] / prof["Employees"]).round(1)
            show(px.scatter(Z, x="Overtime", y="EngagementScore", color="Persona", title=f"k-means employee personas (k={best[1]}, silhouette {best[0]:.2f})"), "Groups of similar employees found automatically from behaviour (no age, gender or marital status).", "Each colour is a persona; the table gives its average profile and attrition.", "")
            st.dataframe(prof, hide_index=True)
    except ImportError:
        st.info("Install scikit-learn to see employee personas (k-means).")
    au = []
    for mt in mets:
        s = F[[mt, "Left"]].dropna()
        if s["Left"].nunique() == 2:
            r = s[mt].rank()
            n1, n0 = int(s["Left"].sum()), int((1 - s["Left"]).sum())
            au.append({"Metric": mt, "Strength": abs((r[s["Left"] == 1].sum() - n1 * (n1 + 1) / 2) / (n1 * n0) - 0.5) * 2})
    if au:
        au = pd.DataFrame(au).sort_values("Strength")
        show(px.bar(au, x="Strength", y="Metric", orientation="h", title="Driver strength (0 = no signal, 1 = perfect separation)", color_discrete_sequence=[NAVY]), "How well each measure alone separates leavers from stayers.", "Longer bar = stronger driver.", "Bars near 1.0 signal perfect separation (see the warning above).")

# ------------------------------------------------------------------------------------------------ 7 data quality
with tabs[6]:
    st.subheader("Data quality")
    st.markdown("*Automatic checks on the workbook. PASS = no issues found; WARN = review.*")
    e0, dep, mg, at = T["Employees_Master"], T["Departments_Master"], T["Managers_Master"], T["Attendance"]
    chk = []
    add = lambda name, bad: chk.append({"Check": name, "Issues": int(bad), "Result": "PASS" if bad == 0 else "WARN"})
    add("Duplicate EmployeeID", e0["EmployeeID"].duplicated().sum())
    add("Employees with unknown department", (~e0["DepartmentID"].isin(dep["DepartmentID"])).sum())
    add("Employees with unknown manager", (~e0["ReportingManagerID"].isin(mg["ManagerID"])).sum())
    add("Exit date present but not flagged as left (or reverse)", ((e0["AttritionFlag"].astype(str).str.lower() == "yes") != e0["ExitDate"].notna()).sum())
    add("Salary outside job-level band", sum(not (BANDS[l][0] <= s <= BANDS[l][1]) for l, s in zip(e0["JobLevel"], e0["CurrentSalary"]) if l in BANDS))
    add("Absent days != working - present", (at["AbsentDays"] != at["WorkingDays"] - at["PresentDays"]).sum())
    add("Attendance % differs from recomputed", ((100 * at["PresentDays"] / at["WorkingDays"]).round(1) - at["AttendancePercent"]).abs().gt(0.11).sum())
    add("Attendance after exit month", int(at.merge(E[["EmployeeID", "ExitDate"]], on="EmployeeID").pipe(lambda d: (pd.to_datetime(d["MonthYear"], format="%b-%Y") > d["ExitDate"]).sum())))
    add("Engagement scores outside 1-10", (~T["Engagement_Surveys"][SURVEY].stack().between(1, 10)).sum())
    st.dataframe(pd.DataFrame(chk), hide_index=True)
    prof = pd.DataFrame([{"Sheet": s, "Rows": len(d), "Columns": d.shape[1], "Missing cells": int(d.isna().sum().sum()), "Duplicate rows": int(d.duplicated().sum())} for s, d in T.items() if s in REQUIRED])
    st.markdown("**Sheet profile**")
    st.dataframe(prof, hide_index=True)

# ------------------------------------------------------------------------------------------------ 8 explorer
with tabs[7]:
    st.subheader("Data explorer")
    st.markdown("*Build your own view: choose columns, search, and download. Sensitive columns are masked in privacy mode.*")
    base = F.copy()
    if mask:
        base["FullName"] = "hidden"
        base["CurrentSalary"] = base["SalaryBand"]
    show_cols = st.multiselect("Columns", list(base.columns), default=["EmployeeID", "DepartmentName", "JobRole", "JobLevel", "Tenure", "Overtime", "EngagementScore", "RiskCategory", "Status"])
    q = st.text_input("Search (any text)")
    view = base[show_cols] if show_cols else base
    if q:
        view = view[view.astype(str).apply(lambda r: r.str.contains(q, case=False, regex=False).any(), axis=1)]
    st.dataframe(view, hide_index=True)
    st.download_button("Download this view (CSV)", view.to_csv(index=False).encode("utf-8"), file_name="hr_view.csv", mime="text/csv")
    st.markdown("**Pivot**")
    pv_row = st.selectbox("Rows", ["DepartmentName", "JobLevel", "WorkLocation", "TenureBand", "PayPosition", "AgeBand"])
    pv_val = st.selectbox("Measure", ["Attrition %", "Employees", "Avg overtime", "Avg engagement", "Avg tenure"])
    pt = F.groupby(pv_row).agg(Employees=("EmployeeID", "count"), Left=("Left", "sum"), OT=("Overtime", "mean"), ENG=("EngagementScore", "mean"), TEN=("Tenure", "mean")).reset_index()
    pt["Attrition %"], pt["Avg overtime"], pt["Avg engagement"], pt["Avg tenure"] = (100 * pt["Left"] / pt["Employees"]).round(1), pt["OT"].round(1), pt["ENG"].round(2), pt["TEN"].round(2)
    show(px.bar(pt, x=pv_row, y=pv_val, title=f"{pv_val} by {pv_row}", color_discrete_sequence=[TEAL]), "Your chosen measure for the chosen grouping.", "Change the dropdowns to explore.", "")

# ------------------------------------------------------------------------------------------------ 9 guide
with tabs[8]:
    st.subheader("How to read this dashboard")
    st.markdown("""
**Structure (like a Power BI report).** The sidebar holds the *slicers* (department, location, level, status): every number and chart on every tab responds to them. Each tab answers one question:

| Tab | Question it answers |
|---|---|
| 1 Overview | How big is the problem, where, when and why do people say they left? |
| 2 Attrition drivers | Which groups (tenure, pay position, overtime, manager...) have the highest attrition? |
| 3 People & risk | Who is at risk right now and what is driving the flag? |
| 4 Workload & engagement | Are people overworked or disengaged, and where? |
| 5 Career & pay | Do pay position, promotions and training relate to leaving? |
| 6 Statistics & advanced | Are the differences real (tests), how long do people stay (survival), are there natural personas (k-means)? |
| 7 Data quality | Can I trust the data? |
| 8 Explorer | Build my own table or pivot and download it. |

**Analyst method.** (1) Describe (KPIs) -> (2) Diagnose (drivers) -> (3) Test (statistics) -> (4) Segment and survive (personas, Kaplan-Meier) -> (5) Act (watch-list, recommendations) -> (6) Verify the data first (data quality).

**Key terms.** *Attrition rate* = leavers / employees. *Compa-ratio* = salary / midpoint of the job level's band. *p-value* = chance of seeing a difference this big if there were no real effect. *Cliff's delta* = how separated two groups are (+/-1 = no overlap). *Kaplan-Meier* = share still employed over time, handling people who have not left yet.

**Responsible use.** A risk flag starts a supportive conversation; it must never decide pay, promotion or dismissal. Privacy mode is on by default. The data looks synthetic (perfect separation), so conclusions are a demonstration of method until re-run on real, approved data.
""")

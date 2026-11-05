"""Mavericks Player Marketability: player, team, leaderboard, compare, similar-player and data views.
The Excel workbook is the database (see data_layer.py).  Run:  streamlit run app.py
"""
from __future__ import annotations

import hashlib
import html
import re
import tempfile
import time
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from sklearn.metrics import roc_auc_score

import data_layer as dl
import knn

st.set_page_config(page_title="Mavs Marketability", layout="wide")

# ---------------------------------------------------------------- palette (Mavericks: royal blue #00538C, navy, silver, on white)
INK, MUTED, RULE = "#0E2240", "#5B6B7A", "#DCE3EA"
GROUND, SURFACE, SAND = "#FFFFFF", "#FFFFFF", "#C9D3DC"
ACCENT, ACCENT_SOFT = "#00538C", "#E1ECF5"
POS, NEG, NOTE = "#00538C", "#B9472E", "#7A8794"                 # strong (Mavs blue) / weak (orange-red) / neutral
LOW_RGB, MID_RGB, HIGH_RGB = (185, 71, 46), (232, 237, 241), (0, 83, 140)
FONT = '"Work Sans", -apple-system, "Helvetica Neue", Arial, sans-serif'
BUNDLED = Path(__file__).resolve().parent / "data" / "Mavs_model_slim.xlsx"

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Work+Sans:wght@400;500;600;700&display=swap');
#MainMenu, footer, [data-testid="stAppDeployButton"], [data-testid="stDecoration"] { display: none; }
[data-testid="stHeader"] { background: transparent; }
/* Streamlit's invisible header sits on top of the band: give it no height (its few buttons still overflow and show)
   and let clicks through it, so it can never swallow clicks meant for the tabs */
[data-testid="stHeader"] { pointer-events: none; height: 0 !important; min-height: 0 !important; overflow: visible; }
[data-testid="stHeader"] button, [data-testid="stHeader"] a { pointer-events: auto; color: #FFFFFF; }
/* the (empty) element holding this stylesheet would otherwise leave a gap above the band */
[data-testid="stElementContainer"]:has(style) { display: none; }
.block-container { padding-top: 0; max-width: 1280px; container-type: inline-size; }
.stApp, .stApp p, .stApp li, .stApp label, .stApp input, .stApp textarea, .stApp button, .stApp td, .stApp th, .stApp div[data-baseweb] {
  font-family: 'Work Sans', -apple-system, 'Helvetica Neue', Arial, sans-serif; }
.stApp h1, .stApp h2, .stApp h3 { font-family: 'Barlow Condensed', sans-serif; letter-spacing: .01em; }

/* top bar: the main tab list itself is the full-bleed blue band; the brand is drawn into its left padding
   (::before = mark + title, ::after = subtitle, both set further down), so nothing is layered over the tabs. */
/* role-based selectors work on both the older (baseweb) and newer Streamlit tab markup */
.stTabs [role="tablist"] { position: relative; gap: 2px; background: #00538C; min-height: 84px; align-items: center; justify-content: flex-end;
                           padding: 0 0 0 330px; flex-wrap: nowrap; border: none !important; box-shadow: 0 0 0 100vmax #00538C !important;
                           clip-path: inset(0 -100vmax); }
/* newer Streamlit uses these pseudo-elements for its own underline, so reset its geometry and background first */
.stTabs [role="tablist"]::before, .stTabs [role="tablist"]::after { position: absolute; left: 4px; right: auto; bottom: auto; width: auto; height: auto;
  background: none; border: none; color: #FFFFFF; white-space: nowrap; pointer-events: none; }
.stTabs [role="tablist"]::before { top: 18px; height: 48px; padding-left: 68px; line-height: 1;
  font-family: 'Barlow Condensed', sans-serif; font-weight: 700; font-size: 1.7rem; letter-spacing: .03em; text-transform: uppercase;
  /* NBA logo (logoman + white wordmark, made for dark backgrounds) straight on the blue band */
  background: url("https://cdn.nba.com/logos/leagues/logo-nba.svg") no-repeat left center / 58px 34px; }
.stTabs [role="tablist"]::after { top: 54px; left: 69px; font-size: .8rem; line-height: 1.2; color: #CFE0EE; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }
.stTabs [role="tab"] { color: #FFFFFF !important; background: transparent; border-radius: 999px; padding: 9px 16px !important; height: auto; white-space: nowrap;
                       border: none !important; box-shadow: none !important; margin: 0 !important; }
.stTabs [role="tab"] p { color: inherit !important; font-weight: 600; font-size: 1rem; }
@container (max-width: 1040px) {   /* content area too narrow for tabs beside the brand: tabs wrap onto rows under it */
  .stTabs [role="tablist"] { padding: 80px 0 12px; justify-content: flex-start; flex-wrap: wrap; }
}
.stTabs [role="tab"]::after, .stTabs [role="tab"]::before { display: none !important; }
.stTabs [role="tab"]:hover { background: #00467A; color: #FFFFFF !important; }
.stTabs [role="tab"][aria-selected="true"] { background: #FFFFFF !important; color: #00538C !important; }
/* nested tabs (Profile / Recommendations, For agents / team / player): underline style */
.stTabs .stTabs { margin-top: 0; }
.stTabs .stTabs [role="tablist"] { background: transparent; padding: 0; border-radius: 0; border-bottom: 1px solid #DCE3EA !important; gap: 2px;
                                   min-height: 0; justify-content: flex-start; flex-wrap: wrap; box-shadow: none !important; clip-path: none; }
.stTabs .stTabs [role="tablist"]::before, .stTabs .stTabs [role="tablist"]::after { content: none; }
.stTabs .stTabs [role="tab"] p { font-size: .95rem; }
.stTabs .stTabs [role="tab"] { color: #5B6B7A !important; border-radius: 8px 8px 0 0; padding: 8px 16px !important; }
.stTabs .stTabs [role="tab"]:hover { background: #EEF3F8; color: #0E2240 !important; }
.stTabs .stTabs [role="tab"][aria-selected="true"] { background: transparent !important; color: #00538C !important; box-shadow: inset 0 -3px 0 #00538C !important; }

/* player / team header card */
.ph { display: flex; justify-content: space-between; align-items: center; gap: 24px; flex-wrap: wrap;
      background: #FFFFFF; border: 1px solid #DCE3EA; border-radius: 14px; padding: 20px 24px; margin: 6px 0 4px; }
.ph-left { display: flex; align-items: center; gap: 18px; }
.avatar { width: 72px; height: 72px; border-radius: 50%; background: #E1ECF5; display: flex; align-items: center; justify-content: center;
          font-family: 'Barlow Condensed', sans-serif; font-weight: 600; font-size: 1.8rem; color: #5B6B7A; flex: none;
          position: relative; overflow: hidden; }
/* headshot over the initials; if it can't load (offline, unknown player) the initials underneath still show */
.avatar.logo { background: #FFFFFF; border: 1px solid #DCE3EA; }
.avatar.logo img { padding: 10px; object-fit: contain; object-position: center; box-sizing: border-box; }
.avatar img { position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; object-position: center top; }
/* browsers draw ::before only on an image that failed to load: cover the broken-image icon with the initials again */
.avatar img::before { content: attr(data-initials); position: absolute; inset: 0; display: flex; align-items: center; justify-content: center;
                      background: #E1ECF5; color: #5B6B7A; }
.ph-name { font-family: 'Barlow Condensed', sans-serif; font-size: 2.6rem; font-weight: 700; color: #0E2240; line-height: 1; }
.ph-meta { color: #5B6B7A; font-size: .92rem; margin-top: 6px; }
.ph-meta b { color: #0E2240; font-weight: 600; }
.chips { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 8px; }
.chip { display: inline-block; padding: 3px 10px; border-radius: 999px; font-size: .78rem; font-weight: 600; }
.scorebox { text-align: right; min-width: 220px; }
.scorebox .n { font-family: 'Barlow Condensed', sans-serif; font-size: 4.2rem; font-weight: 600; line-height: 1; color: #00538C; font-variant-numeric: tabular-nums; }
.scorebox .l { font-size: .72rem; color: #5B6B7A; text-transform: uppercase; letter-spacing: .07em; font-weight: 600; }

/* stat tiles */
.tiles { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 14px; margin: 6px 0 4px; }
.tile { background: #FFFFFF; border: 1px solid #DCE3EA; border-radius: 14px; padding: 14px 18px; }
.tile .l { font-size: .72rem; color: #5B6B7A; text-transform: uppercase; letter-spacing: .07em; font-weight: 600; }
.tile .n { font-family: 'Barlow Condensed', sans-serif; font-size: 2.4rem; font-weight: 600; line-height: 1.1; color: #0E2240; font-variant-numeric: tabular-nums; }
.tile .s { font-size: .82rem; color: #5B6B7A; }

.sec { font-family: 'Barlow Condensed', sans-serif; font-size: 1.45rem; font-weight: 600; color: #0E2240;
       border-bottom: 1px solid #DCE3EA; padding-bottom: 4px; margin: 24px 0 10px; }
.sec small { font-family: 'Work Sans', sans-serif; font-weight: 400; color: #5B6B7A; font-size: .8rem; margin-left: 8px; }
table.kv, table.grid { border-collapse: collapse; width: 100%; font-size: .88rem; font-variant-numeric: tabular-nums; }
table.kv td { padding: 7px 4px; border-bottom: 1px solid #E8EDF1; color: #0E2240; }
table.kv td:first-child { color: #5B6B7A; }
table.kv td:last-child { text-align: right; font-weight: 600; }
table.grid th { text-align: right; font-weight: 600; color: #5B6B7A; font-size: .72rem; text-transform: uppercase;
                letter-spacing: .05em; border-bottom: 1px solid #0E2240; padding: 6px 8px; }
table.grid th:first-child, table.grid td:first-child { text-align: left; }
table.grid td { text-align: right; padding: 7px 8px; border-bottom: 1px solid #E8EDF1; color: #0E2240; }
.pr { display: grid; grid-template-columns: 165px 1fr 72px; align-items: center; gap: 14px; padding: 8px 0;
      border-bottom: 1px solid #E8EDF1; font-size: .9rem; color: #0E2240; }
.pr .lab span { display: block; color: #5B6B7A; font-size: .74rem; }
.pr .track { position: relative; height: 8px; background: #E8EDF1; border-radius: 99px; margin: 0 13px; }
.pr .fill { position: absolute; left: 0; top: 0; height: 8px; border-radius: 99px; }
.pr .dot { position: absolute; top: 50%; transform: translate(-50%, -50%); width: 28px; height: 28px; border-radius: 50%;
           font-size: .74rem; font-weight: 700; display: flex; align-items: center; justify-content: center; border: 2px solid #FFFFFF; }
.pr .val { text-align: right; color: #5B6B7A; font-size: .8rem; font-variant-numeric: tabular-nums; }
.pr.off { opacity: .5; }
.pc { display: inline-block; min-width: 30px; text-align: center; padding: 1px 6px; font-weight: 600; border-radius: 4px; }
ul.recs { list-style: none; padding: 0; margin: 0; }
ul.recs li { display: grid; grid-template-columns: 86px 1fr; gap: 12px; padding: 10px 0; border-bottom: 1px solid #E8EDF1;
             font-size: .92rem; line-height: 1.55; color: #0E2240; }
ul.recs .tag { font-size: .68rem; font-weight: 700; text-transform: uppercase; letter-spacing: .06em; padding-top: 4px; }
ul.drv { margin: 4px 0 12px; padding-left: 18px; line-height: 1.7; font-size: .92rem; }
.eyebrow { font-size: .72rem; letter-spacing: .07em; text-transform: uppercase; font-weight: 700; margin: 0; }
.stApp .muted, .stApp p.muted { color: #5B6B7A; font-size: .82rem !important; line-height: 1.5; }
.scale span { display: inline-block; width: 18px; height: 9px; vertical-align: middle; border-radius: 2px; }
.check { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 8px; margin: 6px 0 10px; }
.check div { display: flex; align-items: center; gap: 8px; padding: 8px 12px; border-radius: 8px; font-size: .86rem; }
.check b { width: 14px; }
.check span { margin-left: auto; font-size: .74rem; }
/* override Streamlit's default markdown sizes / table borders for our HTML blocks */
.stApp p.eyebrow { font-size: .72rem !important; margin: 0 !important; }
.stApp ul.drv li, .stApp ul.recs li > span:last-child { font-size: .9rem !important; }
.stApp ul.recs .tag { font-size: .68rem !important; }
.stApp table.kv, .stApp table.grid { border: none !important; }
.stApp table.kv td, .stApp table.grid td, .stApp table.grid th { border-left: none !important; border-right: none !important; border-top: none !important; }
.stApp table.kv tr, .stApp table.grid tr { background: transparent !important; }
/* newer Streamlit draws a selected-tab indicator bar inside the tab list; the pill already shows selection */
.stTabs .react-aria-SelectionIndicator { display: none !important; }
</style>""", unsafe_allow_html=True)


# ---------------------------------------------------------------- small helpers
def e(x):
    return html.escape(str(x))


def md(s: str):                                   # escape $ so dollar amounts never render as LaTeX
    st.markdown(s.replace("$", "&#36;"), unsafe_allow_html=True)


def isnum(x) -> bool:
    try:
        return x is not None and not pd.isna(x)
    except (TypeError, ValueError):
        return False


def pct_color(p):
    """Diverging scale: burnt clay (bottom of league) -> warm sand (middle) -> dusk blue (top)."""
    if not isnum(p):
        return "#E8EDF1", MUTED
    p = min(max(float(p), 0.0), 1.0)
    a, b, t = (LOW_RGB, MID_RGB, p / .5) if p < .5 else (MID_RGB, HIGH_RGB, (p - .5) / .5)
    rgb = tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))
    return "#%02X%02X%02X" % rgb, ("white" if abs(p - .5) > .28 else INK)


def pctl(p) -> str:
    return "–" if not isnum(p) else str(min(99, max(1, int(round(float(p) * 100)))))


def pc(p) -> str:
    bg, fg = pct_color(p)
    return f'<span class="pc" style="background:{bg};color:{fg}">{pctl(p)}</span>'


def ordinal(n) -> str:
    n = int(n)
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def place(p) -> str:                              # "top 2%" for strong players, "bottom 3%" for weak ones
    if p >= 0.5:
        return f"top {max(1, int(round((1 - p) * 100 + 0.5))) if p < 1 else 1}%"
    return f"bottom {max(1, int(round(p * 100)))}%"


def money(v) -> str:
    if not isnum(v):
        return "–"
    v = float(v)
    return f"${v / 1e6:.1f}M" if abs(v) >= 1e5 else f"${v:,.0f}"


def big(v) -> str:
    if not isnum(v):
        return "–"
    v = float(v)
    return f"{v / 1e6:.1f}M" if v >= 1e6 else f"{v / 1e3:.0f}K" if v >= 1e3 else f"{v:.0f}"


def initials(name) -> str:
    parts = [p for p in re.split(r"[\s\-]+", str(name)) if p]
    return "".join(p[0] for p in parts[:2]).upper() or "?"


@st.cache_data
def photo_ids() -> dict:
    """Player name -> NBA.com person id (data/player_ids.csv), used for headshot URLs."""
    f = BUNDLED.parent / "player_ids.csv"
    return dict(pd.read_csv(f).itertuples(index=False)) if f.exists() else {}


def avatar(name) -> str:
    pid, ini = photo_ids().get(name), e(initials(name))
    img = (f'<img src="https://cdn.nba.com/headshots/nba/latest/260x190/{int(pid)}.png" alt="" data-initials="{ini}">'
           if pid else "")
    return f'<div class="avatar">{ini}{img}</div>'


# NBA.com team ids, used for logo URLs (cdn.nba.com/logos/nba/<id>/...)
TEAM_IDS = {"Atlanta Hawks": 1610612737, "Boston Celtics": 1610612738, "Brooklyn Nets": 1610612751, "Charlotte Hornets": 1610612766,
            "Chicago Bulls": 1610612741, "Cleveland Cavaliers": 1610612739, "Dallas Mavericks": 1610612742, "Denver Nuggets": 1610612743,
            "Detroit Pistons": 1610612765, "Golden State Warriors": 1610612744, "Houston Rockets": 1610612745, "Indiana Pacers": 1610612754,
            "LA Clippers": 1610612746, "Los Angeles Lakers": 1610612747, "Memphis Grizzlies": 1610612763, "Miami Heat": 1610612748,
            "Milwaukee Bucks": 1610612749, "Minnesota Timberwolves": 1610612750, "New Orleans Pelicans": 1610612740,
            "New York Knicks": 1610612752, "Oklahoma City Thunder": 1610612760, "Orlando Magic": 1610612753,
            "Philadelphia 76ers": 1610612755, "Phoenix Suns": 1610612756, "Portland Trail Blazers": 1610612757,
            "Sacramento Kings": 1610612758, "San Antonio Spurs": 1610612759, "Toronto Raptors": 1610612761, "Utah Jazz": 1610612762,
            "Washington Wizards": 1610612764}


def team_logo(team) -> str:
    tid, ini = TEAM_IDS.get(team), e(initials(team))
    img = f'<img src="https://cdn.nba.com/logos/nba/{tid}/primary/L/logo.svg" alt="" data-initials="{ini}">' if tid else ""
    return f'<div class="avatar logo">{ini}{img}</div>'


def sec(title, note=""):
    md(f'<div class="sec">{e(title)}{f"<small>{e(note)}</small>" if note else ""}</div>')


def tiles(items):
    """items: (label, value, sub, value_color or None)"""
    md('<div class="tiles">' + "".join(
        f'<div class="tile"><div class="l">{e(l)}</div><div class="n" style="color:{c or INK}">{v}</div><div class="s">{s}</div></div>'
        for l, v, s, c in items) + "</div>")


def kv_table(rows) -> str:
    return '<table class="kv">' + "".join(f"<tr><td>{k}</td><td>{v}</td></tr>" for k, v in rows) + "</table>"


def name_list(names, n=4) -> str:
    names = [e(x) for x in names]
    return ", ".join(names[:n]) + (f" and {len(names) - n} more" if len(names) > n else "")


def chart(fig, h, xt="", yt="", legend=False):
    fig.update_layout(template="simple_white", height=h, margin=dict(l=10, r=10, t=10, b=40), showlegend=legend,
                      font=dict(family=FONT, size=12, color=INK), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                      xaxis_title=xt, yaxis_title=yt)
    fig.update_xaxes(linecolor=SAND, tickcolor=SAND)
    fig.update_yaxes(linecolor=SAND, tickcolor=SAND)
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def radar(traces, h=400):
    """traces: (name, {factor: percentile 0-1}, color, dashed, filled). Plots the five scored factors as a profile shape."""
    labs = [dl.LABEL[f] for f in SCORED]
    fig = go.Figure()
    for name, vals, color, dashed, filled in traces:
        r = [float(vals[f]) * 100 if isnum(vals[f]) else 0 for f in SCORED]
        rgb = tuple(int(color[i:i + 2], 16) for i in (1, 3, 5))
        fig.add_trace(go.Scatterpolar(r=r + r[:1], theta=labs + labs[:1], name=name, mode="lines+markers",
                                      line=dict(color=color, width=2.5 if filled else 2, dash="dash" if dashed else "solid"),
                                      marker=dict(size=5, color=color), fill="toself" if filled else "none",
                                      fillcolor=f"rgba({rgb[0]},{rgb[1]},{rgb[2]},0.16)",
                                      hovertemplate="%{theta}: %{r:.0f}th percentile<extra>" + e(name) + "</extra>"))
    fig.update_layout(height=h, margin=dict(l=110, r=110, t=30, b=40), paper_bgcolor="rgba(0,0,0,0)",
                      font=dict(family=FONT, size=12, color=INK), showlegend=True,
                      legend=dict(orientation="h", yanchor="bottom", y=-0.16, xanchor="center", x=0.5),
                      polar=dict(bgcolor="rgba(0,0,0,0)",
                                 radialaxis=dict(range=[0, 100], tickvals=[25, 50, 75, 100], showticklabels=False, gridcolor=RULE, linecolor=RULE),
                                 angularaxis=dict(gridcolor=RULE, linecolor=SAND)))
    st.plotly_chart(fig, width="stretch", config={"displayModeBar": False})


def style_pct(v):
    if not isnum(v):
        return ""
    bg, fg = pct_color(v / 100)
    return f"background-color:{bg};color:{fg};font-weight:600"


# ---------------------------------------------------------------- data
@st.cache_data(show_spinner="Reading the workbook...")
def get_model(path: str, mtime: float):
    return dl.load(path)


REQUIRED_SHEETS = ["Control Panel", "Players", "Outcome", "Endorsement Leads"]
OPTIONAL_SHEETS = ["Regression", "Team Brand Health", "Data Dictionary", "Scores", "Draw Factor"]
PLAYER_COLS = ["Player", "Team", "Instagram", "TikTok", "X followers", "2K", "Sentiment", "Global Reach",
               "Availability", "Age", "Seasons", "Salary", "Data flags"]          # header prefixes data_layer.load() looks for


@st.cache_data
def workbook_check(path: str, mtime: float):
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets, head = list(wb.sheetnames), []
    if "Players" in sheets:
        first = next(wb["Players"].iter_rows(min_row=1, max_row=1, values_only=True), ())
        head = [str(h) for h in first if h]
    wb.close()
    return sheets, head


# the uploader lives in the Data tab; its value is in session_state from the previous run onward
up = st.session_state.get("upload")
path, source, load_error = BUNDLED, "Bundled dataset", None
if up is not None:
    raw = up.getvalue()
    path = Path(tempfile.gettempdir()) / f"mavs_upload_{hashlib.md5(raw).hexdigest()[:12]}.xlsx"
    if not path.exists():
        path.write_bytes(raw)
    source = up.name
try:
    if not path.exists():
        raise FileNotFoundError("bundled workbook not found in data/")
    model = get_model(str(path), path.stat().st_mtime)
except Exception as err:
    load_error = f"Could not read {source}: {err}"
    if up is not None and BUNDLED.exists():
        path, source = BUNDLED, "Bundled dataset"
        model = get_model(str(path), path.stat().st_mtime)
    else:
        st.error(load_error)
        st.stop()

with st.sidebar:
    st.markdown("**Model settings**")
    scen = st.selectbox("Weighting scenario", model.scenarios,
                        index=model.scenarios.index(model.active_scenario) if model.active_scenario in model.scenarios else 0,
                        help="Same choices as the Control Panel tab in Excel.")
    with st.expander("Thresholds"):
        hi = st.slider("Strong at or above (percentile)", 0.5, 0.9, 0.70, 0.05)
        lo = st.slider("Weak below (percentile)", 0.1, 0.6, 0.40, 0.05)
    st.caption("Upload a new season's workbook in the **Data** tab.")

mt = model.meta
d = dl.rescored(model, scen)
N = len(d)
names = sorted(d.player)
active_w = model.weights[scen]
W_TOT = active_w.sum()
SCORED = [f for f in dl.FACTORS[:5]]
END_MED = {f: d.loc[d.endorsed == 1, "s_" + f].median() for f in dl.FACTORS}
END_PCT = {f: d.loc[d.endorsed == 1, "p_" + f].median() for f in dl.FACTORS}          # typical endorsed player, as a percentile
TEAM_TBL = (d.groupby("team").agg(Players=("player", "count"), Avg=("score", "mean"), Endorsed=("endorsed", "sum"), Gap=("gap", "mean"))
            .sort_values("Avg", ascending=False))
TEAM_TBL["Rank"] = np.arange(1, len(TEAM_TBL) + 1)
_top = d.sort_values("score").groupby("team").tail(1).set_index("team")
TEAM_TBL["Top player"] = _top["player"]
TEAM_TBL["Top score"] = _top["score"]
DEFAULT_PLAYER = "Cooper Flagg" if "Cooper Flagg" in names else names[0]


def kind(p):
    return "Strength" if p >= hi else "Risk" if p < lo else "Note"


def tier(p):
    return ("Elite", ACCENT_SOFT, "#003A63") if p >= .9 else ("Strong", "#E1ECF5", POS) if p >= hi else \
           ("Mid-pack", "#E8EDF1", MUTED) if p >= lo else ("Developing", "#F8E4DC", NEG)


TAGC = {"Strength": POS, "Risk": NEG, "Action": INK, "Note": NOTE, "Priority": ACCENT}


def show_recs(items):
    if not items:
        md('<p class="muted">Nothing stands out.</p>')
        return
    lis = "".join(f'<li><span class="tag" style="color:{TAGC[k]}">{k}</span><span>{re.sub(r"[*][*](.+?)[*][*]", r"<b>\1</b>", t)}</span></li>'
                  for k, t in items)
    md(f'<ul class="recs">{lis}</ul>')


def css_str(s) -> str:
    return '"' + str(s).replace("\\", "\\\\").replace('"', '\\"').replace("<", "\\3C ") + '"'


# brand text for the top band (drawn into the main tab list, see the "top bar" CSS)
md(f'''<style>
.stTabs [role="tablist"]::before {{ content: "Marketability Index"; }}
.stTabs [role="tablist"]::after {{ content: {css_str(f"Dallas Mavericks · {N} players")}; }}
</style>''')


# ---------------------------------------------------------------- player pieces
def raw_value(row, f) -> str:
    g = row.get
    try:
        if f == "social":
            fol = g("followers")
            if not isnum(fol):
                fol = sum(float(g(c, 0) or 0) for c in ("instagram", "tiktok", "x"))
            return big(fol)
        if f == "talent":
            return f"{g('rating_2k'):.0f} 2K" if isnum(g("rating_2k")) else "–"
        if f == "sentiment":
            return f"{g('sentiment_score'):.2f}" if isnum(g("sentiment_score")) else "–"
        if f == "global_reach":
            return f"{g('global_reach_score'):.2f}" if isnum(g("global_reach_score")) else "–"
        if f == "availability":
            v = g("availability_raw")
            return "–" if not isnum(v) else f"{v:.0%} GP" if v <= 1.5 else f"{v:.0f} GP"
        if f == "brand_health":
            return f"${g('franchise_value'):.1f}B" if isnum(g("franchise_value")) else "–"
    except (TypeError, ValueError):
        pass
    return "–"


def pct_rows(row) -> str:
    out = ""
    for f in dl.FACTORS:
        p, w = row["p_" + f], active_w[f]
        bg, fg = pct_color(p)
        x = min(max(float(p), 0), 1) * 100 if isnum(p) else 0
        sub = f"{w / W_TOT * 100:.0f}% of score" if w > 0 else "context only, not scored"
        out += (f'<div class="pr{"" if w > 0 else " off"}" title="{e(dl.EXPLAIN[f])}">'
                f'<div class="lab">{e(dl.LABEL[f])}<span>{sub}</span></div>'
                f'<div class="track"><div class="fill" style="width:{x:.0f}%;background:{bg}"></div>'
                f'<div class="dot" style="left:{x:.0f}%;background:{bg};color:{fg}">{pctl(p)}</div></div>'
                f'<div class="val">{raw_value(row, f)}</div></div>')
    return out


def score_hist(row):
    fig = go.Figure(go.Histogram(x=d.score, nbinsx=30, marker_color=SAND, marker_line_color=GROUND, marker_line_width=1))
    fig.add_vline(x=row["score"], line_color=ACCENT, line_width=2.5, annotation_text=f"{row['player']}  {row['score']:.1f}",
                  annotation_position="top right", annotation_font_size=11, annotation_font_color=ACCENT)
    chart(fig, 250, f"Marketability score, all {N} players", "Players")


def pay_scatter(row):
    fig = go.Figure(go.Scatter(x=d.pay_pct, y=d.score_pct, mode="markers", marker=dict(size=5, color="#C3CDD6"),
                               text=d.player, hovertemplate="%{text}<extra></extra>"))
    fig.add_trace(go.Scatter(x=[0, 1], y=[0, 1], mode="lines", line=dict(color=MUTED, dash="dot", width=1), hoverinfo="skip"))
    if isnum(row["pay_pct"]):
        fig.add_trace(go.Scatter(x=[row["pay_pct"]], y=[row["score_pct"]], mode="markers+text", text=[row["player"]],
                                 textposition="top center", marker=dict(size=12, color=ACCENT, line=dict(color=SURFACE, width=2))))
    fig.update_xaxes(range=[0, 1.02], tickformat=".0%")
    fig.update_yaxes(range=[0, 1.06], tickformat=".0%")
    chart(fig, 250, "Salary percentile", "Marketability percentile")


def drivers(row) -> str:
    ranked = sorted([f for f in SCORED if active_w[f] > 0 and isnum(row["p_" + f])], key=lambda f: -row["p_" + f])
    li = lambda fs: "".join(f"<li>{e(dl.LABEL[f])}: {ordinal(pctl(row['p_' + f]))} percentile ({place(row['p_' + f])})</li>" for f in fs)
    return (f'<p class="eyebrow" style="color:{POS}">Strengths</p><ul class="drv">{li(ranked[:2])}</ul>'
            f'<p class="eyebrow" style="color:{NEG}">Holding back</p><ul class="drv">{li(ranked[::-1][:2])}</ul>')


CATEGORY = {"social": "digital-first brands and creator-style campaigns (apparel drops, beverages, apps) that pay for reach",
            "global_reach": "brands with an international footprint: global apparel, airlines, consumer tech, overseas launches",
            "talent": "performance categories where on-court credibility sells: footwear, sports nutrition, training gear, video games",
            "sentiment": "trust-sensitive categories: banks, insurance, healthcare, family brands, community programs",
            "availability": "long-term ambassador roles that need him on the floor and in the market for years"}
ACTION = {"social": "Grow and diversify his audience (the sheet counts Instagram, TikTok and X).",
          "global_reach": "Build an international audience: foreign-language content, overseas tours, regional partners.",
          "talent": "On-court production is the base for everything else.",
          "sentiment": "A more positive, consistent media presence: interviews, community work.",
          "availability": "Stay on the court (health, load management) and lock in long-term deals."}


def recs_team(row):
    out, g, age = [], row["gap"], row["age"]
    if isnum(g):
        if g >= 0.25:
            out.append(("Strength", f"**Good value.** His brand ranks well above his pay (marketability {row['score_pct']:.0%} vs salary {row['pay_pct']:.0%} percentile)."))
        elif g <= -0.25:
            out.append(("Risk", f"**Pay ahead of brand.** Salary {row['pay_pct']:.0%} vs marketability {row['score_pct']:.0%} percentile. The contract case rests on performance, not marketing."))
        else:
            out.append(("Note", f"**Pay in line with brand** (salary {row['pay_pct']:.0%} vs marketability {row['score_pct']:.0%} percentile)."))
    if row["p_games"] < lo:
        out.append(("Risk", "**Availability.** He has missed a lot of games recently, which limits appearances and ticket-driven value."))
    elif row["p_games"] >= hi:
        out.append(("Strength", "**Reliable.** Among the most available players, so campaigns can be built around him."))
    if isnum(age):
        if row["p_runway"] < lo:
            out.append(("Risk", f"**Short runway.** At {int(age)} his contract and endorsement window are closer to ending. Plan for succession."))
        elif row["p_runway"] >= hi:
            out.append(("Strength", f"**Long runway.** At {int(age)} he has many seasons ahead, so multi-year investment makes sense."))
    if row["p_sentiment"] < lo:
        out.append(("Risk", "**Reputation watch.** News coverage is less positive or sparser than most players'."))
    best = max(SCORED, key=lambda f: row["p_" + f])
    if row["p_" + best] >= lo:
        out.append(("Action", f"**Lead campaigns with his {dl.LABEL[best].lower()}**, his strongest asset ({place(row['p_' + best])})."))
    else:
        out.append(("Risk", f"**No standout strength yet.** Even his best factor ({dl.LABEL[best].lower()}) is in the {place(row['p_' + best])}."))
    if isnum(row["metro_pop"]):
        out.append(("Note", f"**Market context.** {e(row['team'])} play in a metro of {row['metro_pop']:.1f}M and are valued at ${row['franchise_value']:.1f}B "
                            f"({place(row['p_brand_health'])}). Market size did not predict endorsements, so it is not scored."))
    return out


def recs_player(row):
    out = []
    for f in sorted(SCORED, key=lambda f: row["p_" + f])[:2]:
        out.append(("Action", f"**Improve {dl.LABEL[f].lower()}** ({place(row['p_' + f])}). Typical players with a qualifying endorsement score {END_MED[f]:.2f}; he scores {row['s_' + f]:.2f}. {ACTION[f]}"))
    best = max(SCORED, key=lambda f: row["p_" + f])
    if row["p_" + best] >= lo:
        out.append(("Strength", f"**Strongest asset: {dl.LABEL[best].lower()}** ({place(row['p_' + best])})."))
    if row["tiktok"] == 0 and row["instagram"] > 1e6:
        out.append(("Note", "TikTok shows 0 followers despite a large Instagram audience: either no account or not yet counted."))
    if row["x"] == 0:
        out.append(("Note", "No X followers are counted; any audience there is missing from the score."))
    return out


def recs_agent(row, sim):
    out, g, age, sp = [], row["gap"], row["age"], row["score_pct"]
    pe = row.get("p_endorse", np.nan)
    # 1. overall priority
    if row["endorsed"] == 0 and isnum(pe) and pe >= 0.5:
        out.append(("Priority", f"**Pursue now.** His profile matches players who already have deals (model likelihood {pe:.0%}), but none is recorded."))
    elif sp >= hi:
        out.append(("Priority", f"**Premium client.** Marketability in the {place(sp)} of the league: protect the brand and aim for flagship categories."))
    elif sp < lo and (not isnum(pe) or pe < 0.3):
        out.append(("Priority", "**Develop before pitching.** The numbers don't support a major deal yet; build audience and performance first."))
    else:
        out.append(("Priority", f"**Build.** Mid-pack profile ({place(sp)}): target regional and category-specific deals before national flagship ones."))
    # 2. category fit from his two strongest scored factors
    for f in [f for f in sorted(SCORED, key=lambda f: -row["p_" + f])[:2] if row["p_" + f] >= lo]:
        out.append(("Action", f"**Pitch {CATEGORY[f]}.** Fits his {dl.LABEL[f].lower()} ({place(row['p_' + f])})."))
    # 3. comparables
    if sim is not None and len(sim):
        en = sim.loc[sim.endorsed == 1, "player"].tolist()
        if en:
            out.append(("Note", f"**Comparables.** {len(en)} of his {len(sim)} closest profiles have deals ({name_list(en)}). Use their terms as benchmarks."))
        else:
            out.append(("Risk", f"**No comparables with qualifying deals.** None of his {len(sim)} closest profiles has a qualifying endorsement, so the pitch rests on his own numbers."))
    if row["endorsed"] == 1 and isnum(pe):
        if pe < 0.3:
            out.append(("Strength", f"**Over-delivers.** Qualifying endorsement although the model likelihood is only {pe:.0%}: evidence he sells beyond his numbers. Use it in renewals."))
        else:
            out.append(("Note", f"Qualifying endorsement, as his profile suggests (model likelihood {pe:.0%})."))
    elif row["endorsed"] == 0 and isnum(pe) and pe < 0.5:
        out.append(("Note", f"No qualifying endorsement; model likelihood {pe:.0%}."))
    # 4. biggest lever: points gained by lifting a factor to the typical endorsed player's level
    gains = {f: max(0.0, END_MED[f] - row["s_" + f]) * active_w[f] / W_TOT * 100 for f in SCORED if active_w[f] > 0 and isnum(END_MED[f])}
    if gains:
        bf = max(gains, key=gains.get)
        if gains[bf] >= 0.5:
            out.append(("Action", f"**Biggest lever: {dl.LABEL[bf].lower()}.** Reaching the level of a typical player with a qualifying endorsement would add about {gains[bf]:.1f} points to his score. {ACTION[bf]}"))
    # 5. deal structure
    if isnum(age):
        if row["p_runway"] >= hi:
            out.append(("Action", f"**Structure.** At {int(age)} he can anchor multi-year deals. Ask for escalators tied to performance milestones and consider equity for part of the fee."))
        elif row["p_runway"] < lo:
            out.append(("Action", f"**Structure.** At {int(age)} the window is narrowing. Favour shorter deals with more guaranteed money up front, and start on post-career licensing and media."))
    if row["p_games"] < lo:
        out.append(("Risk", "**Availability clauses.** He has missed many games, so brands will ask for games-played or appearance terms. Push for a guaranteed base and put the upside in bonuses."))
    # 6. mismatches between reach and play
    if row["p_talent"] >= hi and row["p_social"] < lo:
        out.append(("Action", "**Audience lags his play.** Invest in content and a social team before the next deal cycle; his on-court value isn't reaching sponsors yet."))
    if row["p_social"] >= hi and row["p_talent"] < lo:
        out.append(("Note", "**Reach without results.** His audience outpaces his production, so lifestyle and creator deals fit better than performance brands."))
    if row["p_sentiment"] < lo:
        out.append(("Risk", "**Reputation.** Media tone is weaker than most players'. Run a media plan before pitching trust-sensitive brands and expect tighter morality clauses."))
    # 7. home market
    mp = row["metro_pop"]
    if isnum(mp):
        q1, q3 = d.metro_pop.quantile(0.25), d.metro_pop.quantile(0.75)
        if mp <= q1 and sp >= hi:
            out.append(("Note", f"**Small home market** ({mp:.1f}M). Local deals cap out quickly, so lead with national brands."))
        elif mp >= q3:
            out.append(("Note", f"**Large home market** ({mp:.1f}M). Stack local partnerships (dealerships, banks, restaurants) on top of national deals."))
    # 8. contract angle, data quality
    if isnum(g):
        out.append(("Strength" if g >= 0.25 else "Risk" if g <= -0.25 else "Note",
                    f"**Contract angle.** Brand minus pay is {g:+.2f}. " + ("Argue the brand justifies more." if g >= 0.25
                    else "Expect pushback: pay already outranks the brand." if g <= -0.25 else "Pay roughly matches the brand.")))
    missing = [n for n, c in (("TikTok", "tiktok"), ("X", "x")) if row[c] == 0]
    if missing:
        out.append(("Note", f"**Fix the data first.** No {' or '.join(missing)} followers are counted, so his reach may be understated. Update the counts before using this score in talks."))
    if str(row["endorse_source"]).startswith("Lead"):
        out.append(("Risk", "**Unverified endorsement.** It comes from a web-search lead nobody has confirmed. Verify it before using it in a negotiation."))
    return out


def sponsor_names(s) -> str:                      # lead text mixes names with research notes: keep the names only
    parts = (p.strip() for p in re.sub(r"\s*\([^)]*\)", "", str(s)).split(";"))
    return "; ".join(p for p in parts if p and not re.search(r"no shoe deal|aggregator|without a reported", p))


# ---------------------------------------------------------------- tabs
tab_p, tab_t, tab_l, tab_c, tab_s, tab_d, tab_m = st.tabs(["Player", "Team", "Leaderboard", "Compare", "Similar players", "Data", "Method"])
if load_error:   # after the tabs, so the band stays at the very top of the page
    st.error(load_error + " Showing the bundled dataset instead. Check the Data tab for which sheets or columns are missing.")

with tab_p:
    pick = st.selectbox("Player", names, index=names.index(DEFAULT_PLAYER), key="pick",
                        label_visibility="collapsed", placeholder="Search for a player")
    row = d.loc[pick]
    age_txt = f"Age {int(row['age'])}" if isnum(row["age"]) else "Age n/a"
    t_name, t_bg, t_fg = tier(row["score_pct"])
    if row["endorsed"]:
        en_chip = ("Qualifying endorsement · unverified lead", "#E8EDF1", MUTED) if str(row["endorse_source"]).startswith("Lead") else ("Qualifying endorsement", "#E1ECF5", POS)
    else:
        en_chip = ("No qualifying endorsement", "#E8EDF1", MUTED)
    md(f'''<div class="ph"><div class="ph-left">{avatar(row["player"])}<div>
        <div class="ph-name">{e(row["player"])}</div>
        <div class="ph-meta"><b>{e(row["team"])}</b> · {age_txt} · Salary {money(row["salary"])}</div>
        <div class="chips"><span class="chip" style="background:{t_bg};color:{t_fg}">Tier: {t_name}</span>
        <span class="chip" style="background:{en_chip[1]};color:{en_chip[2]}">{e(en_chip[0])}</span></div></div></div>
        <div class="scorebox"><div class="l">Marketability · {e(scen)}</div><div class="n">{row["score"]:.1f}</div>
        <div class="l">Rank {row["rank"]} of {N} · {ordinal(pctl(row["score_pct"]))} percentile</div></div></div>''')

    k = int(st.session_state.get("k", 5))
    sim = knn.similar_players(d, pick, k=k)

    view_profile, view_recs = st.tabs(["Profile", "Recommendations"])
    with view_profile:
        left, right = st.columns([1, 1.15], gap="large")
        with left:
            sec("Profile shape", "league percentile on each scored factor")
            radar([(pick, {f: row["p_" + f] for f in SCORED}, ACCENT, False, True),
                   ("Typical player with a qualifying endorsement", END_PCT, NOTE, True, False)])
        with right:
            sec("Percentile rankings", f"vs all {N} players")
            md(pct_rows(row))
            md('<p class="muted" style="margin-top:6px"><span class="scale">'
               + "".join(f'<span style="background:{pct_color(v)[0]}"></span>' for v in (0.02, 0.2, 0.5, 0.8, 0.98))
               + "</span>&nbsp; 1 = bottom of the league, 99 = top. Hover a row for what it measures.</p>")

        c1, c2 = st.columns([1.15, 1], gap="large")
        with c1:
            sec("Summary")
            pe = row.get("p_endorse", np.nan)
            rows = [(f"Score · {e(s)}", f"{row['score_' + s]:.1f}") for s in model.scenarios if "score_" + s in row.index]
            rows += [("Qualifying endorsement likelihood (model)", f"{pe:.0%}" if isnum(pe) else "–"),
                     ("Salary percentile", pc(row["pay_pct"])),
                     ("Brand minus pay", f"{row['gap']:+.2f}" if isnum(row["gap"]) else "–"),
                     ("Games played", pc(row["p_games"])),
                     ("Career runway", pc(row["p_runway"])),
                     ("Road draw factor (not in score)", f"{row['draw_factor']:.3f} · {row['road_games']:.0f} road games" if isnum(row["draw_factor"]) else "–"),
                     ("Draw vs own team (not in score)", f"{row['draw_team_adj']:+.3f}" if isnum(row["draw_team_adj"]) else "–"),
                     ("Sponsors on record", e(row["sponsors_listed"]) if row["sponsors_listed"] else "None")]
            if row["lead_sponsors"]:
                links = " ".join(f'<a href="{e(u.strip())}" target="_blank">[{i + 1}]</a>' for i, u in enumerate(str(row["lead_links"]).split("|")) if u.strip().startswith("http"))
                rows.append((f"Web lead ({e(row['lead_tier'])} conf., unverified)", f"{e(row['lead_sponsors'])} {links}"))
            md(kv_table(rows))
        with c2:
            sec("What drives the score")
            md(drivers(row))

        c1, c2 = st.columns(2, gap="large")
        with c1:
            sec("Where he sits in the league")
            score_hist(row)
        with c2:
            sec("Brand vs pay", "above the dotted line = brand outranks salary")
            pay_scatter(row)

        sec("Comparable players", "closest profiles across the five scored factors")
        st.slider("Number of comparables", 3, 10, 5, key="k")
        st.caption(f"{sim.endorsed.mean():.0%} of these {len(sim)} players have a qualifying endorsement. For look-alikes matched on performance only, see the Similar players tab.")
        show = sim[["player", "team", "similarity", "score", "rank", "endorsed", "followers", "rating_2k"]].rename(
            columns={"player": "Player", "team": "Team", "similarity": "Similarity", "score": "Score", "rank": "Rk",
                     "endorsed": "Qualifying endorsement", "followers": "Followers", "rating_2k": "2K"})
        st.dataframe(show, hide_index=True, width="stretch",
                     column_config={"Similarity": st.column_config.NumberColumn(format="%.0f%%"), "Score": st.column_config.NumberColumn(format="%.1f"),
                                    "Qualifying endorsement": st.column_config.CheckboxColumn(), "Followers": st.column_config.NumberColumn(format="%d")})

        with st.expander("Raw inputs"):
            raw = pd.DataFrame({"Instagram": [row["instagram"]], "TikTok": [row["tiktok"]], "X": [row["x"]], "2K rating": [row["rating_2k"]],
                                "Sentiment": [row["sentiment_score"]], "Global reach": [row["global_reach_score"]], "Availability (games/82)": [row["availability_raw"]],
                                "Age": [row["age"]], "Games-played score": [row["s_games"]], "Runway score": [row["s_runway"]], "Team metro (M)": [row["metro_pop"]],
                                "Franchise value ($B)": [row["franchise_value"]], "Salary": [row["salary"]]})
            st.dataframe(raw, hide_index=True, width="stretch")
            if row["flags"]:
                st.warning(f"Data flags: {row['flags']}")

    with view_recs:
        left, right = st.columns([1.7, 1], gap="large")
        with left:
            ra, rt, rp = st.tabs(["For agents", "For the team", "For the player"])
            with ra:
                show_recs(recs_agent(row, sim))
            with rt:
                show_recs(recs_team(row))
            with rp:
                show_recs(recs_player(row))
        with right:
            sec("At a glance")
            pe = row.get("p_endorse", np.nan)
            md(kv_table([("Marketability", f"{row['score']:.1f} · {ordinal(pctl(row['score_pct']))} pct"),
                         ("Qualifying endorsement likelihood", f"{pe:.0%}" if isnum(pe) else "–"),
                         ("Brand minus pay", f"{row['gap']:+.2f}" if isnum(row["gap"]) else "–"),
                         ("Comparables with a qualifying endorsement", f"{int(sim.endorsed.sum())} of {len(sim)}")]))
            md(drivers(row))
            st.caption("Recommendations are rule-based on this player's factor percentiles, the weighting scenario and the thresholds in the sidebar.")


# ---------------------------------------------------------------- team view
def team_recs(tm, roster):
    out, tr = [], TEAM_TBL.loc[tm]
    out.append(("Note", f"**League standing.** Average marketability {tr['Avg']:.1f}, {ordinal(tr['Rank'])} of {len(TEAM_TBL)} teams."))
    face = roster.iloc[0]
    out.append(("Strength", f"**Face of the franchise: {e(face['player'])}** ({face['score']:.1f}, {place(face['score_pct'])} of the league). Build flagship campaigns around him."))
    under = roster[roster.gap >= 0.25].player.tolist()
    if under:
        out.append(("Strength", f"**Brand ahead of pay:** {name_list(under)}. Marketing value the payroll isn't reflecting."))
    over = roster[roster.gap <= -0.25].player.tolist()
    if over:
        out.append(("Risk", f"**Pay ahead of brand:** {name_list(over)}. These contracts rest on performance, not marketing."))
    if "p_endorse" in roster.columns:
        tgt = roster[(roster.endorsed == 0) & (roster.p_endorse >= 0.5)].player.tolist()
        if tgt:
            out.append(("Action", f"**Sponsor outreach:** {name_list(tgt)} look like players with a qualifying endorsement but have none on record. Pitch them to team partners."))
    core = roster.head(5)
    hurt = core[core.p_games < lo].player.tolist()
    if hurt:
        out.append(("Risk", f"**Availability among top names:** {name_list(hurt)} missed many games. Don't build every campaign around one player."))
    old = core[core.p_runway < lo].player.tolist()
    if old:
        out.append(("Action", f"**Succession:** {name_list(old)} are late in their careers. Start featuring the next face of the team now."))
    young = roster[(roster.p_runway >= hi) & (roster.score_pct >= 0.5)].player.tolist()
    if young:
        out.append(("Strength", f"**Long-term assets:** {name_list(young)} combine above-average marketability with a long runway."))
    leads = roster[roster.endorse_source.astype(str).str.startswith("Lead")].player.tolist()
    if leads:
        out.append(("Note", f"**Unverified endorsements:** {name_list(leads)}. Confirm before using in partner pitches."))
    return out


with tab_t:
    teams = list(TEAM_TBL.sort_index().index)
    default_team = d.loc[DEFAULT_PLAYER, "team"] if DEFAULT_PLAYER in d.index else next((t for t in teams if "Dallas" in str(t)), teams[0])
    tm = st.selectbox("Team", teams, index=teams.index(default_team) if default_team in teams else 0, key="team", label_visibility="collapsed")
    roster = d[d.team == tm].sort_values("score", ascending=False)
    tr = TEAM_TBL.loc[tm]
    payroll = roster.salary.sum(min_count=1)
    md(f'''<div class="ph"><div class="ph-left">{team_logo(tm)}<div><div class="ph-name">{e(tm)}</div>
        <div class="ph-meta">{len(roster)} players · {int(tr["Endorsed"])} qualifying endorsements · Payroll {money(payroll)}
        {f" · Metro {roster.metro_pop.iloc[0]:.1f}M · Valued at ${roster.franchise_value.iloc[0]:.1f}B" if isnum(roster.metro_pop.iloc[0]) else ""}</div></div></div>
        <div class="scorebox"><div class="l">Avg marketability · {e(scen)}</div><div class="n">{tr["Avg"]:.1f}</div>
        <div class="l">{ordinal(tr["Rank"])} of {len(TEAM_TBL)} teams</div></div></div>''')

    face = roster.iloc[0]
    tiles([("Most marketable", e(face["player"]), f"{face['score']:.1f} · rank {face['rank']} in league", ACCENT),
           ("League median score", f"{d.score.median():.1f}", f"team avg {tr['Avg']:.1f}", None),
           ("Qualifying endorsements", f"{int(tr['Endorsed'])}", f"of {len(roster)} on the roster", None),
           ("Avg brand − pay", f"{tr['Gap']:+.2f}" if isnum(tr["Gap"]) else "–", "positive = brand outranks salaries", None)])

    left, right = st.columns([1.35, 1], gap="large")
    with left:
        sec("Roster by marketability", "dotted line = league median")
        rr = roster.iloc[::-1]
        fig = go.Figure(go.Bar(y=rr.player, x=rr.score, orientation="h", marker_color=[pct_color(p)[0] for p in rr.score_pct],
                               marker_line_color=SAND, marker_line_width=0.5, text=[f"{v:.1f}" for v in rr.score], textposition="outside",
                               hovertemplate="%{y}: %{x:.1f}<extra></extra>"))
        fig.add_vline(x=d.score.median(), line_color=MUTED, line_dash="dot", line_width=1)
        fig.update_xaxes(range=[0, max(100, rr.score.max() * 1.1)])
        chart(fig, max(260, 26 * len(rr) + 60), "Marketability score")
    with right:
        sec("Takeaways")
        show_recs(team_recs(tm, roster))

    sec("Roster", "factor columns are league percentiles")
    cols = {"player": "Player", "age": "Age", "score": "Score", "rank": "Rk", **{"p_" + f: dl.LABEL[f] for f in dl.FACTORS},
            "endorsed": "Qualifying endorsement", "gap": "Brand − pay", "salary": "Salary"}
    if "p_endorse" in roster.columns:
        cols["p_endorse"] = "Qualifying endorsement likelihood"
    rt_ = roster[list(cols)].rename(columns=cols)
    for f in dl.FACTORS:
        rt_[dl.LABEL[f]] = (rt_[dl.LABEL[f]] * 100).round(0)
    rt_["Qualifying endorsement"] = rt_["Qualifying endorsement"].map({1: "Yes", 0: "", True: "Yes", False: ""})
    fmt = {"Age": "{:.0f}", "Score": "{:.1f}", "Brand − pay": "{:+.2f}", "Salary": lambda v: money(v), **{dl.LABEL[f]: "{:.0f}" for f in dl.FACTORS}}
    if "Qualifying endorsement likelihood" in rt_.columns:
        fmt["Qualifying endorsement likelihood"] = "{:.0%}"
    st.dataframe(rt_.style.map(style_pct, subset=[dl.LABEL[f] for f in dl.FACTORS]).format(fmt, na_rep="–"),
                 hide_index=True, width="stretch", height=min(640, 36 * len(rt_) + 40))
    st.download_button("Download roster (CSV)", rt_.to_csv(index=False).encode(), f"{tm}_marketability.csv".replace(" ", "_"), "text/csv")

    sec("All teams", "ranked by average player marketability")
    lt = TEAM_TBL.reset_index().rename(columns={"team": "Team", "Avg": "Avg score", "Gap": "Avg brand − pay", "Endorsed": "Qualifying endorsements"})
    lt = lt[["Rank", "Team", "Players", "Avg score", "Top player", "Top score", "Qualifying endorsements", "Avg brand − pay"]]
    hl = lambda r: [f"background-color:{ACCENT_SOFT};font-weight:600" if r["Team"] == tm else "" for _ in r]
    st.dataframe(lt.style.apply(hl, axis=1).format({"Avg score": "{:.1f}", "Top score": "{:.1f}", "Avg brand − pay": "{:+.2f}", "Qualifying endorsements": "{:.0f}"}, na_rep="–"),
                 hide_index=True, width="stretch", height=min(1100, 36 * len(lt) + 40))


# ---------------------------------------------------------------- leaderboard
with tab_l:
    RANK_BY = {"Marketability score": "score", **{dl.LABEL[f]: "p_" + f for f in dl.FACTORS}}
    if "p_endorse" in d.columns:
        RANK_BY["Qualifying endorsement likelihood"] = "p_endorse"
    RANK_BY["Brand − pay"] = "gap"
    f0, f1, f2 = st.columns([1.4, 1, 0.8])
    by = f0.selectbox("Rank by", list(RANK_BY), key="lb_by")
    order = f1.selectbox("Order", ["Highest first", "Lowest first"], key="lb_order")
    n = f2.number_input("Show top", 5, N, min(25, N), 5)
    g1, g2, g3 = st.columns([2, 1, 1])
    tms = g1.multiselect("Teams", sorted(d.team.unique()))
    en = g2.selectbox("Qualifying endorsement", ["All", "Qualifying", "Not qualifying"])
    q = g3.text_input("Name contains")
    col = RANK_BY[by]
    t = d.copy()
    if tms:
        t = t[t.team.isin(tms)]
    if en != "All":
        t = t[t.endorsed == (1 if en == "Qualifying" else 0)]
    if q:
        t = t[t.player.str.contains(q, case=False, regex=False)]
    t = t.sort_values([col, "score"], ascending=[order == "Lowest first", False], na_position="last").head(int(n))
    board = pd.DataFrame({"#": np.arange(1, len(t) + 1), "Player": t.player, "Team": t.team})
    if col not in ("score",):
        board[by] = (t[col] * 100).round(0) if col.startswith("p_") else t[col]
    board["Score"] = t.score
    board["Overall rk"] = t["rank"]
    for f in dl.FACTORS:
        if dl.LABEL[f] != by:
            board[dl.LABEL[f]] = (t["p_" + f] * 100).round(0)
    board["Qualifying endorsement"] = t.endorsed.map({1: "Yes", 0: "", True: "Yes", False: ""})
    board["Sponsors"] = t.sponsors_listed.where(t.sponsors_listed != "", t.lead_sponsors.map(sponsor_names))
    if by != "Brand − pay":
        board["Brand − pay"] = t.gap
    board["Draw factor"] = t.draw_factor
    board["Draw vs team"] = t.draw_team_adj
    board["Road games"] = t.road_games
    pct_cols = [c for c in board.columns if c in [dl.LABEL[f] for f in dl.FACTORS]]
    key_style = lambda s: [f"background-color:{ACCENT_SOFT};color:{INK};font-weight:700" for _ in s]
    sty = board.style.map(style_pct, subset=pct_cols).apply(key_style, subset=[by if by in board.columns else "Score"])
    fmt = {"Score": "{:.1f}", "Brand − pay": "{:+.2f}", "Draw factor": "{:.3f}", "Draw vs team": "{:+.3f}", "Road games": "{:.0f}", **{c: "{:.0f}" for c in pct_cols}}
    if by in board.columns and by != "Brand − pay":
        fmt[by] = "{:.0f}" if col.startswith("p_") and col != "p_endorse" else "{:.1f}"
    if by == "Qualifying endorsement likelihood":
        fmt[by] = "{:.0f}%"
    st.dataframe(sty.format(fmt, na_rep="–"), hide_index=True, width="stretch", height=640)
    st.download_button("Download this table (CSV)", board.to_csv(index=False).encode(), "mavs_leaderboard.csv", "text/csv")
    st.caption("Highlighted column = what the table is ranked by. Factor columns are league percentiles (99 = best). Draw columns are context only and do not affect the score or rank.")


# ---------------------------------------------------------------- compare
with tab_c:
    c1, c2 = st.columns(2)
    pa = c1.selectbox("Player 1", names, index=names.index(DEFAULT_PLAYER), key="cmpA")
    pb_default = "Luka Dončić" if "Luka Dončić" in names else names[min(1, len(names) - 1)]
    pb = c2.selectbox("Player 2", names, index=names.index(pb_default), key="cmpB")
    ra_, rb_ = d.loc[pa], d.loc[pb]
    lead_a = sum(ra_["p_" + f] > rb_["p_" + f] for f in SCORED)
    lead_b = sum(rb_["p_" + f] > ra_["p_" + f] for f in SCORED)
    tiles([(pa, f"{ra_['score']:.1f}", f"{e(ra_['team'])} · rank {ra_['rank']} · leads {lead_a} of {len(SCORED)} factors", ACCENT),
           (pb, f"{rb_['score']:.1f}", f"{e(rb_['team'])} · rank {rb_['rank']} · leads {lead_b} of {len(SCORED)} factors", NOTE)])
    left, right = st.columns([1.1, 1], gap="large")
    with left:
        sec("Profile shapes", "league percentiles")
        radar([(pa, {f: ra_["p_" + f] for f in SCORED}, ACCENT, False, True),
               (pb, {f: rb_["p_" + f] for f in SCORED}, NOTE, True, False)], h=430)
    with right:
        sec("Head to head", "percentiles vs the league")
        body = "".join(f"<tr><td>{e(dl.LABEL[f])}</td><td>{pc(ra_['p_' + f])}</td><td>{pc(rb_['p_' + f])}</td></tr>" for f in dl.FACTORS)
        body += (f"<tr><td><b>Marketability score</b></td><td><b>{ra_['score']:.1f}</b></td><td><b>{rb_['score']:.1f}</b></td></tr>"
                 f"<tr><td>Rank</td><td>{ra_['rank']}</td><td>{rb_['rank']}</td></tr>"
                 f"<tr><td>Age</td><td>{int(ra_['age']) if isnum(ra_['age']) else '–'}</td><td>{int(rb_['age']) if isnum(rb_['age']) else '–'}</td></tr>"
                 f"<tr><td>Salary</td><td>{money(ra_['salary'])}</td><td>{money(rb_['salary'])}</td></tr>"
                 f"<tr><td>Brand minus pay</td><td>{ra_['gap']:+.2f}</td><td>{rb_['gap']:+.2f}</td></tr>"
                 f"<tr><td>Qualifying endorsement</td><td>{'Yes' if ra_['endorsed'] else 'No'}</td><td>{'Yes' if rb_['endorsed'] else 'No'}</td></tr>")
        md(f'<table class="grid"><tr><th></th><th style="color:{ACCENT}">{e(pa)}</th><th>{e(pb)}</th></tr>{body}</table>')


# ---------------------------------------------------------------- similar players (KNN on performance)
with tab_s:
    md('<p class="muted" style="font-size:.95rem;margin:8px 0 4px">Look-alikes are matched on <b>performance</b> (2K-based talent score) by default. '
       'Comparing their marketability shows who is under-marketed for how well he plays.</p>')
    s1, s2, s3 = st.columns([1.5, 1, 1.6])
    sp_pick = s1.selectbox("Player", names, index=names.index(st.session_state.get("pick", DEFAULT_PLAYER)), key="sim_pick")
    sk = s2.slider("Look-alikes (k)", 3, 15, 5, key="sim_k")
    match_labels = s3.multiselect("Match on", [dl.LABEL[f] for f in SCORED], default=[dl.LABEL["talent"]], key="sim_feats",
                                  help="Performance only answers: 'plays like them, but is he as marketable?'. Add factors to match on more.")
    feats = ["s_" + f for f in SCORED if dl.LABEL[f] in match_labels] or knn.PERFORMANCE
    me = d.loc[sp_pick]
    nb = knn.similar_players(d, sp_pick, k=sk, features=feats)
    peer_avg = nb.score.mean()
    gap = me["score"] - peer_avg
    performance_only = list(feats) == list(knn.PERFORMANCE)
    matched_on = "performance" if performance_only else dl.LABEL[feats[0][2:]].lower() if len(feats) == 1 else "matched factors"
    verdict, vcol = (f"Under-marketed for his {matched_on}", NEG) if gap <= -8 else \
                    ("Out-marketing his look-alikes", POS) if gap >= 8 else ("In line with his look-alikes", MUTED)
    tiles([(f"{sp_pick} · 2K rating", f"{me['rating_2k']:.0f}" if isnum(me["rating_2k"]) else "–", e(me["team"]), None),
           (f"{sp_pick} · marketability", f"{me['score']:.1f}", f"rank {me['rank']} of {N}", ACCENT),
           ("Look-alike average", f"{peer_avg:.1f}", f"{len(nb)} players · {int(nb.endorsed.sum())} qualifying endorsements", None),
           ("Marketability gap", f"{gap:+.1f}", f'<b style="color:{vcol}">{verdict}</b>', vcol)])

    # one factor fits beside the chart; several factor columns need the full width
    left, right = st.columns([1.15, 1], gap="large") if len(feats) == 1 else (st.container(), st.container())
    with left:
        sec("Performance vs marketability", "shaded = look-alike 2K range")
        others = d[~d.player.isin(list(nb.player) + [sp_pick])]
        fig = go.Figure()
        band = pd.concat([nb.rating_2k, pd.Series([me["rating_2k"]])]).dropna()
        if len(band):
            fig.add_vrect(x0=band.min() - 0.5, x1=band.max() + 0.5, fillcolor=ACCENT, opacity=0.08, line_width=0)
        fig.add_trace(go.Scatter(x=others.rating_2k, y=others.score, mode="markers", name="Other players",
                                 marker=dict(size=6, color="#C3CDD6"), text=others.player, hovertemplate="%{text}: 2K %{x}, score %{y:.1f}<extra></extra>"))
        fig.add_trace(go.Scatter(x=nb.rating_2k, y=nb.score, mode="markers", name="Look-alikes", text=nb.player,
                                 marker=dict(size=11, color=SURFACE, line=dict(color=INK, width=2.5)),
                                 hovertemplate="%{text}: 2K %{x}, score %{y:.1f}<extra></extra>"))
        fig.add_trace(go.Scatter(x=[me["rating_2k"]], y=[me["score"]], mode="markers+text", name=sp_pick, text=[f"  {sp_pick}"], textposition="middle right",
                                 textfont=dict(size=12, color=ACCENT), marker=dict(size=15, color=ACCENT, line=dict(color=SURFACE, width=2)),
                                 hovertemplate="%{text}: 2K %{x}, score %{y:.1f}<extra></extra>"))
        fig.add_hline(y=peer_avg, line_color=INK, line_dash="dot", line_width=1, annotation_text=f"look-alike avg {peer_avg:.1f}",
                      annotation_position="top left", annotation_font_size=11)
        fig.update_layout(legend=dict(orientation="h", yanchor="bottom", y=1.0, x=0))
        chart(fig, 420, "NBA 2K rating", "Marketability score", legend=True)
    with right:
        sec("Nearest look-alikes")
        row_order = [sp_pick] + list(nb.player)                                          # selected player first, for reference
        tbl = pd.DataFrame({"Player": [f"{sp_pick} (selected)"] + list(nb.player), "Team": d.loc[row_order, "team"].to_numpy(),
                            "Similarity": [None] + list(nb.similarity)})
        factor_cols = {dl.LABEL[c[2:]]: c for c in feats}
        for label, c in factor_cols.items():
            tbl[label] = d.loc[row_order, c].to_numpy()
        tbl["2K"] = d.loc[row_order, "rating_2k"].to_numpy()
        tbl["Score"] = d.loc[row_order, "score"].to_numpy()
        tbl[f"Score vs {sp_pick}"] = [None] + list(nb.score - me["score"])
        tbl["Qualifying endorsement"] = d.loc[row_order, "endorsed"].astype(bool).to_numpy()
        st.dataframe(tbl, hide_index=True, width="stretch",
                     column_config={"Similarity": st.column_config.NumberColumn(format="%.0f%%"), "2K": st.column_config.NumberColumn(format="%.0f"),
                                    **{label: st.column_config.NumberColumn(format="%.2f") for label in factor_cols},
                                    "Score": st.column_config.NumberColumn(format="%.1f"), f"Score vs {sp_pick}": st.column_config.NumberColumn(format="%+.1f"),
                                    "Qualifying endorsement": st.column_config.CheckboxColumn()})
        st.caption(f"Factor columns are the 0-1 scores used for the match. Positive 'Score vs' = the look-alike is more marketable than {sp_pick}.")

    sec("League-wide", f"each player's score minus the average of his {sk} look-alikes ({', '.join(match_labels) or 'Performance (2K)'})")
    gaps = knn.marketability_gap(d, k=sk, features=feats).join(d[["player", "team", "rating_2k", "score", "endorsed"]])
    gt = lambda g: pd.DataFrame({"Player": g.player, "Team": g.team, "2K": g.rating_2k, "Score": g.score, "Look-alike avg": g.peer_avg,
                                 "Gap": g.gap, "Qualifying endorsement": g.endorsed.astype(bool)})
    cfg = {"2K": st.column_config.NumberColumn(format="%.0f"), "Score": st.column_config.NumberColumn(format="%.1f"),
           "Look-alike avg": st.column_config.NumberColumn(format="%.1f"), "Gap": st.column_config.NumberColumn(format="%+.1f"),
           "Qualifying endorsement": st.column_config.CheckboxColumn()}
    u1, u2 = st.columns(2, gap="large")
    with u1:
        md(f'<p class="eyebrow" style="color:{NEG}">Most under-marketed</p>')
        st.dataframe(gt(gaps.nsmallest(15, "gap")), hide_index=True, width="stretch", column_config=cfg, height=420)
    with u2:
        md(f'<p class="eyebrow" style="color:{POS}">Most out-marketing their {"play" if performance_only else "look-alikes"}</p>')
        st.dataframe(gt(gaps.nlargest(15, "gap")), hide_index=True, width="stretch", column_config=cfg, height=420)
    st.download_button("Download all gaps (CSV)", gt(gaps.sort_values("gap")).to_csv(index=False).encode(), "mavs_marketability_gaps.csv", "text/csv")


# ---------------------------------------------------------------- data (yearly upload)
with tab_d:
    left, right = st.columns([1.5, 1], gap="large")
    with left:
        sec("Load a new season", "same sheets and columns as the model workbook")
        st.file_uploader("Upload the season's workbook (.xlsx)", type=["xlsx"], key="upload",
                         help="Used for this browser session. To make it the default for everyone, replace data/Mavs_model_slim.xlsx in the repo.")
        sec("Workbook check", e(source))
        try:
            sheets, head = workbook_check(str(path), path.stat().st_mtime)
            ok_style, bad_style, opt_style = (f"background:#E1ECF5;color:{POS}", f"background:#F8E4DC;color:{NEG}", f"background:#E8EDF1;color:{MUTED}")
            cells = "".join(f'<div style="{ok_style if s in sheets else bad_style}"><b>{"✓" if s in sheets else "✗"}</b>{e(s)}<span>{"Found" if s in sheets else "Missing · required"}</span></div>'
                            for s in REQUIRED_SHEETS)
            cells += "".join(f'<div style="{ok_style if s in sheets else opt_style}"><b>{"✓" if s in sheets else "–"}</b>{e(s)}<span>{"Found" if s in sheets else "Optional"}</span></div>'
                             for s in OPTIONAL_SHEETS)
            md('<p class="eyebrow" style="color:#5B6B7A">Sheets</p><div class="check">' + cells + "</div>")
            have = lambda p: any(h.lower().startswith(p.lower()) for h in head)
            cells = "".join(f'<div style="{ok_style if have(c) else bad_style}"><b>{"✓" if have(c) else "✗"}</b>{e(c)}<span>{"Found" if have(c) else "Missing"}</span></div>'
                            for c in PLAYER_COLS)
            md('<p class="eyebrow" style="color:#5B6B7A">Players sheet columns</p><div class="check">' + cells + "</div>")
        except Exception as err:
            st.warning(f"Could not inspect the workbook: {err}")
        if BUNDLED.exists():
            st.download_button("Download the current model workbook (use it as the template)", BUNDLED.read_bytes(), BUNDLED.name,
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with right:
        sec("Active dataset")
        cc = mt["crosscheck"]
        md(kv_table([("Source", e(source)),
                     ("Players", f"{mt['n']}"),
                     ("Qualifying endorsements", f"{mt['n_endorsed']} ({mt['n_endorsed_orig']} recorded + {mt['n_leads_included']} unverified leads)"),
                     ("Updated", time.strftime("%d %b %Y %H:%M", time.localtime(mt["modified"]))),
                     ("Scenario in use", e(scen)),
                     ("Excel cross-check", "Matches the workbook" if cc is not None and cc < 1e-6 else "Computed from inputs" if cc is None else f"Differs by up to {cc:.3f}")]))
        sec("Each new season")
        md('<ol style="margin:0;padding-left:20px;line-height:1.8;font-size:.92rem">'
           "<li>Download the current workbook above and paste the new season's numbers into it.</li>"
           "<li>Upload it here and check every row above shows ✓.</li>"
           "<li>Every tab re-scores instantly, KNN look-alikes included.</li>"
           "<li>Happy with it? Replace <code>data/Mavs_model_slim.xlsx</code> so it becomes the default.</li></ol>")


# ---------------------------------------------------------------- method
with tab_m:
    wf = pd.DataFrame({"Factor": [dl.LABEL[f] for f in dl.FACTORS], "Weight (%)": (active_w / W_TOT * 100).round(1).to_numpy(),
                       "What it measures": [dl.EXPLAIN[f] for f in dl.FACTORS]})
    sec("How the score is built", e(scen))
    c1, c2 = st.columns([1, 1.4], gap="large")
    with c1:
        fig = go.Figure(go.Bar(y=wf.Factor[::-1], x=wf["Weight (%)"][::-1], orientation="h", marker_color=ACCENT,
                               text=[f"{v:.0f}%" for v in wf["Weight (%)"][::-1]], textposition="outside"))
        chart(fig, 280, "Weight (%)")
    with c2:
        st.dataframe(wf, hide_index=True, width="stretch")
    st.markdown(f"Each factor score is a 0–1 percentile rank among the {N} players. The Marketability Score multiplies each by its weight, adds them up and scales the total to 0–100.")

    sec("All weighting scenarios", "points")
    st.dataframe(model.weights.rename(index=dl.LABEL).round(1), width="stretch")

    sec("How well the score lines up with qualifying endorsements")
    y = d.endorsed
    auc_rows = []
    for s in model.scenarios:
        try:
            auc_rows.append((s, roc_auc_score(y, d["score_" + s])))
        except (ValueError, KeyError):
            auc_rows.append((s, np.nan))
    reg = model.regression
    a1, a2 = st.columns([1, 1.4], gap="large")
    a1.dataframe(pd.DataFrame(auc_rows, columns=["Weights", "AUC vs qualifying endorsement"]).round(3), hide_index=True, width="stretch")
    a2.markdown(f"AUC is how often a random player with a qualifying endorsement outscores a random player without one (0.5 = coin flip, 1.0 = perfect). "
                f"The regression's cross-validated AUC is **{reg.get('cv_auc', float('nan')):.3f}**. Scenarios other than the regression "
                "were not fitted to the data, so their AUC is a fair test.")

    sec("Nearest-neighbour check")
    kb = knn.best_k(d)
    k1, k2 = st.columns([1, 1.4], gap="large")
    k1.dataframe(kb.round(3), hide_index=True, width="stretch")
    k2.markdown("For each player, the share of his *k* nearest neighbours who have a qualifying endorsement, tested leave-one-out (a player is never his own neighbour). "
                "The AUC shows how well that share separates players with a qualifying endorsement from those without one.")

    sec("Limits")
    for line in [f"Qualifying endorsements are the outcome. The sheet records mostly well-documented deals; {mt['n_leads_included']} more come from web-search leads nobody has verified.",
                 "Factors overlap (a famous player is high on several), so individual weights are uncertain; the ranking changes little across reasonable weights.",
                 "Sentiment is media tone from headlines, not fan opinion. Global reach uses Wikipedia and Google Trends and is weak for China.",
                 "Team brand health (market size and franchise value) is shown at weight 0: it explains about 1% of endorsements. Availability is 50% games played, 50% career runway.",
                 "Social media counts followers only, not engagement or growth.",
                 "Similar-player look-alikes match on factor scores only; two players with the same 2K rating can still differ in role, position or market."]:
        st.markdown(f"- {line}")
    if len(model.dictionary):
        with st.expander("Data dictionary"):
            st.dataframe(model.dictionary, hide_index=True, width="stretch")

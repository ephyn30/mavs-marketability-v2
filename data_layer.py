"""Data layer: the Excel workbook is the DATABASE; this module reads it and recomputes every score.

Why recompute instead of reading the workbook's formula results?  A workbook saved by a script (not opened in Excel) has no stored
formula results. Recomputing from the raw inputs (Players, Outcome, Endorsement Leads, Control Panel, AHP tabs) always works, and
`crosscheck()` compares against Excel's own numbers whenever they are stored, so the two can never silently disagree.

Everything mirrors the workbook formulas:
  factor score = percentile rank (average rank / n) of the input; Talent can use the rating-based scale (Control Panel B26/B27)
  Marketability Score = sum(score x weight) / sum(weights) x 100
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd

HERE = Path(__file__).resolve().parent
DEFAULT_PATH = Path(os.environ.get("MAVS_WORKBOOK", HERE.parent / "outputs" / "Mavs_model_updated.xlsx"))

FACTORS = ["social", "global_reach", "talent", "sentiment", "availability", "brand_health"]
LABEL = {"social": "Social media size", "global_reach": "Global reach", "talent": "Performance (2K)", "sentiment": "Sentiment",
         "availability": "Availability & runway", "brand_health": "Brand health (team)"}
PILLAR = {"social": "Audience", "global_reach": "Audience", "talent": "Performance", "sentiment": "Perception", "availability": "Durability", "brand_health": "Brand health"}
EXPLAIN = {"social": "Instagram + TikTok + X followers combined, ranked against all 505 players",
           "global_reach": "Fame outside the US and Canada (Wikipedia readers abroad + countries where people search for him)",
           "talent": "NBA 2K rating (on-court quality)",
           "sentiment": "How much and how positively news headlines cover him (VADER on Google News)",
           "availability": "Can he keep playing, and for how long? Games played (up to 3 seasons) linked with age: an older player is closer to losing his contract and his endorsements",
           "brand_health": "How big his team's brand is: market size by location (metro population, GDP, Nielsen rank) and franchise value. Weight 0: tested, adds about 1%"}
TEAM_NAME = {"ATL": "Atlanta Hawks", "BOS": "Boston Celtics", "BRK": "Brooklyn Nets", "CHI": "Chicago Bulls", "CHO": "Charlotte Hornets", "CLE": "Cleveland Cavaliers",
             "DAL": "Dallas Mavericks", "DEN": "Denver Nuggets", "DET": "Detroit Pistons", "GSW": "Golden State Warriors", "HOU": "Houston Rockets", "IND": "Indiana Pacers",
             "LAC": "LA Clippers", "LAL": "Los Angeles Lakers", "MEM": "Memphis Grizzlies", "MIA": "Miami Heat", "MIL": "Milwaukee Bucks", "MIN": "Minnesota Timberwolves",
             "NOP": "New Orleans Pelicans", "NYK": "New York Knicks", "OKC": "Oklahoma City Thunder", "ORL": "Orlando Magic", "PHI": "Philadelphia 76ers", "PHO": "Phoenix Suns",
             "POR": "Portland Trail Blazers", "SAC": "Sacramento Kings", "SAS": "San Antonio Spurs", "TOR": "Toronto Raptors", "UTA": "Utah Jazz", "WAS": "Washington Wizards"}
UP = [(0, 1), (0, 2), (0, 3), (1, 2), (1, 3), (2, 3)]
PILLARS = ["Audience", "Performance", "Perception", "Durability"]


@dataclass
class Model:
    players: pd.DataFrame                    # one row per player: inputs, factor scores, totals, ranks, endorsement
    weights: pd.DataFrame                    # factors x scenarios (points; normalised inside the score)
    scenarios: list
    active_scenario: str
    leads: pd.DataFrame
    regression: dict
    dictionary: pd.DataFrame
    meta: dict = field(default_factory=dict)


def _rows(ws, min_row=1, max_col=None):
    return list(ws.iter_rows(min_row=min_row, max_col=max_col, values_only=True))


def _merged_rows(wb, wv, name, min_row=1, max_col=None):
    """Rows of a sheet with each formula cell replaced by the result Excel stored for it (when the file was last saved by Excel).
    Typed inputs are kept as they are, and a formula with no stored result stays a formula (read as missing / recomputed)."""
    from itertools import zip_longest
    a, b = _rows(wb[name], min_row, max_col), _rows(wv[name], min_row, max_col)
    return [tuple(v if isinstance(f, str) and f.startswith("=") and v is not None else f for f, v in zip_longest(rf, rv or ()))
            for rf, rv in zip_longest(a, b)]


def _num(v):
    if v is None or v == "": return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if s.startswith("="):                                  # typed fractions such as =1/3 in a file that Excel has not saved
        if re.fullmatch(r"=[0-9.+\-*/() ]+", s):
            try: return float(eval(s[1:], {"__builtins__": {}}))
            except Exception: return np.nan
        return np.nan
    try: return float(s)
    except ValueError: return np.nan


def ahp_pillars(wb) -> dict:
    """Group AHP pillar weights from the member tabs (geometric mean of complete matrices), as in the workbook."""
    mats = []
    for name in wb.sheetnames:
        if not name.startswith("AHP ") or name.split("AHP ", 1)[1] in ("Group", "Compare", "Example", "Read me"):
            continue
        ws = wb[name]; grid = {(r, c): v for r, row in enumerate(_rows(ws, 1, 6), 1) for c, v in enumerate(row, 1)}
        A = np.ones((4, 4)); ok = True
        for i, j in UP:
            v = _num(grid.get((6 + i, 2 + j)))
            if np.isnan(v) or v <= 0: ok = False; break
            A[i, j], A[j, i] = v, 1 / v
        if ok: mats.append(A)
    if not mats: return {}
    G = np.exp(np.mean([np.log(A) for A in mats], axis=0))
    w = (G / G.sum(0)).mean(1)
    return dict(zip(PILLARS, w.tolist()))


def brand_table(wb):
    """Team brand health (0-1 rank among the 30 teams), recomputed from the 'Team Brand Health' tab exactly as its formulas do."""
    if "Team Brand Health" not in wb.sheetnames: return None
    rows = [r for r in _rows(wb["Team Brand Health"], 2, 8) if r and r[1] and str(r[1]) in TEAM_NAME]
    t = pd.DataFrame(rows, columns=["team", "code", "valuation", "revenue", "dma", "dma_rank", "metro_pop", "metro_gdp"]).set_index("code")
    for c in ("valuation", "revenue", "metro_pop", "metro_gdp"): t[c] = pd.to_numeric(t[c], errors="coerce")
    pop, gdp = t.metro_pop.rank(pct=True), t.metro_gdp.rank(pct=True)
    dma = (-pd.to_numeric(t.dma_rank, errors="coerce")).rank(pct=True)
    t["market_size"] = pd.concat([pop, gdp, dma], axis=1).mean(axis=1, skipna=True).rank(pct=True)
    t["brand_health"] = pd.concat([t.market_size, t.valuation.rank(pct=True)], axis=1).mean(axis=1).rank(pct=True)
    return t


def _rank_pct(s: pd.Series, ascending=True):
    return s.rank(method="average", ascending=ascending) / s.notna().sum()


def load(path: str | Path = DEFAULT_PATH) -> Model:
    path = Path(path)
    wb = openpyxl.load_workbook(path, read_only=True, data_only=False)       # formulas and typed inputs
    wv = openpyxl.load_workbook(path, read_only=True, data_only=True)        # stored results (may be empty)
    # ---------- Control Panel
    cp = {(r, c): v for r, row in enumerate(_merged_rows(wb, wv, "Control Panel", 1, 10), 1) for c, v in enumerate(row, 1)}
    scen_names = [cp[(5, c)] for c in range(4, 11) if cp.get((5, c))]
    active = cp.get((3, 2)) or scen_names[0]
    split = _num(cp.get((22, 2))); split = 0.5 if np.isnan(split) else split
    talent_switch = int(_num(cp.get((26, 2))) or 1); floor = _num(cp.get((27, 2))); floor = 60 if np.isnan(floor) else floor
    age_share = _num(cp.get((23, 2))); age_share = 0.5 if np.isnan(age_share) else age_share
    tier_on = {str(cp.get((30 + k, 1))): str(cp.get((30 + k, 2))).upper() == "Y" for k in range(3)}
    W = pd.DataFrame(index=FACTORS, columns=scen_names, dtype=float)
    for k, f in enumerate(FACTORS):
        for c, s in enumerate(scen_names):
            W.loc[f, s] = _num(cp.get((6 + k, 4 + c)))
    pill = ahp_pillars(wb)
    if pill:
        ahp = {"social": pill["Audience"] * split, "global_reach": pill["Audience"] * (1 - split), "talent": pill["Performance"],
               "sentiment": pill["Perception"], "availability": pill["Durability"], "brand_health": 0.0}
        W["AHP (team)"] = [100 * ahp[f] for f in FACTORS]
        if "Regression (data)" in W:
            a = W["AHP (team)"]; r = W["Regression (data)"]
            W["Blend (AHP + regression)"] = (a / a.sum() * 100 + r / r.sum() * 100) / 2
    W = W.fillna(0.0)
    # ---------- Players (inputs)
    prow = _merged_rows(wb, wv, "Players"); head = [str(h) for h in prow[0]]
    P = pd.DataFrame([r[:len(head)] for r in prow[1:] if r[0]], columns=head)
    pick = lambda start: next(c for c in head if c.lower().startswith(start.lower()))
    d = pd.DataFrame({"player": P[pick("Player")].astype(str), "team_code": P[pick("Team")], "instagram": P[pick("Instagram")], "tiktok": P[pick("TikTok")], "x": P[pick("X followers")],
                      "rating_2k": P[pick("2K")], "sentiment_score": P[pick("Sentiment")], "global_reach_score": P[pick("Global Reach")], "availability_raw": P[pick("Availability")],
                      "age": P[pick("Age")], "seasons": P[pick("Seasons")], "salary": P[pick("Salary")], "flags": P[pick("Data flags")]})
    for c in ["instagram", "tiktok", "x", "rating_2k", "sentiment_score", "global_reach_score", "availability_raw", "age", "seasons", "salary"]:
        d[c] = pd.to_numeric(d[c], errors="coerce")
    d["team"] = d.team_code.map(lambda t: TEAM_NAME.get(str(t).split("/")[0], str(t)))
    d["followers"] = d[["instagram", "tiktok", "x"]].fillna(0).sum(axis=1)
    # ---------- factor scores
    n = len(d)
    d["s_social"] = _rank_pct(d.followers); d["s_global_reach"] = _rank_pct(d.global_reach_score); d["s_sentiment"] = _rank_pct(d.sentiment_score)
    d["s_talent"] = _rank_pct(d.rating_2k) if talent_switch == 1 else ((d.rating_2k - floor).clip(lower=0) / (d.rating_2k.max() - floor))
    d["s_games"] = _rank_pct(d.availability_raw).fillna(0.5); d["s_runway"] = _rank_pct(d.age, ascending=False).fillna(0.5)          # younger = higher
    d["s_availability"] = (1 - age_share) * d.s_games + age_share * d.s_runway          # Availability & runway: games played linked with age
    bt = brand_table(wb)
    first = d.team_code.astype(str).str.split("/").str[0]
    d["s_brand_health"] = first.map(bt.brand_health) if bt is not None else 0.5
    d["s_brand_health"] = d.s_brand_health.fillna(0.5)
    d["metro_pop"] = first.map(bt.metro_pop) if bt is not None else np.nan; d["franchise_value"] = first.map(bt.valuation) if bt is not None else np.nan
    for f in FACTORS + ["games", "runway"]:             # where the player stands in the league on this score (for the traffic light)
        d["p_" + f] = _rank_pct(d["s_" + f])
    S = d[["s_" + f for f in FACTORS]].to_numpy()
    for s in W.columns:
        w = W[s].to_numpy(); d["score_" + s] = S @ w / w.sum() * 100 if w.sum() > 0 else np.nan
    d["score"] = d["score_" + active]; d["rank"] = d.score.rank(ascending=False, method="min").astype(int)
    d["score_pct"] = _rank_pct(d.score); d["pay_pct"] = _rank_pct(d.salary); d["gap"] = d.score_pct - d.pay_pct
    # ---------- endorsement outcome (original sheet + leads)
    orow = _merged_rows(wb, wv, "Outcome"); O = pd.DataFrame([r[:10] for r in orow[1:] if r[0]], columns=["player", "orig", "orig_score", "n_sponsors", "sponsors", "t", "li", "ls", "e", "src"])
    O = O.set_index("player")
    lrows = _merged_rows(wb, wv, "Endorsement Leads"); heads = ["player", "team", "tier", "sponsors", "status", "evidence", "date", "links", "verified", "verified_sponsors", "notes"]
    L = pd.DataFrame([r[:11] for r in lrows[1:] if r[0]], columns=heads)
    def include(r):
        v = str(r.verified or "").strip().lower()
        return "N" if v == "no" else "Y" if v == "yes" else ("Y" if tier_on.get(r.tier, False) else "N")
    L["include"] = L.apply(include, axis=1)
    inc = set(L.player[L.include == "Y"])
    d["endorsed_orig"] = d.player.map(O.orig).fillna(0).astype(int)
    d["endorsed"] = ((d.endorsed_orig == 1) | d.player.isin(inc)).astype(int)
    d["endorse_source"] = np.where(d.endorsed_orig == 1, "Original sheet", np.where(d.player.isin(inc), "Lead (not yet verified)", "none"))
    d["sponsors_listed"] = d.player.map(O.sponsors).fillna("")
    d["n_sponsors"] = d.player.map(O.n_sponsors)
    lead_map = L.set_index("player")
    d["lead_sponsors"] = d.player.map(lambda p: lead_map.sponsors.get(p, "") if p in inc else "")
    d["lead_tier"] = d.player.map(lambda p: lead_map.tier.get(p, "") if p in inc else "")
    d["lead_links"] = d.player.map(lambda p: lead_map.links.get(p, "") if p in inc else "")
    # ---------- regression tab (static coefficients) -> endorsement likelihood
    rg = {(r, c): v for r, row in enumerate(_rows(wb["Regression"], 1, 8), 1) for c, v in enumerate(row, 1)} if "Regression" in wb.sheetnames else {}
    reg = {}
    if rg:
        mean = [_num(rg[(12 + k, 2)]) for k in range(5)]; sd = [_num(rg[(12 + k, 3)]) for k in range(5)]; coef = [_num(rg[(12 + k, 4)]) for k in range(5)]
        icpt = _num(rg[(17, 4)]); cv = _num(rg.get((24, 2)))
        R = np.column_stack([_rank_pct(d.followers), _rank_pct(d.global_reach_score), _rank_pct(d.rating_2k), _rank_pct(d.sentiment_score), d.s_availability])
        z = (R - np.array(mean)) / np.array(sd); d["p_endorse"] = 1 / (1 + np.exp(-(icpt + z @ np.array(coef))))
        reg = {"coef": coef, "mean": mean, "sd": sd, "intercept": icpt, "cv_auc": cv}
    # ---------- data dictionary
    dd = pd.DataFrame([r[:5] for r in _rows(wb["Data Dictionary"])[1:] if r[0]], columns=["Item", "Where", "Meaning", "How it is made", "Source"]) if "Data Dictionary" in wb.sheetnames else pd.DataFrame()
    d = d.set_index("player", drop=False)
    meta = {"age_share": age_share, "path": str(path), "modified": os.path.getmtime(path), "n": n, "talent_switch": talent_switch, "floor": floor, "tier_on": tier_on, "ahp_pillars": pill,
            "n_endorsed": int(d.endorsed.sum()), "n_endorsed_orig": int(d.endorsed_orig.sum()), "n_leads_included": len(inc), "crosscheck": crosscheck(wv, d)}
    return Model(d, W, list(W.columns), active, L, reg, dd, meta)


def crosscheck(wv, d: pd.DataFrame):
    """Compare our scores with the numbers Excel stored in Scores!I (None if the workbook has no stored results)."""
    try:
        rows = _rows(wv["Scores"], 12, 9)
        vals = [(r[0], r[8]) for r in rows if r and r[0] and isinstance(r[8], (int, float))]
        if len(vals) < 100: return None
        diff = max(abs(v - d.score[p]) for p, v in vals if p in d.index)
        return float(diff)
    except Exception:
        return None


def rescored(m: Model, scenario: str) -> pd.DataFrame:
    """Same players table but with a different scenario's weights as the 'score' (for the sidebar scenario picker)."""
    d = m.players.copy(); d["score"] = d["score_" + scenario]; d["rank"] = d.score.rank(ascending=False, method="min").astype(int)
    d["score_pct"] = _rank_pct(d.score); d["gap"] = d.score_pct - d.pay_pct
    return d


if __name__ == "__main__":
    m = load()
    print(m.meta); print(m.weights.round(1)); print(m.players.sort_values("rank").head(5)[["player", "team", "score", "rank", "endorsed"]])
    f = m.players.loc["Cooper Flagg"]; print(f[["team", "score", "rank", "endorsed", "p_endorse", "gap"]])

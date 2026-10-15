"""KNN starter: 'player twins' = the k players whose factor scores are closest to the chosen player's.

This is a SIMPLE, transparent starting point for the team's KNN work, not a finished analysis. What it does:
  1. describes each player by his five factor scores (social, global reach, performance, sentiment, availability), each on a 0-1 scale
  2. finds the k nearest players by straight-line (Euclidean) distance
  3. reports how many of those neighbours are endorsed, as a rough 'players like him' benchmark
Ideas to extend it:  weight the features by the regression weights; add age or position; use distance-weighted voting;
tune k with cross-validation; predict rookies (needs pre-draft features); compare KNN's endorsement prediction with the regression's.
"""
from __future__ import annotations

import numpy as np
import pandas as pd
from sklearn.metrics import roc_auc_score
from sklearn.neighbors import NearestNeighbors

FEATURES = ["s_social", "s_global_reach", "s_talent", "s_sentiment", "s_availability"]
PERFORMANCE = ["s_talent"]          # 'performance look-alikes': matched on the 2K-based talent score only


def _X(d: pd.DataFrame, features) -> np.ndarray:
    """Feature matrix with any missing score filled by the league median (NearestNeighbors can't take NaN)."""
    f = list(features)
    return d[f].fillna(d[f].median()).fillna(0.5).to_numpy()


def similar_players(d: pd.DataFrame, name: str, k: int = 5, features=FEATURES) -> pd.DataFrame:
    features = list(features)
    X = _X(d, features)
    nn = NearestNeighbors(n_neighbors=min(k + 1, len(d))).fit(X)
    dist, idx = nn.kneighbors(X[[d.index.get_loc(name)]])
    rows = [(d.index[j], dd) for j, dd in zip(idx[0], dist[0]) if d.index[j] != name][:k]
    out = d.loc[[r[0] for r in rows], ["player", "team", "score", "rank", "endorsed", "followers", "rating_2k"]].copy()
    out["distance"] = [r[1] for r in rows]
    out["similarity"] = 100 * (1 - out.distance / np.sqrt(len(features)))            # 100 = identical profile
    return out.reset_index(drop=True)


def neighbour_endorsement_rate(d: pd.DataFrame, k: int = 10, features=FEATURES) -> pd.Series:
    """For every player: the share of his k nearest neighbours (excluding himself) who are endorsed."""
    X = _X(d, features)
    _, idx = NearestNeighbors(n_neighbors=k + 1).fit(X).kneighbors(X)
    e = d.endorsed.to_numpy()
    rates = [np.mean([e[j] for j in row if j != i][:k]) for i, row in enumerate(idx)]
    return pd.Series(rates, index=d.index)


def marketability_gap(d: pd.DataFrame, k: int = 5, features=PERFORMANCE) -> pd.DataFrame:
    """For every player: the average marketability score of his k nearest look-alikes (excluding himself) and his gap to it.
    With the default features (performance only), a large negative gap = plays like his peers but is far less marketable."""
    X = _X(d, features)
    _, idx = NearestNeighbors(n_neighbors=min(k + 1, len(d))).fit(X).kneighbors(X)
    sc = d.score.to_numpy()
    peer = [np.mean([sc[j] for j in row if j != i][:k]) for i, row in enumerate(idx)]
    out = pd.DataFrame({"peer_avg": peer}, index=d.index)
    out["gap"] = d.score - out.peer_avg
    return out


def loo_auc(d: pd.DataFrame, k: int = 10, features=FEATURES) -> float:
    """Leave-one-out check: does 'share of my neighbours who are endorsed' separate endorsed from non-endorsed players? (AUC; 0.5 = no)"""
    return float(roc_auc_score(d.endorsed, neighbour_endorsement_rate(d, k, features)))


def best_k(d: pd.DataFrame, ks=(3, 5, 7, 10, 15, 20, 30), features=FEATURES) -> pd.DataFrame:
    return pd.DataFrame({"k": ks, "leave-one-out AUC": [loo_auc(d, k, features) for k in ks]})


if __name__ == "__main__":
    import data_layer as dl
    m = dl.load(); d = m.players
    print(similar_players(d, "Cooper Flagg").round(2).to_string(index=False)); print(best_k(d).round(3).to_string(index=False))

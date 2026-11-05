# Mavs Player Marketability Scorecard (private)

Interactive scorecard for the Dallas Mavericks 2026 Sports Analytics Case. Search a player, see his marketability score,
profile shape, factor-by-factor percentiles, endorsements, and recommendations for agents, the team and the player.

**Tabs:** Player (Profile · Recommendations) · Team · Leaderboard · Compare · Similar players · Data · Method

- `app.py`: the Streamlit interface (warm-neutral theme; palette constants at the top of the file)
- `data_layer.py`: reads the workbook and recomputes every score
- `knn.py`: KNN look-alikes; `marketability_gap()` powers the Similar players tab (matched on performance by default)
- `.streamlit/config.toml`: theme colors used by Streamlit's own widgets and tables
- `data/Mavs_model_slim.xlsx`: the database (trimmed copy of the team workbook: only the tabs the app reads)

Run locally: `pip install -r requirements.txt` then `streamlit run app.py`.

**New season:** open the Data tab, upload the new workbook (same sheets and columns), check the workbook check is all ✓.
Uploads last for the browser session; replace `data/Mavs_model_slim.xlsx` to make a season the default.

Note: 16 of the 114 players with a qualifying endorsement come from web-search leads that have not been verified by the team (the app says so).

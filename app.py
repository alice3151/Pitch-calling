import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

st.set_page_config(page_title="Keio Pitch Calling Support", layout="wide")
DATA_FILE = Path(__file__).parent / "trackman.xlsx"

TEAM_NAMES = {
    "KEI_KEI": "Keio University",
    "RIK": "Rikkyo University",
    "MEI_MEI": "Meiji University",
    "TOK": "Hosei University",
    "WAS_EDA": "Waseda University",
    "HOS_HOS": "Hosei University",
}
TEAM_ORDER = ["KEI_KEI", "RIK", "MEI_MEI", "TOK", "WAS_EDA", "HOS_HOS"]


def clean_text(x):
    return "" if pd.isna(x) else str(x).strip()


def valid_person_name(x):
    s = clean_text(x)
    return bool(s) and "," in s and not s.isdigit()


def valid_numeric_id(x):
    s = clean_text(x)
    return s.isdigit()


@st.cache_data
def load_data():
    xl = pd.ExcelFile(DATA_FILE)
    raw = []
    for sheet in xl.sheet_names:
        try:
            x = pd.read_excel(DATA_FILE, sheet_name=sheet)
            if "Pitcher" in x.columns and "Catcher" in x.columns:
                x["SourceSheet"] = sheet
                raw.append(x)
        except Exception:
            pass
    df = pd.concat(raw, ignore_index=True)

    # ---- Repair malformed catcher fields using valid ID/name pairs ----
    catcher_by_id = {}
    catcher_team_by_id = {}
    for _, r in df.iterrows():
        c, cid, team = r.get("Catcher"), r.get("CatcherId"), r.get("CatcherTeam")
        if valid_person_name(c) and valid_numeric_id(cid):
            key = str(int(float(cid)))
            catcher_by_id.setdefault(key, clean_text(c))
            if clean_text(team):
                catcher_team_by_id.setdefault(key, clean_text(team))

    fixed_c, fixed_id, fixed_team = [], [], []
    for _, r in df.iterrows():
        c, cid, cteam = r.get("Catcher"), r.get("CatcherId"), r.get("CatcherTeam")
        if valid_person_name(c) and valid_numeric_id(cid):
            fixed_c.append(clean_text(c)); fixed_id.append(str(int(float(cid))))
            fixed_team.append(clean_text(cteam)); continue
        if valid_numeric_id(c):
            key = str(int(float(c)))
            if key in catcher_by_id:
                fixed_c.append(catcher_by_id[key]); fixed_id.append(key)
                fixed_team.append(catcher_team_by_id.get(key, "")); continue
        fixed_c.append(""); fixed_id.append(""); fixed_team.append("")

    df["Catcher"], df["CatcherId"], df["CatcherTeam"] = fixed_c, fixed_id, fixed_team

    core = [
        "Pitcher", "PitcherId", "PitcherTeam", "Catcher", "CatcherId", "CatcherTeam",
        "Batter", "BatterId", "BatterTeam", "BatterSide", "TaggedPitchType",
        "AutoPitchType", "PitchCall", "PlayResult", "RunnerState", "GameID",
        "PitchUID", "PlayID", "Top/Bottom", "Inning", "PAofInning", "PitchofPA",
    ]
    for c in core:
        if c in df.columns:
            df[c] = df[c].map(clean_text)

    df["PitchType"] = df["TaggedPitchType"].where(
        df["TaggedPitchType"].ne(""), df["AutoPitchType"]
    )

    for c in ["Balls", "Strikes", "PitchofPA"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["Balls"] = df["Balls"].fillna(0).astype(int)
    df["Strikes"] = df["Strikes"].fillna(0).astype(int)
    df["PitchofPA"] = df["PitchofPA"].fillna(0).astype(int)

    # Actual location -> same 3x3 grid used by the UI.
    if "PlateLocSide" in df.columns and "PlateLocHeight" in df.columns:
        side = pd.to_numeric(df["PlateLocSide"], errors="coerce")
        height = pd.to_numeric(df["PlateLocHeight"], errors="coerce")
        col = np.select([side < -0.23, side <= 0.23], ["Inner", "Middle"], default="Outer")
        row = np.select([height < 2.15, height <= 3.15], ["Low", "Middle"], default="High")
        df["ActualZone"] = np.where(side.notna() & height.notna(), row + "-" + col, "Unknown")
    else:
        df["ActualZone"] = "Unknown"

    # A PA is uniquely identified by game + inning half + PA number.
    df["PAKey"] = (
        df["GameID"].astype(str) + "|" + df["Inning"].astype(str) + "|" +
        df["Top/Bottom"].astype(str) + "|" + df["PAofInning"].astype(str)
    )

    # Keep only rows where pitcher/catcher identity is trustworthy.
    df = df[df["Pitcher"].map(valid_person_name) & df["Catcher"].map(valid_person_name)].copy()
    return df


df = load_data()


def team_label(code):
    return TEAM_NAMES.get(code, code)


def finish_count(balls, strikes, pitch_call):
    pc = str(pitch_call)
    if pc in {"InPlay", "HitByPitch"}:
        return None
    if pc == "BallCalled" or pc == "Ball":
        return None if balls >= 3 else (balls + 1, strikes)
    if pc in {"StrikeCalled", "StrikeSwinging"}:
        return None if strikes >= 2 else (balls, strikes + 1)
    if "Foul" in pc:
        return (balls, strikes) if strikes >= 2 else (balls, strikes + 1)
    if "Strike" in pc:
        return None if strikes >= 2 else (balls, strikes + 1)
    return balls, strikes


def normalize_side(x):
    s = clean_text(x).lower()
    if s.startswith("r"):
        return "Right"
    if s.startswith("l"):
        return "Left"
    return "Unknown"


def result_label(pc):
    return {
        "BallCalled": "ボール",
        "StrikeCalled": "見逃しストライク",
        "StrikeSwinging": "空振り",
        "FoulBall": "ファウル",
        "InPlay": "インプレー",
        "HitByPitch": "死球",
    }.get(str(pc), str(pc) if clean_text(pc) else "その他")


def history_match_score(row, history):
    """How many most-recent pitches in the live PA match this historical PA?"""
    if not history:
        return 0
    hist = history[-len(history):]
    n = min(len(hist), int(row["PitchofPA"]) - 1)
    if n <= 0:
        return -1
    # This function is not used for row-by-row matching; kept for clarity.
    return n


def get_candidate_data(base, current_b, current_s, history):
    """Hierarchical next-pitch search.

    1) Same pitcher/catcher + handedness + current count + exact prior sequence
    2) Same + current count + most recent prior pitch
    3) Same + current count + handedness
    4) Same + handedness
    """
    count_df = base[base["Balls"].eq(current_b) & base["Strikes"].eq(current_s)].copy()
    if count_df.empty:
        count_df = base.copy()

    # If there is live history, find historical PAs whose preceding pitches match
    # the same pitch type + result. Zone is used when available, but not required,
    # so the model does not collapse when the sample is small.
    if history:
        seq = []
        for _, r in count_df.iterrows():
            pa = base[base["PAKey"].eq(r["PAKey"])].sort_values("PitchofPA")
            if pa.empty or int(r["PitchofPA"]) <= len(history):
                continue
            preceding = pa[pa["PitchofPA"] < r["PitchofPA"]].sort_values("PitchofPA").tail(len(history))
            if len(preceding) != len(history):
                continue
            ok = True
            for (_, pr), h in zip(preceding.iterrows(), history[-len(preceding):]):
                if clean_text(pr["PitchType"]) != clean_text(h["球種"]):
                    ok = False; break
                if result_label(pr["PitchCall"]) != clean_text(h["結果"]):
                    ok = False; break
            if ok:
                seq.append(r)
        if seq:
            return pd.DataFrame(seq), "直前までの配球履歴が一致する過去打席"

        # Relax to the most recent pitch only.
        h = history[-1]
        seq = []
        for _, r in count_df.iterrows():
            pa = base[base["PAKey"].eq(r["PAKey"])].sort_values("PitchofPA")
            prev = pa[pa["PitchofPA"] < r["PitchofPA"]].sort_values("PitchofPA").tail(1)
            if len(prev) == 1:
                pr = prev.iloc[0]
                if clean_text(pr["PitchType"]) == clean_text(h["球種"]) and result_label(pr["PitchCall"]) == clean_text(h["結果"]):
                    seq.append(r)
        if seq:
            return pd.DataFrame(seq), "直前の1球が一致する過去打席"

    return count_df, "現在のカウント・左右が一致する過去投球"


# ---------- Session state ----------
def reset_at_bat():
    st.session_state.balls = 0
    st.session_state.strikes = 0
    st.session_state.pitch_history = []

for k, v in {"balls": 0, "strikes": 0, "pitch_history": []}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------- Game setup ----------
st.title("⚾ Keio Pitch Calling Support")
st.caption("投手×捕手×打者左右の過去配球から、1球ごとに次球候補を更新するサポートアプリ")

with st.sidebar:
    st.header("Game setup")
    team_values = set(df["PitcherTeam"]) | set(df["CatcherTeam"])
    available_teams = [t for t in TEAM_ORDER if t in team_values]
    others = sorted(t for t in team_values if t and t not in available_teams)
    team_options = available_teams + others
    team_code = st.selectbox("大学（投手・捕手側）", team_options, format_func=team_label)

    pitcher_df = df[df["PitcherTeam"].eq(team_code)]
    pitcher_options = sorted(pitcher_df["Pitcher"].dropna().unique())
    pitcher_options = [x for x in pitcher_options if valid_person_name(x)]
    if not pitcher_options:
        st.warning("この大学の投手データがありません。"); st.stop()
    pitcher = st.selectbox("投手", pitcher_options)

    catcher_df = pitcher_df[
        pitcher_df["Pitcher"].eq(pitcher) &
        pitcher_df["CatcherTeam"].eq(team_code) &
        pitcher_df["Catcher"].ne("")
    ]
    catcher_options = sorted(catcher_df["Catcher"].unique())
    if not catcher_options:
        st.warning("この投手に紐づく捕手データがありません。"); st.stop()
    catcher = st.selectbox("捕手", catcher_options)

    matchup = df[
        df["Pitcher"].eq(pitcher) &
        df["Catcher"].eq(catcher) &
        df["PitcherTeam"].eq(team_code)
    ].copy()
    opponent_options = sorted(x for x in matchup["BatterTeam"].unique() if x and x != team_code)
    opponent = None
    if opponent_options:
        opponent = st.selectbox("相手大学", opponent_options, format_func=team_label)
        matchup = matchup[matchup["BatterTeam"].eq(opponent)]

    # IMPORTANT: choose handedness, not a specific batter.
    sides = [s for s in ["Right", "Left"] if s in set(matchup["BatterSide"].map(normalize_side))]
    if not sides:
        sides = ["Right", "Left"]
    batter_side = st.radio("打者", sides, horizontal=True)

    if st.button("🔄 新しい打席", use_container_width=True):
        reset_at_bat(); st.rerun()

# Current historical pool: pitcher + catcher + opponent + batter handedness.
base = df[
    df["Pitcher"].eq(pitcher) &
    df["Catcher"].eq(catcher) &
    df["BatterSide"].map(normalize_side).eq(batter_side)
].copy()
if opponent:
    base = base[base["BatterTeam"].eq(opponent)]

if base.empty:
    st.error("この投手×捕手×打者左右の組み合わせに過去データがありません。")
    st.stop()

st.markdown(f"### {pitcher} × {catcher} × {batter_side}打者")
if opponent:
    st.caption(f"{team_label(team_code)}  |  {team_label(opponent)}")
else:
    st.caption(team_label(team_code))

current_b, current_s = st.session_state.balls, st.session_state.strikes
st.metric("現在のカウント", f"{current_b} - {current_s}")

# ---------- Live history ----------
if st.session_state.pitch_history:
    st.subheader("この打席の履歴")
    st.dataframe(pd.DataFrame(st.session_state.pitch_history), use_container_width=True, hide_index=True)

# ---------- Candidate generation ----------
candidates, basis = get_candidate_data(base, current_b, current_s, st.session_state.pitch_history)
st.subheader("次球候補")
st.caption(f"候補の根拠：{basis}（{len(candidates)}球）")

pitch_counts = candidates["PitchType"].replace("", np.nan).dropna().value_counts()
if not pitch_counts.empty:
    pitch_pct = (pitch_counts / pitch_counts.sum() * 100).round(1)
    rec = pd.DataFrame({
        "球種": pitch_counts.index,
        "過去投球数": pitch_counts.values,
        "使用率(%)": [pitch_pct[x] for x in pitch_counts.index],
    })
    st.dataframe(rec, use_container_width=True, hide_index=True)

    # Top candidates as compact cards.
    cols = st.columns(min(4, len(rec)))
    for i, row in rec.head(4).iterrows():
        with cols[i % len(cols)]:
            st.metric(str(row["球種"]), f'{row["使用率(%)"]:.1f}%', f'{int(row["過去投球数"])}球')
else:
    st.warning("この条件では球種データがありません。")

# Location tendency from actual TrackMan locations.
# Show a 3x3 heatmap-like grid, and let the selected pitch type change the grid.
zone_order = [
    ["High-Inner", "High-Middle", "High-Outer"],
    ["Middle-Inner", "Middle-Middle", "Middle-Outer"],
    ["Low-Inner", "Low-Middle", "Low-Outer"],
]
zone_labels = {
    "High-Inner": "高め・内", "High-Middle": "高め・真ん中", "High-Outer": "高め・外",
    "Middle-Inner": "中・内", "Middle-Middle": "中・真ん中", "Middle-Outer": "中・外",
    "Low-Inner": "低め・内", "Low-Middle": "低め・真ん中", "Low-Outer": "低め・外",
}

st.markdown("#### 過去のコース分布")

# The pitch type selector is shared by the location analysis and the actual iPitch input.
pitch_types = sorted(x for x in candidates["PitchType"].dropna().unique() if clean_text(x))
if not pitch_types:
    pitch_types = sorted(x for x in base["PitchType"].dropna().unique() if clean_text(x))
selected_type = st.selectbox("球種（コース分布を確認）", pitch_types) if pitch_types else ""

all_loc = candidates[candidates["ActualZone"].isin(zone_labels)].copy()
pitch_loc = all_loc[all_loc["PitchType"].eq(selected_type)] if selected_type else all_loc

def render_zone_grid(data, title, key_prefix, clickable=False):
    counts = data["ActualZone"].value_counts() if not data.empty else pd.Series(dtype=int)
    total = int(counts.sum())
    st.markdown(f"**{title}**  {'（' + str(total) + '球）' if total else '（データなし）'}")
    for r, row in enumerate(zone_order):
        cols = st.columns(3)
        for c, zone in enumerate(row):
            n = int(counts.get(zone, 0))
            pct = (n / total * 100) if total else 0.0
            label = f"{zone_labels[zone]}\n{pct:.1f}%\n{n}球"
            with cols[c]:
                if clickable:
                    selected = st.session_state.selected_zone == zone
                    button_label = ("✓ " if selected else "") + label.replace("\n", "  ")
                    if st.button(button_label, key=f"{key_prefix}_{zone}", use_container_width=True):
                        st.session_state.selected_zone = zone
                        st.rerun()
                else:
                    st.button(label.replace("\n", "  "), key=f"{key_prefix}_{zone}", use_container_width=True, disabled=True)
    return counts, total

render_zone_grid(pitch_loc, f"{selected_type} のコース分布", "stats_zone")

# ---------- Actual iPitch input/result ----------
st.divider()
st.subheader("iPitchで指定した球を記録")

# 3x3 location picker for the actual iPitch call.
st.divider()
st.subheader("iPitchで指定した球を記録")

if not pitch_types:
    st.warning("この条件では球種データがありません。")

# 3x3 location picker for the actual iPitch call.
zone_grid = [
    [("High-Inner", "高め・内"), ("High-Middle", "高め・真ん中"), ("High-Outer", "高め・外")],
    [("Middle-Inner", "中・内"), ("Middle-Middle", "中・真ん中"), ("Middle-Outer", "中・外")],
    [("Low-Inner", "低め・内"), ("Low-Middle", "低め・真ん中"), ("Low-Outer", "低め・外")],
]

if "selected_zone" not in st.session_state:
    st.session_state.selected_zone = "Middle-Middle"

st.markdown("**コース（iPitchで指定した位置）**")
st.caption("9マスから、実際にiPitchへ入力したコースをタップ")

for r, row in enumerate(zone_grid):
    cols = st.columns(3)
    for c, (zone_value, zone_label) in enumerate(row):
        with cols[c]:
            is_selected = st.session_state.selected_zone == zone_value
            label = f"✓ {zone_label}" if is_selected else zone_label
            if st.button(
                label,
                key=f"zone_{zone_value}",
                use_container_width=True,
            ):
                st.session_state.selected_zone = zone_value
                st.rerun()

selected_zone = st.session_state.selected_zone
st.caption(f"選択中：**{dict(sum(zone_grid, []))[selected_zone]}**")
st.caption("ここで選んだ球種・コースを実際にiPitchへ入力し、投球後に結果をタップします。")

results = {
    "BallCalled": "ボール",
    "StrikeCalled": "見逃しストライク",
    "StrikeSwinging": "空振り",
    "FoulBall": "ファウル",
    "InPlay": "インプレー",
    "HitByPitch": "死球",
}
cols = st.columns(3)
for i, (result, label) in enumerate(results.items()):
    with cols[i % 3]:
        if st.button(label, key=f"result_{result}", use_container_width=True):
            st.session_state.pitch_history.append({
                "球": len(st.session_state.pitch_history) + 1,
                "カウント": f"{current_b}-{current_s}",
                "球種": selected_type,
                "コース": selected_zone,
                "結果": label,
            })
            nxt = finish_count(current_b, current_s, result)
            if nxt is None:
                st.session_state.balls = 0
                st.session_state.strikes = 0
                st.success("打席終了。「新しい打席」で次の打者へ進めます。")
            else:
                st.session_state.balls, st.session_state.strikes = nxt
            st.rerun()

# ---------- Data evidence ----------
st.divider()
st.subheader("候補生成に使った過去データ")
show_cols = [c for c in [
    "Date", "Pitcher", "Catcher", "Batter", "BatterSide", "BatterTeam",
    "Balls", "Strikes", "PitchofPA", "PitchType", "PitchCall", "PlayResult",
    "RelSpeed", "SpinRate", "InducedVertBreak", "HorzBreak", "Extension",
    "PlateLocSide", "PlateLocHeight", "VAA", "ActualZone",
] if c in candidates.columns]
st.dataframe(
    candidates[show_cols].tail(100),
    use_container_width=True,
    hide_index=True
)

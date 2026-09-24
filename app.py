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


def valid_person_name(x):
    if pd.isna(x):
        return False
    s = str(x).strip()
    # TrackMan names in this file are "First, Last".
    return bool(s) and "," in s and not s.isdigit()


def valid_numeric_id(x):
    if pd.isna(x):
        return False
    s = str(x).strip()
    return s.isdigit()


@st.cache_data
def load_data():
    sheets = pd.ExcelFile(DATA_FILE).sheet_names
    raw = []

    # Read all sheets first. We need the valid rows from other games to
    # repair the malformed Catcher columns in a few September sheets.
    for sheet in sheets:
        try:
            x = pd.read_excel(DATA_FILE, sheet_name=sheet)
            if "Pitcher" in x.columns and "Catcher" in x.columns:
                x["SourceSheet"] = sheet
                raw.append(x)
        except Exception:
            continue

    df = pd.concat(raw, ignore_index=True)

    # ---------- Build global person registries from valid rows ----------
    catcher_by_id = {}
    catcher_team_by_id = {}
    for _, r in df.iterrows():
        c = r.get("Catcher")
        cid = r.get("CatcherId")
        team = r.get("CatcherTeam")
        if valid_person_name(c) and valid_numeric_id(cid):
            key = str(int(float(cid)))
            catcher_by_id.setdefault(key, c)
            if pd.notna(team):
                catcher_team_by_id.setdefault(key, str(team).strip())

    # ---------- Repair Catcher fields ----------
    # Some sheets have a shifted/corrupted Catcher block:
    # Catcher=<CatcherId>, CatcherId=<ThrowHand>, CatcherThrows=<Team>,
    # CatcherTeam=<UUID>. In those rows, recover the catcher name from the
    # global ID registry.
    repaired_catcher = []
    repaired_catcher_id = []
    repaired_catcher_team = []

    for _, r in df.iterrows():
        c = r.get("Catcher")
        cid = r.get("CatcherId")
        cteam = r.get("CatcherTeam")

        if valid_person_name(c) and valid_numeric_id(cid):
            repaired_catcher.append(c)
            repaired_catcher_id.append(str(int(float(cid))))
            repaired_catcher_team.append(str(cteam).strip() if pd.notna(cteam) else "")
            continue

        # Corrupted pattern observed in 2026-09 sheets.
        if valid_numeric_id(c):
            inferred_id = str(int(float(c)))
            inferred_name = catcher_by_id.get(inferred_id)
            inferred_team = catcher_team_by_id.get(inferred_id)

            if inferred_name:
                repaired_catcher.append(inferred_name)
                repaired_catcher_id.append(inferred_id)
                repaired_catcher_team.append(
                    inferred_team or (
                        str(r.get("CatcherThrows")).strip()
                        if pd.notna(r.get("CatcherThrows")) else ""
                    )
                )
                continue

        # Unknown / unusable catcher row: keep blank rather than displaying an ID.
        repaired_catcher.append("")
        repaired_catcher_id.append("")
        repaired_catcher_team.append("")

    df["Catcher"] = repaired_catcher
    df["CatcherId"] = repaired_catcher_id
    df["CatcherTeam"] = repaired_catcher_team

    # ---------- Normalize core columns ----------
    for c in [
        "Pitcher", "PitcherId", "PitcherTeam",
        "Catcher", "CatcherId", "CatcherTeam",
        "Batter", "BatterId", "BatterTeam", "BatterSide",
        "TaggedPitchType", "AutoPitchType", "PitchCall",
        "PlayResult", "RunnerState",
    ]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    df["PitchType"] = df["TaggedPitchType"].where(
        df["TaggedPitchType"].ne(""), df["AutoPitchType"]
    )

    for c in ["Balls", "Strikes", "PitchofPA"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Balls"] = df["Balls"].fillna(0).astype(int)
    df["Strikes"] = df["Strikes"].fillna(0).astype(int)

    # Keep only rows with a valid pitcher/catcher relationship.
    df = df[
        df["Pitcher"].map(valid_person_name)
        & df["Catcher"].map(valid_person_name)
    ].copy()

    return df


df = load_data()


def team_label(code):
    return TEAM_NAMES.get(code, code)


def finish_count(balls, strikes, pitch_call):
    pc = str(pitch_call)
    if pc in {"InPlay", "HitByPitch"}:
        return None
    if pc == "BallCalled" or "Ball" == pc:
        return None if balls >= 3 else (balls + 1, strikes)
    if pc in {"StrikeCalled", "StrikeSwinging"}:
        return None if strikes >= 2 else (balls, strikes + 1)
    if "Foul" in pc:
        return balls, strikes if strikes >= 2 else (balls, strikes + 1)
    if "Strike" in pc:
        return None if strikes >= 2 else (balls, strikes + 1)
    return balls, strikes


def zone_label(side, height):
    try:
        s, h = float(side), float(height)
    except (TypeError, ValueError):
        return "Unknown"
    col = "Inner" if s < -0.23 else ("Middle" if s <= 0.23 else "Outer")
    row = "Low" if h < 2.15 else ("Middle" if h <= 3.15 else "High")
    return f"{row}-{col}"


# ---------- Session state ----------
defaults = {
    "balls": 0,
    "strikes": 0,
    "pitch_history": [],
}
for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ---------- Game setup ----------
st.title("⚾ Keio Pitch Calling Support")
st.caption("TrackMan過去データから、投手×捕手×打者の状況に応じて次球候補を表示")

with st.sidebar:
    st.header("Game setup")

    available_teams = [
        t for t in TEAM_ORDER
        if t in set(df["PitcherTeam"]) or t in set(df["CatcherTeam"])
    ]
    other_teams = sorted(set(df["PitcherTeam"]).union(df["CatcherTeam"]).union(df["BatterTeam"]))

    team_options = available_teams + [t for t in other_teams if t not in available_teams]
    team_code = st.selectbox(
        "大学（投手・捕手側）",
        team_options,
        format_func=team_label,
    )

    pitcher_df = df[df["PitcherTeam"].eq(team_code)]
    pitcher_options = sorted(
        pitcher_df[["Pitcher", "PitcherId"]].drop_duplicates()["Pitcher"].tolist()
    )
    if not pitcher_options:
        st.warning("この大学の投手データがありません。")
        st.stop()
    pitcher = st.selectbox("投手", pitcher_options)

    # Only catchers actually paired with this pitcher AND on the same team.
    catcher_df = pitcher_df[
        pitcher_df["CatcherTeam"].eq(team_code)
        & pitcher_df["Pitcher"].eq(pitcher)
        & pitcher_df["Catcher"].ne("")
    ]
    catcher_options = sorted(catcher_df["Catcher"].drop_duplicates().tolist())

    if not catcher_options:
        st.warning("この投手に紐づく捕手データがありません。")
        st.stop()
    catcher = st.selectbox("捕手", catcher_options)

    # Opponent is inferred from the batter data against this pitcher/catcher.
    matchup = df[
        df["Pitcher"].eq(pitcher)
        & df["Catcher"].eq(catcher)
        & df["PitcherTeam"].eq(team_code)
    ]
    opponent_options = sorted(
        [x for x in matchup["BatterTeam"].unique() if x and x != team_code]
    )
    if opponent_options:
        opponent = st.selectbox(
            "相手大学",
            opponent_options,
            format_func=team_label,
        )
        matchup = matchup[matchup["BatterTeam"].eq(opponent)]
    else:
        opponent = None

    batter_options = sorted([x for x in matchup["Batter"].unique() if x])
    if not batter_options:
        st.warning("この組み合わせの打者データがありません。")
        st.stop()

    batter = st.selectbox("打者", batter_options)

    if st.button("🔄 新しい打席", use_container_width=True):
        st.session_state.balls = 0
        st.session_state.strikes = 0
        st.session_state.pitch_history = []
        st.rerun()

base = df[
    df["Pitcher"].eq(pitcher)
    & df["Catcher"].eq(catcher)
    & df["Batter"].eq(batter)
].copy()

if base.empty:
    st.error("選択した投手・捕手・打者の組み合わせにデータがありません。")
    st.stop()

st.markdown(
    f"### {pitcher} × {catcher} × {batter}"
)
st.caption(
    f"{team_label(team_code)}  |  {team_label(base['BatterTeam'].iloc[0])}"
)
st.metric("現在のカウント", f"{st.session_state.balls} - {st.session_state.strikes}")

# ---------- Current situation ----------
current_b = st.session_state.balls
current_s = st.session_state.strikes

situation = base[
    base["Balls"].eq(current_b) & base["Strikes"].eq(current_s)
].copy()

fallback = situation.empty
if fallback:
    situation = base.copy()

if st.session_state.pitch_history:
    st.subheader("この打席の履歴")
    st.dataframe(
        pd.DataFrame(st.session_state.pitch_history),
        use_container_width=True,
        hide_index=True,
    )

st.subheader("次球候補")
if fallback:
    st.info("このカウントの過去データがないため、同じ投手×捕手×打者の全投球から候補を表示しています。")

pitch_counts = situation["PitchType"].replace("", np.nan).dropna().value_counts()
if not pitch_counts.empty:
    pitch_pct = (pitch_counts / pitch_counts.sum() * 100).round(1)
    rec = pd.DataFrame({
        "球種": pitch_counts.index,
        "過去投球数": pitch_counts.values,
        "使用率(%)": [pitch_pct[x] for x in pitch_counts.index],
    })
    st.dataframe(rec, use_container_width=True, hide_index=True)

# ---------- Actual pitch entry ----------
st.divider()
st.subheader("iPitchで指定した球を記録")

pitch_types = sorted([x for x in situation["PitchType"].dropna().unique() if str(x)])
if not pitch_types:
    pitch_types = sorted([x for x in base["PitchType"].dropna().unique() if str(x)])

selected_type = st.selectbox("球種", pitch_types)
selected_zone = st.selectbox(
    "コース",
    [
        "High-Inner", "High-Middle", "High-Outer",
        "Middle-Inner", "Middle-Middle", "Middle-Outer",
        "Low-Inner", "Low-Middle", "Low-Outer",
    ],
)

st.caption("ここではiPitchで実際に指定した球を記録します。サイン自体はiPitch側で出します。")

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

st.divider()
st.subheader("選択条件に対応する過去データ")

show_cols = [
    c for c in [
        "Date", "Pitcher", "Catcher", "Batter", "BatterSide",
        "Balls", "Strikes", "PitchType", "PitchCall", "PlayResult",
        "RelSpeed", "SpinRate", "InducedVertBreak", "HorzBreak",
        "Extension", "PlateLocSide", "PlateLocHeight", "VAA",
        "PitcherTeam", "CatcherTeam", "BatterTeam",
    ] if c in situation.columns
]
st.dataframe(situation[show_cols].tail(100), use_container_width=True, hide_index=True)

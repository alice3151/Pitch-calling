
import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path

st.set_page_config(page_title="Keio Pitch Calling Support", layout="wide")

DATA_FILE = Path(__file__).parent / "trackman.xlsx"

@st.cache_data
def load_data():
    sheets = pd.ExcelFile(DATA_FILE).sheet_names
    frames = []
    for s in sheets:
        try:
            x = pd.read_excel(DATA_FILE, sheet_name=s)
            if "Pitcher" in x.columns and "Catcher" in x.columns:
                x["SourceSheet"] = s
                frames.append(x)
        except Exception:
            pass
    df = pd.concat(frames, ignore_index=True)

    # Normalize strings
    for c in ["Pitcher","Catcher","Batter","BatterSide","TaggedPitchType",
              "AutoPitchType","PitchCall","PlayResult","RunnerState"]:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str).str.strip()

    # Use TrackMan's tagged pitch type first; fall back to auto type.
    df["PitchType"] = df["TaggedPitchType"].where(
        df["TaggedPitchType"].ne(""), df["AutoPitchType"]
    )

    # Stable pitch order inside each PA.
    if "PitchofPA" in df.columns:
        df["PitchofPA"] = pd.to_numeric(df["PitchofPA"], errors="coerce")
    if "Balls" in df.columns:
        df["Balls"] = pd.to_numeric(df["Balls"], errors="coerce").fillna(0).astype(int)
    if "Strikes" in df.columns:
        df["Strikes"] = pd.to_numeric(df["Strikes"], errors="coerce").fillna(0).astype(int)

    return df

df = load_data()

# ---------- Helpers ----------
def finish_count(balls, strikes, pitch_call):
    """Return count after the selected result. None means PA ended."""
    pc = str(pitch_call)

    if pc in {"InPlay", "HitByPitch"}:
        return None

    # Common TrackMan result labels.
    if pc == "BallCalled":
        if balls >= 3:
            return None  # walk
        return balls + 1, strikes

    if pc in {"StrikeCalled", "StrikeSwinging"}:
        if strikes >= 2:
            return None  # strikeout
        return balls, strikes + 1

    if pc in {"FoulBall", "FoulBallNotFieldable"}:
        # Foul with two strikes does not advance the count.
        return balls, min(strikes + 1, 2) if strikes < 2 else strikes

    # Generic fallbacks
    if "Foul" in pc:
        return balls, min(strikes + 1, 2) if strikes < 2 else strikes
    if "Ball" in pc:
        if balls >= 3:
            return None
        return balls + 1, strikes
    if "Strike" in pc:
        if strikes >= 2:
            return None
        return balls, strikes + 1

    return balls, strikes

def zone_label(side, height):
    if pd.isna(side) or pd.isna(height):
        return "Unknown"
    # Approximate 3x3 target grid. Thresholds can be tuned later.
    s = float(side)
    h = float(height)
    col = "Inner" if s < -0.23 else ("Middle" if s <= 0.23 else "Outer")
    row = "Low" if h < 2.15 else ("Middle" if h <= 3.15 else "High")
    return f"{row}-{col}"

def result_group(row):
    pc = str(row.get("PitchCall",""))
    if pc == "BallCalled" or pc == "Ball":
        return "Ball"
    if pc == "StrikeCalled":
        return "Called Strike"
    if pc == "StrikeSwinging":
        return "Swinging Strike"
    if "Foul" in pc:
        return "Foul"
    if pc == "InPlay":
        return "In Play"
    if pc == "HitByPitch":
        return "HBP"
    return pc or "Other"

# ---------- Sidebar ----------
st.title("⚾ Keio Pitch Calling Support")
st.caption("TrackMan過去データから、投手×捕手×打者の状況に応じて次球候補を表示する参照アプリ")

with st.sidebar:
    st.header("Game setup")
    pitchers = sorted([x for x in df["Pitcher"].unique() if x])
    pitcher = st.selectbox("投手", pitchers)

    catchers = sorted([x for x in df.loc[df["Pitcher"].eq(pitcher), "Catcher"].unique() if x])
    catcher = st.selectbox("捕手", catchers)

    batters = sorted([x for x in df.loc[
        df["Pitcher"].eq(pitcher) & df["Catcher"].eq(catcher), "Batter"
    ].unique() if x])
    batter_mode = st.radio("打者", ["打者を指定", "左右のみ"], horizontal=True)

    if batter_mode == "打者を指定":
        batter = st.selectbox("打者名", batters if batters else [""])
        batter_side = None
    else:
        batter = None
        batter_side = st.selectbox("打者左右", ["Right", "Left"])

    if st.button("🔄 新しい打席", use_container_width=True):
        st.session_state.balls = 0
        st.session_state.strikes = 0
        st.session_state.pitch_history = []
        st.session_state.current_pitch = None
        st.rerun()

# ---------- Session state ----------
for key, default in [
    ("balls", 0), ("strikes", 0), ("pitch_history", []),
    ("current_pitch", None), ("selected_pitch_type", None),
    ("selected_zone", None)
]:
    if key not in st.session_state:
        st.session_state[key] = default

# Base filter
base = df[df["Pitcher"].eq(pitcher) & df["Catcher"].eq(catcher)].copy()
if batter is not None:
    base = base[base["Batter"].eq(batter)]
elif batter_side is not None:
    base = base[base["BatterSide"].eq(batter_side)]

st.markdown(
    f"### {pitcher} × {catcher}"
    + (f" × {batter}" if batter else f" × {batter_side}打者")
)
st.metric("現在のカウント", f"{st.session_state.balls} - {st.session_state.strikes}")

if base.empty:
    st.warning("この条件のTrackManデータがありません。投手・捕手・打者条件を変更してください。")
    st.stop()

# History shown as the actual pitch-by-pitch interaction.
if st.session_state.pitch_history:
    st.subheader("この打席の履歴")
    hist = pd.DataFrame(st.session_state.pitch_history)
    st.dataframe(hist, use_container_width=True, hide_index=True)

# Current situation subset: internally determined by previous results.
current_b = st.session_state.balls
current_s = st.session_state.strikes
situation = base[
    base["Balls"].eq(current_b) & base["Strikes"].eq(current_s)
].copy()

# If no exact count sample, use all matching P/C/batter data rather than inventing.
fallback = False
if situation.empty:
    situation = base.copy()
    fallback = True

st.subheader("次球候補")
if fallback:
    st.info("このカウントの過去サンプルが少ないため、同じ投手×捕手×打者条件の全投球から候補を表示しています。")

# ---------- Pitch recommendation ----------
pitch_counts = situation["PitchType"].replace("", np.nan).dropna().value_counts()
if not pitch_counts.empty:
    pitch_pct = (pitch_counts / pitch_counts.sum() * 100).round(1)
    cols = st.columns(min(4, len(pitch_pct)))
    for i, (ptype, n) in enumerate(pitch_counts.items()):
        with cols[i % len(cols)]:
            st.metric(ptype, f"{pitch_pct[ptype]:.1f}%", f"{n}球")

    # Candidate table
    rec = pd.DataFrame({
        "球種": pitch_counts.index,
        "過去投球数": pitch_counts.values,
        "使用率(%)": [pitch_pct[x] for x in pitch_counts.index],
    })
    st.dataframe(rec, use_container_width=True, hide_index=True)

# ---------- Location recommendation ----------
loc = situation.copy()
loc["Zone"] = loc.apply(lambda r: zone_label(r.get("PlateLocSide"), r.get("PlateLocHeight")), axis=1)
loc_counts = loc["Zone"].replace("Unknown", np.nan).dropna().value_counts()

if not loc_counts.empty:
    st.markdown("#### コース傾向")
    loc_pct = (loc_counts / loc_counts.sum() * 100).round(1)
    grid = {
        "High-Inner":0, "High-Middle":0, "High-Outer":0,
        "Middle-Inner":0, "Middle-Middle":0, "Middle-Outer":0,
        "Low-Inner":0, "Low-Middle":0, "Low-Outer":0
    }
    for z, n in loc_counts.items():
        if z in grid:
            grid[z] = int(n)

    gcols = st.columns(3)
    for j, rowname in enumerate(["High","Middle","Low"]):
        with gcols[j]:
            st.write(f"**{rowname}**")
            for cname in ["Inner","Middle","Outer"]:
                z = f"{rowname}-{cname}"
                st.button(
                    f"{cname}: {grid[z]}球",
                    key=f"loc_{rowname}_{cname}",
                    use_container_width=True,
                    disabled=True
                )

# ---------- Select what was actually called ----------
st.divider()
st.subheader("iPitchで指定した球を記録")

pitch_types = sorted([x for x in situation["PitchType"].dropna().unique() if str(x)])
if not pitch_types:
    pitch_types = sorted([x for x in base["PitchType"].dropna().unique() if str(x)])

selected_type = st.selectbox("球種", pitch_types)
selected_zone = st.selectbox(
    "コース（記録用）",
    ["High-Inner","High-Middle","High-Outer",
     "Middle-Inner","Middle-Middle","Middle-Outer",
     "Low-Inner","Low-Middle","Low-Outer"]
)

st.caption("このアプリではサイン自体は出しません。ここで球種・コースを確認し、実際のiPitchで入力した球を記録します。")

result_options = [
    "BallCalled", "StrikeCalled", "StrikeSwinging",
    "FoulBall", "InPlay", "HitByPitch"
]
result_labels = {
    "BallCalled":"ボール",
    "StrikeCalled":"見逃しストライク",
    "StrikeSwinging":"空振り",
    "FoulBall":"ファウル",
    "InPlay":"インプレー",
    "HitByPitch":"死球"
}

st.markdown("#### 結果をタップ")
result_cols = st.columns(3)
for i, result in enumerate(result_options):
    with result_cols[i % 3]:
        if st.button(result_labels[result], key=f"result_{result}", use_container_width=True):
            new_entry = {
                "球": len(st.session_state.pitch_history) + 1,
                "カウント": f"{current_b}-{current_s}",
                "球種": selected_type,
                "コース": selected_zone,
                "結果": result_labels[result]
            }
            st.session_state.pitch_history.append(new_entry)

            nxt = finish_count(current_b, current_s, result)
            if nxt is None:
                st.session_state.balls = 0
                st.session_state.strikes = 0
                st.session_state.current_pitch = None
                st.success("打席終了。『新しい打席』で次の打者へ進めます。")
            else:
                st.session_state.balls, st.session_state.strikes = nxt
            st.rerun()

# ---------- Evidence ----------
st.divider()
st.subheader("この条件の過去データ")
show_cols = [c for c in [
    "Date","Pitcher","Catcher","Batter","BatterSide",
    "Balls","Strikes","PitchType","PitchCall","PlayResult",
    "RelSpeed","SpinRate","InducedVertBreak","HorzBreak",
    "Extension","PlateLocSide","PlateLocHeight","VAA"
] if c in situation.columns]
st.dataframe(situation[show_cols].tail(100), use_container_width=True, hide_index=True)

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path


# =========================================================
# Streamlit
# =========================================================

st.set_page_config(
    page_title="Keio Pitch Calling Support",
    layout="wide",
)


# =========================================================
# Responsive 3x3 pitch zone UI
# =========================================================

st.markdown(
    """
    <style>

    /* =========================================
       iPitch 3x3 zone
       ========================================= */

    .st-key-pitch-zone [data-testid="stHorizontalBlock"] {
        flex-wrap: nowrap !important;
        gap: 6px !important;
    }

    .st-key-pitch-zone [data-testid="column"] {
        min-width: 0 !important;
        width: 33.333% !important;
        flex: 1 1 0 !important;
    }

    .st-key-pitch-zone button {
        width: 100% !important;
        min-height: 0 !important;
        height: clamp(75px, 25vw, 130px) !important;
        padding: 4px !important;
        font-size: clamp(11px, 3.2vw, 17px) !important;
        line-height: 1.2 !important;
        white-space: normal !important;
    }

    /* =========================================
       Smartphone
       ========================================= */

    @media (max-width: 640px) {

        .st-key-pitch-zone [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            gap: 5px !important;
        }

        .st-key-pitch-zone button {
            height: 27vw !important;
            min-height: 75px !important;
            max-height: 105px !important;
            font-size: 12px !important;
            padding: 2px !important;
        }
    }

    /* =========================================
       Selected zone
       ========================================= */

    .selected-zone-label {
        text-align: center;
        font-size: 16px;
        font-weight: 700;
        margin-top: 10px;
        margin-bottom: 4px;
    }

    /* =========================================
       Candidate cards
       ========================================= */

    .candidate-card {
        padding: 10px 12px;
        border-radius: 10px;
        border: 1px solid rgba(128,128,128,0.25);
        margin-bottom: 8px;
    }

    .candidate-main {
        font-size: 17px;
        font-weight: 700;
    }

    .candidate-sub {
        font-size: 13px;
        opacity: 0.75;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# Data
# =========================================================

DATA_FILE = Path(__file__).parent / "trackman.xlsx"


# =========================================================
# Team names
# =========================================================

TEAM_NAMES = {
    "KEI_KEI": "Keio University",
    "RIK": "Rikkyo University",
    "MEI_MEI": "Meiji University",
    "TOK": "Hosei University",
    "WAS_EDA": "Waseda University",
    "HOS_HOS": "Hosei University",
}


TEAM_ORDER = [
    "KEI_KEI",
    "RIK",
    "MEI_MEI",
    "TOK",
    "WAS_EDA",
    "HOS_HOS",
]


# =========================================================
# Helper
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def valid_person_name(value):
    """
    正常な選手名かどうか。
    TrackManの壊れた行ではIDなどがName列に入ることがあるため除外。
    """

    value = clean_text(value)

    if not value:
        return False

    if value.replace(".", "").isdigit():
        return False

    if "," not in value:
        return False

    return True


def valid_numeric_id(value):
    """
    IDが数値として扱えるか。
    """

    value = clean_text(value)

    if not value:
        return False

    try:
        float(value)
        return True
    except Exception:
        return False


def normalize_side(value):
    """
    Right / Left を統一。
    """

    value = clean_text(value).lower()

    if value in ["right", "r", "rh", "右", "右打"]:
        return "Right"

    if value in ["left", "l", "lh", "左", "左打"]:
        return "Left"

    return ""


def result_label(value):
    """
    結果表示用。
    """

    mapping = {
        "BallCalled": "ボール",
        "Ball": "ボール",
        "StrikeCalled": "見逃しストライク",
        "StrikeSwinging": "空振り",
        "FoulBall": "ファウル",
        "Foul": "ファウル",
        "InPlay": "インプレー",
        "HitByPitch": "死球",
    }

    return mapping.get(value, value)


# =========================================================
# Zone calculation
# =========================================================

def make_actual_zone(row):
    """
    PlateLocSide / PlateLocHeight から9分割ゾーンを作る。

    Side:
        < -0.23      Inner
        -0.23〜0.23  Middle
        > 0.23       Outer

    Height:
        < 2.15       Low
        2.15〜3.15    Middle
        > 3.15       High
    """

    try:
        side = float(row["PlateLocSide"])
        height = float(row["PlateLocHeight"])
    except Exception:
        return ""

    if pd.isna(side) or pd.isna(height):
        return ""

    if side < -0.23:
        side_name = "Inner"
    elif side <= 0.23:
        side_name = "Middle"
    else:
        side_name = "Outer"

    if height < 2.15:
        height_name = "Low"
    elif height <= 3.15:
        height_name = "Middle"
    else:
        height_name = "High"

    return f"{height_name}-{side_name}"


# =========================================================
# Load data
# =========================================================

@st.cache_data
def load_data():

    if not DATA_FILE.exists():
        st.error(f"データファイルが見つかりません: {DATA_FILE}")
        return pd.DataFrame()

    try:
        sheets = pd.read_excel(
            DATA_FILE,
            sheet_name=None
        )
    except Exception as e:
        st.error(f"Excelの読み込みに失敗しました: {e}")
        return pd.DataFrame()

    frames = []

    for sheet_name, sheet_df in sheets.items():

        if sheet_df is None or sheet_df.empty:
            continue

        columns = set(sheet_df.columns)

        if "Pitcher" not in columns:
            continue

        if "Catcher" not in columns:
            continue

        temp = sheet_df.copy()
        temp["SourceSheet"] = sheet_name

        frames.append(temp)

    if not frames:
        return pd.DataFrame()

    df = pd.concat(
        frames,
        ignore_index=True
    )

    # -----------------------------------------------------
    # 必要列を作る
    # -----------------------------------------------------

    required_columns = [
        "Pitcher",
        "PitcherId",
        "PitcherTeam",
        "Catcher",
        "CatcherId",
        "CatcherTeam",
        "BatterSide",
        "PitchType",
        "TaggedPitchType",
        "AutoPitchType",
        "Balls",
        "Strikes",
        "PitchofPA",
        "GameID",
        "Inning",
        "Top/Bottom",
        "PAofInning",
        "PlateLocSide",
        "PlateLocHeight",
        "PlayResult",
    ]

    for col in required_columns:
        if col not in df.columns:
            df[col] = ""

    # -----------------------------------------------------
    # Catcher正常データから復元辞書を作る
    # -----------------------------------------------------

    catcher_lookup = {}

    for _, row in df.iterrows():

        catcher_name = clean_text(row["Catcher"])
        catcher_id = clean_text(row["CatcherId"])
        catcher_team = clean_text(row["CatcherTeam"])

        if (
            valid_person_name(catcher_name)
            and valid_numeric_id(catcher_id)
            and catcher_team
        ):
            catcher_lookup[catcher_id] = (
                catcher_name,
                catcher_team,
            )

    # -----------------------------------------------------
    # 壊れたCatcherデータを修復
    # -----------------------------------------------------

    fixed_catcher = []
    fixed_catcher_id = []
    fixed_catcher_team = []

    for _, row in df.iterrows():

        catcher_name = clean_text(row["Catcher"])
        catcher_id = clean_text(row["CatcherId"])
        catcher_team = clean_text(row["CatcherTeam"])

        # 正常
        if valid_person_name(catcher_name):

            fixed_catcher.append(catcher_name)
            fixed_catcher_id.append(catcher_id)
            fixed_catcher_team.append(catcher_team)

            continue

        # Catcher列にIDが入っている場合
        if valid_numeric_id(catcher_name):

            lookup_id = catcher_name

            if lookup_id in catcher_lookup:

                name, team = catcher_lookup[lookup_id]

                fixed_catcher.append(name)
                fixed_catcher_id.append(lookup_id)
                fixed_catcher_team.append(team)

                continue

        # CatcherIdにIDが入っている場合
        if valid_numeric_id(catcher_id):

            if catcher_id in catcher_lookup:

                name, team = catcher_lookup[catcher_id]

                fixed_catcher.append(name)
                fixed_catcher_id.append(catcher_id)
                fixed_catcher_team.append(team)

                continue

        # 復元できない
        fixed_catcher.append(catcher_name)
        fixed_catcher_id.append(catcher_id)
        fixed_catcher_team.append(catcher_team)

    df["Catcher"] = fixed_catcher
    df["CatcherId"] = fixed_catcher_id
    df["CatcherTeam"] = fixed_catcher_team

    # -----------------------------------------------------
    # 文字列列
    # -----------------------------------------------------

    text_columns = [
        "Pitcher",
        "PitcherId",
        "PitcherTeam",
        "Catcher",
        "CatcherId",
        "CatcherTeam",
        "BatterSide",
        "TaggedPitchType",
        "AutoPitchType",
        "GameID",
        "Inning",
        "Top/Bottom",
        "PAofInning",
        "PlayResult",
    ]

    for col in text_columns:
        df[col] = df[col].map(clean_text)

    # -----------------------------------------------------
    # PitchType
    # -----------------------------------------------------

    df["PitchType"] = df["TaggedPitchType"]

    empty_pitch = df["PitchType"].eq("")

    df.loc[empty_pitch, "PitchType"] = df.loc[
        empty_pitch,
        "AutoPitchType"
    ]

    # -----------------------------------------------------
    # Numeric
    # -----------------------------------------------------

    for col in [
        "Balls",
        "Strikes",
        "PitchofPA",
        "PlateLocSide",
        "PlateLocHeight",
    ]:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        )

    # -----------------------------------------------------
    # Batter side
    # -----------------------------------------------------

    df["BatterSideNormalized"] = (
        df["BatterSide"]
        .map(normalize_side)
    )

    # -----------------------------------------------------
    # ActualZone
    # -----------------------------------------------------

    df["ActualZone"] = df.apply(
        make_actual_zone,
        axis=1
    )

    # -----------------------------------------------------
    # PA key
    # -----------------------------------------------------

    df["PAKey"] = (
        df["GameID"].astype(str)
        + "|"
        + df["Inning"].astype(str)
        + "|"
        + df["Top/Bottom"].astype(str)
        + "|"
        + df["PAofInning"].astype(str)
    )

    # -----------------------------------------------------
    # Pitcher / catcher nameが正常な行だけ
    # -----------------------------------------------------

    df = df[
        df["Pitcher"].map(valid_person_name)
        &
        df["Catcher"].map(valid_person_name)
    ].copy()

    # -----------------------------------------------------
    # 並び順
    # -----------------------------------------------------

    sort_columns = []

    if "GameID" in df.columns:
        sort_columns.append("GameID")

    if "Inning" in df.columns:
        sort_columns.append("Inning")

    if "Top/Bottom" in df.columns:
        sort_columns.append("Top/Bottom")

    if "PAofInning" in df.columns:
        sort_columns.append("PAofInning")

    if "PitchofPA" in df.columns:
        sort_columns.append("PitchofPA")

    if sort_columns:
        df = df.sort_values(
            sort_columns,
            kind="stable"
        )

    return df.reset_index(drop=True)


# =========================================================
# Count update
# =========================================================

def finish_count(
    balls,
    strikes,
    result
):
    """
    投球結果から次のカウントを計算。

    戻り値:
        next_balls
        next_strikes
        plate_appearance_finished
    """

    balls = int(balls)
    strikes = int(strikes)

    result = clean_text(result)

    # ---------------------------------------------
    # 打席終了
    # ---------------------------------------------

    if result in [
        "InPlay",
        "HitByPitch",
    ]:
        return 0, 0, True

    # ---------------------------------------------
    # ボール
    # ---------------------------------------------

    if result in [
        "BallCalled",
        "Ball",
    ]:

        balls += 1

        if balls >= 4:
            return 0, 0, True

        return balls, strikes, False

    # ---------------------------------------------
    # 見逃し / 空振り
    # ---------------------------------------------

    if result in [
        "StrikeCalled",
        "StrikeSwinging",
    ]:

        strikes += 1

        if strikes >= 3:
            return 0, 0, True

        return balls, strikes, False

    # ---------------------------------------------
    # ファウル
    # ---------------------------------------------

    if result in [
        "Foul",
        "FoulBall",
    ]:

        # 2ストライクなら増えない
        if strikes >= 2:
            return balls, strikes, False

        strikes += 1

        return balls, strikes, False

    # ---------------------------------------------
    # その他
    # ---------------------------------------------

    return balls, strikes, False


# =========================================================
# Candidate data
# =========================================================

def get_candidate_data(
    base,
    current_balls,
    current_strikes,
    history,
):
    """
    現在の条件に合う過去投球から
    次球候補を作る。

    優先順位：

    1. 現在のカウント
       +
       直前の配球履歴完全一致

    2. 現在のカウント
       +
       直前1球一致

    3. 現在のカウントのみ
    """

    if base.empty:
        return base.copy()

    count_df = base[
        (base["Balls"] == current_balls)
        &
        (base["Strikes"] == current_strikes)
    ].copy()

    if count_df.empty:
        return count_df

    if not history:
        return count_df

    # -----------------------------------------------------
    # 同じPA内の次球を探す
    # -----------------------------------------------------

    rows = []

    for pa_key, pa_df in count_df.groupby("PAKey"):

        pa_df = pa_df.sort_values(
            "PitchofPA"
        )

        for _, row in pa_df.iterrows():

            pitch_no = row["PitchofPA"]

            if pd.isna(pitch_no):
                continue

            try:
                pitch_no = int(pitch_no)
            except Exception:
                continue

            if pitch_no <= len(history):
                continue

            # この球の直前までの履歴
            previous = pa_df[
                pa_df["PitchofPA"] < pitch_no
            ].sort_values("PitchofPA")

            if previous.empty:
                continue

            previous = previous.tail(
                len(history)
            )

            if len(previous) < len(history):
                continue

            match = True

            for hist, (_, prev_row) in zip(
                history,
                previous.iterrows()
            ):

                hist_pitch = clean_text(
                    hist.get("球種", "")
                )

                hist_zone = clean_text(
                    hist.get("コース", "")
                )

                hist_result = clean_text(
                    hist.get("結果", "")
                )

                prev_pitch = clean_text(
                    prev_row.get("PitchType", "")
                )

                prev_zone = clean_text(
                    prev_row.get("ActualZone", "")
                )

                prev_result = clean_text(
                    prev_row.get("PitchResult", "")
                )

                # 球種
                if (
                    hist_pitch
                    and prev_pitch
                    and hist_pitch != prev_pitch
                ):
                    match = False
                    break

                # コース
                if (
                    hist_zone
                    and prev_zone
                    and hist_zone != prev_zone
                ):
                    match = False
                    break

                # 結果
                if (
                    hist_result
                    and prev_result
                    and hist_result != prev_result
                ):
                    match = False
                    break

            if match:
                rows.append(row)

    # -----------------------------------------------------
    # 完全一致
    # -----------------------------------------------------

    if rows:

        result = pd.DataFrame(rows)

        if not result.empty:
            return result

    # -----------------------------------------------------
    # 直前1球一致
    # -----------------------------------------------------

    if history:

        last = history[-1]

        last_pitch = clean_text(
            last.get("球種", "")
        )

        last_zone = clean_text(
            last.get("コース", "")
        )

        last_result = clean_text(
            last.get("結果", "")
        )

        matched = count_df.copy()

        matched["PreviousPitchType"] = (
            matched.groupby("PAKey")["PitchType"]
            .shift(0)
        )

        # 直前球の情報をPA内で照合
        rows = []

        for pa_key, pa_df in count_df.groupby("PAKey"):

            pa_df = pa_df.sort_values("PitchofPA")

            for _, row in pa_df.iterrows():

                pitch_no = row["PitchofPA"]

                if pd.isna(pitch_no):
                    continue

                try:
                    pitch_no = int(pitch_no)
                except Exception:
                    continue

                previous = pa_df[
                    pa_df["PitchofPA"] < pitch_no
                ].sort_values("PitchofPA")

                if previous.empty:
                    continue

                prev_row = previous.iloc[-1]

                pitch_match = (
                    not last_pitch
                    or clean_text(prev_row["PitchType"])
                    == last_pitch
                )

                zone_match = (
                    not last_zone
                    or clean_text(prev_row["ActualZone"])
                    == last_zone
                )

                result_match = (
                    not last_result
                    or clean_text(prev_row["PitchResult"])
                    == last_result
                )

                if (
                    pitch_match
                    and zone_match
                    and result_match
                ):
                    rows.append(row)

        if rows:

            result = pd.DataFrame(rows)

            if not result.empty:
                return result

    # -----------------------------------------------------
    # カウントのみ
    # -----------------------------------------------------

    return count_df


# =========================================================
# Zone definitions
# =========================================================

ZONE_GRID = [
    [
        ("High-Inner", "高め・内"),
        ("High-Middle", "高め・真ん中"),
        ("High-Outer", "高め・外"),
    ],
    [
        ("Middle-Inner", "中・内"),
        ("Middle-Middle", "中・真ん中"),
        ("Middle-Outer", "中・外"),
    ],
    [
        ("Low-Inner", "低め・内"),
        ("Low-Middle", "低め・真ん中"),
        ("Low-Outer", "低め・外"),
    ],
]

ZONE_NAME_MAP = dict(
    sum(ZONE_GRID, [])
)

ZONE_ORDER = [
    "High-Inner",
    "High-Middle",
    "High-Outer",
    "Middle-Inner",
    "Middle-Middle",
    "Middle-Outer",
    "Low-Inner",
    "Low-Middle",
    "Low-Outer",
]


# =========================================================
# Load
# =========================================================

df = load_data()

if df.empty:
    st.stop()


# =========================================================
# Session state
# =========================================================

if "balls" not in st.session_state:
    st.session_state.balls = 0

if "strikes" not in st.session_state:
    st.session_state.strikes = 0

if "pitch_history" not in st.session_state:
    st.session_state.pitch_history = []

if "selected_zone" not in st.session_state:
    st.session_state.selected_zone = "Middle-Middle"


# =========================================================
# Header
# =========================================================

st.title("⚾ Keio Pitch Calling Support")

st.caption(
    "過去のTrackManデータから、投手×捕手×打者左右×カウント×配球履歴を考慮して次球候補を表示します。"
)


# =========================================================
# Sidebar
# =========================================================

st.sidebar.header("条件設定")


# ---------------------------------------------------------
# University
# ---------------------------------------------------------

available_teams = [
    code
    for code in TEAM_ORDER
    if code in set(df["PitcherTeam"])
    or code in set(df["CatcherTeam"])
]

if not available_teams:
    available_teams = sorted(
        set(
            df["PitcherTeam"].dropna()
        )
        |
        set(
            df["CatcherTeam"].dropna()
        )
    )

team_code = st.sidebar.selectbox(
    "大学（投手・捕手側）",
    available_teams,
    format_func=lambda x: TEAM_NAMES.get(x, x),
)


# ---------------------------------------------------------
# Pitcher
# ---------------------------------------------------------

pitcher_df = df[
    df["PitcherTeam"] == team_code
].copy()

pitchers = sorted(
    pitcher_df["Pitcher"]
    .dropna()
    .unique()
    .tolist()
)

if not pitchers:
    st.warning(
        "この大学の投手データがありません。"
    )
    st.stop()

pitcher = st.sidebar.selectbox(
    "投手",
    pitchers,
)


# ---------------------------------------------------------
# Catcher
# ---------------------------------------------------------

catcher_df = pitcher_df[
    pitcher_df["Pitcher"] == pitcher
].copy()

catcher_df = catcher_df[
    catcher_df["CatcherTeam"] == team_code
]

catchers = sorted(
    catcher_df["Catcher"]
    .dropna()
    .unique()
    .tolist()
)

if not catchers:
    st.warning(
        "この投手と組んでいる捕手データがありません。"
    )
    st.stop()

catcher = st.sidebar.selectbox(
    "捕手",
    catchers,
)


# ---------------------------------------------------------
# Opponent
# ---------------------------------------------------------

opponent_candidates = sorted(
    set(
        df[
            df["PitcherTeam"] == team_code
        ]["CatcherTeam"]
        .dropna()
        .tolist()
    )
)

opponent_candidates = [
    x
    for x in opponent_candidates
    if x != team_code
]

opponent_options = ["すべて"] + opponent_candidates

opponent = st.sidebar.selectbox(
    "相手大学",
    opponent_options,
    format_func=lambda x: (
        "すべて"
        if x == "すべて"
        else TEAM_NAMES.get(x, x)
    ),
)


# ---------------------------------------------------------
# Batter side
# ---------------------------------------------------------

batter_side = st.sidebar.radio(
    "打者",
    ["Right", "Left"],
    format_func=lambda x: (
        "右打者"
        if x == "Right"
        else "左打者"
    ),
)


# ---------------------------------------------------------
# Reset
# ---------------------------------------------------------

if st.sidebar.button(
    "この打席をリセット",
    use_container_width=True,
):

    st.session_state.balls = 0
    st.session_state.strikes = 0
    st.session_state.pitch_history = []
    st.session_state.selected_zone = "Middle-Middle"

    st.rerun()


# =========================================================
# Current condition
# =========================================================

st.subheader("現在の状況")

count_col1, count_col2, count_col3 = st.columns(3)

with count_col1:
    st.metric(
        "投手",
        pitcher,
    )

with count_col2:
    st.metric(
        "捕手",
        catcher,
    )

with count_col3:
    st.metric(
        "カウント",
        f"{st.session_state.balls}-{st.session_state.strikes}",
    )


st.write(
    f"**打者：** {'右打者' if batter_side == 'Right' else '左打者'}"
)

if opponent != "すべて":
    st.write(
        f"**相手：** {TEAM_NAMES.get(opponent, opponent)}"
    )


# =========================================================
# Base data
# =========================================================

base = df[
    (df["Pitcher"] == pitcher)
    &
    (df["Catcher"] == catcher)
    &
    (df["BatterSideNormalized"] == batter_side)
].copy()


if opponent != "すべて":
    base = base[
        base["CatcherTeam"] == opponent
    ].copy()


# =========================================================
# Current candidates
# =========================================================

candidates = get_candidate_data(
    base=base,
    current_balls=st.session_state.balls,
    current_strikes=st.session_state.strikes,
    history=st.session_state.pitch_history,
)


# =========================================================
# Overall pitch usage
# =========================================================

st.subheader("過去の球種傾向")

if base.empty:

    st.warning(
        "現在の条件に一致する過去データがありません。"
    )

else:

    pitch_counts = (
        base["PitchType"]
        .replace("", np.nan)
        .dropna()
        .value_counts()
    )

    if not pitch_counts.empty:

        usage_df = pitch_counts.rename(
            "投球数"
        ).reset_index()

        usage_df.columns = [
            "球種",
            "投球数",
        ]

        usage_df["使用率"] = (
            usage_df["投球数"]
            / usage_df["投球数"].sum()
            * 100
        ).round(1)

        st.dataframe(
            usage_df,
            use_container_width=True,
            hide_index=True,
        )

    else:

        st.info(
            "球種データがありません。"
        )


# =========================================================
# Current count candidate distribution
# =========================================================

st.subheader(
    f"このカウントでの次球候補"
)

if candidates.empty:

    st.info(
        "この条件・カウントに一致する過去データがありません。"
    )

else:

    candidate_pairs = (
        candidates[
            candidates["PitchType"].ne("")
        ]
        .groupby(
            ["PitchType", "ActualZone"],
            dropna=False
        )
        .size()
        .reset_index(name="投球数")
        .sort_values(
            "投球数",
            ascending=False
        )
    )

    if not candidate_pairs.empty:

        total_candidates = (
            candidate_pairs["投球数"].sum()
        )

        candidate_pairs["割合"] = (
            candidate_pairs["投球数"]
            / total_candidates
            * 100
        ).round(1)

        candidate_pairs = candidate_pairs.head(10)

        for rank, (_, row) in enumerate(
            candidate_pairs.iterrows(),
            start=1,
        ):

            pitch_type = clean_text(
                row["PitchType"]
            )

            zone = clean_text(
                row["ActualZone"]
            )

            zone_label = (
                ZONE_NAME_MAP.get(
                    zone,
                    zone
                )
            )

            count = int(
                row["投球数"]
            )

            percentage = float(
                row["割合"]
            )

            st.markdown(
                f"""
                <div class="candidate-card">
                    <div class="candidate-main">
                        {rank}. {pitch_type} × {zone_label}
                    </div>
                    <div class="candidate-sub">
                        {count}球 / {percentage:.1f}%
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================
# Location tendency by pitch type
# =========================================================

st.subheader("球種別コース分布")

pitch_types = sorted(
    [
        x
        for x in base["PitchType"]
        .dropna()
        .unique()
        if clean_text(x)
    ]
)

if not pitch_types:

    st.info(
        "この条件では球種データがありません。"
    )

else:

    selected_analysis_type = st.selectbox(
        "球種",
        pitch_types,
        key="analysis_pitch_type",
    )

    location_df = base[
        base["PitchType"]
        == selected_analysis_type
    ].copy()

    zone_counts = (
        location_df["ActualZone"]
        .value_counts()
        .reindex(
            ZONE_ORDER,
            fill_value=0
        )
    )

    zone_total = zone_counts.sum()

    # -----------------------------------------------------
    # 9 zone display
    # -----------------------------------------------------

    location_rows = []

    for row_index in range(3):

        cols = st.columns(3)

        for col_index in range(3):

            zone = ZONE_GRID[
                row_index
            ][col_index][0]

            label = ZONE_GRID[
                row_index
            ][col_index][1]

            count = int(
                zone_counts.get(
                    zone,
                    0
                )
            )

            percentage = (
                count / zone_total * 100
                if zone_total > 0
                else 0
            )

            with cols[col_index]:

                st.metric(
                    label,
                    f"{percentage:.1f}%",
                    f"{count}球",
                )


# =========================================================
# iPitch actual input
# =========================================================

st.divider()

st.subheader(
    "iPitchで指定した球を記録"
)

st.markdown(
    "**① 球種を選択 → ② コースを3×3からタップ → "
    "③ iPitchに入力 → ④ 実際の結果をタップ**"
)


# =========================================================
# Pitch type
# =========================================================

if not pitch_types:

    st.warning(
        "この条件では球種データがありません。"
    )

    selected_type = ""

else:

    selected_type = st.selectbox(
        "iPitchで指定する球種",
        pitch_types,
        key="ipitch_pitch_type",
    )


# =========================================================
# 3x3 Zone
# =========================================================

st.markdown("**コース**")

with st.container(
    border=True,
    key="pitch-zone",
):

    for row in ZONE_GRID:

        cols = st.columns(
            3,
            gap="small"
        )

        for col, (
            zone_value,
            zone_label,
        ) in zip(
            cols,
            row,
        ):

            with col:

                is_selected = (
                    st.session_state.selected_zone
                    == zone_value
                )

                if is_selected:

                    button_label = (
                        f"✓\n{zone_label}"
                    )

                else:

                    button_label = (
                        zone_label
                    )

                if st.button(
                    button_label,
                    key=f"zone_{zone_value}",
                    use_container_width=True,
                ):

                    st.session_state.selected_zone = (
                        zone_value
                    )

                    st.rerun()

    st.markdown(
        f"""
        <div class="selected-zone-label">
            選択中：
            {ZONE_NAME_MAP[
                st.session_state.selected_zone
            ]}
        </div>
        """,
        unsafe_allow_html=True,
    )


st.caption(
    "選択した球種・コースをiPitchへ入力してから、"
    "実際の投球結果を下からタップしてください。"
)


# =========================================================
# Pitch summary
# =========================================================

st.markdown("### 今回の指定")

summary_col1, summary_col2 = st.columns(2)

with summary_col1:

    st.info(
        f"**球種**\n\n{selected_type}"
    )

with summary_col2:

    st.info(
        f"**コース**\n\n"
        f"{ZONE_NAME_MAP[st.session_state.selected_zone]}"
    )


# =========================================================
# Result input
# =========================================================

st.markdown("### 実際の投球結果")

result_options = [
    ("BallCalled", "ボール"),
    ("StrikeCalled", "見逃しストライク"),
    ("StrikeSwinging", "空振り"),
    ("FoulBall", "ファウル"),
    ("InPlay", "インプレー"),
    ("HitByPitch", "死球"),
]


result_cols = st.columns(3)

for i, (
    result_value,
    result_text,
) in enumerate(result_options):

    with result_cols[i % 3]:

        if st.button(
            result_text,
            key=f"result_{result_value}",
            use_container_width=True,
        ):

            # ---------------------------------------------
            # 履歴に追加
            # ---------------------------------------------

            st.session_state.pitch_history.append(
                {
                    "球種": selected_type,
                    "コース": st.session_state.selected_zone,
                    "結果": result_value,
                    "Balls": st.session_state.balls,
                    "Strikes": st.session_state.strikes,
                }
            )

            # ---------------------------------------------
            # カウント更新
            # ---------------------------------------------

            (
                next_balls,
                next_strikes,
                pa_finished,
            ) = finish_count(
                st.session_state.balls,
                st.session_state.strikes,
                result_value,
            )

            # ---------------------------------------------
            # 打席終了
            # ---------------------------------------------

            if pa_finished:

                st.session_state.balls = 0
                st.session_state.strikes = 0
                st.session_state.pitch_history = []

            # ---------------------------------------------
            # 継続
            # ---------------------------------------------

            else:

                st.session_state.balls = (
                    next_balls
                )

                st.session_state.strikes = (
                    next_strikes
                )

            st.rerun()


# =========================================================
# Current pitch history
# =========================================================

if st.session_state.pitch_history:

    st.divider()

    st.subheader(
        "この打席の配球履歴"
    )

    history_rows = []

    for i, item in enumerate(
        st.session_state.pitch_history,
        start=1,
    ):

        history_rows.append(
            {
                "球": i,
                "カウント": (
                    f"{item['Balls']}-"
                    f"{item['Strikes']}"
                ),
                "球種": item["球種"],
                "コース": ZONE_NAME_MAP.get(
                    item["コース"],
                    item["コース"],
                ),
                "結果": result_label(
                    item["結果"]
                ),
            }
        )

    history_df = pd.DataFrame(
        history_rows
    )

    st.dataframe(
        history_df,
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# Candidate evidence
# =========================================================

if not candidates.empty:

    st.divider()

    st.subheader(
        "候補球の根拠データ"
    )

    show_cols = [
        "GameID",
        "Inning",
        "Top/Bottom",
        "PAofInning",
        "PitchofPA",
        "Balls",
        "Strikes",
        "PitchType",
        "ActualZone",
        "PlayResult",
    ]

    show_cols = [
        col
        for col in show_cols
        if col in candidates.columns
    ]

    st.dataframe(
        candidates[
            show_cols
        ].tail(100),
        use_container_width=True,
        hide_index=True,
    )

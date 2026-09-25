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
# Responsive pitch-zone UI
# =========================================================

st.markdown(
    """
    <style>

    /* =========================================
       Pitch zone container
       ========================================= */

    .st-key-pitch-zone {
        width: min(100%, 360px) !important;
        max-width: 360px !important;
        margin-left: auto !important;
        margin-right: auto !important;
    }

    .st-key-pitch-zone [data-testid="column"] {
        min-width: 0 !important;
    }

    /* =========================================
       Pitch zone buttons
       ========================================= */

    .st-key-pitch-zone button {
        width: 100% !important;
        aspect-ratio: 1 / 1 !important;
        min-height: 0 !important;
        height: auto !important;
        padding: 2px !important;

        font-size: 13px !important;
        line-height: 1.15 !important;

        white-space: normal !important;
        overflow: hidden !important;
    }

    /* =========================================
       Selected zone
       ========================================= */

    .selected-zone-label {
        text-align: center;
        font-size: 15px;
        font-weight: 700;
        margin-top: 10px;
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

    /* =========================================
       Smartphone
       ========================================= */

    @media (max-width: 640px) {

        .st-key-pitch-zone {
            width: min(calc(100vw - 40px), 340px) !important;
            max-width: 340px !important;
        }

        .st-key-pitch-zone button {
            font-size: 12px !important;
        }
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
# Zone definitions
# =========================================================

ZONE_GRID = [
    [
        ("High-Inner", "高・内"),
        ("High-Middle", "高・中"),
        ("High-Outer", "高・外"),
    ],
    [
        ("Middle-Inner", "中・内"),
        ("Middle-Middle", "中・中"),
        ("Middle-Outer", "中・外"),
    ],
    [
        ("Low-Inner", "低・内"),
        ("Low-Middle", "低・中"),
        ("Low-Outer", "低・外"),
    ],
]


ZONE_NAME_MAP = dict(
    sum(ZONE_GRID, [])
)


ZONE_ORDER = [
    zone
    for row in ZONE_GRID
    for zone, _ in row
]


# =========================================================
# Helper functions
# =========================================================

def clean_text(value):
    if pd.isna(value):
        return ""

    return str(value).strip()


def valid_person_name(value):
    """
    正常な選手名かどうか判定。
    TrackManの壊れた行ではIDなどが名前列に入ることがある。
    """

    value = clean_text(value)

    if not value:
        return False

    if value.replace(".", "").isdigit():
        return False

    return "," in value


def valid_numeric_id(value):
    value = clean_text(value)

    if not value:
        return False

    try:
        float(value)
        return True

    except Exception:
        return False


def normalize_side(value):
    value = clean_text(value).lower()

    if value in {
        "right",
        "r",
        "rh",
        "右",
        "右打",
    }:
        return "Right"

    if value in {
        "left",
        "l",
        "lh",
        "左",
        "左打",
    }:
        return "Left"

    return ""


def result_label(value):
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

    return mapping.get(
        value,
        value,
    )


def first_existing_column(
    df,
    names,
):
    """
    複数候補の中から、DataFrameに存在する最初の列を返す。
    """

    for name in names:

        if name in df.columns:
            return name

    return None


# =========================================================
# Actual zone calculation
# =========================================================

def make_actual_zone(row):
    """
    PlateLocSide / PlateLocHeight から9分割ゾーンを作る。
    """

    try:

        side = float(
            row.get(
                "PlateLocSide",
                np.nan,
            )
        )

        height = float(
            row.get(
                "PlateLocHeight",
                np.nan,
            )
        )

    except Exception:

        return ""

    if pd.isna(side) or pd.isna(height):
        return ""

    # ---------------------------------------------
    # 左右
    # ---------------------------------------------

    if side < -0.23:

        side_name = "Inner"

    elif side <= 0.23:

        side_name = "Middle"

    else:

        side_name = "Outer"

    # ---------------------------------------------
    # 高低
    # ---------------------------------------------

    if height < 2.15:

        height_name = "Low"

    elif height <= 3.15:

        height_name = "Middle"

    else:

        height_name = "High"

    return (
        f"{height_name}-{side_name}"
    )


# =========================================================
# Load TrackMan data
# =========================================================

@st.cache_data
def load_data():

    if not DATA_FILE.exists():

        st.error(
            f"データファイルが見つかりません: {DATA_FILE}"
        )

        return pd.DataFrame()

    try:

        sheets = pd.read_excel(
            DATA_FILE,
            sheet_name=None,
        )

    except Exception as e:

        st.error(
            f"Excelの読み込みに失敗しました: {e}"
        )

        return pd.DataFrame()

    frames = []

    # =====================================================
    # 全シート読み込み
    # =====================================================

    for sheet_name, sheet_df in sheets.items():

        if sheet_df is None or sheet_df.empty:
            continue

        if "Pitcher" not in sheet_df.columns:
            continue

        if "Catcher" not in sheet_df.columns:
            continue

        temp = sheet_df.copy()

        temp["SourceSheet"] = sheet_name

        frames.append(temp)

    if not frames:

        return pd.DataFrame()

    df = pd.concat(
        frames,
        ignore_index=True,
    )

    # =====================================================
    # 必要列
    # =====================================================

    required_columns = [
        "Pitcher",
        "PitcherId",
        "PitcherTeam",

        "Catcher",
        "CatcherId",
        "CatcherTeam",

        "BatterSide",

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


    # =====================================================
    # Catcher正常データから
    # ID → 名前 / 大学
    # =====================================================

    catcher_lookup = {}

    for _, row in df.iterrows():

        catcher_name = clean_text(
            row["Catcher"]
        )

        catcher_id = clean_text(
            row["CatcherId"]
        )

        catcher_team = clean_text(
            row["CatcherTeam"]
        )

        if (
            valid_person_name(catcher_name)
            and valid_numeric_id(catcher_id)
            and catcher_team
        ):

            catcher_lookup[
                catcher_id
            ] = (
                catcher_name,
                catcher_team,
            )


    # =====================================================
    # 壊れたCatcherデータを修復
    # =====================================================

    fixed_names = []
    fixed_ids = []
    fixed_teams = []

    for _, row in df.iterrows():

        catcher_name = clean_text(
            row["Catcher"]
        )

        catcher_id = clean_text(
            row["CatcherId"]
        )

        catcher_team = clean_text(
            row["CatcherTeam"]
        )

        # ---------------------------------------------
        # 正常
        # ---------------------------------------------

        if valid_person_name(
            catcher_name
        ):

            fixed_names.append(
                catcher_name
            )

            fixed_ids.append(
                catcher_id
            )

            fixed_teams.append(
                catcher_team
            )

            continue


        # ---------------------------------------------
        # Catcher列にID
        # ---------------------------------------------

        if valid_numeric_id(
            catcher_name
        ):

            lookup_id = catcher_name

            if lookup_id in catcher_lookup:

                name, team = catcher_lookup[
                    lookup_id
                ]

                fixed_names.append(name)
                fixed_ids.append(lookup_id)
                fixed_teams.append(team)

                continue


        # ---------------------------------------------
        # CatcherIdにID
        # ---------------------------------------------

        if valid_numeric_id(
            catcher_id
        ):

            if catcher_id in catcher_lookup:

                name, team = catcher_lookup[
                    catcher_id
                ]

                fixed_names.append(name)
                fixed_ids.append(catcher_id)
                fixed_teams.append(team)

                continue


        # ---------------------------------------------
        # 復元不能
        # ---------------------------------------------

        fixed_names.append(
            catcher_name
        )

        fixed_ids.append(
            catcher_id
        )

        fixed_teams.append(
            catcher_team
        )


    df["Catcher"] = fixed_names
    df["CatcherId"] = fixed_ids
    df["CatcherTeam"] = fixed_teams


    # =====================================================
    # Text columns
    # =====================================================

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

        df[col] = df[col].map(
            clean_text
        )


    # =====================================================
    # PitchType
    # =====================================================

    df["PitchType"] = (
        df["TaggedPitchType"]
    )

    empty_pitch = (
        df["PitchType"] == ""
    )

    df.loc[
        empty_pitch,
        "PitchType"
    ] = df.loc[
        empty_pitch,
        "AutoPitchType"
    ]


    # =====================================================
    # Numeric
    # =====================================================

    numeric_columns = [
        "Balls",
        "Strikes",
        "PitchofPA",
        "PlateLocSide",
        "PlateLocHeight",
    ]

    for col in numeric_columns:

        df[col] = pd.to_numeric(
            df[col],
            errors="coerce",
        )


    # =====================================================
    # Batter side
    # =====================================================

    df["BatterSideNormalized"] = (
        df["BatterSide"].map(
            normalize_side
        )
    )


    # =====================================================
    # Actual zone
    # =====================================================

    df["ActualZone"] = df.apply(
        make_actual_zone,
        axis=1,
    )


    # =====================================================
    # PA key
    # =====================================================

    df["PAKey"] = (
        df["GameID"].astype(str)
        + "|"
        + df["Inning"].astype(str)
        + "|"
        + df["Top/Bottom"].astype(str)
        + "|"
        + df["PAofInning"].astype(str)
    )


    # =====================================================
    # 正常なPitcher / Catcherだけ残す
    # =====================================================

    df = df[
        df["Pitcher"].map(
            valid_person_name
        )
        &
        df["Catcher"].map(
            valid_person_name
        )
    ].copy()


    # =====================================================
    # Sort
    # =====================================================

    sort_columns = [
        col
        for col in [
            "GameID",
            "Inning",
            "Top/Bottom",
            "PAofInning",
            "PitchofPA",
        ]
        if col in df.columns
    ]

    if sort_columns:

        df = df.sort_values(
            sort_columns,
            kind="stable",
        )


    return df.reset_index(
        drop=True
    )


# =========================================================
# Count update
# =========================================================

def finish_count(
    balls,
    strikes,
    result,
):
    """
    投球結果から次のカウントを計算。
    """

    balls = int(balls)
    strikes = int(strikes)

    result = clean_text(
        result
    )


    # =====================================================
    # 打席終了
    # =====================================================

    if result in [
        "InPlay",
        "HitByPitch",
    ]:

        return (
            0,
            0,
            True,
        )


    # =====================================================
    # Ball
    # =====================================================

    if result in [
        "BallCalled",
        "Ball",
    ]:

        balls += 1

        if balls >= 4:

            return (
                0,
                0,
                True,
            )

        return (
            balls,
            strikes,
            False,
        )


    # =====================================================
    # Strike
    # =====================================================

    if result in [
        "StrikeCalled",
        "StrikeSwinging",
    ]:

        strikes += 1

        if strikes >= 3:

            return (
                0,
                0,
                True,
            )

        return (
            balls,
            strikes,
            False,
        )


    # =====================================================
    # Foul
    # =====================================================

    if result in [
        "Foul",
        "FoulBall",
    ]:

        # 2ストライク後は据え置き
        if strikes >= 2:

            return (
                balls,
                strikes,
                False,
            )

        strikes += 1

        return (
            balls,
            strikes,
            False,
        )


    return (
        balls,
        strikes,
        False,
    )


# =========================================================
# IMPORTANT:
# TrackMan result column compatibility
# =========================================================

def get_previous_result(row):
    """
    過去の投球結果を取得する。

    現在のTrackManデータでは PlayResult を使用。
    古いコードで使っていた PitchResult を直接参照しない。

    これによって KeyError を防ぐ。
    """

    if "PlayResult" in row.index:

        return clean_text(
            row["PlayResult"]
        )

    if "PitchResult" in row.index:

        return clean_text(
            row["PitchResult"]
        )

    if "Result" in row.index:

        return clean_text(
            row["Result"]
        )

    return ""


# =========================================================
# History matching
# =========================================================

def history_matches(
    previous_rows,
    history,
):
    """
    現在の配球履歴と過去PAの直前履歴を比較。
    """

    if not history:

        return False

    if len(previous_rows) < len(history):

        return False

    previous_rows = previous_rows.tail(
        len(history)
    )


    for hist, (_, prev_row) in zip(
        history,
        previous_rows.iterrows(),
    ):

        hist_pitch = clean_text(
            hist.get(
                "球種",
                "",
            )
        )

        hist_zone = clean_text(
            hist.get(
                "コース",
                "",
            )
        )

        hist_result = clean_text(
            hist.get(
                "結果",
                "",
            )
        )


        prev_pitch = clean_text(
            prev_row.get(
                "PitchType",
                "",
            )
        )

        prev_zone = clean_text(
            prev_row.get(
                "ActualZone",
                "",
            )
        )

        # ここが今回のKeyError対策
        prev_result = get_previous_result(
            prev_row
        )


        # ---------------------------------------------
        # 球種
        # ---------------------------------------------

        if (
            hist_pitch
            and prev_pitch
            and hist_pitch != prev_pitch
        ):

            return False


        # ---------------------------------------------
        # コース
        # ---------------------------------------------

        if (
            hist_zone
            and prev_zone
            and hist_zone != prev_zone
        ):

            return False


        # ---------------------------------------------
        # 結果
        # ---------------------------------------------

        if (
            hist_result
            and prev_result
            and hist_result != prev_result
        ):

            return False


    return True


# =========================================================
# Candidate generation
# =========================================================

def get_candidate_data(
    base,
    current_balls,
    current_strikes,
    history,
):
    """
    次球候補を作る。

    優先順位：

    1. 現在カウント
       +
       直前までの配球履歴完全一致

    2. 現在カウント
       +
       直前1球一致

    3. 現在カウントのみ
    """

    if base.empty:

        return base.copy()


    # =====================================================
    # Current count
    # =====================================================

    count_df = base[
        (base["Balls"] == current_balls)
        &
        (base["Strikes"] == current_strikes)
    ].copy()


    if count_df.empty:

        return count_df


    # 履歴なしならカウントだけ
    if not history:

        return count_df


    # =====================================================
    # 完全一致
    # =====================================================

    full_matches = []


    for _, pa_df in count_df.groupby(
        "PAKey",
        sort=False,
    ):

        pa_df = pa_df.sort_values(
            "PitchofPA"
        )


        for _, row in pa_df.iterrows():

            pitch_no = row[
                "PitchofPA"
            ]


            if pd.isna(pitch_no):

                continue


            try:

                pitch_no = int(
                    pitch_no
                )

            except (
                TypeError,
                ValueError,
            ):

                continue


            previous = pa_df[
                pa_df["PitchofPA"]
                < pitch_no
            ].sort_values(
                "PitchofPA"
            )


            if history_matches(
                previous,
                history,
            ):

                full_matches.append(
                    row
                )


    if full_matches:

        return pd.DataFrame(
            full_matches
        ).reset_index(
            drop=True
        )


    # =====================================================
    # 直前1球一致
    # =====================================================

    last = history[-1]


    last_pitch = clean_text(
        last.get(
            "球種",
            "",
        )
    )

    last_zone = clean_text(
        last.get(
            "コース",
            "",
        )
    )

    last_result = clean_text(
        last.get(
            "結果",
            "",
        )
    )


    one_pitch_matches = []


    for _, pa_df in count_df.groupby(
        "PAKey",
        sort=False,
    ):

        pa_df = pa_df.sort_values(
            "PitchofPA"
        )


        for _, row in pa_df.iterrows():

            pitch_no = row[
                "PitchofPA"
            ]


            if pd.isna(pitch_no):

                continue


            try:

                pitch_no = int(
                    pitch_no
                )

            except (
                TypeError,
                ValueError,
            ):

                continue


            previous = pa_df[
                pa_df["PitchofPA"]
                < pitch_no
            ].sort_values(
                "PitchofPA"
            )


            if previous.empty:

                continue


            prev_row = previous.iloc[-1]


            prev_pitch = clean_text(
                prev_row.get(
                    "PitchType",
                    "",
                )
            )

            prev_zone = clean_text(
                prev_row.get(
                    "ActualZone",
                    "",
                )
            )

            # ★ ここもPitchResultを直接参照しない
            prev_result = get_previous_result(
                prev_row
            )


            if (
                last_pitch
                and prev_pitch
                and last_pitch != prev_pitch
            ):

                continue


            if (
                last_zone
                and prev_zone
                and last_zone != prev_zone
            ):

                continue


            if (
                last_result
                and prev_result
                and last_result != prev_result
            ):

                continue


            one_pitch_matches.append(
                row
            )


    if one_pitch_matches:

        return pd.DataFrame(
            one_pitch_matches
        ).reset_index(
            drop=True
        )


    # =====================================================
    # カウントのみ
    # =====================================================

    return count_df.reset_index(
        drop=True
    )


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

    st.session_state.selected_zone = (
        "Middle-Middle"
    )


# =========================================================
# Header
# =========================================================

st.title(
    "⚾ Keio Pitch Calling Support"
)

st.caption(
    "投手×捕手×打者左右×カウント×配球履歴から次球候補を表示"
)


# =========================================================
# Sidebar
# =========================================================

st.sidebar.header(
    "条件設定"
)


# =========================================================
# University
# =========================================================

available_teams = [
    code
    for code in TEAM_ORDER
    if (
        code in set(
            df["PitcherTeam"]
        )
        or
        code in set(
            df["CatcherTeam"]
        )
    )
]


if not available_teams:

    available_teams = sorted(
        set(
            df["PitcherTeam"]
            .dropna()
        )
        |
        set(
            df["CatcherTeam"]
            .dropna()
        )
    )


team_code = st.sidebar.selectbox(
    "大学（投手・捕手側）",
    available_teams,
    format_func=lambda x:
        TEAM_NAMES.get(
            x,
            x,
        ),
)


# =========================================================
# Pitcher
# =========================================================

pitcher_df = df[
    df["PitcherTeam"]
    == team_code
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


# =========================================================
# Catcher
# =========================================================

catcher_df = pitcher_df[
    (
        pitcher_df["Pitcher"]
        == pitcher
    )
    &
    (
        pitcher_df["CatcherTeam"]
        == team_code
    )
].copy()


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


# =========================================================
# Opponent
# =========================================================

opponent_col = first_existing_column(
    df,
    [
        "BatterTeam",
        "BattingTeam",
        "OpponentTeam",
        "Opponent",
        "HitterTeam",
    ],
)


if opponent_col:

    opponent_candidates = sorted(
        x
        for x in (
            df[opponent_col]
            .dropna()
            .map(clean_text)
            .unique()
        )
        if x
        and x != team_code
    )

else:

    opponent_candidates = sorted(
        x
        for x in (
            df["CatcherTeam"]
            .dropna()
            .map(clean_text)
            .unique()
        )
        if x
        and x != team_code
    )


opponent_options = [
    "すべて"
] + opponent_candidates


opponent = st.sidebar.selectbox(
    "相手大学",
    opponent_options,
    format_func=lambda x:
        (
            "すべて"
            if x == "すべて"
            else TEAM_NAMES.get(
                x,
                x,
            )
        ),
)


# =========================================================
# Batter side
# =========================================================

batter_side = st.sidebar.radio(
    "打者",
    [
        "Right",
        "Left",
    ],
    format_func=lambda x:
        (
            "右打者"
            if x == "Right"
            else "左打者"
        ),
)


# =========================================================
# Reset
# =========================================================

if st.sidebar.button(
    "この打席をリセット",
    use_container_width=True,
):

    st.session_state.balls = 0
    st.session_state.strikes = 0

    st.session_state.pitch_history = []

    st.session_state.selected_zone = (
        "Middle-Middle"
    )

    st.rerun()


# =========================================================
# Current condition
# =========================================================

st.subheader(
    "現在の状況"
)


col1, col2, col3 = st.columns(3)


with col1:

    st.metric(
        "投手",
        pitcher,
    )


with col2:

    st.metric(
        "捕手",
        catcher,
    )


with col3:

    st.metric(
        "カウント",
        (
            f"{st.session_state.balls}"
            "-"
            f"{st.session_state.strikes}"
        ),
    )


st.write(
    "**打者：** "
    +
    (
        "右打者"
        if batter_side == "Right"
        else "左打者"
    )
)


if opponent != "すべて":

    st.write(
        "**相手：** "
        +
        TEAM_NAMES.get(
            opponent,
            opponent,
        )
    )


# =========================================================
# Base data
# =========================================================

base = df[
    (df["Pitcher"] == pitcher)
    &
    (df["Catcher"] == catcher)
    &
    (
        df["BatterSideNormalized"]
        == batter_side
    )
].copy()


# =========================================================
# Opponent filter
# =========================================================

if opponent != "すべて":

    if opponent_col:

        base = base[
            base[
                opponent_col
            ].map(
                clean_text
            )
            == opponent
        ].copy()

    else:

        # 相手大学列が存在しないデータで
        # 誤った絞り込みをしない
        base = base.iloc[0:0].copy()


# =========================================================
# Candidates
# =========================================================

candidates = get_candidate_data(
    base=base,
    current_balls=(
        st.session_state.balls
    ),
    current_strikes=(
        st.session_state.strikes
    ),
    history=(
        st.session_state.pitch_history
    ),
)


# =========================================================
# Pitch usage
# =========================================================

st.subheader(
    "過去の球種傾向"
)


if base.empty:

    st.warning(
        "現在の条件に一致する過去データがありません。"
    )

else:

    pitch_counts = (
        base["PitchType"]
        .replace(
            "",
            np.nan,
        )
        .dropna()
        .value_counts()
    )


    if not pitch_counts.empty:

        usage_df = (
            pitch_counts
            .rename("投球数")
            .reset_index()
        )


        usage_df.columns = [
            "球種",
            "投球数",
        ]


        usage_df["使用率"] = (
            usage_df["投球数"]
            /
            usage_df["投球数"].sum()
            *
            100
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
# Next pitch candidates
# =========================================================

st.subheader(
    "このカウントでの次球候補"
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
            [
                "PitchType",
                "ActualZone",
            ],
            dropna=False,
        )
        .size()
        .reset_index(
            name="投球数"
        )
        .sort_values(
            "投球数",
            ascending=False,
        )
    )


    if not candidate_pairs.empty:

        total_candidates = (
            candidate_pairs[
                "投球数"
            ].sum()
        )


        candidate_pairs["割合"] = (
            candidate_pairs["投球数"]
            /
            total_candidates
            *
            100
        ).round(1)


        for rank, (
            _,
            row,
        ) in enumerate(
            candidate_pairs.head(
                10
            ).iterrows(),
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
                    zone,
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
                        {rank}.
                        {pitch_type}
                        ×
                        {zone_label}
                    </div>

                    <div class="candidate-sub">
                        {count}球 /
                        {percentage:.1f}%
                    </div>

                </div>
                """,
                unsafe_allow_html=True,
            )


# =========================================================
# Pitch type location tendency
# =========================================================

st.subheader(
    "球種別コース分布"
)


pitch_types = sorted(
    [
        x
        for x in (
            base["PitchType"]
            .dropna()
            .unique()
        )
        if clean_text(x)
    ]
)


if not pitch_types:

    st.info(
        "この条件では球種データがありません。"
    )

else:

    selected_analysis_type = (
        st.selectbox(
            "球種",
            pitch_types,
            key="analysis_pitch_type",
        )
    )


    location_df = base[
        base["PitchType"]
        == selected_analysis_type
    ].copy()


    zone_counts = (
        location_df[
            "ActualZone"
        ]
        .value_counts()
        .reindex(
            ZONE_ORDER,
            fill_value=0,
        )
    )


    zone_total = int(
        zone_counts.sum()
    )


    for row in ZONE_GRID:

        cols = st.columns(
            3,
            gap="small",
            vertical_alignment="center",
            wrap=False,
        )


        for col, (
            zone,
            label,
        ) in zip(
            cols,
            row,
        ):

            count = int(
                zone_counts.get(
                    zone,
                    0,
                )
            )


            percentage = (
                count
                /
                zone_total
                *
                100
                if zone_total
                else 0
            )


            with col:

                st.metric(
                    label,
                    f"{percentage:.1f}%",
                    f"{count}球",
                )


# =========================================================
# iPitch input
# =========================================================

st.divider()


st.subheader(
    "iPitchで指定した球を記録"
)


st.markdown(
    """
    **① 球種を選択 → ② コースを3×3からタップ
    → ③ iPitchに入力 → ④ 実際の結果をタップ**
    """
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
# 3x3 pitch zone
# =========================================================

st.markdown(
    "**コース**"
)


with st.container(
    border=True,
    key="pitch-zone",
):

    for row in ZONE_GRID:

        cols = st.columns(
            3,
            gap="small",
            vertical_alignment="center",
            wrap=False,
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
                    ==
                    zone_value
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
                    key=(
                        f"zone_"
                        f"{zone_value}"
                    ),
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
            {
                ZONE_NAME_MAP[
                    st.session_state.selected_zone
                ]
            }
        </div>
        """,
        unsafe_allow_html=True,
    )


st.caption(
    "選択した球種・コースをiPitchへ入力してから、"
    "実際の投球結果をタップしてください。"
)


# =========================================================
# Pitch summary
# =========================================================

st.markdown(
    "### 今回の指定"
)


summary_col1, summary_col2 = (
    st.columns(2)
)


with summary_col1:

    st.info(
        f"""
        **球種**

        {selected_type}
        """
    )


with summary_col2:

    st.info(
        f"""
        **コース**

        {
            ZONE_NAME_MAP[
                st.session_state.selected_zone
            ]
        }
        """
    )


# =========================================================
# Result input
# =========================================================

st.markdown(
    "### 実際の投球結果"
)


result_options = [
    (
        "BallCalled",
        "ボール",
    ),
    (
        "StrikeCalled",
        "見逃しストライク",
    ),
    (
        "StrikeSwinging",
        "空振り",
    ),
    (
        "FoulBall",
        "ファウル",
    ),
    (
        "InPlay",
        "インプレー",
    ),
    (
        "HitByPitch",
        "死球",
    ),
]


result_cols = st.columns(3)


for i, (
    result_value,
    result_text,
) in enumerate(
    result_options
):

    with result_cols[
        i % 3
    ]:

        if st.button(
            result_text,
            key=(
                f"result_"
                f"{result_value}"
            ),
            use_container_width=True,
        ):

            # =============================================
            # 履歴保存
            # =============================================

            st.session_state.pitch_history.append(
                {
                    "球種": selected_type,

                    "コース": (
                        st.session_state.selected_zone
                    ),

                    "結果": result_value,

                    "Balls": (
                        st.session_state.balls
                    ),

                    "Strikes": (
                        st.session_state.strikes
                    ),
                }
            )


            # =============================================
            # カウント更新
            # =============================================

            (
                next_balls,
                next_strikes,
                pa_finished,
            ) = finish_count(
                st.session_state.balls,
                st.session_state.strikes,
                result_value,
            )


            # =============================================
            # 打席終了
            # =============================================

            if pa_finished:

                st.session_state.balls = 0

                st.session_state.strikes = 0

                st.session_state.pitch_history = []

                st.session_state.selected_zone = (
                    "Middle-Middle"
                )


            # =============================================
            # 打席継続
            # =============================================

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
                    f"{item['Balls']}"
                    "-"
                    f"{item['Strikes']}"
                ),

                "球種": item[
                    "球種"
                ],

                "コース": (
                    ZONE_NAME_MAP.get(
                        item["コース"],
                        item["コース"],
                    )
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

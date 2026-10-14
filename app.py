import numbers
from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components


# ==================================================
# 설정
# ==================================================
PURPLE = "#5E2BB8"
MINT = "#08D1D9"
DARK = "#252238"
MUTED = "#777487"
BORDER = "#E9E6F1"
LIGHT_PURPLE = "#F5F1FC"

BASE_DIR = Path(__file__).resolve().parent
FILE_PATH = BASE_DIR / "weekly_data.xlsx"

REAL_SALES_SHEET = "실매출"
ADJUSTED_SALES_SHEET = "보정매출"
AGENCY_SHEET = "대리점 리스트"
MULTIPLIER_SHEET = "매출 배수적용"

st.set_page_config(
    page_title="HNT Promotion Dashboard",
    layout="wide",
)

st.markdown(
    f"""
    <style>
    .stApp {{
        background: #FFFFFF;
        color: {DARK};
    }}

    .block-container {{
        max-width: 1440px;
        padding-top: 2.2rem;
        padding-bottom: 3rem;
    }}

    .dashboard-header {{
        border-bottom: 1px solid {BORDER};
        padding-bottom: 18px;
        margin-bottom: 22px;
    }}

    .dashboard-title {{
        color: {DARK};
        font-size: 2.15rem;
        font-weight: 750;
        letter-spacing: -0.04em;
        line-height: 1.2;
        margin: 0;
    }}

    .dashboard-title .purple {{
        color: {PURPLE};
    }}

    .brand-line {{
        width: 66px;
        height: 4px;
        border-radius: 4px;
        background: linear-gradient(90deg, {PURPLE}, {MINT});
        margin-top: 14px;
    }}

    .section-heading {{
        color: {DARK};
        font-size: 1.15rem;
        font-weight: 700;
        border-left: 4px solid {PURPLE};
        padding-left: 10px;
        margin: 18px 0 14px 0;
    }}

    .subsection-heading {{
        color: {DARK};
        font-size: 1rem;
        font-weight: 700;
        margin: 18px 0 10px 0;
    }}

    .kpi-card {{
        background: #FFFFFF;
        border: 1px solid {BORDER};
        border-top: 3px solid {PURPLE};
        border-radius: 12px;
        padding: 18px 20px;
        min-height: 112px;
        text-align: center;
    }}

    .kpi-card.mint {{
        border-top-color: {MINT};
    }}

    .kpi-label {{
        color: {MUTED};
        font-size: 0.88rem;
        font-weight: 600;
        margin-bottom: 12px;
    }}

    .kpi-value {{
        color: {DARK};
        font-size: 1.65rem;
        font-weight: 750;
        letter-spacing: -0.035em;
        line-height: 1.15;
    }}

    .kpi-note {{
        color: {MUTED};
        font-size: 0.78rem;
        margin-top: 7px;
        line-height: 1.5;
    }}

    .info-note {{
        color: {MUTED};
        font-size: 0.88rem;
        background: {LIGHT_PURPLE};
        border-left: 3px solid {MINT};
        border-radius: 6px;
        padding: 10px 13px;
        margin: 10px 0 14px 0;
        line-height: 1.65;
    }}

    .stTabs [data-baseweb="tab-list"] {{
        gap: 8px;
        border-bottom: 1px solid {BORDER};
    }}

    .stTabs [data-baseweb="tab"] {{
        height: 46px;
        color: {MUTED};
        font-weight: 650;
        padding: 0 18px;
        border-radius: 8px 8px 0 0;
    }}

    .stTabs [aria-selected="true"] {{
        color: {PURPLE} !important;
        background-color: {LIGHT_PURPLE};
    }}

    .stTabs [data-baseweb="tab-highlight"] {{
        background-color: {PURPLE};
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


# ==================================================
# 공통 함수
# ==================================================
def clean_code(series):
    return series.astype("string").str.strip()


def to_numeric(series):
    cleaned = (
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    return pd.to_numeric(cleaned, errors="coerce").fillna(0)


def format_won(value):
    return f"{value:,.0f}원"


def find_column(dataframe, candidates, sheet_name, purpose):
    columns = {
        str(column).strip(): column
        for column in dataframe.columns
    }

    for candidate in candidates:
        if candidate in columns:
            return columns[candidate]

    st.error(
        f"'{sheet_name}' 시트에서 {purpose} 컬럼을 찾지 못했습니다."
    )
    st.caption(
        "확인된 컬럼명: "
        + ", ".join(map(str, dataframe.columns))
    )
    st.stop()


def load_sales_sheet(
    file_path,
    sheet_name,
    revenue_candidates,
    revenue_name,
):
    try:
        dataframe = pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            header=1,
        )
    except Exception as error:
        st.error(f"'{sheet_name}' 시트를 읽지 못했습니다: {error}")
        st.stop()

    dataframe.columns = [
        str(column).strip()
        for column in dataframe.columns
    ]

    code_column = find_column(
        dataframe,
        ["대리점코드", "대리점 코드", "채널코드"],
        sheet_name,
        "대리점 코드",
    )
    revenue_column = find_column(
        dataframe,
        revenue_candidates,
        sheet_name,
        revenue_name,
    )

    rename_map = {
        code_column: "_대리점코드",
        revenue_column: "_원본매출",
    }

    aliases = {
        "담당실": ["담당실", "영업실"],
        "담당팀": ["담당팀", "영업팀"],
        "담당세일즈": ["담당세일즈", "담당 세일즈", "세일즈"],
        "대리점명": ["대리점한글명", "대리점명", "대리점 이름"],
    }

    for standard_name, candidates in aliases.items():
        for candidate in candidates:
            if candidate in dataframe.columns:
                rename_map[candidate] = standard_name
                break

    dataframe = dataframe.rename(columns=rename_map).copy()
    dataframe = dataframe.dropna(how="all")

    fill_columns = [
        column
        for column in ["담당실", "담당팀", "담당세일즈"]
        if column in dataframe.columns
    ]
    if fill_columns:
        dataframe[fill_columns] = dataframe[fill_columns].ffill()

    dataframe["_대리점코드"] = clean_code(
        dataframe["_대리점코드"]
    )
    dataframe["매출"] = to_numeric(dataframe["_원본매출"])

    dataframe = dataframe[
        dataframe["_대리점코드"].notna()
        & dataframe["_대리점코드"].ne("")
    ].copy()

    if dataframe.empty:
        st.error(f"'{sheet_name}' 시트에 집계할 데이터가 없습니다.")
        st.stop()

    return dataframe


def render_kpi(label, value, note="", accent="purple"):
    accent_class = "mint" if accent == "mint" else ""
    note_html = (
        f'<div class="kpi-note">{escape(str(note))}</div>'
        if note
        else ""
    )

    st.markdown(
        f"""
        <div class="kpi-card {accent_class}">
            <div class="kpi-label">{escape(str(label))}</div>
            <div class="kpi-value">{escape(str(value))}</div>
            {note_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def safe_display_value(value):
    """숫자는 천 단위 쉼표로 표시하고 문자열은 그대로 출력합니다."""
    if pd.isna(value):
        return ""

    if isinstance(value, numbers.Integral) and not isinstance(value, bool):
        display = f"{value:,}"
    elif isinstance(value, numbers.Real) and not isinstance(value, bool):
        numeric_value = float(value)
        if numeric_value.is_integer():
            display = f"{int(numeric_value):,}"
        else:
            display = f"{numeric_value:,.10f}".rstrip("0").rstrip(".")
    else:
        display = str(value)

    return escape(display)


def render_html_table(
    dataframe,
    right_align_columns=None,
    left_align_columns=None,
    column_widths=None,
    min_table_width=None,
    max_height=560,
):
    """독립 스크롤을 가진 HTML 표를 렌더링합니다."""
    right_align_columns = set(right_align_columns or [])
    left_align_columns = set(left_align_columns or [])
    column_widths = column_widths or {}

    colgroup = []
    for column in dataframe.columns:
        width = column_widths.get(column)
        if width:
            colgroup.append(
                f'<col style="width:{escape(str(width))};">'
            )
        else:
            colgroup.append("<col>")

    headers = [
        f"<th>{escape(str(column))}</th>"
        for column in dataframe.columns
    ]

    rows = []
    for _, row in dataframe.iterrows():
        cells = []

        for column in dataframe.columns:
            if column in right_align_columns:
                alignment = "right"
            elif column in left_align_columns:
                alignment = "left"
            else:
                alignment = "center"

            cells.append(
                f'<td class="align-{alignment}">'
                f"{safe_display_value(row[column])}"
                "</td>"
            )

        rows.append(f"<tr>{''.join(cells)}</tr>")

    table_min_width = min_table_width or "100%"

    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <style>
            * {{ box-sizing: border-box; }}

            html, body {{
                margin: 0;
                padding: 0;
                width: 100%;
                height: 100%;
                font-family: Pretendard, -apple-system, BlinkMacSystemFont,
                             "Segoe UI", sans-serif;
                color: {DARK};
                background: transparent;
            }}

            .table-scroll {{
                width: 100%;
                height: 100%;
                overflow-x: auto;
                overflow-y: auto;
                scrollbar-gutter: stable;
                border: 1px solid {BORDER};
                border-radius: 11px;
                background: #FFFFFF;
            }}

            .table-scroll::-webkit-scrollbar {{
                width: 12px;
                height: 12px;
            }}

            .table-scroll::-webkit-scrollbar-track {{
                background: #F1EFF5;
                border-radius: 6px;
            }}

            .table-scroll::-webkit-scrollbar-thumb {{
                background: #AAA5B3;
                border: 2px solid #F1EFF5;
                border-radius: 6px;
            }}

            table {{
                width: max-content;
                min-width: {table_min_width};
                border-collapse: collapse;
                table-layout: auto;
                font-size: 14px;
            }}

            thead th {{
                position: sticky;
                top: 0;
                z-index: 1;
                padding: 11px 12px;
                background: #F8F7FA;
                color: {MUTED};
                font-weight: 600;
                text-align: center;
                border-bottom: 1px solid {BORDER};
                border-right: 1px solid {BORDER};
                white-space: nowrap;
            }}

            tbody td {{
                padding: 10px 12px;
                background: #FFFFFF;
                border-bottom: 1px solid #E8E8EB;
                border-right: 1px solid #E8E8EB;
                vertical-align: middle;
                line-height: 1.5;
                overflow: hidden;
                text-overflow: ellipsis;
                white-space: nowrap;
            }}

            tbody tr:last-child td {{ border-bottom: 0; }}
            thead th:last-child, tbody td:last-child {{ border-right: 0; }}
            tbody td.align-center {{ text-align: center; }}
            tbody td.align-left {{ text-align: left; }}
            tbody td.align-right {{ text-align: right; }}
        </style>
    </head>
    <body>
        <div class="table-scroll">
            <table>
                <colgroup>{''.join(colgroup)}</colgroup>
                <thead><tr>{''.join(headers)}</tr></thead>
                <tbody>{''.join(rows)}</tbody>
            </table>
        </div>
    </body>
    </html>
    """

    content_height = min(
        max_height,
        44 + len(dataframe) * 43 + 4,
    )

    components.html(
        html,
        height=content_height,
        scrolling=False,
    )


def render_promotion_overview():
    rows = [
        ("최초예약일", "2026년 10월 1일 ~ 2026년 11월 30일"),
        ("출발일", "2026년 10월 1일 ~ 2027년 2월 28일"),
        ("속성", "ALL 속성 (Q 제외)"),
        (
            "제외",
            "인센티브, 코브랜드, 닷컴, 랜드온리, 업무포탈, 항공, "
            "FND, 호텔, 유아, 취소예약 제외",
        ),
        (
            "선정",
            "상품속성 및 휴일 판매 여부에 따라 인정배수를 차등 적용",
        ),
    ]

    payout_lines = [
        "1) 프로모션 종료 후 인정배수 적용 매출을 기준으로 매출 및 성장률 순위 산정",
        "2) 매출순위 및 매출성장률 기준 1위 ~ 10위 대리점 대상 VI 지급",
        "3) 최소 매출 허들 131,612,240원 미달 대리점은 지급 대상에서 제외",
        "4) 매출순위와 매출성장률 기준을 동시에 충족하면 지급액이 높은 기준 1개 적용",
    ]

    body = "".join(
        f"<tr><td>{escape(label)}</td><td>{escape(content)}</td></tr>"
        for label, content in rows
    )
    payout = "<br>".join(escape(line) for line in payout_lines)
    body += f"<tr><td>지급</td><td>{payout}</td></tr>"

    st.markdown(
        f"""
        <div style="width:100%;overflow-x:auto;border:1px solid {BORDER};
                    border-radius:10px;margin-bottom:12px;">
            <table style="width:100%;border-collapse:collapse;color:{DARK};
                          background:#FFFFFF;font-size:14px;">
                <thead>
                    <tr style="background:{LIGHT_PURPLE};">
                        <th style="padding:10px 14px;">구분</th>
                        <th style="padding:10px 14px;">내용</th>
                    </tr>
                </thead>
                <tbody>{body}</tbody>
            </table>
        </div>
        <style>
            table td {{
                padding:10px 14px;
                border-top:1px solid {BORDER};
                vertical-align:middle;
                line-height:1.6;
            }}
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_multiplier_rules():
    st.markdown(
        f"""
        <div style="width:100%;overflow-x:auto;border:1px solid {BORDER};
                    border-radius:10px;margin-bottom:12px;">
            <table style="width:100%;border-collapse:collapse;color:{DARK};
                          background:#FFFFFF;font-size:14px;text-align:center;">
                <thead>
                    <tr style="background:{LIGHT_PURPLE};">
                        <th style="padding:10px;">상품 속성</th>
                        <th style="padding:10px;">브랜드</th>
                        <th style="padding:10px;">인정 배수</th>
                        <th style="padding:10px;">휴일 예약</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="padding:10px;">Z</td>
                        <td style="padding:10px;">제우스</td>
                        <td rowspan="2" style="padding:10px;">2배</td>
                        <td rowspan="4" style="padding:10px;">1.5배</td>
                    </tr>
                    <tr>
                        <td style="padding:10px;">U</td>
                        <td style="padding:10px;">우리끼리</td>
                    </tr>
                    <tr>
                        <td style="padding:10px;">P</td>
                        <td style="padding:10px;">스탠다드, 프리미엄</td>
                        <td style="padding:10px;">1.5배</td>
                    </tr>
                    <tr>
                        <td style="padding:10px;">ALL (Q 제외)</td>
                        <td style="padding:10px;">-</td>
                        <td style="padding:10px;">1배</td>
                    </tr>
                </tbody>
            </table>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ==================================================
# 페이지 제목 및 프로모션 정보
# ==================================================
st.markdown(
    f"""
    <div class="dashboard-header">
        <h1 class="dashboard-title">
            HNT <span class="purple">Promotion Dashboard</span>
        </h1>
        <div class="brand-line"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-heading">2026 유통점 영업 활성화 프로모션</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="subsection-heading">프로모션 개요</div>',
    unsafe_allow_html=True,
)
render_promotion_overview()

st.markdown(
    '<div class="subsection-heading">상품 속성별 인정 배수</div>',
    unsafe_allow_html=True,
)
render_multiplier_rules()

st.markdown(
    '<div class="subsection-heading">VI 지급 기준</div>',
    unsafe_allow_html=True,
)

vi_df = pd.DataFrame(
    [
        ("매출", "1위", "3,500,000원", "1개", "3,500,000원"),
        ("매출", "2위", "3,000,000원", "1개", "3,000,000원"),
        ("매출", "3위", "2,000,000원", "1개", "2,000,000원"),
        ("매출", "4위", "1,500,000원", "1개", "1,500,000원"),
        ("매출", "5위", "1,000,000원", "1개", "1,000,000원"),
        ("매출", "6~10위", "600,000원", "5개", "3,000,000원"),
        ("성장률", "1~5위", "700,000원", "5개", "3,500,000원"),
        ("성장률", "6~10위", "500,000원", "5개", "2,500,000원"),
    ],
    columns=["순위 기준", "순위", "지급 금액", "개수", "지급액"],
)
render_html_table(vi_df, max_height=420)

st.divider()


# ==================================================
# 엑셀 파일 읽기
# ==================================================
if not FILE_PATH.exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {FILE_PATH}")
    st.stop()

real_sales_df = load_sales_sheet(
    FILE_PATH,
    REAL_SALES_SHEET,
    ["실매출", "매출", "실매출액"],
    "실매출",
)

adjusted_sales_df = load_sales_sheet(
    FILE_PATH,
    ADJUSTED_SALES_SHEET,
    ["보정매출", "총보정매출", "보정 매출"],
    "보정매출",
)

try:
    multiplier_source_df = pd.read_excel(
        FILE_PATH,
        sheet_name=MULTIPLIER_SHEET,
        header=1,
    ).dropna(how="all")
    multiplier_source_error = None
except Exception as error:
    multiplier_source_df = None
    multiplier_source_error = str(error)


# ==================================================
# 대리점 채널 정보
# ==================================================
try:
    agency_df = pd.read_excel(
        FILE_PATH,
        sheet_name=AGENCY_SHEET,
        header=1,
    )
except Exception as error:
    st.error(f"'{AGENCY_SHEET}' 시트를 읽지 못했습니다: {error}")
    st.stop()

agency_df.columns = [
    str(column).strip()
    for column in agency_df.columns
]

required_columns = ["대리점코드", "채널구분"]
missing_columns = [
    column
    for column in required_columns
    if column not in agency_df.columns
]

if missing_columns:
    st.error(
        f"'{AGENCY_SHEET}' 시트에서 필요한 컬럼을 찾지 못했습니다: "
        + ", ".join(missing_columns)
    )
    st.caption(
        "확인된 컬럼명: "
        + ", ".join(map(str, agency_df.columns))
    )
    st.stop()

agency_df["_대리점코드"] = clean_code(
    agency_df["대리점코드"]
)

channel_map = agency_df[
    ["_대리점코드", "채널구분"]
].copy()

channel_map["채널구분"] = (
    channel_map["채널구분"]
    .fillna("미분류")
    .astype(str)
    .str.strip()
)
channel_map.loc[
    channel_map["채널구분"].eq(""),
    "채널구분",
] = "미분류"

channel_map = channel_map.drop_duplicates(
    subset="_대리점코드",
    keep="first",
)

real_sales_df = real_sales_df.merge(
    channel_map,
    on="_대리점코드",
    how="left",
)
real_sales_df["채널구분"] = (
    real_sales_df["채널구분"]
    .fillna("미분류")
    .astype(str)
    .str.strip()
)


# ==================================================
# 매출 순위 데이터 구성
# 보정매출 시트 기준으로 집계합니다.
# 담당팀을 담당실 바로 다음 컬럼으로 표시합니다.
# ==================================================
aggregations = {
    "보정매출": ("매출", "sum"),
}

for column in ["담당실", "담당팀", "담당세일즈", "대리점명"]:
    if column in adjusted_sales_df.columns:
        aggregations[column] = (column, "first")

sales_ranking = (
    adjusted_sales_df
    .groupby("_대리점코드", as_index=False)
    .agg(**aggregations)
    .sort_values("보정매출", ascending=False)
    .reset_index(drop=True)
)

sales_ranking.insert(
    0,
    "순위",
    range(1, len(sales_ranking) + 1),
)

for column in ["대리점명", "담당실", "담당팀", "담당세일즈"]:
    if column not in sales_ranking.columns:
        if column == "대리점명":
            sales_ranking[column] = sales_ranking["_대리점코드"]
        else:
            sales_ranking[column] = "미분류"

sales_ranking["대리점코드"] = (
    sales_ranking["_대리점코드"]
    .fillna("")
    .astype(str)
    .str.replace(r"\.0$", "", regex=True)
)
sales_ranking["대리점명"] = (
    sales_ranking["대리점명"]
    .fillna("")
    .astype(str)
    .str.strip()
)


# ==================================================
# KPI — 실매출 기준
# ==================================================
total_agencies = real_sales_df["_대리점코드"].nunique()

if "담당세일즈" in real_sales_df.columns:
    total_staff = (
        real_sales_df["담당세일즈"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .nunique()
    )
else:
    total_staff = 0

total_real_sales = real_sales_df["매출"].sum()
average_real_sales = (
    total_real_sales / total_agencies
    if total_agencies
    else 0
)

st.markdown(
    '<div class="section-heading">주요 실적</div>',
    unsafe_allow_html=True,
)

kpi1, kpi2, kpi3, kpi4 = st.columns(4, gap="medium")

with kpi1:
    render_kpi(
        "총 대리점 수",
        f"{total_agencies:,}개",
        "유통점 (현대/홈플 제외) + 직계약 유통점 대상",
    )

with kpi2:
    render_kpi(
        "총 담당자 수",
        f"{total_staff:,}명",
        "담당 세일즈 수",
        "mint",
    )

with kpi3:
    render_kpi(
        "매출",
        format_won(total_real_sales),
        "실매출 탭 기준",
    )

with kpi4:
    render_kpi(
        "대리점당 평균 매출",
        format_won(average_real_sales),
        "",
        "mint",
    )


# ==================================================
# 탭 순서: 매출 순위를 첫 번째 탭으로 배치
# ==================================================
st.markdown("<br>", unsafe_allow_html=True)

tab_ranking, tab_office, tab_channel = st.tabs(
    [
        "매출 순위",
        "영업실별 매출현황",
        "유통채널별 매출현황",
    ]
)


# ==================================================
# 매출 순위 — 첫 번째 탭
# ==================================================
with tab_ranking:
    st.markdown(
        '<div class="section-heading">매출 순위</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-note">
            매출 순위는 보정매출 탭의 보정매출을 기준으로 산출합니다.
            그 외 주요 실적 및 매출 현황은 실매출 탭 기준입니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    ranking_search = st.text_input(
        "매출 순위 검색",
        placeholder="대리점코드, 대리점명, 담당실, 담당팀 또는 담당세일즈 검색",
        key="ranking_search",
    )

    ranking_columns = [
        "순위",
        "대리점코드",
        "대리점명",
        "담당실",
        "담당팀",
        "담당세일즈",
        "보정매출",
    ]

    ranking_table = sales_ranking[ranking_columns].copy()

    if ranking_search.strip():
        search_text = ranking_search.strip().casefold()
        searchable_columns = [
            "대리점코드",
            "대리점명",
            "담당실",
            "담당팀",
            "담당세일즈",
        ]

        search_mask = ranking_table[searchable_columns].apply(
            lambda column: column.astype(str)
            .str.casefold()
            .str.contains(search_text, na=False, regex=False)
        ).any(axis=1)

        ranking_table = ranking_table[search_mask].copy()

    ranking_table["보정매출"] = ranking_table["보정매출"].map(
        format_won
    )

    if ranking_search.strip():
        st.caption(f"검색 결과 {len(ranking_table):,}건")

    if ranking_table.empty:
        st.info("검색 조건에 맞는 매출 순위가 없습니다.")
    else:
        render_html_table(
            ranking_table,
            left_align_columns=["대리점명"],
            right_align_columns=["보정매출"],
            column_widths={
                "순위": "80px",
                "대리점코드": "150px",
                "대리점명": "330px",
                "담당실": "150px",
                "담당팀": "150px",
                "담당세일즈": "160px",
                "보정매출": "190px",
            },
            min_table_width="1210px",
            max_height=560,
        )

    st.markdown(
        '<div class="subsection-heading">원본 데이터 (매출 배수적용)</div>',
        unsafe_allow_html=True,
    )

    if multiplier_source_df is None:
        st.error(
            f"'{MULTIPLIER_SHEET}' 시트를 읽지 못했습니다: "
            f"{multiplier_source_error}"
        )
    elif multiplier_source_df.empty:
        st.info(
            f"'{MULTIPLIER_SHEET}' 시트에 표시할 데이터가 없습니다."
        )
    else:
        source_display_df = multiplier_source_df.copy()
        source_display_df.columns = [
            str(column).strip()
            for column in source_display_df.columns
        ]

        source_column_width = 180
        source_min_width = max(
            1200,
            source_column_width * len(source_display_df.columns),
        )
        source_widths = {
            column: f"{source_column_width}px"
            for column in source_display_df.columns
        }

        render_html_table(
            source_display_df,
            column_widths=source_widths,
            min_table_width=f"{source_min_width}px",
            max_height=560,
        )


# ==================================================
# 영업실별 매출현황
# ==================================================
with tab_office:
    st.markdown(
        '<div class="section-heading">영업실별 매출현황</div>',
        unsafe_allow_html=True,
    )

    if "담당실" not in real_sales_df.columns:
        st.warning(
            "'실매출' 시트에 담당실 컬럼이 없어 "
            "영업실별 매출현황을 표시할 수 없습니다."
        )
    else:
        office_summary = (
            real_sales_df
            .groupby("담당실", dropna=False, as_index=False)["매출"]
            .sum()
            .sort_values("매출", ascending=False)
        )

        office_summary["담당실"] = (
            office_summary["담당실"]
            .fillna("미분류")
            .astype(str)
            .str.strip()
        )

        office_chart = px.bar(
            office_summary,
            x="담당실",
            y="매출",
            text="매출",
            color="담당실",
            color_discrete_sequence=[
                PURPLE,
                MINT,
                "#8B6BD1",
                "#66DEE2",
                "#B8A4E6",
                "#9AECEF",
            ],
        )

        office_chart.update_traces(
            texttemplate="%{text:,.0f}원",
            textposition="outside",
            marker_line_width=0,
            hovertemplate=(
                "<b>%{x}</b><br>"
                "매출: %{y:,.0f}원"
                "<extra></extra>"
            ),
        )

        office_chart.update_layout(
            showlegend=False,
            xaxis_title="영업실",
            yaxis_title="매출",
            yaxis_tickformat=",.0f",
            plot_bgcolor="#FFFFFF",
            paper_bgcolor="#FFFFFF",
            font=dict(color=DARK),
            margin=dict(l=20, r=20, t=25, b=20),
            xaxis=dict(showgrid=False, linecolor=BORDER),
            yaxis=dict(gridcolor="#F0EDF5", zeroline=False),
        )

        st.plotly_chart(
            office_chart,
            use_container_width=True,
        )

        office_table = office_summary.rename(
            columns={"담당실": "영업실"}
        ).copy()
        office_table["매출"] = office_table["매출"].map(format_won)

        render_html_table(
            office_table,
            right_align_columns=["매출"],
            max_height=440,
        )


# ==================================================
# 유통채널별 매출현황
# ==================================================
def normalize_channel_name(value):
    if pd.isna(value):
        return ""

    return (
        str(value)
        .strip()
        .replace(" ", "")
        .replace("\u3000", "")
    )


def classify_channel(value):
    channel = normalize_channel_name(value)

    if not channel:
        return "기타"
    if "롯데백화점" in channel:
        return "롯데백화점"
    if "롯데마트" in channel:
        return "롯데마트"
    if "이마트" in channel:
        return "이마트"
    if "하나로마트" in channel:
        return "하나로마트"
    if (
        channel == "NC"
        or "NC백화점" in channel
        or channel.startswith("NC")
    ):
        return "NC"

    return "기타"


with tab_channel:
    st.markdown(
        '<div class="section-heading">유통채널별 매출현황</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-note">
            유통채널별 매출은 실매출 탭 기준으로 집계합니다.
            롯데백화점과 롯데백화점(몰)은 합산하며,
            지정되지 않은 채널은 기타로 합산합니다.
            그래프의 기타 영역에 마우스를 올리면 세부 채널별 구성비를 확인할 수 있습니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    channel_source = real_sales_df.copy()

    # 원래 채널명을 보존하여 기타 항목의 세부 내역으로 활용합니다.
    channel_source["세부채널"] = (
        channel_source["채널구분"]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    channel_source.loc[
        channel_source["세부채널"].eq(""),
        "세부채널",
    ] = "미분류"

    channel_source["유통채널"] = channel_source["채널구분"].apply(
        classify_channel
    )

    channel_summary = (
        channel_source
        .groupby("유통채널", as_index=False)["매출"]
        .sum()
    )

    channel_order = [
        "이마트",
        "롯데마트",
        "롯데백화점",
        "하나로마트",
        "NC",
        "기타",
    ]

    channel_summary["_정렬순서"] = channel_summary["유통채널"].map(
        {
            channel: index
            for index, channel in enumerate(channel_order)
        }
    )

    channel_summary = (
        channel_summary
        .sort_values("_정렬순서")
        .drop(columns="_정렬순서")
        .reset_index(drop=True)
    )

    channel_colors = {
        "이마트": "#FFD600",
        "롯데마트": "#DA291C",
        "롯데백화점": "#C9A227",
        "하나로마트": "#8BC34A",
        "NC": "#FFFFFF",
        "기타": "#A9A5B3",
    }

    channel_revenue_total = channel_summary["매출"].sum()

    if channel_summary.empty or channel_revenue_total == 0:
        st.info("표시할 유통채널별 매출 집계 결과가 없습니다.")
    else:
        # 기타에 해당하는 원래 채널명별 매출을 집계합니다.
        other_detail = (
            channel_source[
                channel_source["유통채널"].eq("기타")
            ]
            .groupby("세부채널")["매출"]
            .sum()
            .sort_values(ascending=False)
        )

        other_revenue_total = other_detail.sum()

        if other_revenue_total > 0:
            other_detail_lines = []

            for detail_name, detail_revenue in other_detail.items():
                share_of_total = (
                    detail_revenue / channel_revenue_total * 100
                )
                share_of_other = (
                    detail_revenue / other_revenue_total * 100
                )

                other_detail_lines.append(
                    f"{escape(str(detail_name))}: "
                    f"전체 매출 {share_of_total:.2f}% · "
                    f"기타 내 {share_of_other:.2f}%"
                )

            other_hover_detail = (
                "<b>기타 세부 구성</b><br>"
                + "<br>".join(other_detail_lines)
            )
        else:
            other_hover_detail = ""

        # 기타 조각의 툴팁에만 세부 구성 정보를 전달합니다.
        channel_summary["hover_detail"] = channel_summary[
            "유통채널"
        ].apply(
            lambda channel: (
                other_hover_detail
                if channel == "기타"
                else ""
            )
        )

        channel_chart = px.pie(
            channel_summary,
            names="유통채널",
            values="매출",
            hole=0.56,
            category_orders={
                "유통채널": channel_order,
            },
            color="유통채널",
            color_discrete_map=channel_colors,
            custom_data=["hover_detail"],
        )

        channel_chart.update_traces(
            textposition="inside",
            texttemplate="%{label}<br>%{percent}",
            marker=dict(
                colors=[
                    channel_colors[channel]
                    for channel in channel_summary["유통채널"]
                ],
                line=dict(
                    color="#A9A5B3",
                    width=1.5,
                ),
            ),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "매출: %{value:,.0f}원<br>"
                "전체 구성비: %{percent}"
                "<br>%{customdata[0]}"
                "<extra></extra>"
            ),
        )

        channel_chart.update_layout(
            plot_bgcolor="#FFFFFF",
            paper_bgcolor="#FFFFFF",
            font=dict(color=DARK),
            legend_title_text="유통채널",
            margin=dict(l=20, r=20, t=20, b=20),
        )

        st.plotly_chart(
            channel_chart,
            use_container_width=True,
        )

        # 툴팁용 세부 정보는 제외하고 기존 표 열만 표시합니다.
        channel_table = channel_summary[
            ["유통채널", "매출"]
        ].copy()

        channel_table["매출"] = channel_table["매출"].map(format_won)
        channel_table["구성비"] = (
            channel_summary["매출"]
            .div(channel_revenue_total)
            .mul(100)
            .map(lambda value: f"{value:.1f}%")
        )

        render_html_table(
            channel_table,
            right_align_columns=["매출", "구성비"],
            max_height=440,
        )

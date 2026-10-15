from pathlib import Path
import numbers
from html import escape
import hmac

import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as components


# 페이지 설정은 다른 Streamlit 명령보다 먼저, 한 번만 호출합니다.
st.set_page_config(
    page_title="HNT Promotion Dashboard",
    layout="wide",
)


# =========================================================
# 설정
# =========================================================
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
HOLIDAY_SHEET = "휴일예약"


# =========================================================
# 인증 설정
# =========================================================
# 기본 비밀번호 (필요시 원하는 비밀번호로 변경하세요. .streamlit/secrets.toml 설정 시 해당 비밀번호가 우선 적용됩니다)
DEFAULT_PASSWORD = "hana2026"
AUTH_KEY = "hnt_dashboard_authenticated"


# =========================================================
# 스타일 (CSS) - HanaTour Brand Design System
# =========================================================
st.markdown(
    f"""
    <link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
    <style>
        /* 폰트 및 기본 테마 설정 */
        html, body, [class*="css"], .stApp {{
            font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
            background: #FAFAFC;
            color: {DARK};
        }}

        .block-container {{
            max-width: 1440px;
            padding-top: 1.8rem;
            padding-bottom: 3.5rem;
        }}

        /* 대시보드 헤더 배너 */
        .dashboard-header {{
            background: linear-gradient(135deg, #FFFFFF 0%, {LIGHT_PURPLE} 100%);
            border: 1px solid {BORDER};
            border-radius: 16px;
            padding: 26px 32px 24px 32px;
            margin-bottom: 24px;
            box-shadow: 0 4px 20px rgba(94, 43, 184, 0.04);
            position: relative;
            overflow: hidden;
        }}

        .dashboard-header::after {{
            content: "";
            position: absolute;
            top: -20px;
            right: -20px;
            width: 240px;
            height: 180px;
            background: radial-gradient(circle, rgba(8, 209, 217, 0.15) 0%, rgba(94, 43, 184, 0.05) 50%, transparent 75%);
            pointer-events: none;
        }}

        .header-top-tag {{
            display: inline-flex;
            align-items: center;
            gap: 7px;
            background: #FFFFFF;
            border: 1px solid {BORDER};
            color: {PURPLE};
            font-size: 0.8rem;
            font-weight: 700;
            padding: 4px 12px;
            border-radius: 20px;
            margin-bottom: 10px;
            box-shadow: 0 2px 6px rgba(94, 43, 184, 0.04);
        }}

        .live-dot {{
            width: 8px;
            height: 8px;
            border-radius: 50%;
            background: {MINT};
            box-shadow: 0 0 8px {MINT};
        }}

        .dashboard-title {{
            color: {DARK};
            font-size: 2.15rem;
            font-weight: 800;
            letter-spacing: -0.04em;
            line-height: 1.2;
            margin: 0;
        }}

        .dashboard-title .brand-title {{
            background: linear-gradient(135deg, {PURPLE} 0%, #834BE8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .dashboard-subtitle {{
            color: {MUTED};
            font-size: 0.92rem;
            margin-top: 6px;
            font-weight: 500;
            line-height: 1.5;
        }}

        .brand-line {{
            width: 60px;
            height: 4px;
            border-radius: 4px;
            background: linear-gradient(90deg, {PURPLE}, {MINT});
            margin-top: 10px;
        }}

        /* 섹션 및 서브섹션 타이틀 */
        .section-heading {{
            display: flex;
            align-items: center;
            gap: 10px;
            color: {DARK};
            font-size: 1.25rem;
            font-weight: 800;
            letter-spacing: -0.03em;
            margin: 28px 0 16px 0;
        }}

        .section-heading::before {{
            content: "";
            display: inline-block;
            width: 5px;
            height: 20px;
            border-radius: 3px;
            background: linear-gradient(180deg, {PURPLE} 0%, {MINT} 100%);
        }}

        .subsection-heading {{
            color: {DARK};
            font-size: 1.05rem;
            font-weight: 750;
            letter-spacing: -0.02em;
            margin: 22px 0 12px 0;
            padding-left: 2px;
        }}

        /* KPI 카드 */
        .kpi-card {{
            background: #FFFFFF;
            border: 1px solid {BORDER};
            border-radius: 14px;
            padding: 22px 20px 20px 20px;
            min-height: 118px;
            text-align: center;
            position: relative;
            overflow: hidden;
            box-shadow: 0 2px 10px rgba(94, 43, 184, 0.03);
            transition: all 0.25s cubic-bezier(0.16, 1, 0.3, 1);
        }}

        .kpi-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 10px 24px rgba(94, 43, 184, 0.08);
            border-color: #D3C9EF;
        }}

        .kpi-card::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, {PURPLE}, #7F4BE8);
        }}

        .kpi-card.mint::before {{
            background: linear-gradient(90deg, {MINT}, #33E5EC);
        }}

        .kpi-label {{
            color: {MUTED};
            font-size: 0.86rem;
            font-weight: 600;
            letter-spacing: -0.01em;
            margin-bottom: 10px;
        }}

        .kpi-value {{
            color: {DARK};
            font-size: 1.72rem;
            font-weight: 800;
            letter-spacing: -0.04em;
            line-height: 1.15;
            font-variant-numeric: tabular-nums;
        }}

        .kpi-note {{
            display: inline-block;
            color: {MUTED};
            font-size: 0.76rem;
            font-weight: 500;
            margin-top: 10px;
            padding: 3px 9px;
            background: #F4F3F8;
            border-radius: 6px;
            line-height: 1.4;
        }}

        /* 안내 박스 (Info Note) */
        .info-note {{
            color: #49455B;
            font-size: 0.88rem;
            background: linear-gradient(135deg, #F8F5FF 0%, #F3FCFC 100%);
            border: 1px solid #E3DCF5;
            border-left: 4px solid {PURPLE};
            border-radius: 10px;
            padding: 12px 16px;
            margin: 12px 0 16px 0;
            line-height: 1.65;
            box-shadow: 0 1px 4px rgba(94, 43, 184, 0.02);
        }}

        /* 탭 네비게이션 고도화 (Modern Segmented Navigation) */
        .stTabs [data-baseweb="tab-list"] {{
            gap: 8px;
            background: #F2EFF8;
            padding: 6px;
            border-radius: 12px;
            border: 1px solid {BORDER};
            margin-bottom: 12px;
        }}

        .stTabs [data-baseweb="tab"] {{
            height: 42px;
            color: {MUTED};
            font-weight: 700;
            font-size: 0.92rem;
            padding: 0 18px;
            border: none !important;
            border-radius: 8px !important;
            background-color: transparent !important;
            transition: all 0.2s ease;
        }}

        .stTabs [data-baseweb="tab"]:hover {{
            color: {PURPLE} !important;
            background-color: rgba(94, 43, 184, 0.08) !important;
        }}

        .stTabs [data-baseweb="tab"][aria-selected="true"] {{
            color: #FFFFFF !important;
            background: linear-gradient(135deg, {PURPLE} 0%, #763FE6 100%) !important;
            box-shadow: 0 4px 12px rgba(94, 43, 184, 0.25) !important;
        }}

        .stTabs [data-baseweb="tab-highlight"],
        .stTabs [data-baseweb="tab-border"] {{
            display: none !important;
        }}

        /* 폼 위젯 스타일링 (Input, Select, Button) */
        div[data-testid="stTextInput"] input {{
            border: 1px solid {BORDER};
            border-radius: 10px;
            background: #FFFFFF;
            color: {DARK};
            font-size: 0.92rem;
            padding: 9px 14px;
            transition: all 0.2s ease;
        }}

        div[data-testid="stTextInput"] input:focus {{
            border-color: {PURPLE};
            box-shadow: 0 0 0 3px rgba(94, 43, 184, 0.12);
        }}

        div[data-testid="stSelectbox"] > div > div {{
            border: 1px solid {BORDER};
            border-radius: 10px;
            background: #FFFFFF;
            transition: all 0.2s ease;
        }}

        div[data-testid="stSelectbox"] > div > div:focus-within {{
            border-color: {PURPLE};
            box-shadow: 0 0 0 3px rgba(94, 43, 184, 0.12);
        }}

        button[kind="primary"], .stButton > button {{
            border-radius: 10px;
            font-weight: 700;
            transition: all 0.2s ease;
        }}

        button[kind="primary"] {{
            background: linear-gradient(135deg, {PURPLE} 0%, #7539DC 100%) !important;
            border: none !important;
            color: #FFFFFF !important;
            box-shadow: 0 4px 14px rgba(94, 43, 184, 0.22);
        }}

        button[kind="primary"]:hover {{
            transform: translateY(-1px);
            box-shadow: 0 6px 18px rgba(94, 43, 184, 0.32);
        }}

        hr {{
            border: none;
            height: 1px;
            background: linear-gradient(90deg, transparent, {BORDER}, transparent);
            margin: 28px 0;
        }}

        /* 로그인 카드 스타일 */
        .login-card {{
            background: #FFFFFF;
            border: 1px solid {BORDER};
            border-radius: 16px;
            padding: 34px 30px 26px 30px;
            box-shadow: 0 12px 36px rgba(94, 43, 184, 0.08);
            margin-top: 45px;
            margin-bottom: 20px;
            position: relative;
            overflow: hidden;
        }}

        .login-card::before {{
            content: "";
            position: absolute;
            top: 0;
            left: 0;
            right: 0;
            height: 4px;
            background: linear-gradient(90deg, {PURPLE}, {MINT});
        }}

        .login-title {{
            color: {DARK};
            font-size: 1.55rem;
            font-weight: 800;
            letter-spacing: -0.03em;
            margin: 0;
        }}

        .login-title .brand-title {{
            background: linear-gradient(135deg, {PURPLE} 0%, #834BE8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
        }}

        .login-subtitle {{
            color: {MUTED};
            font-size: 0.9rem;
            margin-top: 8px;
            margin-bottom: 4px;
            line-height: 1.5;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


def check_password(entered_password: str) -> bool:
    """비밀번호 검증 (secrets.toml 설정 시 우선 적용, 없을 시 DEFAULT_PASSWORD 사용)"""
    saved_password = None
    try:
        if "auth" in st.secrets and "password" in st.secrets["auth"]:
            saved_password = str(st.secrets["auth"]["password"])
        elif "password" in st.secrets:
            saved_password = str(st.secrets["password"])
    except Exception:
        saved_password = None

    if not saved_password:
        saved_password = DEFAULT_PASSWORD

    return hmac.compare_digest(entered_password.strip(), saved_password.strip())


def require_password():
    """비밀번호 인증 화면 처리"""
    if st.session_state.get(AUTH_KEY, False):
        with st.sidebar:
            st.markdown(
                f"""
                <div style="padding: 10px 14px; background: {LIGHT_PURPLE}; border-radius: 10px; border: 1px solid #E3DCF5; margin-bottom: 12px;">
                    <div style="color: {PURPLE}; font-weight: 750; font-size: 0.88rem;">🔒 로그인 인증 완료</div>
                    <div style="color: {MUTED}; font-size: 0.78rem; margin-top: 2px;">정상 접근 상태입니다.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button("로그아웃", use_container_width=True):
                st.session_state[AUTH_KEY] = False
                st.rerun()
        return

    # 미인증 시 중앙 로그인 폼 노출
    _, center_col, _ = st.columns([1, 1.4, 1])

    with center_col:
        st.markdown(
            f"""
            <div class="login-card">
                <div class="header-top-tag" style="margin-bottom: 14px;">
                    <span class="live-dot"></span> SECURE ACCESS
                </div>
                <div class="login-title"><span class="brand-title">HNT</span> Promotion Dashboard</div>
                <div class="brand-line" style="margin-top: 8px; margin-bottom: 14px;"></div>
                <div class="login-subtitle">대시보드를 확인하려면 비밀번호를 입력해 주세요.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

        with st.form("password_login_form"):
            entered_password = st.text_input(
                "비밀번호",
                type="password",
                placeholder="비밀번호를 입력하세요",
            )
            submitted = st.form_submit_button("대시보드 접속", use_container_width=True, type="primary")

        if submitted:
            if entered_password and check_password(entered_password):
                st.session_state[AUTH_KEY] = True
                st.success("인증되었습니다. 대시보드로 이동합니다...")
                st.rerun()
            else:
                st.error("비밀번호가 올바르지 않습니다.")

    st.stop()


require_password()


# =========================================================
# 공통 함수
# =========================================================
def find_column(dataframe, candidates, sheet_name, purpose):
    """후보 목록에서 실제 데이터프레임의 컬럼명을 찾습니다."""
    column_lookup = {
        str(column).strip(): column
        for column in dataframe.columns
    }

    for candidate in candidates:
        if candidate in column_lookup:
            return column_lookup[candidate]

    st.error(
        f"'{sheet_name}' 시트에서 {purpose} 컬럼을 찾지 못했습니다."
    )
    st.caption(
        "확인된 컬럼명: "
        + ", ".join(map(str, dataframe.columns))
    )
    st.stop()


def numeric_series(series):
    """쉼표, 원 기호 등이 포함된 값을 숫자로 변환합니다."""
    text = series.astype("string").str.strip()
    text = text.str.replace(",", "", regex=False)
    text = text.str.replace("원", "", regex=False)
    text = text.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(text, errors="coerce")


def clean_code(series):
    """대리점코드를 문자열로 정리합니다."""
    return (
        series.astype("string")
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def format_won(value):
    return f"{float(value):,.0f}원"


def format_count(value):
    if pd.isna(value) or str(value).strip() == "":
        return "-"
    try:
        return f"{int(round(float(value))):,}명"
    except (ValueError, TypeError):
        return str(value)


def format_number(value):
    if pd.isna(value):
        return ""

    if isinstance(value, numbers.Integral):
        return f"{value:,}"

    if isinstance(value, numbers.Real):
        value = float(value)
        if value.is_integer():
            return f"{int(value):,}"
        return f"{value:,.2f}".rstrip("0").rstrip(".")

    return str(value)


def safe_display_value(value):
    """HTML 표 셀에 표시할 값을 안전한 문자열로 반환합니다."""
    if pd.isna(value):
        return ""
    return escape(format_number(value))


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


def load_sales_sheet(
    file_path,
    sheet_name,
    revenue_candidates,
    revenue_name,
):
    """
    실매출·보정매출·휴일예약 시트를 읽어 공통 컬럼명으로 정리합니다.
    실매출 시트는 header=1 기준의 새 컬럼 구조를 지원합니다.
    """
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
    dataframe = dataframe.dropna(how="all").copy()

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

    # 새 실매출 시트와 기존 시트에서 쓰는 컬럼명을 함께 지원합니다.
    aliases = {
        "본부": ["본부", "담당본부"],
        "담당실": ["담당실", "영업실", "담당영업실"],
        "담당팀": ["담당팀", "영업팀", "담당영업팀"],
        "담당세일즈": ["담당세일즈", "담당 세일즈", "세일즈"],
        "대리점명": ["대리점한글명", "대리점명", "대리점 이름"],
        "예약인원": ["예약인원", "예약 인원", "예약인원수", "인원수", "예약자수", "인원"],
    }

    rename_map = {
        code_column: "_대리점코드",
        revenue_column: "_원본매출",
    }

    for standard_name, candidates in aliases.items():
        for candidate in candidates:
            if candidate in dataframe.columns:
                rename_map[candidate] = standard_name
                break

    dataframe = dataframe.rename(columns=rename_map).copy()

    if "_대리점코드" not in dataframe.columns:
        st.error(f"'{sheet_name}' 시트에 대리점코드가 없습니다.")
        st.stop()

    dataframe["_대리점코드"] = clean_code(
        dataframe["_대리점코드"]
    )
    dataframe["매출"] = numeric_series(
        dataframe["_원본매출"]
    ).fillna(0)

    if "예약인원" in dataframe.columns:
        dataframe["예약인원"] = numeric_series(
            dataframe["예약인원"]
        ).fillna(0)

    # 본부/실/팀 등 구분 컬럼에 병합된 셀이 있으면 위 값을 이어받습니다.
    hierarchy_columns = [
        column
        for column in [
            "본부",
            "담당실",
            "담당팀",
            "담당세일즈",
        ]
        if column in dataframe.columns
    ]

    if hierarchy_columns:
        dataframe[hierarchy_columns] = (
            dataframe[hierarchy_columns].ffill()
        )

    dataframe = dataframe[
        dataframe["_대리점코드"].notna()
        & dataframe["_대리점코드"].ne("")
    ].copy()

    for column in [
        "본부",
        "담당실",
        "담당팀",
        "담당세일즈",
        "대리점명",
    ]:
        if column in dataframe.columns:
            dataframe[column] = (
                dataframe[column]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    if dataframe.empty:
        st.error(f"'{sheet_name}' 시트에 표시할 데이터가 없습니다.")
        st.stop()

    return dataframe


def filter_agency_search(
    dataframe,
    search_text,
    searchable_columns,
):
    """
    검색어가 '영업1실'처럼 '실'로 끝나면 담당실을 정확히 비교합니다.
    그 외 검색어는 지정 컬럼에서 부분 일치로 검색합니다.
    """
    search_text = str(search_text).strip().casefold()

    if not search_text:
        return dataframe.copy()

    office_column = None
    for candidate in ["담당실", "영업실"]:
        if candidate in dataframe.columns:
            office_column = candidate
            break

    if search_text.endswith("실") and office_column:
        office_values = (
            dataframe[office_column]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.casefold()
        )
        return dataframe[office_values.eq(search_text)].copy()

    available_columns = [
        column
        for column in searchable_columns
        if column in dataframe.columns
    ]

    if not available_columns:
        return dataframe.iloc[0:0].copy()

    mask = dataframe[available_columns].apply(
        lambda column: (
            column.fillna("")
            .astype(str)
            .str.casefold()
            .str.contains(
                search_text,
                na=False,
                regex=False,
            )
        )
    ).any(axis=1)

    return dataframe[mask].copy()


def render_html_table(
    dataframe,
    right_align_columns=None,
    left_align_columns=None,
    column_widths=None,
    min_table_width=None,
    max_height=560,
    fit_width=False,
    compact=False,
):
    """
    표를 HTML로 표시합니다.
    fit_width=True인 경우 표를 화면 너비에 맞추고 가로 스크롤을 숨깁니다.
    """
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

    if fit_width:
        table_width_css = """
            width: 100%;
            min-width: 100%;
            table-layout: fixed;
        """
        overflow_x = "hidden"
        cell_wrap_css = """
            white-space: normal;
            overflow-wrap: anywhere;
            word-break: keep-all;
        """
    else:
        table_min_width = min_table_width or "100%"
        table_width_css = f"""
            width: max-content;
            min-width: {table_min_width};
            table-layout: auto;
        """
        overflow_x = "auto"
        cell_wrap_css = "white-space: nowrap;"

    if compact:
        font_size = "12px"
        header_padding = "8px 5px"
        cell_padding = "7px 5px"
    else:
        font_size = "14px"
        header_padding = "11px 12px"
        cell_padding = "10px 12px"

    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
        <style>
            * {{
                box-sizing: border-box;
            }}

            html, body {{
                margin: 0;
                padding: 0;
                width: 100%;
                height: 100%;
                font-family: Pretendard, -apple-system,
                             BlinkMacSystemFont, "Segoe UI", sans-serif;
                color: {DARK};
                background: transparent;
            }}

            .table-scroll {{
                width: 100%;
                height: 100%;
                overflow-x: {overflow_x};
                overflow-y: auto;
                scrollbar-gutter: stable;
                border: 1px solid {BORDER};
                border-radius: 12px;
                background: #FFFFFF;
                box-shadow: 0 2px 10px rgba(94, 43, 184, 0.03);
            }}

            table {{
                {table_width_css}
                border-collapse: collapse;
                font-size: {font_size};
            }}

            thead th {{
                position: sticky;
                top: 0;
                z-index: 1;
                padding: {header_padding};
                background: #F8F7FC;
                color: {PURPLE};
                font-weight: 750;
                text-align: center;
                border-bottom: 2px solid {BORDER};
                border-right: 1px solid #ECE8F5;
                white-space: normal;
                line-height: 1.35;
                overflow-wrap: anywhere;
                word-break: keep-all;
                letter-spacing: -0.01em;
            }}

            tbody td {{
                padding: {cell_padding};
                background: #FFFFFF;
                border-bottom: 1px solid #F0EDF5;
                border-right: 1px solid #F0EDF5;
                vertical-align: middle;
                line-height: 1.45;
                {cell_wrap_css}
                transition: background-color 0.12s ease;
            }}

            tbody tr:hover td {{
                background-color: #FAF8FE !important;
            }}

            tbody tr:last-child td {{
                border-bottom: 0;
            }}

            thead th:last-child,
            tbody td:last-child {{
                border-right: 0;
            }}

            tbody td.align-center {{
                text-align: center;
            }}

            tbody td.align-left {{
                text-align: left;
            }}

            tbody td.align-right {{
                text-align: right;
                font-variant-numeric: tabular-nums;
            }}

            .table-scroll::-webkit-scrollbar {{
                width: 7px;
                height: 7px;
            }}

            .table-scroll::-webkit-scrollbar-track {{
                background: #F8F7FA;
                border-radius: 6px;
            }}

            .table-scroll::-webkit-scrollbar-thumb {{
                background: #CBC5D5;
                border-radius: 6px;
            }}

            .table-scroll::-webkit-scrollbar-thumb:hover {{
                background: {PURPLE};
            }}
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

    row_height = 35 if compact else 43
    content_height = min(
        max_height,
        44 + len(dataframe) * row_height + 4,
    )

    components.html(
        html,
        height=content_height,
        scrolling=False,
    )


def sort_controls(
    dataframe,
    numeric_columns,
    key_prefix,
    label="정렬",
    default_column=None,
):
    """숫자형 정렬 기준과 오름차순/내림차순 컨트롤을 표시합니다."""
    available_columns = [
        column
        for column in numeric_columns
        if column in dataframe.columns
    ]

    if not available_columns:
        return dataframe.copy(), None, "내림차순"

    if default_column not in available_columns:
        default_column = available_columns[0]

    left, right = st.columns([2, 1])

    with left:
        sort_column = st.selectbox(
            f"{label} 기준",
            options=available_columns,
            index=available_columns.index(default_column),
            key=f"{key_prefix}_sort_column",
        )

    with right:
        sort_order = st.selectbox(
            f"{label} 순서",
            options=["내림차순", "오름차순"],
            index=0,
            key=f"{key_prefix}_sort_order",
        )

    result = dataframe.copy()
    result["_정렬용값"] = numeric_series(result[sort_column])

    result = (
        result
        .sort_values(
            "_정렬용값",
            ascending=(sort_order == "오름차순"),
            na_position="last",
            kind="stable",
        )
        .drop(columns="_정렬용값")
    )

    return result, sort_column, sort_order


def render_overview_table():
    """프로모션 개요 표를 별도 HTML 프레임으로 표시합니다."""
    overview = [
        ("최초예약일", "2026년 10월 1일 ~ 2026년 11월 30일"),
        ("출발일", "2026년 10월 1일 ~ 2027년 2월 28일"),
        ("속성", "ALL 속성 (Q 제외)"),
        (
            "제외",
            "인센티브, 코브랜드, 닷컴, 랜드온리, 업무포탈, "
            "항공, FND, 호텔, 유아, 취소예약",
        ),
        (
            "지급",
            [
                "프로모션 종료 후 인정배수 적용 매출을 기준으로 매출 및 성장률 순위 산정",
                "매출순위 및 매출성장률 기준 1위 ~ 10위 대리점 대상 VI 지급",
                "최소 매출 허들 131,612,240원(동기간 2025년 평균 매출) 미달성 대리점은 지급 대상에서 제외",
                "매출순위와 매출 성장률 기준을 동시에 충족한 경우 총 지급액이 높은 1개의 기준만 적용",
            ],
        ),
    ]

    rows = []

    for label, content in overview:
        if label == "지급":
            content_html = (
                "<ol>"
                + "".join(
                    f"<li>{escape(item)}</li>"
                    for item in content
                )
                + "</ol>"
            )
        else:
            content_html = escape(content)

        rows.append(
            "<tr>"
            f"<td class='label'>{escape(label)}</td>"
            f"<td class='content'>{content_html}</td>"
            "</tr>"
        )

    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
        <style>
            * {{
                box-sizing: border-box;
            }}

            html, body {{
                margin: 0;
                padding: 0;
                width: 100%;
                font-family: Pretendard, -apple-system,
                             BlinkMacSystemFont, "Segoe UI", sans-serif;
                color: {DARK};
                background: transparent;
            }}

            .wrap {{
                width: 100%;
                margin: 0;
                padding: 0;
                overflow: hidden;
                border: 1px solid {BORDER};
                border-radius: 12px;
                background: #FFFFFF;
                box-shadow: 0 2px 10px rgba(94, 43, 184, 0.03);
            }}

            table {{
                width: 100%;
                margin: 0;
                border-collapse: collapse;
                font-size: 13.5px;
            }}

            th {{
                padding: 12px 14px;
                text-align: center;
                background: {LIGHT_PURPLE};
                color: {PURPLE};
                font-weight: 750;
                border-bottom: 2px solid {BORDER};
            }}

            td {{
                padding: 10px 16px;
                border-top: 1px solid #F0EDF5;
                vertical-align: middle;
                line-height: 1.55;
            }}

            tbody tr:hover td {{
                background-color: #FAF8FE;
            }}

            tbody tr:last-child td {{
                border-bottom: 0;
            }}

            td.label {{
                width: 130px;
                text-align: center;
                white-space: nowrap;
                font-weight: 700;
                color: {PURPLE};
                background: #FAF9FE;
                border-right: 1px solid #F0EDF5;
            }}

            td.content {{
                text-align: left;
                color: {DARK};
            }}

            ol {{
                margin: 0;
                padding-left: 20px;
            }}

            li {{
                margin: 5px 0;
            }}
        </style>
    </head>
    <body>
        <div class="wrap">
            <table>
                <thead>
                    <tr><th>구분</th><th>내용</th></tr>
                </thead>
                <tbody>{''.join(rows)}</tbody>
            </table>
        </div>
    </body>
    </html>
    """

    components.html(html, height=330, scrolling=False)



def render_multiplier_table():
    """상품 속성별 인정 배수 표를 별도 HTML 프레임으로 표시합니다."""
    html = f"""
    <!DOCTYPE html>
    <html lang="ko">
    <head>
        <meta charset="utf-8">
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
        <style>
            * {{
                box-sizing: border-box;
            }}

            html, body {{
                margin: 0;
                padding: 0;
                width: 100%;
                font-family: Pretendard, -apple-system,
                             BlinkMacSystemFont, "Segoe UI", sans-serif;
                color: {DARK};
                background: transparent;
            }}

            .wrap {{
                width: 100%;
                margin: 0;
                padding: 0;
                overflow: hidden;
                border: 1px solid {BORDER};
                border-radius: 12px;
                background: #FFFFFF;
                box-shadow: 0 2px 10px rgba(94, 43, 184, 0.03);
            }}

            table {{
                width: 100%;
                margin: 0;
                border-collapse: separate;
                border-spacing: 0;
                table-layout: fixed;
                font-size: 13.5px;
            }}

            th, td {{
                padding: 11px 12px;
                text-align: center !important;
                vertical-align: middle !important;
                border-right: 1px solid #F0EDF5;
            }}

            th {{
                background: {LIGHT_PURPLE};
                color: {PURPLE};
                font-weight: 750;
                border-bottom: 2px solid {BORDER};
            }}

            td {{
                border-top: 1px solid #F0EDF5;
            }}

            tbody tr:hover td {{
                background-color: #FAF8FE;
            }}

            th:last-child,
            td.last-column {{
                border-right: 0;
            }}

            .badge-purple {{
                display: inline-block;
                padding: 3px 10px;
                background: {LIGHT_PURPLE};
                color: {PURPLE};
                font-weight: 750;
                border-radius: 6px;
                border: 1px solid #E4DCF5;
            }}

            .badge-mint {{
                display: inline-block;
                padding: 3px 10px;
                background: #E6FAFB;
                color: #00878C;
                font-weight: 750;
                border-radius: 6px;
                border: 1px solid #B8F2F5;
            }}
        </style>
    </head>
    <body>
        <div class="wrap">
            <table>
                <thead>
                    <tr>
                        <th>상품 속성</th>
                        <th>브랜드</th>
                        <th>인정 배수</th>
                        <th>휴일 예약</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td style="font-weight:700;">Z</td>
                        <td>제우스</td>
                        <td rowspan="2"><span class="badge-purple">2배</span></td>
                        <td rowspan="4" class="last-column"><span class="badge-mint">1.5배</span></td>
                    </tr>
                    <tr>
                        <td style="font-weight:700;">U</td>
                        <td>우리끼리</td>
                    </tr>
                    <tr>
                        <td style="font-weight:700;">P</td>
                        <td>스탠다드, 프리미엄</td>
                        <td><span class="badge-mint">1.5배</span></td>
                    </tr>
                    <tr>
                        <td style="font-weight:700;">ALL (Q 제외)</td>
                        <td>-</td>
                        <td style="font-weight:600;">1배</td>
                    </tr>
                </tbody>
            </table>
        </div>
    </body>
    </html>
    """

    components.html(html, height=220, scrolling=False)


# =========================================================
# 제목 및 프로모션 안내
# =========================================================
st.markdown(
    f"""
    <div class="dashboard-header">
        <div class="header-top-tag">
            <span class="live-dot"></span> 2026 유통점 영업 활성화 프로모션
        </div>
        <h1 class="dashboard-title">
            <span class="brand-title">HNT</span> Promotion Dashboard
        </h1>
        <div class="dashboard-subtitle">
            대리점별 실매출·보정매출 현황 및 프로모션 달성도 모니터링
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-heading">프로모션 개요 및 가이드</div>',
    unsafe_allow_html=True,
)
st.markdown(
    '<div class="subsection-heading">프로모션 상세 개요</div>',
    unsafe_allow_html=True,
)
render_overview_table()

st.markdown(
    '<div class="subsection-heading">상품 속성별 인정 배수</div>',
    unsafe_allow_html=True,
)
render_multiplier_table()

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

vi_rows = []

for criterion, group in vi_df.groupby("순위 기준", sort=False):
    group_size = len(group)

    for row_index, (_, row) in enumerate(group.iterrows()):
        cells = []

        if row_index == 0:
            cells.append(
                f"""
                <td rowspan="{group_size}"
                    style="text-align:center; vertical-align:middle;
                           font-weight:700; background:{LIGHT_PURPLE};">
                    {escape(str(criterion))}
                </td>
                """
            )

        for column in ["순위", "지급 금액", "개수", "지급액"]:
            cells.append(
                "<td style='text-align:center; vertical-align:middle;'>"
                + escape(str(row[column]))
                + "</td>"
            )

        vi_rows.append(
            "<tr>" + "".join(cells) + "</tr>"
        )

vi_table_html = f"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8">
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css" />
    <style>
        * {{
            box-sizing: border-box;
        }}

        html, body {{
            margin: 0;
            padding: 0;
            font-family: Pretendard, -apple-system,
                         BlinkMacSystemFont, "Segoe UI", sans-serif;
            color: {DARK};
            background: transparent;
        }}

        .vi-table-wrap {{
            width: 100%;
            border: 1px solid {BORDER};
            border-radius: 12px;
            overflow: hidden;
            background: #FFFFFF;
            box-shadow: 0 2px 10px rgba(94, 43, 184, 0.03);
        }}

        table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13.5px;
        }}

        th {{
            padding: 11px 12px;
            text-align: center;
            background: {LIGHT_PURPLE};
            color: {PURPLE};
            font-weight: 750;
            border-right: 1px solid #ECE8F5;
            border-bottom: 2px solid {BORDER};
        }}

        td {{
            padding: 10px 12px;
            border-top: 1px solid #F0EDF5;
            border-right: 1px solid #F0EDF5;
            line-height: 1.5;
            font-variant-numeric: tabular-nums;
        }}

        tbody tr:hover td {{
            background-color: #FAF8FE;
        }}

        th:last-child,
        td:last-child {{
            border-right: 0;
        }}
    </style>
</head>
<body>
    <div class="vi-table-wrap">
        <table>
            <thead>
                <tr>
                    <th>순위 기준</th>
                    <th>순위</th>
                    <th>지급 금액</th>
                    <th>개수</th>
                    <th>지급액</th>
                </tr>
            </thead>
            <tbody>
                {''.join(vi_rows)}
            </tbody>
        </table>
    </div>
</body>
</html>
"""

components.html(
    vi_table_html,
    height=420,
    scrolling=False,
)

st.divider()


# =========================================================
# 엑셀 데이터 불러오기
# =========================================================
if not FILE_PATH.exists():
    st.error(f"엑셀 파일을 찾을 수 없습니다: {FILE_PATH}")
    st.stop()

real_sales_df = load_sales_sheet(
    FILE_PATH,
    REAL_SALES_SHEET,
    ["실매출", "실매출액", "매출"],
    "실매출",
)

adjusted_sales_df = load_sales_sheet(
    FILE_PATH,
    ADJUSTED_SALES_SHEET,
    ["총보정매출", "보정매출", "보정 매출"],
    "보정매출",
)

holiday_sales_df = load_sales_sheet(
    FILE_PATH,
    HOLIDAY_SHEET,
    ["휴일 매출", "휴일매출", "매출"],
    "휴일 매출",
)

try:
    multiplier_source_df = pd.read_excel(
        FILE_PATH,
        sheet_name=MULTIPLIER_SHEET,
        header=1,
    ).dropna(how="all")

    multiplier_source_df.columns = [
        str(column).strip()
        for column in multiplier_source_df.columns
    ]
    multiplier_source_error = None

except Exception as error:
    multiplier_source_df = None
    multiplier_source_error = str(error)


# =========================================================
# 대리점 리스트 및 유통채널 정보
# =========================================================
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

agency_code_column = find_column(
    agency_df,
    ["대리점코드", "대리점 코드"],
    AGENCY_SHEET,
    "대리점 코드",
)
channel_column = find_column(
    agency_df,
    ["채널구분", "유통채널"],
    AGENCY_SHEET,
    "채널 구분",
)

channel_map = agency_df[
    [agency_code_column, channel_column]
].copy()

channel_map = channel_map.rename(
    columns={
        agency_code_column: "_대리점코드",
        channel_column: "채널구분",
    }
)

channel_map["_대리점코드"] = clean_code(
    channel_map["_대리점코드"]
)
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


# =========================================================
# 보정매출 기준 순위 데이터
# =========================================================
ranking_aggregations = {
    "보정매출": ("매출", "sum"),
}

for column in [
    "대리점명",
    "담당실",
    "담당팀",
    "담당세일즈",
]:
    if column in adjusted_sales_df.columns:
        ranking_aggregations[column] = (column, "first")

sales_ranking = (
    adjusted_sales_df
    .groupby("_대리점코드", as_index=False)
    .agg(**ranking_aggregations)
    .sort_values("보정매출", ascending=False)
    .reset_index(drop=True)
)

# 실매출 시트의 예약인원을 대리점코드별로 합산해 순위 데이터에 연결합니다.
if "예약인원" in real_sales_df.columns:
    reservation_by_agency = (
        real_sales_df
        .groupby("_대리점코드", as_index=False)
        .agg(예약인원=("예약인원", "sum"))
    )
    sales_ranking = sales_ranking.merge(
        reservation_by_agency,
        on="_대리점코드",
        how="left",
    )
    sales_ranking["예약인원"] = (
        numeric_series(sales_ranking["예약인원"])
        .fillna(0)
    )
else:
    sales_ranking["예약인원"] = 0

sales_ranking.insert(
    0,
    "순위",
    range(1, len(sales_ranking) + 1),
)

for column in ["대리점명", "담당실", "담당팀", "담당세일즈"]:
    if column not in sales_ranking.columns:
        sales_ranking[column] = ""

sales_ranking["대리점코드"] = (
    sales_ranking["_대리점코드"]
    .fillna("")
    .astype(str)
)
sales_ranking["대리점명"] = (
    sales_ranking["대리점명"]
    .fillna("")
    .astype(str)
    .str.strip()
)


# =========================================================
# KPI — 실매출 기준
# =========================================================
# 총 대리점 수의 계산 방식은 기존 코드와 동일하게 유지합니다.
total_agencies = sales_ranking["_대리점코드"].nunique()

if "담당세일즈" in sales_ranking.columns:
    total_staff = (
        sales_ranking["담당세일즈"]
        .dropna()
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .nunique()
    )
else:
    total_staff = 0

# 총매출과 총 예약인원은 실매출 시트 기준으로 계산합니다.
total_real_sales = real_sales_df["매출"].sum()
total_reservations = (
    numeric_series(real_sales_df["예약인원"])
    .fillna(0)
    .sum()
    if "예약인원" in real_sales_df.columns
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
        "매출발생 대리점 기준",
    )

with kpi2:
    render_kpi(
        "총 담당자 수",
        f"{total_staff:,}명",
        "담당 세일즈 수",
        accent="mint",
    )

with kpi3:
    render_kpi(
        "실매출",
        format_won(total_real_sales),
        "대리점별 실매출 합계",
    )

with kpi4:
    render_kpi(
        "예약인원",
        format_count(total_reservations),
        "실매출 시트 예약인원 합계",
        accent="mint",
    )


# =========================================================
# 탭 구성
# =========================================================
(
    tab_ranking,
    tab_office,
    tab_channel,
    tab_real_sales,
    tab_holiday,
) = st.tabs(
    [
        "매출 순위",
        "영업실별 매출현황",
        "유통채널별 매출현황",
        "실매출",
        "휴일예약 매출 비중",
    ]
)


# =========================================================
# 매출 순위 탭
# =========================================================
with tab_ranking:
    st.markdown(
        '<div class="section-heading">매출 순위</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-note">
            매출 순위는 보정매출 시트의 보정매출을 기준으로 표시합니다.
            대리점코드, 대리점명, 담당실, 담당팀, 담당세일즈로 검색할 수 있습니다.
            담당실 검색은 정확히 일치하는 결과만 표시합니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    ranking_search = st.text_input(
        "매출 순위 검색",
        placeholder="대리점코드, 대리점명, 담당실, 담당팀 또는 담당세일즈 검색",
        key="ranking_search",
    )

    ranking_table = sales_ranking[
        [
            "순위",
            "대리점코드",
            "대리점명",
            "담당실",
            "담당팀",
            "담당세일즈",
            "보정매출",
            "예약인원",
        ]
    ].copy()

    if ranking_search.strip():
        ranking_table = filter_agency_search(
            ranking_table,
            ranking_search,
            [
                "대리점코드",
                "대리점명",
                "담당실",
                "담당팀",
                "담당세일즈",
            ],
        )

    ranking_table, ranking_sort_column, ranking_sort_order = sort_controls(
        ranking_table,
        numeric_columns=["순위", "보정매출", "예약인원"],
        key_prefix="ranking",
        label="정렬",
        default_column="보정매출",
    )

    st.caption(
        f"검색 결과 {len(ranking_table):,}건 · "
        f"정렬 기준: {ranking_sort_column or '없음'} · "
        f"정렬 순서: {ranking_sort_order}"
    )

    if ranking_table.empty:
        st.info("검색 조건에 맞는 매출 순위가 없습니다.")
    else:
        ranking_display = ranking_table.copy()
        ranking_display["보정매출"] = (
            ranking_display["보정매출"].map(format_won)
        )
        ranking_display["예약인원"] = (
            ranking_display["예약인원"].map(format_count)
        )

        render_html_table(
            ranking_display,
            left_align_columns=["대리점명"],
            right_align_columns=["보정매출", "예약인원"],
            column_widths={
                "순위": "70px",
                "대리점코드": "130px",
                "대리점명": "280px",
                "담당실": "130px",
                "담당팀": "130px",
                "담당세일즈": "140px",
                "보정매출": "190px",
                "예약인원": "120px",
            },
            min_table_width="1190px",
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
        st.info("매출 배수적용 시트에 표시할 데이터가 없습니다.")
    else:
        source_width = 180
        source_min_width = max(
            1200,
            source_width * len(multiplier_source_df.columns),
        )
        source_widths = {
            column: f"{source_width}px"
            for column in multiplier_source_df.columns
        }

        render_html_table(
            multiplier_source_df,
            column_widths=source_widths,
            min_table_width=f"{source_min_width}px",
            max_height=560,
        )


# =========================================================
# 영업실별 매출현황 탭
# =========================================================
with tab_office:
    st.markdown(
        '<div class="section-heading">영업실별 매출현황</div>',
        unsafe_allow_html=True,
    )

    if "담당실" not in real_sales_df.columns:
        st.warning(
            "실매출 시트에서 영업실 컬럼을 찾지 못해 "
            "영업실별 현황을 표시할 수 없습니다."
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
                "실매출: %{y:,.0f}원"
                "<extra></extra>"
            ),
        )

        office_chart.update_layout(
            showlegend=False,
            xaxis_title="",
            yaxis_title="",
            yaxis_tickformat=",.0f",
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color=DARK, family="Pretendard, -apple-system, sans-serif"),
            margin=dict(l=20, r=20, t=28, b=20),
            xaxis=dict(showgrid=False, linecolor=BORDER, tickfont=dict(size=13, color=DARK)),
            yaxis=dict(gridcolor="#EFEAF7", zeroline=False, tickfont=dict(size=12, color=MUTED)),
            bargap=0.36,
        )

        st.plotly_chart(
            office_chart,
            use_container_width=True,
        )

        office_table = office_summary.rename(
            columns={
                "담당실": "영업실",
                "매출": "실매출",
            }
        ).copy()
        office_table["실매출"] = office_table["실매출"].map(format_won)

        render_html_table(
            office_table,
            right_align_columns=["실매출"],
            max_height=440,
        )


# =========================================================
# 유통채널별 매출현황 탭
# =========================================================
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
            유통채널별 매출은 실매출 시트를 기준으로 집계합니다.
            그래프의 기타 영역에 마우스를 올리면 세부 채널 구성비를 확인할 수 있습니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    channel_source = real_sales_df.copy()
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

    channel_source["유통채널"] = (
        channel_source["채널구분"].apply(classify_channel)
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

    channel_summary["_정렬순서"] = channel_summary[
        "유통채널"
    ].map(
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

    total_channel_sales = channel_summary["매출"].sum()

    if channel_summary.empty or total_channel_sales == 0:
        st.info("표시할 유통채널별 매출 데이터가 없습니다.")
    else:
        other_detail = (
            channel_source[
                channel_source["유통채널"].eq("기타")
            ]
            .groupby("세부채널")["매출"]
            .sum()
            .sort_values(ascending=False)
        )

        total_other_sales = other_detail.sum()

        if total_other_sales > 0:
            detail_lines = []

            for name, value in other_detail.items():
                total_share = value / total_channel_sales * 100
                other_share = value / total_other_sales * 100

                detail_lines.append(
                    f"{escape(str(name))}: "
                    f"전체 매출 {total_share:.2f}% · "
                    f"기타 내 {other_share:.2f}%"
                )

            other_hover = (
                "<b>기타 세부 구성</b><br>"
                + "<br>".join(detail_lines)
            )
        else:
            other_hover = ""

        channel_summary["기타 상세"] = channel_summary[
            "유통채널"
        ].apply(
            lambda name: other_hover if name == "기타" else ""
        )

        channel_chart = px.pie(
            channel_summary,
            names="유통채널",
            values="매출",
            hole=0.56,
            category_orders={"유통채널": channel_order},
            color="유통채널",
            color_discrete_map=channel_colors,
            custom_data=["기타 상세"],
        )

        channel_chart.update_traces(
            textposition="inside",
            texttemplate="%{label}<br>%{percent}",
            marker=dict(
                colors=[
                    channel_colors[name]
                    for name in channel_summary["유통채널"]
                ],
                line=dict(color="#A9A5B3", width=1.5),
            ),
            hovertemplate=(
                "<b>%{label}</b><br>"
                "실매출: %{value:,.0f}원<br>"
                "전체 구성비: %{percent}"
                "<br>%{customdata[0]}"
                "<extra></extra>"
            ),
        )

        channel_chart.update_layout(
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            font=dict(color=DARK, family="Pretendard, -apple-system, sans-serif"),
            legend_title_text="유통채널",
            legend=dict(font=dict(size=12, color=DARK)),
            margin=dict(l=20, r=20, t=20, b=20),
        )

        st.plotly_chart(
            channel_chart,
            use_container_width=True,
        )

        channel_table = channel_summary[
            ["유통채널", "매출"]
        ].copy()
        channel_table = channel_table.rename(
            columns={"매출": "실매출"}
        )
        channel_table["구성비"] = (
            channel_summary["매출"]
            .div(total_channel_sales)
            .mul(100)
            .map(lambda value: f"{value:.1f}%")
        )
        channel_table["실매출"] = channel_table[
            "실매출"
        ].map(format_won)

        render_html_table(
            channel_table,
            right_align_columns=["실매출", "구성비"],
            max_height=440,
        )


# =========================================================
# 실매출 탭: 검색 + 숫자 기준 정렬
# =========================================================
with tab_real_sales:
    st.markdown(
        '<div class="section-heading">실매출 데이터</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-note">
            엑셀 실매출 시트의 본부, 영업실, 영업팀, 담당세일즈,
            대리점코드, 대리점한글명, 실매출, 예약인원 데이터를 표시합니다.
            검색어와 정렬 기준을 함께 사용할 수 있습니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 새 실매출 시트의 컬럼 순서로 표시용 데이터를 생성합니다.
    real_sales_display = pd.DataFrame(index=real_sales_df.index)

    display_column_map = [
        ("본부", "본부"),
        ("담당실", "영업실"),
        ("담당팀", "영업팀"),
        ("담당세일즈", "담당세일즈"),
        ("_대리점코드", "대리점코드"),
        ("대리점명", "대리점한글명"),
        ("매출", "실매출"),
        ("예약인원", "예약인원"),
    ]

    for source_column, display_column in display_column_map:
        if source_column in real_sales_df.columns:
            real_sales_display[display_column] = real_sales_df[
                source_column
            ]

    # 검색은 표시되는 전체 컬럼에서 부분 일치로 수행합니다.
    real_sales_search = st.text_input(
        "실매출 데이터 검색",
        placeholder="본부, 영업실, 영업팀, 담당자, 대리점코드 또는 대리점명 검색",
        key="real_sales_search",
    )

    real_sales_filtered = real_sales_display.copy()

    if real_sales_search.strip():
        real_sales_filtered = filter_agency_search(
            real_sales_filtered,
            real_sales_search,
            list(real_sales_filtered.columns),
        )

    # 매출 및 예약인원 숫자 기준으로 정렬할 수 있도록 내부 정렬 컬럼을 유지합니다.
    available_numeric = [
        column
        for column in ["실매출", "예약인원"]
        if column in real_sales_filtered.columns
    ]
    real_sales_filtered, real_sort_column, real_sort_order = sort_controls(
        real_sales_filtered,
        numeric_columns=available_numeric,
        key_prefix="real_sales",
        label="숫자 정렬",
        default_column="실매출",
    )

    st.caption(
        f"검색 결과 {len(real_sales_filtered):,}건 · "
        f"정렬 기준: {real_sort_column or '없음'} · "
        f"정렬 순서: {real_sort_order}"
    )

    if real_sales_filtered.empty:
        st.info("검색 조건에 맞는 실매출 데이터가 없습니다.")
    else:
        real_sales_table = real_sales_filtered.copy()
        real_sales_table["실매출"] = real_sales_table[
            "실매출"
        ].map(format_won)
        if "예약인원" in real_sales_table.columns:
            real_sales_table["예약인원"] = real_sales_table[
                "예약인원"
            ].map(format_count)

        render_html_table(
            real_sales_table,
            left_align_columns=["대리점한글명"],
            right_align_columns=["실매출", "예약인원"],
            column_widths={
                "본부": "8%",
                "영업실": "9%",
                "영업팀": "11%",
                "담당세일즈": "9%",
                "대리점코드": "10%",
                "대리점한글명": "22%",
                "실매출": "17%",
                "예약인원": "14%",
            },
            max_height=620,
            fit_width=True,
        )


# =========================================================
# 휴일예약 매출 비중 탭: 검색 + 정렬 기준
# =========================================================
with tab_holiday:
    st.markdown(
        '<div class="section-heading">휴일예약 매출 비중</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        <div class="info-note">
            대리점별 실매출과 휴일예약 시트의 매출을 비교합니다.
            휴일예약 매출 비중은 휴일예약 매출을 실매출로 나누어 계산합니다.
            휴일예약 시트에 매출이 없는 대리점은 휴일예약 매출을 0원으로 표시합니다.
        </div>
        """,
        unsafe_allow_html=True,
    )

    holiday_aggregations = {
        "실매출": ("매출", "sum"),
    }

    for column in [
        "본부",
        "담당실",
        "담당팀",
        "담당세일즈",
        "대리점명",
    ]:
        if column in real_sales_df.columns:
            holiday_aggregations[column] = (column, "first")

    holiday_comparison = (
        real_sales_df
        .groupby("_대리점코드", as_index=False)
        .agg(**holiday_aggregations)
    )

    holiday_by_agency = (
        holiday_sales_df
        .groupby("_대리점코드", as_index=False)
        .agg(휴일예약매출=("매출", "sum"))
    )

    holiday_comparison = holiday_comparison.merge(
        holiday_by_agency,
        on="_대리점코드",
        how="left",
    )

    holiday_comparison["휴일예약매출"] = (
        holiday_comparison["휴일예약매출"].fillna(0)
    )

    for column in [
        "본부",
        "담당실",
        "담당팀",
        "담당세일즈",
        "대리점명",
    ]:
        if column not in holiday_comparison.columns:
            holiday_comparison[column] = ""

        holiday_comparison[column] = (
            holiday_comparison[column]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    holiday_comparison["대리점코드"] = clean_code(
        holiday_comparison["_대리점코드"]
    ).fillna("")

    holiday_comparison["_휴일예약 매출 비중"] = 0.0
    has_real_sales = holiday_comparison["실매출"].ne(0)

    holiday_comparison.loc[
        has_real_sales,
        "_휴일예약 매출 비중",
    ] = (
        holiday_comparison.loc[
            has_real_sales,
            "휴일예약매출",
        ]
        .div(
            holiday_comparison.loc[
                has_real_sales,
                "실매출",
            ]
        )
        .mul(100)
    )

    holiday_search = st.text_input(
        "휴일예약 매출 검색",
        placeholder="대리점코드, 대리점명, 담당실, 담당팀 또는 담당세일즈 검색",
        key="holiday_search",
    )

    holiday_filtered = holiday_comparison.copy()

    if holiday_search.strip():
        holiday_filtered = filter_agency_search(
            holiday_filtered,
            holiday_search,
            [
                "대리점코드",
                "대리점명",
                "담당실",
                "담당팀",
                "담당세일즈",
            ],
        )

    holiday_filtered = holiday_filtered.rename(
        columns={
            "_휴일예약 매출 비중": "휴일예약 매출 비중",
            "휴일예약매출": "휴일예약 매출",
        }
    )

    holiday_filtered, holiday_sort_column, holiday_sort_order = sort_controls(
        holiday_filtered,
        numeric_columns=[
            "실매출",
            "휴일예약 매출",
            "휴일예약 매출 비중",
        ],
        key_prefix="holiday",
        label="정렬",
        default_column="휴일예약 매출 비중",
    )

    st.caption(
        f"검색 결과 {len(holiday_filtered):,}건 · "
        f"정렬 기준: {holiday_sort_column or '없음'} · "
        f"정렬 순서: {holiday_sort_order}"
    )

    if holiday_filtered.empty:
        st.info("검색 조건에 맞는 대리점이 없습니다.")
    else:
        holiday_table = holiday_filtered[
            [
                "대리점코드",
                "대리점명",
                "담당실",
                "담당팀",
                "담당세일즈",
                "실매출",
                "휴일예약 매출",
                "휴일예약 매출 비중",
            ]
        ].copy()

        holiday_table["실매출"] = holiday_table[
            "실매출"
        ].map(format_won)
        holiday_table["휴일예약 매출"] = holiday_table[
            "휴일예약 매출"
        ].map(format_won)
        holiday_table["휴일예약 매출 비중"] = holiday_table[
            "휴일예약 매출 비중"
        ].map(lambda value: f"{value:.1f}%")

        render_html_table(
            holiday_table,
            left_align_columns=["대리점명"],
            right_align_columns=[
                "실매출",
                "휴일예약 매출",
                "휴일예약 매출 비중",
            ],
            column_widths={
                "대리점코드": "9%",
                "대리점명": "17%",
                "담당실": "10%",
                "담당팀": "10%",
                "담당세일즈": "11%",
                "실매출": "15%",
                "휴일예약 매출": "15%",
                "휴일예약 매출 비중": "13%",
            },
            max_height=560,
            fit_width=True,
            compact=True,
        )

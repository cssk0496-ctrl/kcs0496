from __future__ import annotations

import time
import re
from datetime import date
from io import BytesIO

import pandas as pd
import plotly.express as px
import requests
import streamlit as st


st.set_page_config(
    page_title="팀 예산 관리 시스템",
    page_icon="📊",
    layout="wide",
)

API_URL = (
    "https://script.google.com/macros/s/"
    "AKfycbw7NN0cP98boE3IjD7-a7yC9l2sMS81y0oMuX87eAZxusCDQtlh2DlIXpaWR4X55bs5GQ/"
    "exec"
)
CATEGORIES = ["복리후생비", "수선비", "출장비", "접대비"]
TEAMS = ["기획팀", "개발팀", "제조팀", "경영팀"]
MEMBERS = ["팀장", "팀원"]
COLUMNS = ["ID", "연월일", "팀명", "팀원", "항목", "금액"]


def empty_data() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def normalize_date(value: object) -> str:
    text = str(value).strip()
    iso_match = re.match(r"^(\d{4})-(\d{2})(?:-(\d{2}))?", text)
    if iso_match:
        day = iso_match.group(3) or "01"
        return f"{iso_match.group(1)}-{iso_match.group(2)}-{day}"

    english_match = re.search(
        r"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\s+(\d{1,2})\s+(\d{4})",
        text,
    )
    if english_match:
        month_number = {
            "Jan": "01", "Feb": "02", "Mar": "03", "Apr": "04",
            "May": "05", "Jun": "06", "Jul": "07", "Aug": "08",
            "Sep": "09", "Oct": "10", "Nov": "11", "Dec": "12",
        }[english_match.group(1)]
        return (
            f"{english_match.group(3)}-{month_number}-"
            f"{int(english_match.group(2)):02d}"
        )
    return text


def date_label(value: object) -> str:
    normalized = normalize_date(value)
    if re.match(r"^\d{4}-\d{2}-\d{2}$", normalized):
        year, month, day = normalized.split("-")
        return f"{year}년 {month}월 {day}일"
    return normalized


def parse_api_response(response: requests.Response) -> dict:
    response.raise_for_status()
    try:
        result = response.json()
    except requests.JSONDecodeError as exc:
        raise RuntimeError(
            "Apps Script가 JSON을 반환하지 않았습니다. 웹 앱 배포 URL과 "
            "액세스 권한(모든 사용자)을 확인하세요."
        ) from exc
    if not result.get("success"):
        raise RuntimeError(result.get("error", "Apps Script 요청에 실패했습니다."))
    return result


def call_api(payload: dict) -> dict:
    response = requests.post(
        API_URL,
        json=payload,
        timeout=30,
        allow_redirects=True,
    )
    return parse_api_response(response)


def load_data() -> pd.DataFrame:
    response = requests.get(
        API_URL,
        params={"action": "list"},
        timeout=30,
        allow_redirects=True,
    )
    result = parse_api_response(response)
    records = result.get("data", [])
    if not records:
        return empty_data()
    for record in records:
        if "연월일" not in record and "연월" in record:
            record["연월일"] = record["연월"]
    result = pd.DataFrame(records, columns=COLUMNS)
    result["ID"] = result["ID"].astype(str)
    result["연월일"] = result["연월일"].map(normalize_date)
    result["팀명"] = (
        result["팀명"]
        .fillna("미지정")
        .astype(str)
        .replace({"None": "미지정", "nan": "미지정", "": "미지정"})
    )
    result["금액"] = pd.to_numeric(result["금액"], errors="coerce").fillna(0).astype(int)
    return result


def normalize_data(data: pd.DataFrame) -> pd.DataFrame:
    if "연월일" not in data.columns and "연월" in data.columns:
        data = data.rename(columns={"연월": "연월일"})
    missing = set(COLUMNS) - set(data.columns)
    if missing:
        raise ValueError(f"필수 열이 없습니다: {', '.join(sorted(missing))}")

    result = data[COLUMNS].copy()
    result["ID"] = result["ID"].astype(str)
    result["금액"] = pd.to_numeric(result["금액"], errors="raise").astype("int64")
    result["연월일"] = result["연월일"].map(normalize_date)
    result["팀명"] = result["팀명"].astype(str)
    result["팀원"] = result["팀원"].astype(str)
    result["항목"] = result["항목"].astype(str)
    if (~result["연월일"].str.match(r"^\d{4}-\d{2}-\d{2}$")).any():
        raise ValueError("연월일은 YYYY-MM-DD 형식이어야 합니다.")
    if (result["금액"] < 0).any():
        raise ValueError("금액은 0원 이상이어야 합니다.")
    return result


def replace_all_data(data: pd.DataFrame) -> None:
    records = [
        {
            "ID": str(row["ID"]),
            "연월일": str(row["연월일"]),
            "팀명": str(row["팀명"]),
            "팀원": str(row["팀원"]),
            "항목": str(row["항목"]),
            "금액": int(row["금액"]),
        }
        for row in data.to_dict("records")
    ]
    call_api({"action": "replaceAll", "data": records})


def delete_entries(ids: set[str]) -> None:
    call_api({"action": "delete", "ids": list(ids)})


def to_excel(data: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="예산내역")
    return output.getvalue()


st.markdown(
    """
    <style>
    .block-container {max-width: 1180px; padding-top: 2rem;}
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 18px;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.05);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 팀 예산 관리 시스템")
st.caption("부장님 보고용 월별 예산 취합 및 대시보드 · Google Sheets 실시간 연동")

try:
    data = load_data()
except Exception as exc:
    st.error("Google Sheets에 연결하지 못했습니다.")
    st.info(
        "Apps Script가 웹 앱으로 배포되어 있고 액세스 권한이 "
        "'모든 사용자'인지 확인하세요."
    )
    st.code(str(exc))
    st.stop()

input_tab, dashboard_tab = st.tabs(["📝 데이터 입력", "📈 전체 대시보드"])

with input_tab:
    form_col, history_col = st.columns([1, 2], gap="large")

    with form_col:
        st.subheader("내역 입력")
        with st.form("budget_form", clear_on_submit=True):
            team = st.selectbox("팀명", TEAMS)
            member = st.selectbox("팀원 선택", MEMBERS)
            selected_date = st.date_input("사용일", value=date.today())
            category = st.selectbox("예산 항목", CATEGORIES)
            amount = st.number_input(
                "사용 금액 (원)",
                min_value=0,
                step=10_000,
                format="%d",
            )
            submitted = st.form_submit_button(
                "기록 저장하기",
                type="primary",
                use_container_width=True,
            )

        if submitted:
            call_api(
                {
                    "action": "append",
                    "ID": str(time.time_ns()),
                    "연월일": selected_date.strftime("%Y-%m-%d"),
                    "팀명": team,
                    "팀원": member,
                    "항목": category,
                    "금액": int(amount),
                }
            )
            st.success("Google Sheets에 정상적으로 저장되었습니다.")
            st.rerun()

        st.divider()
        st.subheader("파일로 불러오기")
        uploaded_file = st.file_uploader(
            "이 앱에서 내려받은 CSV 파일을 선택하세요.",
            type=["csv"],
        )
        if uploaded_file is not None and st.button(
            "CSV 데이터로 전체 교체",
            use_container_width=True,
        ):
            try:
                uploaded_data = normalize_data(pd.read_csv(uploaded_file))
                replace_all_data(uploaded_data)
                st.success("Google Sheets 데이터를 교체했습니다.")
                st.rerun()
            except Exception as exc:
                st.error(f"파일을 불러오지 못했습니다: {exc}")

    with history_col:
        st.subheader("📂 최근 입력 내역")
        if data.empty:
            st.info("등록된 데이터가 없습니다.")
        else:
            display_data = data.copy()
            display_data["연월일"] = display_data["연월일"].map(date_label)
            display_data["금액"] = display_data["금액"].map(
                lambda value: f"{int(value):,}원"
            )
            st.dataframe(
                display_data.drop(columns=["ID"]),
                hide_index=True,
                use_container_width=True,
            )

            delete_options = {
                f"{date_label(row['연월일'])} | {row['팀명']} | "
                f"{row['팀원']} | {row['항목']} | "
                f"{int(row['금액']):,}원": str(row["ID"])
                for _, row in data.iterrows()
            }
            selected_entries = st.multiselect(
                "삭제할 내역 선택",
                options=list(delete_options),
            )
            delete_col, clear_col = st.columns(2)
            with delete_col:
                if st.button(
                    "선택 내역 삭제",
                    disabled=not selected_entries,
                    use_container_width=True,
                ):
                    delete_entries(
                        {delete_options[label] for label in selected_entries}
                    )
                    st.rerun()
            with clear_col:
                clear_confirmed = st.checkbox("전체 초기화 확인")
                if st.button(
                    "모든 데이터 초기화",
                    disabled=not clear_confirmed,
                    use_container_width=True,
                ):
                    replace_all_data(empty_data())
                    st.rerun()

            csv_data = data.to_csv(index=False).encode("utf-8-sig")
            download_col1, download_col2 = st.columns(2)
            with download_col1:
                st.download_button(
                    "CSV 다운로드",
                    data=csv_data,
                    file_name="team_budget_data.csv",
                    mime="text/csv",
                    use_container_width=True,
                )
            with download_col2:
                st.download_button(
                    "Excel 다운로드",
                    data=to_excel(data),
                    file_name="team_budget_data.xlsx",
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True,
                )

with dashboard_tab:
    total = int(data["금액"].sum()) if not data.empty else 0
    category_totals = (
        data.groupby("항목", as_index=False)["금액"].sum()
        if not data.empty
        else pd.DataFrame(columns=["항목", "금액"])
    )
    top_category = (
        category_totals.loc[category_totals["금액"].idxmax()]
        if not category_totals.empty
        else None
    )

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("전체 누적 사용액", f"{total:,}원")
    metric2.metric(
        "이번 달 최대 사용 항목",
        (
            f"{top_category['항목']} ({int(top_category['금액']):,}원)"
            if top_category is not None
            else "-"
        ),
    )
    metric3.metric("데이터 건수", f"{len(data):,}건")

    if data.empty:
        st.info("데이터를 입력하면 차트와 월별 취합표가 표시됩니다.")
    else:
        summary_data = data.copy()
        summary_data["연월"] = summary_data["연월일"].str.slice(0, 7)
        chart_col1, chart_col2 = st.columns(2, gap="large")
        with chart_col1:
            st.subheader("항목별 예산 분포")
            fig_category = px.pie(
                category_totals,
                names="항목",
                values="금액",
                hole=0.65,
                color="항목",
                color_discrete_map={
                    "복리후생비": "#3b82f6",
                    "수선비": "#10b981",
                    "출장비": "#8b5cf6",
                    "접대비": "#f59e0b",
                },
            )
            fig_category.update_layout(
                margin=dict(l=10, r=10, t=20, b=10),
                legend_title_text="",
            )
            st.plotly_chart(fig_category, use_container_width=True)

        with chart_col2:
            st.subheader("팀원별 누적 사용액")
            member_totals = data.groupby("팀원", as_index=False)["금액"].sum()
            fig_member = px.bar(
                member_totals,
                x="팀원",
                y="금액",
                text_auto=",.0f",
                color_discrete_sequence=["#60a5fa"],
            )
            fig_member.update_layout(
                margin=dict(l=10, r=10, t=20, b=10),
                yaxis_title="사용 금액(원)",
                xaxis_title="",
                showlegend=False,
            )
            st.plotly_chart(fig_member, use_container_width=True)

        st.subheader("📅 월별·항목별 요약 테이블")
        summary = (
            pd.pivot_table(
                summary_data,
                index="연월",
                columns="항목",
                values="금액",
                aggfunc="sum",
                fill_value=0,
            )
            .reindex(columns=CATEGORIES, fill_value=0)
            .sort_index(ascending=False)
        )
        summary["합계"] = summary.sum(axis=1)
        summary.index = summary.index.map(
            lambda value: f"{value[:4]}년 {value[5:7]}월"
        )
        summary.index.name = "연월"
        st.dataframe(
            summary.reset_index(),
            hide_index=True,
            use_container_width=True,
        )

st.caption("입력한 데이터는 연결된 Google 스프레드시트에 실시간 저장됩니다.")

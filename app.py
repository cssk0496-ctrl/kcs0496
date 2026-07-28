from __future__ import annotations

from datetime import date
from io import BytesIO

import pandas as pd
import plotly.express as px
import streamlit as st


st.set_page_config(
    page_title="팀 예산 관리 시스템",
    page_icon="📊",
    layout="wide",
)

CATEGORIES = ["수선유지비", "비품", "개량공사"]
MEMBERS = ["부장님", "팀원1", "팀원2", "팀원3", "팀원4"]
COLUMNS = ["ID", "연월", "팀원", "항목", "금액"]


def empty_data() -> pd.DataFrame:
    return pd.DataFrame(columns=COLUMNS)


def normalize_data(data: pd.DataFrame) -> pd.DataFrame:
    missing = set(COLUMNS) - set(data.columns)
    if missing:
        raise ValueError(f"필수 열이 없습니다: {', '.join(sorted(missing))}")

    result = data[COLUMNS].copy()
    result["ID"] = pd.to_numeric(result["ID"], errors="raise").astype("int64")
    result["금액"] = pd.to_numeric(result["금액"], errors="raise").astype("int64")
    result["연월"] = result["연월"].astype(str)
    result["팀원"] = result["팀원"].astype(str)
    result["항목"] = result["항목"].astype(str)

    if (~result["연월"].str.match(r"^\d{4}-\d{2}$")).any():
        raise ValueError("연월은 YYYY-MM 형식이어야 합니다.")
    if (result["금액"] < 0).any():
        raise ValueError("금액은 0원 이상이어야 합니다.")
    return result


def to_excel(data: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        data.to_excel(writer, index=False, sheet_name="예산내역")
    return output.getvalue()


if "budget_data" not in st.session_state:
    st.session_state.budget_data = empty_data()
if "next_id" not in st.session_state:
    st.session_state.next_id = 1


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
st.caption("부장님 보고용 월별 예산 취합 및 대시보드")

input_tab, dashboard_tab = st.tabs(["📝 데이터 입력", "📈 전체 대시보드"])

with input_tab:
    form_col, history_col = st.columns([1, 2], gap="large")

    with form_col:
        st.subheader("내역 입력")
        with st.form("budget_form", clear_on_submit=True):
            member = st.selectbox("팀원 선택", MEMBERS)
            selected_month = st.date_input(
                "해당 월",
                value=date.today().replace(day=1),
            )
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
            new_row = pd.DataFrame(
                [
                    {
                        "ID": st.session_state.next_id,
                        "연월": selected_month.strftime("%Y-%m"),
                        "팀원": member,
                        "항목": category,
                        "금액": int(amount),
                    }
                ]
            )
            st.session_state.budget_data = pd.concat(
                [new_row, st.session_state.budget_data],
                ignore_index=True,
            )
            st.session_state.next_id += 1
            st.success("예산 데이터가 정상적으로 기록되었습니다.")

        st.divider()
        st.subheader("파일로 불러오기")
        uploaded_file = st.file_uploader(
            "이 앱에서 내려받은 CSV 파일을 선택하세요.",
            type=["csv"],
        )
        if uploaded_file is not None and st.button(
            "CSV 데이터 적용",
            use_container_width=True,
        ):
            try:
                uploaded_data = pd.read_csv(uploaded_file)
                uploaded_data = normalize_data(uploaded_data)
                st.session_state.budget_data = uploaded_data
                st.session_state.next_id = (
                    int(uploaded_data["ID"].max()) + 1
                    if not uploaded_data.empty
                    else 1
                )
                st.success("CSV 데이터를 불러왔습니다.")
                st.rerun()
            except Exception as exc:
                st.error(f"파일을 불러오지 못했습니다: {exc}")

    with history_col:
        st.subheader("📂 최근 입력 내역")
        data = st.session_state.budget_data

        if data.empty:
            st.info("등록된 데이터가 없습니다.")
        else:
            display_data = data.copy()
            display_data["금액"] = display_data["금액"].map(
                lambda value: f"{int(value):,}원"
            )
            st.dataframe(
                display_data.drop(columns=["ID"]),
                hide_index=True,
                use_container_width=True,
            )

            delete_options = {
                f"{row['연월']} | {row['팀원']} | {row['항목']} | "
                f"{int(row['금액']):,}원 (ID {int(row['ID'])})": int(row["ID"])
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
                    selected_ids = {
                        delete_options[label] for label in selected_entries
                    }
                    st.session_state.budget_data = data[
                        ~data["ID"].isin(selected_ids)
                    ].reset_index(drop=True)
                    st.rerun()
            with clear_col:
                if st.button(
                    "모든 데이터 초기화",
                    type="secondary",
                    use_container_width=True,
                ):
                    st.session_state.budget_data = empty_data()
                    st.session_state.next_id = 1
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
    data = st.session_state.budget_data
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
        "최대 사용 항목",
        (
            f"{top_category['항목']} "
            f"({int(top_category['금액']):,}원)"
            if top_category is not None
            else "-"
        ),
    )
    metric3.metric("데이터 건수", f"{len(data):,}건")

    if data.empty:
        st.info("데이터를 입력하면 차트와 월별 취합표가 표시됩니다.")
    else:
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
                    "수선유지비": "#3b82f6",
                    "비품": "#10b981",
                    "개량공사": "#8b5cf6",
                },
            )
            fig_category.update_layout(
                margin=dict(l=10, r=10, t=20, b=10),
                legend_title_text="",
            )
            st.plotly_chart(fig_category, use_container_width=True)

        with chart_col2:
            st.subheader("팀원별 누적 사용액")
            member_totals = data.groupby(
                "팀원",
                as_index=False,
            )["금액"].sum()
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
                data,
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
        summary = summary.reset_index()
        st.dataframe(
            summary,
            hide_index=True,
            use_container_width=True,
            column_config={
                category: st.column_config.NumberColumn(
                    category,
                    format="localized",
                )
                for category in [*CATEGORIES, "합계"]
            },
        )

st.caption(
    "※ Streamlit Community Cloud가 재시작되면 입력 데이터가 초기화될 수 "
    "있으므로 CSV 또는 Excel 파일을 내려받아 보관하세요."
)

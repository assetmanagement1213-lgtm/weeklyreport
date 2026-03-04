
def app():
    import streamlit as st
    import pandas as pd
    import gspread
    from google.oauth2.service_account import Credentials
    from datetime import date
    import requests
    import matplotlib.pyplot as plt
    import plotly.express as px

    q, w = st.columns([4,1])
    with q:
        st.markdown("""
            <style>
                .header-subactivity h1 {
                    margin: 0;
                    font-size: 45px;
                    font-weight: 700;
                    color: #1f2937;
                }   
            </style>
            <div class="header-subactivity">
                <h1>Issues</h1>
            </div>""",unsafe_allow_html=True)
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )

    client = gspread.authorize(creds)

    spreadsheet_name = "ASSET MANAGEMENT 2026"
    sheet = client.open(spreadsheet_name)
    ws= sheet.worksheet("Feedback Issues")
    values = ws.get("A:I")
    df= pd.DataFrame(values[1:], columns=values[0])

    # ================= WEEK FILTER ==================
    weeks = sorted(df["Week"].dropna().unique())

    week1_start = date(2026, 1, 2)
    today = date.today()
    days_since = (today - week1_start).days
    current_week = (days_since // 7)
    default_week = f"Week {current_week}"
    default_week = [default_week] if default_week in weeks else []

    with w:
        week_filter = st.multiselect(
            "Pilih Week",
            weeks,
            default=default_week,
            max_selections=None
        )

    def format_week_title(week_filter: list) -> str:
        if not week_filter:
            return "Semua Week"

        if len(week_filter) == 1:
            return week_filter[0]

        # urutkan biar aman
        week_sorted = sorted(
            week_filter,
            key=lambda x: int(x.replace("Week", "").strip())
        )

    filtered_df = df[df["Week"].isin(week_filter)]

    df_open = df[
        (df["Status"] == "Open") &
        (~df["Week"].isin(week_filter))
    ]

    x1, x2, x3, x4,x5 = st.columns(5)
    st.markdown("""
            <style>

            .metric-card {
                background: black;
                padding : 10px;
                border-radius: 12px;
                border: 1px solid #d9d9d9;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                text-align: left;
            }

            .metric-card h4 {
                margin: 0;
                margin-left : 20px;
                margin-bottom : -30px;
                text-align: left;
                font-size: 16px;
                font-weight: 600;
                color: white;
            }

            .metric-card .value {
                margin-top : -100px;
                margin-left : 20px;
                font-size: 48px;
                text-align: left;
                font-weight: 700;
                margin-top: 0px;
                color: white;
            }

            </style>
            """, unsafe_allow_html=True)
    total_issue = df["Status"].count()
    total_close = (df["Status"] == "Close").sum()

    if total_issue > 0:
        closing_rate = (total_close / total_issue) * 100
    else:
        closing_rate = 0
    with x1:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Closing Rate All Issues</h4>
                <div class="value">{closing_rate:.1f}%</div>
            </div>
        """, unsafe_allow_html=True)    
    st.divider()
    st.markdown("""
        <style>
        .white-box {
            background-color: white;
            padding: 20px 24px;
            border-radius: 16px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            margin-bottom: 24px;
        }
        [class="stVerticalBlock st-emotion-cache-1ne20ew e12zf7d53"] {
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
            text-align: center;
            diplay: inline-block;
        }
        </style>
        """, unsafe_allow_html=True)
    judul_week = format_week_title(week_filter)
    st.markdown(
            f"""
            <style>
                .judul h2 {{
                    font-size: 35px;
                    font-weight: 600;
                    color: black;
                    text-align: center;
                    margin-top :-50px;
                }}
            </style>
            <div class="judul">
                <h2>Issues {judul_week}</h2>
            </div>
            """,
            unsafe_allow_html=True)
    
    st.markdown("""
        <style>
        div[data-baseweb="input"] {
            max-width: 300px;
        }
        </style>
        """, unsafe_allow_html=True)
    from datetime import datetime
    import pandas as pd
    cols = st.columns(3)
    for idx, row in filtered_df.iterrows():

        col = cols[idx % 3]

        with col:
            with st.container(border=True):

                st.markdown(f"### 🚧 {row.Issue}")
                st.write("**Action:**", row.Solusi)

                due_value = row["Due Date"]

                if pd.notna(due_value) and due_value != "":
                    due_value = pd.to_datetime(due_value).date()
                else:
                    due_value = None

                pic_input = st.text_input(
                    "**PIC**",
                    value=row.PIC if row.PIC else "",
                    key=f"pic_{idx}"
                )

                due_date_input = st.date_input(
                    "**Due Date**",
                    value=due_value,
                    key=f"due_{idx}"
                )

                if st.button("Update", key=f"btn_{idx}"):

                    row_number = idx + 2 

                    ws.update_cell(row_number, 7, pic_input)
                    ws.update_cell(row_number, 9, str(due_date_input))

                    st.success("Updated ✅")
    st.divider()
    st.markdown("## Open Issues from Previous Weeks")
    df_open = df_open.reset_index(drop=True)  
    if df_open.empty:
        st.success("Tidak ada issues open dari week sebelumnya ✅")
    else:
        st.dataframe(
            df_open[
                ["Week", "Date", "Issue", "PIC", "Status", "Due Date"]
            ],
            use_container_width=True,hide_index=True
    )


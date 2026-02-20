
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
                <h1>Training Complience Rate</h1>
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

    spreadsheet_name = "TRAINING 2026"
    sheet = client.open(spreadsheet_name)

    # === ambil sheet induksi ===
    nameSheetDok = "DOKUMENTASI"
    sheet_dok = client.open(nameSheetDok)
    ws_training = sheet.worksheet("Training Complience Rate")
    values_training = ws_training.get("E:R")
    df_training = pd.DataFrame(values_training[1:], columns=values_training[0])

    # === ambil sheet dokumentasi ===
    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("D:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    # ================= WEEK FILTER ==================
    weeks = sorted(df_training["Week"].dropna().unique())

    week1_start = date(2026, 1, 2)
    today = date.today()
    days_since = (today - week1_start).days
    current_week = (days_since // 7) + 1
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

    filterTraining = df_training[df_training["Week"].isin(week_filter)]
    training_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Complience Rate")
    ]
    with w : 
        jenisTraining = sorted(filterTraining["Judul Training"].dropna().unique())
        training_filter = st.multiselect(
            "Pilih Training",
            jenisTraining, default=jenisTraining,
            max_selections=None
        ) 

    filtered_training = filterTraining[filterTraining["Judul Training"].isin(training_filter)]
    filtered_training2026 = df_training[df_training["Judul Training"].isin(training_filter)]
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
    with x1:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Peserta Training</h4>
                <div class="value">{len(filtered_training)}</div>
            </div>
        """, unsafe_allow_html=True)
    with x2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Perusahaan</h4>
                <div class="value">{filtered_training["Perusahaan"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x3:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Judul Training</h4>
                <div class="value">{filtered_training["Judul Training"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x4:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Departemen</h4>
                <div class="value">{filtered_training["Departement"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    total_jpl = pd.to_numeric(
        filtered_training["Jam Pelatihan"],
        errors="coerce"
    ).sum()

    with x5:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total JPL</h4>
                <div class="value">{int(total_jpl)}</div>
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

    with st.container(border=True):
        col1, div, col2 = st.columns([3,0.1,3])
        with col1 :
            judul_week = format_week_title(week_filter)
            st.markdown(
                f"""
                <style>
                    .grafik h2 {{
                        margin: 0;
                        font-size: 30px;
                        font-weight: 200;
                        color: black;
                        text-align: center;
                    }}
                </style>
                <div class="grafik">
                    <h2>Data {judul_week} </h2>
                </div>
                """,
                unsafe_allow_html=True
            )
            pivot = filtered_training.pivot_table(
                index="Departement",
                columns="Judul Training",
                values="Nama",
                aggfunc="count",
                fill_value=0
            ).reset_index()

            possible_training = sorted(
                filtered_training["Judul Training"].dropna().unique()
            )

            available_training = [
                jt for jt in possible_training if jt in pivot.columns
            ]

            training_counts = (
                filtered_training["Judul Training"]
                .value_counts()
                .reindex(possible_training, fill_value=0)
                .reset_index()
            )

            training_counts.columns = ["Judul Training", "Total"]

            fig1 = px.bar(
                training_counts,
                x="Total",
                y="Judul Training",
                text="Total",
                orientation="h", template="simple_white"
            )

            fig1.update_traces(
                textposition="outside",
                textfont_size=14, marker_color="#134f5c"
            )

            fig1.update_layout(
                height=420,
                plot_bgcolor="white",
                paper_bgcolor="white",
                yaxis=dict(autorange="reversed")
            )

            st.plotly_chart(fig1, use_container_width=True)

            with st.expander("Data Complience Rate Weekly"):

                if len(available_training) == 0:
                    pivot["Total"] = 0
                else:
                    pivot["Total"] = pivot[available_training].sum(axis=1)

                # --- Urutkan pivot ---
                pivot = pivot.sort_values("Total", ascending=False, ignore_index=True)
                pivot.index = pivot.index + 1
                st.dataframe(pivot, height=400)
        
        with col2:
            st.markdown(
                f"""
                <style>
                    .grafik h2 {{
                        margin: 0;
                        font-size: 30px;
                        font-weight: 200;
                        color: black;
                        text-align: center;
                    }}
                </style>
                <div class="grafik">
                    <h2>Data 2026</h2>
                </div>
                """,
                unsafe_allow_html=True
            )
            pivot2026 = df_training.pivot_table(
                index="Departement",
                columns="Judul Training",
                values="Nama",
                aggfunc="count",
                fill_value=0
            ).reset_index()

            possible_training_2026 = sorted(
                df_training["Judul Training"].dropna().unique()
            )

            training_counts2026 = (
                df_training["Judul Training"]
                .value_counts()
                .reindex(possible_training_2026, fill_value=0)
                .reset_index()
            )

            training_counts2026.columns = ["Judul Training", "Total"]

            fig2 = px.pie(
                training_counts2026,
                names="Judul Training",
                values="Total",
                hole=0.4, color_discrete_sequence=px.colors.sequential.Teal
            )

            fig2.update_traces(
                textinfo="label+value",
                textfont_size=12
            )

            total_all = training_counts2026["Total"].sum()

            fig2.update_layout(
                showlegend=False,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420,
                margin=dict(t=20, b=20, l=20, r=20),
                annotations=[
                    dict(
                        text=f"<b>{total_all}</b>",
                        x=0.5,
                        y=0.5,
                        font=dict(size=26, color="black"),
                        showarrow=False
                    )
                ]
            )

            st.plotly_chart(fig2, use_container_width=True)
            with st.expander("Data Sharing Knowledge 2026"):

                available_training2026 = [bu for bu in possible_training_2026 if bu in pivot2026.columns]

                if len(available_training2026) == 0:
                    pivot2026["Total"] = 0
                else:
                    pivot2026["Total"] = pivot2026[available_training2026].sum(axis=1)

                pivot2026 = pivot2026.sort_values("Total", ascending=False, ignore_index=True)
                pivot2026.index = pivot2026.index + 1
                st.dataframe(pivot2026, height=400)
        
    st.markdown("""
        <style>
            [data-testid="pivot"]:nth-child(2){
                background-color: lightgrey;
            }
        </style>
        """, unsafe_allow_html=True)

    st.divider()
    st.markdown("""
        <style>
        .square-img {
            width: 100%;
            aspect-ratio: 1 / 1;
            overflow: hidden;
            border-radius: 14px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        }

        .square-img img {
            width: 100%;
            height: 100%;
            object-fit: cover;
            transition: 0.3s ease-in-out;
        }

        .square-img img:hover {
            transform: scale(1.05);
        }

        .img-caption {
            text-align: center;
            font-size: 14px;
            margin-top: 6px;
            color: #444;
        }
        </style>
        """, unsafe_allow_html=True)
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
                <h1>📸 Dokumentasi Kegiatan</h1>
            </div>""",unsafe_allow_html=True)
    import requests
    import base64

    cols_per_row = 3
    rows = training_dokumentasi.to_dict("records")

    for i in range(0, len(rows), cols_per_row):
        cols = st.columns(cols_per_row)

        for col, row in zip(cols, rows[i:i+cols_per_row]):
            url = row["url_clean"]

            if url:
                try:
                    response = requests.get(url, stream=True)

                    content_type = response.headers.get("Content-Type", "")

                    with col:
                        if "image" in content_type:
                            img_base64 = base64.b64encode(response.content).decode()
                            st.markdown(f"""
                            <div class="square-img">
                                <img src="data:image/jpeg;base64,{img_base64}">
                            </div>
                            <div class="img-caption">
                                {row.get("Keterangan","")}
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.warning("Link tidak mengembalikan file gambar.")
                            st.write(row["url_clean"])

                except Exception as e:
                    with col:

                        st.error(f"Error load: {e}")


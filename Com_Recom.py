
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
                <h1>Commissioning/Recommissioning</h1>
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

    spreadsheet_name = "COMMISSIONING & RECOMMISSIONING ALL BU 2026"
    sheet = client.open(spreadsheet_name)

    nameSheetDok = "DOKUMENTASI"
    sheet_dok = client.open(nameSheetDok)
    ws_recom = sheet.worksheet("2026")
    values_recom = ws_recom.get("E:V")
    df_recom = pd.DataFrame(values_recom[1:], columns=values_recom[0])

    ws_temuan=sheet.worksheet("Olah Temuan")
    values_temuan=ws_temuan.get("A:H")
    df_temuan = pd.DataFrame(values_temuan[1:],columns=values_temuan[0])

    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("D:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    weeks = sorted(df_recom["Week"].dropna().unique())

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

        return f"{week_sorted[0]} – {week_sorted[-1]}"

    filteredrecom = df_recom[df_recom["Week"].isin(week_filter)]
    recom_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"].isin(["Recommissioning", "Commissining"])) & (dokumentasi["Year"]=="2026")
    ]
    l, m = st.columns(2)

    with l : 
        jenisrecom = sorted(filteredrecom["Keterangan"].dropna().unique())
        recom_filter = st.multiselect(
            "Commissioning / Recommissioning",
            jenisrecom,default=jenisrecom,
            max_selections=None
        ) 

    filtered_recom = filteredrecom[filteredrecom["Keterangan"].isin(recom_filter)]
    filtered_recom2026 = df_recom[df_recom["Keterangan"].isin(recom_filter)]
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
                <h4>Total Unit</h4>
                <div class="value">{len(filtered_recom)}</div>
            </div>
        """, unsafe_allow_html=True)
    with x2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Perusahaan</h4>
                <div class="value">{filtered_recom["Perusahaan"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x3:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Jenis Unit</h4>
                <div class="value">{filtered_recom["Jenis Unit"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x4:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Departemen</h4>
                <div class="value">{filtered_recom["Departemen"].nunique()}</div>
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
    
    #LULUS UJI
    with st.container(border=True):
        st.markdown(
        f"""
        <style>
            .subhead-grafik h2 {{
                margin: 0;
                font-size: 30px;
                font-weight: 600;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="subhead-grafik">
            <h2>LULUS UJI KENDARAAN</h2>
        </div>
        """,
        unsafe_allow_html=True
        )

        col1, div, col2 = st.columns([3,0.1,3])
        with col1 :
            df_lulus=filtered_recom[filtered_recom["LULUS/TIDAK"] == "LULUS UJI KENDARAAN"]
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
                    <h2>Data {judul_week}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

            pivot_lulus = df_lulus.pivot_table(
                index="Departemen",
                columns="BU",
                values="No. Unit",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bu_lulus = ["DCM", "HPAL", "ONC"]
            available_bu_lulus = [bu for bu in possible_bu_lulus if bu in pivot_lulus.columns]
            bu_counts_lulus = (
                df_lulus["BU"]
                .value_counts()
                .reindex(possible_bu_lulus, fill_value=0)
                .reset_index()
            )
            bu_counts_lulus.columns = ["BU", "Total"]
            fig1 = px.bar(
                bu_counts_lulus,
                x="BU",
                y="Total",
                text=bu_counts_lulus["Total"],
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a"
                },
                template="seaborn"
            )
            fig1.update_traces(
                textposition="outside",
                textfont_size=14
            )
            fig1.update_layout(
                xaxis_title="Business Unit",
                yaxis_title="Jumlah",
                showlegend=False,
                bargap=0.3,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420
            )

            st.plotly_chart(fig1, use_container_width=True, key="bar-lulus")

            with st.expander("Data Com/Re-Com Weekly"):


                # --- Kolom BU yang mungkin ---
                possible_bu_lulus = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                possible_bu_lulus = [bu for bu in possible_bu_lulus if bu in pivot_lulus.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bu_lulus) == 0:
                    pivot_lulus["Total"] = 0
                else:
                    pivot_lulus["Total"] = pivot_lulus[available_bu_lulus].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_lulus = pivot_lulus.sort_values("Total", ascending=False, ignore_index=True)
                pivot_lulus.index = pivot_lulus.index + 1
                st.dataframe(pivot_lulus, height=400)
        
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
            df_lulus2026 = filtered_recom2026[filtered_recom2026["LULUS/TIDAK"]=="LULUS UJI KENDARAAN"] 
            pivot_lulus_2026 =df_lulus2026.pivot_table(
                index="Departemen",
                columns="BU",
                values="No. Unit",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bu_lulus_2026 = ["DCM", "HPAL", "ONC"]
            available_bu_lulus_2026 = [bu for bu in possible_bu_lulus_2026 if bu in pivot_lulus_2026.columns]
            bu_totals = {bu: pivot_lulus_2026[bu].sum() for bu in available_bu_lulus_2026}
            sizes  = list(bu_totals.values())
            bu_counts2026 = (
                df_lulus2026["BU"]
                .value_counts()
                .reindex(possible_bu_lulus_2026, fill_value=0)
                .reset_index()
            )
            bu_counts2026.columns = ["BU", "Total"]
            fig2 = px.pie(
                bu_counts2026,
                names="BU",
                values="Total",
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a"
                },
                hole=0.4)
            fig2.update_traces(
                textinfo="label+value",
                textfont_size=14
            )
            
            total_all = sum(sizes)
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
            st.plotly_chart(fig2, use_container_width=True, key="pie_2026")
            with st.expander("Data Com/Re-Com 2026"):


                # --- Kolom BU yang mungkin ---
                possible_bu_lulus_2026 = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bu_lulus_2026 = [bu for bu in possible_bu_lulus_2026 if bu in pivot_lulus_2026.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bu_lulus_2026) == 0:
                    pivot_lulus_2026["Total"] = 0
                else:
                    pivot_lulus_2026["Total"] = pivot_lulus_2026[available_bu_lulus_2026].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_lulus_2026 = pivot_lulus_2026.sort_values("Total", ascending=False, ignore_index=True)
                pivot_lulus_2026.index = pivot_lulus_2026.index + 1
                st.dataframe(pivot_lulus_2026, height=400)

    #BELUM LULUS
    with st.container(border=True):
        st.markdown(
        f"""
        <style>
            .subhead-grafik h2 {{
                margin: 0;
                font-size: 30px;
                font-weight: 600;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="subhead-grafik">
            <h2>BELUM LULUS UJI KENDARAAN</h2>
        </div>
        """,
        unsafe_allow_html=True
        )

        col3, div, col4 = st.columns([3,0.1,3])
        with col3 :
            df_belumlulus=filtered_recom[filtered_recom["LULUS/TIDAK"] == "BELUM LULUS UJI KENDARAAN"]
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
                    <h2>Data {judul_week}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )

            pivot_bl = df_belumlulus.pivot_table(
                index="Departemen",
                columns="BU",
                values="No. Unit",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bu_bl = ["DCM", "HPAL", "ONC"]
            available_bu_bl = [bu for bu in possible_bu_bl if bu in pivot_bl.columns]
            bu_counts_bl = (
                df_belumlulus["BU"]
                .value_counts()
                .reindex(possible_bu_bl, fill_value=0)
                .reset_index()
            )
            bu_counts_bl.columns = ["BU", "Total"]
            fig3 = px.bar(
                bu_counts_bl,
                x="BU",
                y="Total",
                text=bu_counts_bl["Total"],
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a"
                },
                template="seaborn"
            )
            fig3.update_traces(
                textposition="outside",
                textfont_size=14
            )
            fig3.update_layout(
                xaxis_title="Business Unit",
                yaxis_title="Jumlah",
                showlegend=False,
                bargap=0.3,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420
            )

            st.plotly_chart(fig3, use_container_width=True, key="bar-belum_lulus")

            with st.expander("Data Com/Re-Com Weekly"):


                # --- Kolom BU yang mungkin ---
                possible_bu_bl = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bu_bl = [bu for bu in possible_bu_bl if bu in pivot_bl.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bu_bl) == 0:
                    pivot_bl["Total"] = 0
                else:
                    pivot_bl["Total"] = pivot_bl[available_bu_bl].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_bl = pivot_bl.sort_values("Total", ascending=False, ignore_index=True)
                pivot_bl.index = pivot_bl.index + 1
                st.dataframe(pivot_bl, height=400)
        
        with col4:
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
            df_bl_2026 = filtered_recom2026[filtered_recom2026["LULUS/TIDAK"]=="BELUM LULUS UJI KENDARAAN"] 
            pivot_bl_2026 =df_bl_2026.pivot_table(
                index="Departemen",
                columns="BU",
                values="No. Unit",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bu_bl_2026 = ["DCM", "HPAL", "ONC"]
            available_bu_bl_2026 = [bu for bu in possible_bu_bl_2026 if bu in pivot_bl_2026.columns]
            bu_totals_bl = {bu: pivot_bl_2026[bu].sum() for bu in available_bu_bl_2026}
            sizes_bl  = list(bu_totals_bl.values())
            bu_counts_bl2026 = (
                df_bl_2026["BU"]
                .value_counts()
                .reindex(possible_bu_bl_2026, fill_value=0)
                .reset_index()
            )
            bu_counts_bl2026.columns = ["BU", "Total"]
            fig4 = px.pie(
                bu_counts_bl2026,
                names="BU",
                values="Total",
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a"
                },
                hole=0.4)
            fig4.update_traces(
                textinfo="label+value",
                textfont_size=14
            )
            total_all_bl = sum(sizes_bl)
            fig4.update_layout(
                showlegend=False,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420,
                margin=dict(t=20, b=20, l=20, r=20),
                annotations=[
                    dict(
                        text=f"<b>{total_all_bl}</b>",
                        x=0.5,
                        y=0.5,
                        font=dict(size=26, color="black"),
                        showarrow=False
                    )
                ]
            )
            st.plotly_chart(fig4, use_container_width=True, key="piebl2026")
            with st.expander("Data Com/Re-Com 2026"):


                # --- Kolom BU yang mungkin ---
                possible_bu_bl_2026 = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bu_lulus_2026 = [bu for bu in possible_bu_bl_2026 if bu in pivot_bl_2026.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bu_bl_2026) == 0:
                    pivot_bl_2026["Total"] = 0
                else:
                    pivot_bl_2026["Total"] = pivot_bl_2026[available_bu_bl_2026].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_bl_2026 = pivot_bl_2026.sort_values("Total", ascending=False, ignore_index=True)
                pivot_bl_2026.index = pivot_bl_2026.index + 1
                st.dataframe(pivot_bl_2026, height=400)

    #temuan
    with st.container(border=True):
        st.markdown(
        f"""
        <style>
            .subhead-grafik h2 {{
                margin: 0;
                font-size: 30px;
                font-weight: 600;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="subhead-grafik">
            <h2>TOP 10 Temuan Commissioning & Recommissioning</h2>
        </div>
        """,
        unsafe_allow_html=True
        )
        top_temuan = (
                df_temuan
                .groupby("Klasifikasi Temuan", as_index=False)
                .size()
                .rename(columns={"size": "Jumlah"})
                .sort_values("Jumlah", ascending=False)
                .head(10)  
            )
        fig_temuan= px.bar(
                top_temuan,
                x="Jumlah",
                y="Klasifikasi Temuan",
                orientation="h",
                text="Jumlah",
            )

        fig_temuan.update_traces(textposition="outside")
        fig_temuan.update_layout(
            yaxis=dict(
                autorange="reversed",
                tickfont=dict(
                    color="black",
                    size=16
                )
            ),
            plot_bgcolor="white",
            paper_bgcolor="white",
        )


        st.plotly_chart(fig_temuan, use_container_width=True,key="bar_temuan")

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
    import base64
    import requests

    cols_per_row = 3
    rows = recom_dokumentasi.to_dict("records")

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
                            st.warning("Link bukan file gambar.")
                            st.write(url)

                except Exception as e:
                    with col:
                        st.error(f"Error load: {e}")




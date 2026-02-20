
def app():
    import streamlit as st
    import pandas as pd
    import gspread
    from google.oauth2.service_account import Credentials
    from datetime import date
    import requests
    import matplotlib.pyplot as plt
    import plotly_express as px

    q,w =st.columns([4,1])
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
                <h1>Inspeksi & Observasi</h1>
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
    #df observasi
    spreadsheet_observasi = "OBSERVASI ALL BU 2026"
    sheet_obs = client.open(spreadsheet_observasi)
    ws_observasi = sheet_obs.worksheet("Observasi")
    values_observasi = ws_observasi.get("E:O")
    df_observasi = pd.DataFrame(values_observasi[1:], columns=values_observasi[0])

    #df_nilai
    spreadsheet_nilai = "OBSERVASI DCM 2026"
    sheet_nilai=client.open(spreadsheet_nilai)
    ws_nilai = sheet_nilai.worksheet("2026")
    values_nilai = ws_nilai.get("A:M")
    df_nilai = pd.DataFrame(values_nilai[1:],columns=values_nilai[0])
    nilai = df_nilai[["Week","Departemen", "Scores"]].copy()
    nilai["Scores"] = pd.to_numeric(nilai["Scores"], errors="coerce")


    #df observasi
    spreadsheet_inspeksi = "INSPEKSI ALL BU 2026"
    sheet_inspeksi = client.open(spreadsheet_inspeksi)
    ws_inspeksi = sheet_inspeksi.worksheet("2026")
    values_inspeksi = ws_inspeksi.get("E:O")
    df_inspeksi = pd.DataFrame(values_inspeksi[1:], columns=values_inspeksi[0])
    df_inspeksi = df_inspeksi[
        (df_inspeksi["Departement"].notna()) &
        (df_inspeksi["Departement"].str.strip() != "")
    ]

    #temuan
    ws_temuan = sheet_inspeksi.worksheet("Olah Temuan")
    values_temuan = ws_temuan.get("A:G")
    df_temuan = pd.DataFrame(values_temuan[1:],columns=values_temuan[0])

    #dokumentasi
    naemSheetDok = "DOKUMENTASI"
    sheet_dok = client.open(naemSheetDok)
    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("D:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    # ================= WEEK FILTER ==================
    weeks = sorted(df_observasi["Week"].dropna().unique())

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

            week_sorted = sorted(
                week_filter,
                key=lambda x: int(x.replace("Week", "").strip())
            )

            return f"{week_sorted[0]} – {week_sorted[-1]}"

    filtered_observasi = df_observasi[df_observasi["Week"].isin(week_filter)]
    filtered_nilai = nilai[nilai["Week"].isin(week_filter)]
    filtered_inspeksi = df_inspeksi[
        (df_inspeksi["Week"].isin(week_filter)) &
        (df_inspeksi["Departement"].notna()) &
        (df_inspeksi["Departement"].str.strip() != "")
    ]
    filtered_temuan = df_temuan[df_temuan["Week"].isin(week_filter)]

    observasi_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Inspeksi & Observasi") &
        (dokumentasi["Year"]=="2026")
    ]

    x1, x2, x3, x4 = st.columns(4)
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
            .st-emotion-cache-ncvwpv {
                background-color : white; box-shadow: 0 3px 8px rgba(0,0,0,0.08);}
            .st-emotion-cache-165te23 {
                background-color : white;box-shadow: 0 3px 8px rgba(0,0,0,0.08);}
            </style>
            """, unsafe_allow_html=True)
    with x1:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Inspeksi & Observasi</h4>
                <div class="value">{len(filtered_observasi)}</div>
            </div>
        """, unsafe_allow_html=True)
    with x2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Perusahaan</h4>
                <div class="value">{filtered_observasi["Perusahaan"].nunique()}</div>
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
        st.markdown(
        f"""
        <style>
            .subheader h2 {{
                padding-bottom: 30px;
                margin-top: -40px;
                font-size: 35px;
                font-weight: 500;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="subheader">
            <h2>Inspeksi</h2>
        </div>
        """,
        unsafe_allow_html=True)
        a, div, b= st.columns([3,0.1,3])
        pivot_inspeksi = filtered_inspeksi.pivot_table(
            index="Departement",
            columns="BU",
            values="Nomor Lambung",
            aggfunc="count",
            fill_value=0
        ).reset_index()
        with a :
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
            # --- Kolom BU yang mungkin ---
            possible_bus = ["DCM", "HPAL", "ONC"]

            # --- Kolom BU yang benar-benar ada di pivot ---
            available_bus_inspeksi = [bu for bu in possible_bus if bu in pivot_inspeksi.columns]
            bu_counts_inspeksi = (
                filtered_inspeksi["BU"]
                .value_counts()
                .reindex(possible_bus, fill_value=0)
                .reset_index()
            )
            bu_counts_inspeksi.columns = ["BU", "Total"]
            fig_inspeksi = px.bar(
                bu_counts_inspeksi,
                x="BU",
                y="Total",
                text=bu_counts_inspeksi["Total"],
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a",
                    "Lainnya":"#ff9900"
                },
                template="seaborn"
            )
            fig_inspeksi.update_traces(
                textposition="outside",
                textfont_size=14
            )
            fig_inspeksi.update_layout(
                xaxis_title="Business Unit",
                yaxis_title="Jumlah",
                showlegend=False,
                bargap=0.3,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420
            )

            st.plotly_chart(fig_inspeksi, use_container_width=True)

            with st.expander("Data Inspeksi Weekly"):


                # --- Kolom BU yang mungkin ---
                possible_bus = ["DCM", "HPAL", "ONC","Lainnya"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bus_inspeksi = [bu for bu in possible_bus if bu in pivot_inspeksi.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bus_inspeksi) == 0:
                    pivot_inspeksi["Total"] = 0
                else:
                    pivot_inspeksi["Total"] = pivot_inspeksi[available_bus_inspeksi].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_inspeksi = pivot_inspeksi.sort_values("Total", ascending=False, ignore_index=True)
                pivot_inspeksi.index = pivot_inspeksi.index + 1
                st.dataframe(pivot_inspeksi, height=400)
        with b:
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
            pivot2026_inspeksi = df_inspeksi.pivot_table(
                index="Departement",
                columns="BU",
                values="Nomor Lambung",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]
            available_bus_inspeksi2026 = [bu for bu in possible_bus if bu in pivot2026_inspeksi.columns]
            if len(available_bus_inspeksi2026) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals_inspeksi2026 = {bu: pivot2026_inspeksi[bu].sum() for bu in available_bus_inspeksi2026}

                labels_inspeksi2026 = list(bu_totals_inspeksi2026.keys())
                sizess_inspeksi2026  = list(bu_totals_inspeksi2026.values())

                if sum(sizess_inspeksi2026) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    chart_inspeksi = pd.DataFrame({
                        "BU": labels_inspeksi2026,
                        "Total": sizess_inspeksi2026
                    })

                    fig2026 = px.pie(
                        chart_inspeksi,
                        names="BU",
                        values="Total",
                        color="BU",
                        color_discrete_map={
                            "DCM": "#134f5c",
                            "HPAL": "#2f9a7f",
                            "ONC": "#31681a",
                            "Lainnya": "#ff9900"
                        },
                        hole=0.4
                    )

                    fig2026.update_traces(
                        textinfo="label+value",
                        textfont_size=14
                    )
                    total_all_inspeksi = sum(sizess_inspeksi2026)
                    fig2026.update_layout(
                        height=420,
                        showlegend=False,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(t=20, b=20, l=20, r=20),
                        annotations=[
                            dict(
                                text=f"<b>{total_all_inspeksi}</b>",
                                x=0.5,
                                y=0.5,
                                font=dict(size=26, color="black"),
                                showarrow=False
                            )
                        ]
                    )

                    st.plotly_chart(
                        fig2026,
                        use_container_width=True,
                        key="pie_inspeksi_2026"
                    )

            with st.expander("Data Inspeksi 2026"):
                pivot2026_tabel_inspeksi = df_inspeksi.pivot_table(
                    index="Departement",
                    columns="BU",
                    values="Nomor Lambung",
                    aggfunc="count",
                    fill_value=0
                ).reset_index()

                # --- Kolom BU yang mungkin ---
                possible_bus = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bus_inspeksi2026 = [bu for bu in possible_bus if bu in pivot2026_tabel_inspeksi.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bus_inspeksi2026) == 0:
                    pivot2026_tabel_inspeksi["Total"] = 0
                else:
                    pivot2026_tabel_inspeksi["Total"] = pivot2026_tabel_inspeksi[available_bus_inspeksi2026].sum(axis=1)

                # --- Urutkan pivot ---
                pivot2026_tabel_inspeksi = pivot2026_tabel_inspeksi.sort_values("Total", ascending=False, ignore_index=True)
                pivot2026_tabel_inspeksi.index = pivot2026_tabel_inspeksi.index + 1
                st.dataframe(pivot2026_tabel_inspeksi, height=400)
        with st.expander("TOP 10 Temuan Inspeksi"):
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
                yaxis=dict(autorange="reversed")
            )

            st.plotly_chart(fig_temuan, use_container_width=True,key="bar_temuan")


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
        st.markdown(
        f"""
        <style>
            .subheader h2 {{
                margin-top: 0px;
                font-size: 35px;
                font-weight: 500;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="subheader">
            <h2>Observasi</h2>
        </div>
        """,
        unsafe_allow_html=True)
        col1, div, col3= st.columns([3,0.1,3])
        pivot = filtered_observasi.pivot_table(
            index="Perusahaan",
            columns="BU",
            values="Driver/Operator",
            aggfunc="count",
            fill_value=0
        ).reset_index()
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
                    <h2>Data {judul_week}</h2>
                </div>
                """,
                unsafe_allow_html=True
            )
            pivot_tabel = filtered_observasi.pivot_table(
                index="Departemen",
                columns="BU",
                values="Driver/Operator",
                aggfunc="count",
                fill_value=0
            ).reset_index()

            # --- Kolom BU yang mungkin ---
            possible_bus = ["DCM", "HPAL", "ONC"]

            # --- Kolom BU yang benar-benar ada di pivot ---
            available_bus = [bu for bu in possible_bus if bu in pivot.columns]
            bu_counts = (
                filtered_observasi["BU"]
                .value_counts()
                .reindex(possible_bus, fill_value=0)
                .reset_index()
            )
            bu_counts.columns = ["BU", "Total"]
            fig1 = px.bar(
                bu_counts,
                x="BU",
                y="Total",
                text=bu_counts["Total"],
                color="BU",
                color_discrete_map={
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a",
                    "Lainnya":"#ff9900"
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

            st.plotly_chart(fig1, use_container_width=True,key="observasi")

            with st.expander("Data Observasi Weekly"):


                # --- Kolom BU yang mungkin ---
                possible_bus = ["DCM", "HPAL", "ONC","Lainnya"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bus = [bu for bu in possible_bus if bu in pivot_tabel.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bus) == 0:
                    pivot_tabel["Total"] = 0
                else:
                    pivot_tabel["Total"] = pivot_tabel[available_bus].sum(axis=1)

                # --- Urutkan pivot ---
                pivot_tabel = pivot_tabel.sort_values("Total", ascending=False, ignore_index=True)
                pivot_tabel.index = pivot_tabel.index + 1
                st.dataframe(pivot_tabel, height=400)
        with col3:
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
            pivot2026 = df_observasi.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Driver/Operator",
                aggfunc="count",
                fill_value=0
            ).reset_index()
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]
            available_bus = [bu for bu in possible_bus if bu in pivot2026.columns]
            if len(available_bus) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals = {bu: pivot2026[bu].sum() for bu in available_bus}

                labels = list(bu_totals.keys())
                sizes  = list(bu_totals.values())

                if sum(sizes) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    chart_full = pd.DataFrame({
                        "BU": labels,
                        "Total": sizes
                    })

                    fig = px.pie(
                        chart_full,
                        names="BU",
                        values="Total",
                        color="BU",
                        color_discrete_map={
                            "DCM": "#134f5c",
                            "HPAL": "#2f9a7f",
                            "ONC": "#31681a",
                            "Lainnya": "#ff9900"
                        },
                        hole=0.4
                    )

                    fig.update_traces(
                        textinfo="label+value",
                        textfont_size=14
                    )
                    total_all = sum(sizes)
                    fig.update_layout(
                        height=420,
                        showlegend=False,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
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

                    st.plotly_chart(
                        fig,
                        use_container_width=True,
                        key="pie_2026"
                    )

            with st.expander("Data Observasi 2026"):
                pivot2026_tabel = df_observasi.pivot_table(
                    index="Departemen",
                    columns="BU",
                    values="Driver/Operator",
                    aggfunc="count",
                    fill_value=0
                ).reset_index()

                # --- Kolom BU yang mungkin ---
                possible_bus = ["DCM", "HPAL", "ONC"]

                # --- Kolom BU yang benar-benar ada di pivot ---
                available_bus = [bu for bu in possible_bus if bu in pivot2026_tabel.columns]

                # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
                if len(available_bus) == 0:
                    pivot2026_tabel["Total"] = 0
                else:
                    pivot2026_tabel["Total"] = pivot2026_tabel[available_bus].sum(axis=1)

                # --- Urutkan pivot ---
                pivot2026_tabel = pivot2026_tabel.sort_values("Total", ascending=False, ignore_index=True)
                pivot2026_tabel.index = pivot2026_tabel.index + 1
                st.dataframe(pivot2026_tabel, height=400)
        with st.expander("Rata-rata Nilai Observasi"):
            avg_scores = (
                filtered_nilai
                .groupby("Departemen", as_index=False)["Scores"]
                .mean()
                .sort_values("Scores", ascending=False)
            )
            avg_scores["Scores"] = avg_scores["Scores"].round(2)
            fig_nilai = px.bar(
                avg_scores,
                x="Departemen",
                y="Scores",
                title="Average Scores per Departemen",
                text="Scores", color="Departemen"
            )

            fig_nilai.update_traces(
                textposition="outside", textfont_size=14
            )

            fig_nilai.update_layout(
                xaxis_title="Departemen",
                yaxis_title="Average Score",
                yaxis=dict(range=[0, 4]),
                bargap=0.3,
                plot_bgcolor="white",
                paper_bgcolor="white",
                height=420
            )

            st.plotly_chart(fig_nilai, use_container_width=True, key="bar_nilai")
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
    rows = observasi_dokumentasi.to_dict("records")

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

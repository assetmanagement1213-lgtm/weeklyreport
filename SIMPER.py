
from pandas import DataFrame


def app():
    import streamlit as st
    import pandas as pd
    import gspread
    from google.oauth2.service_account import Credentials
    from datetime import date
    import requests
    import matplotlib.pyplot as plt
    import plotly.express as px

    q,w =st.columns([4,1])
    with q:
        st.markdown(
        f"""
        <style>
            .header h2 {{
                margin: 0;
                font-size: 45px;
                font-weight: 700;
                color: #1f2937;
            }}
        </style>
        <div class="header">
            <h2>SIMPER</h2>
        </div>
        """,
        unsafe_allow_html=True)
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )

    client = gspread.authorize(creds)

    spreadsheet_name = "7. DATA SIMPER ALL BU 2026"
    sheet = client.open(spreadsheet_name)

    naemSheetDok = "DOKUMENTASI"
    sheet_dok = client.open(naemSheetDok)
    ws_simper = sheet.worksheet("2026")
    values_simper = ws_simper.get("A:J")
    df_simper = pd.DataFrame(values_simper[1:], columns=values_simper[0])

    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("D:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    weeks = sorted(df_simper["Week 2026"].dropna().unique())

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

        week_sorted = sorted(
            week_filter,
            key=lambda x: int(x.replace("Week", "").strip())
        )

    filtered_simper = df_simper[df_simper["Week 2026"].isin(week_filter)]
    simper_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "simper") &(dokumentasi["Year"]=="2026")
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
            .st-emotion-cache-ncvwpv {
                background-color : white; box-shadow: 0 3px 8px rgba(0,0,0,0.08);}
            .st-emotion-cache-165te23 {
                background-color : white;box-shadow: 0 3px 8px rgba(0,0,0,0.08);}
            </style>
            """, unsafe_allow_html=True)
    with x1:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Simper</h4>
                <div class="value">{len(filtered_simper)}</div>
            </div>
        """, unsafe_allow_html=True)
    with x2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Perusahaan</h4>
                <div class="value">{filtered_simper["Perusahaan"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x3:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Departemen</h4>
                <div class="value">{filtered_simper["Departemen"].nunique()}</div>
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
            .grafik h2 {{
                padding-bottom: 30px;
                margin-top: -40px;
                font-size: 35px;
                font-weight: 500;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="grafik">
            <h2>Data {judul_week}</h2>
        </div>
        """,
        unsafe_allow_html=True)
    
    col1, col2, col3= st.columns(3)
    with col1: 
        with st.container(border=True):
            header1=st.container()
            chart1=st.container(height=420, border=False)
            dataframe1=st.container()
            full=filtered_simper[filtered_simper["Status SIMPER"] == "F"]
            pivot = full.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
        with header1:
            st.markdown("<h3 style='text-align: center;'>SIMPER FULL</h3>", unsafe_allow_html=True)
        with chart1:
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

            available_bus = [bu for bu in possible_bus if bu in pivot.columns]

            if len(available_bus) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals = {bu: pivot[bu].sum() for bu in available_bus}

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
                        key="pie_simper_full"
                    )
                                
        with dataframe1:
            possible_bus = ["DCM", "HPAL", "ONC"]
            available_bus = [bu for bu in possible_bus if bu in pivot.columns]

            if len(available_bus) == 0:
                pivot["Total"] = 0
            else:
                pivot["Total"] = pivot[available_bus].sum(axis=1)

            pivot = pivot.sort_values("Total", ascending=False, ignore_index=True)
            pivot.index = pivot.index + 1
            st.dataframe(pivot, height=300)
    
    with col2:
        with st.container(border=True):
            header2=st.container()
            chart2=st.container(height=420, border=False)
            dataframe2=st.container()
            probation=filtered_simper[filtered_simper["Status SIMPER"] == "P"]
            pivot_prob = probation.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
        with header2:
            st.markdown("<h3 style='text-align: center;'>SIMPER PROBATION</h3>", unsafe_allow_html=True)
        with chart2:
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

            available_bus_prob = [bu for bu in possible_bus if bu in pivot_prob.columns]

            if len(available_bus_prob) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals_prob = {bu: pivot_prob[bu].sum() for bu in available_bus_prob}

                labels_prob = list(bu_totals_prob.keys())
                sizes_prob  = list(bu_totals_prob.values())

                if sum(sizes_prob) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    chart_prob = pd.DataFrame({
                        "BU": labels_prob,
                        "Total": sizes_prob
                    })

                    fig2 = px.pie(
                        chart_prob,
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

                    fig2.update_traces(
                        textinfo="label+value",
                        textfont_size=14
                    )
                    total_all_prob = sum(sizes_prob)
                    fig2.update_layout(
                        height=420,
                        showlegend=False,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(t=20, b=20, l=20, r=20),
                        annotations=[
                            dict(
                                text=f"<b>{total_all_prob}</b>",
                                x=0.5,
                                y=0.5,
                                font=dict(size=26, color="black"),
                                showarrow=False
                            )
                        ]
                    )

                    st.plotly_chart(
                        fig2,
                        use_container_width=True,
                        key="pie_simper_prob"
                    )
        with dataframe2:
            possible_bus = ["DCM", "HPAL", "ONC"]
            available_bus_prob = [bu for bu in possible_bus if bu in pivot_prob.columns]

            if len(available_bus) == 0:
                pivot_prob["Total"] = 0
            else:
                pivot_prob["Total"] = pivot_prob[available_bus_prob].sum(axis=1)

            pivotprob = pivot_prob.sort_values("Total", ascending=False, ignore_index=True)
            pivotprob.index = pivotprob.index + 1
            st.dataframe(pivotprob, height=300)
    
    with col3:
        with st.container(border=True):
            header3=st.container()
            chart3=st.container(height=420, border=False)
            dataframe3=st.container()
            sementara=filtered_simper[filtered_simper["Status SIMPER"] == "T"]
            pivot_sem = sementara.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with header3:
                st.markdown("<h3 style='text-align: center;'>SIMPER SEMENTARA</h3>", unsafe_allow_html=True)
            with chart3:
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bus_sem = [bu for bu in possible_bus if bu in pivot_sem.columns]

                if len(available_bus) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_totals_sem= {bu: pivot_sem[bu].sum() for bu in available_bus_sem}

                    labels_sem = list(bu_totals_sem.keys())
                    sizes_sem  = list(bu_totals_sem.values())

                    if sum(sizes) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_sem = pd.DataFrame({
                            "BU": labels_sem,
                            "Total": sizes_sem
                        })

                        fig3 = px.pie(
                            chart_sem,
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

                        fig3.update_traces(
                            textinfo="label+value",
                            textfont_size=14
                        )
                        total_all_sem = sum(sizes_sem)
                        fig3.update_layout(
                            height=420,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(t=20, b=20, l=20, r=20),
                            annotations=[
                                dict(
                                    text=f"<b>{total_all_sem}</b>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=26, color="black"),
                                    showarrow=False
                                )
                            ]
                        )

                        st.plotly_chart(
                            fig3,
                            use_container_width=True,
                            key="pie_simper_sem"
                        )
            with dataframe3:
                possible_bus = ["DCM", "HPAL", "ONC"]
                available_bus_sem= [bu for bu in possible_bus if bu in pivot_sem.columns]

                if len(available_bus) == 0:
                    pivot_sem["Total"] = 0
                else:
                    pivot_sem["Total"] = pivot_sem[available_bus_sem].sum(axis=1)

                pivotsem = pivot_sem.sort_values("Total", ascending=False, ignore_index=True)
                pivotsem.index = pivotsem.index + 1
                st.dataframe(pivotsem, height=300)

    st.divider()

    st.markdown(
        f"""
        <style>
            .grafik h2 {{
                padding-bottom: 30px;
                margin-top: -40px;
                font-size: 35px;
                font-weight: 500;
                color: black;
                text-align: center;
            }}
        </style>
        <div class="grafik">
            <h2>Data 2026</h2>
        </div>
        """,
        unsafe_allow_html=True)
    
    a, b, c= st.columns(3)
    with a: 
        with st.container(border=True):
            header_a=st.container()
            chart_a=st.container(height=420, border=False)
            dataframe_a=st.container()
            full_a=df_simper[df_simper["Status SIMPER"] == "F"]
            pivot_a = full_a.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
        with header_a:
            st.markdown("<h3 style='text-align: center;'>SIMPER FULL</h3>", unsafe_allow_html=True)
        with chart_a:
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

            available_bus_a = [bu for bu in possible_bus if bu in pivot_a.columns]

            if len(available_bus_a) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals_a = {bu: pivot_a[bu].sum() for bu in available_bus_a}

                labels_a = list(bu_totals_a.keys())
                sizes_a = list(bu_totals_a.values())

                if sum(sizes_a) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    chart_full_a = pd.DataFrame({
                        "BU": labels_a,
                        "Total": sizes_a
                    })

                    fig_a = px.pie(
                        chart_full_a,
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

                    fig_a.update_traces(
                        textinfo="label+value",
                        textfont_size=14
                    )
                    total_all_a = sum(sizes_a)
                    fig_a.update_layout(
                        height=420,
                        showlegend=False,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(t=20, b=20, l=20, r=20),
                        annotations=[
                            dict(
                                text=f"<b>{total_all_a}</b>",
                                x=0.5,
                                y=0.5,
                                font=dict(size=26, color="black"),
                                showarrow=False
                            )
                        ]
                    )

                    st.plotly_chart(
                        fig_a,
                        use_container_width=True,
                        key="pie_simper_full_2026"
                    )
                                
        with dataframe_a:
            possible_bus = ["DCM", "HPAL", "ONC"]
            available_bus_a = [bu for bu in possible_bus if bu in pivot_a.columns]

            if len(available_bus_a) == 0:
                pivot_a["Total"] = 0
            else:
                pivot_a["Total"] = pivot_a[available_bus_a].sum(axis=1)

            pivot_a = pivot_a.sort_values("Total", ascending=False, ignore_index=True)
            pivot_a.index = pivot_a.index + 1
            st.dataframe(pivot_a, height=300)
    
    with b:
        with st.container(border=True):
            header_b =st.container()
            chart_b=st.container(height=420, border=False)
            dataframe_b=st.container()
            probation_b=df_simper[df_simper["Status SIMPER"] == "P"]
            pivot_prob_b = probation_b.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
        with header_b:
            st.markdown("<h3 style='text-align: center;'>SIMPER PROBATION</h3>", unsafe_allow_html=True)
        with chart_b:
            possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

            available_bus_prob_b = [bu for bu in possible_bus if bu in pivot_prob_b.columns]

            if len(available_bus_prob_b) == 0:
                st.info("Tidak ada data untuk ditampilkan.")
            else:
                bu_totals_prob_b = {bu: pivot_prob_b[bu].sum() for bu in available_bus_prob_b}

                labels_prob_b = list(bu_totals_prob_b.keys())
                sizes_prob_b  = list(bu_totals_prob_b.values())

                if sum(sizes_prob_b) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    chart_prob_b = pd.DataFrame({
                        "BU": labels_prob_b,
                        "Total": sizes_prob_b
                    })

                    fig_b = px.pie(
                        chart_prob_b,
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

                    fig_b.update_traces(
                        textinfo="label+value",
                        textfont_size=14
                    )
                    total_all_prob_b = sum(sizes_prob_b)
                    fig_b.update_layout(
                        height=420,
                        showlegend=False,
                        plot_bgcolor="white",
                        paper_bgcolor="white",
                        margin=dict(t=20, b=20, l=20, r=20),
                        annotations=[
                            dict(
                                text=f"<b>{total_all_prob_b}</b>",
                                x=0.5,
                                y=0.5,
                                font=dict(size=26, color="black"),
                                showarrow=False
                            )
                        ]
                    )

                    st.plotly_chart(
                        fig_b,
                        use_container_width=True,
                        key="pie_simper_prob_b"
                    )
        with dataframe_b:
            possible_bus = ["DCM", "HPAL", "ONC"]
            available_bus_prob_b = [bu for bu in possible_bus if bu in pivot_prob_b.columns]

            if len(available_bus_prob_b) == 0:
                pivot_prob_b["Total"] = 0
            else:
                pivot_prob_b["Total"] = pivot_prob_b[available_bus_prob_b].sum(axis=1)

            pivotprob_b = pivot_prob_b.sort_values("Total", ascending=False, ignore_index=True)
            pivotprob_b.index = pivotprob_b.index + 1
            st.dataframe(pivotprob_b, height=300)
    
    with c:
        with st.container(border=True):
            header_c=st.container()
            chart_c=st.container(height=420, border=False)
            dataframe_c=st.container()
            sementara_c=df_simper[df_simper["Status SIMPER"] == "T"]
            pivot_sem_c = sementara_c.pivot_table(
                index="Perusahaan",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with header_c:
                st.markdown("<h3 style='text-align: center;'>SIMPER SEMENTARA</h3>", unsafe_allow_html=True)
            with chart_c:
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bus_sem_c = [bu for bu in possible_bus if bu in pivot_sem_c.columns]

                if len(available_bus_sem_c) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_totals_sem_c= {bu: pivot_sem_c[bu].sum() for bu in available_bus_sem_c}

                    labels_sem_c = list(bu_totals_sem_c.keys())
                    sizes_sem_c  = list(bu_totals_sem_c.values())

                    if sum(sizes_sem_c) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_sem_c = pd.DataFrame({
                            "BU": labels_sem_c,
                            "Total": sizes_sem_c
                        })

                        fig_c = px.pie(
                            chart_sem_c,
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

                        fig_c.update_traces(
                            textinfo="label+value",
                            textfont_size=14
                        )
                        total_all_sem_c = sum(sizes_sem_c)
                        fig_c.update_layout(
                            height=420,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(t=20, b=20, l=20, r=20),
                            annotations=[
                                dict(
                                    text=f"<b>{total_all_sem_c}</b>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=26, color="black"),
                                    showarrow=False
                                )
                            ]
                        )

                        st.plotly_chart(
                            fig_c,
                            use_container_width=True,
                            key="pie_simper_sem_c"
                        )
            with dataframe_c:
                possible_bus = ["DCM", "HPAL", "ONC"]
                available_bus_sem_c= [bu for bu in possible_bus if bu in pivot_sem_c.columns]

                if len(available_bus_sem_c) == 0:
                    pivot_sem_c["Total"] = 0
                else:
                    pivot_sem_c["Total"] = pivot_sem_c[available_bus_sem_c].sum(axis=1)

                pivot_sem_c = pivot_sem_c.sort_values("Total", ascending=False, ignore_index=True)
                pivot_sem_c.index = pivot_sem_c.index + 1
                st.dataframe(pivot_sem_c, height=300)
    st.divider()




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
            <h2>Refresh, Pembekalan, dan Coaching</h2>
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

    #data violation
    sp_vio = "REFRESH VIOLATION 2026"
    sheet_vio = client.open(sp_vio)
    ws_vio = sheet_vio.worksheet("2026")
    values_vio = ws_vio.get("A:M")
    violation = pd.DataFrame(values_vio[1:], columns=values_vio[0])

    #data pembekalan
    sp_ori = "PEMBEKALAN ORIENTASI 2026"
    sheet_ori = client.open(sp_ori)
    ws_ori = sheet_ori.worksheet("2026")
    values_ori = ws_ori.get("A:M")
    pembekalan = pd.DataFrame(values_ori[1:], columns=values_ori[0])

    #data ddc
    sp_ddc = "DEFENSIVE DRIVING ALL BU 2026"
    sheet_ddc = client.open(sp_ddc)
    ws_ddc = sheet_ddc.worksheet("2026")
    values_ddc = ws_ddc.get("A:M")
    ddc = pd.DataFrame(values_ddc[1:], columns=values_ddc[0])

    #data teknik pengoperasian
    sp_refresh = "TRAINING 2026"
    sheet_refresh = client.open(sp_refresh)
    ws_refresh = sheet_refresh.worksheet("Teknik Pengoperasian")
    values_refresh = ws_refresh.get("A:O")
    refresh = pd.DataFrame(values_refresh[1:], columns=values_refresh[0])

    SheetDok = "DOKUMENTASI"
    sheet_dok = client.open(SheetDok)
    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("D:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    weeks = sorted(violation["Week"].dropna().unique())

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

    filtered_vio = violation[violation["Week"].isin(week_filter)]
    filtered_ori = pembekalan[pembekalan["Week"].isin(week_filter)]
    filtered_ddc = ddc[ddc["Week"].isin(week_filter)]
    filtered_refresh = refresh[refresh["Week"].isin(week_filter)]

    refresh_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"].isin(["Refresh","Pembekalan"])& (dokumentasi["Year"]=="2026"))
    ]
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

    #violation
    with st.container(border=True):
        st.markdown(
            f"""
            <style>
                .judul h2 {{
                    font-size: 35px;
                    font-weight: 500;
                    color: black;
                    text-align: center;
                }}
            </style>
            <div class="judul">
                <h2>Refresh Violation</h2>
            </div>
            """,
            unsafe_allow_html=True)
        col1, col2, col3= st.columns([4,0.1,3])
        with col1: 
            header1=st.container()
            pivot_vio = filtered_vio.pivot_table(
                index="Departement",
                columns="BU",
                values="Nama Karyawan",
                aggfunc="count",
                fill_value=0).reset_index()
            with header1:
                st.markdown(f"""<h3 style='text-align: center;'>Data {judul_week}</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bus = [bu for bu in possible_bus if bu in pivot_vio.columns]

                if len(available_bus) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_totals = {bu: pivot_vio[bu].sum() for bu in available_bus}

                    labels = list(bu_totals.keys())
                    sizes  = list(bu_totals.values())

                    if sum(sizes) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_vio = pd.DataFrame({
                            "BU": labels,
                            "Total": sizes
                        })

                        fig = px.bar(
                            chart_vio,
                            x="BU",
                            y="Total",
                            text="Total",
                            color="BU",
                            color_discrete_map={
                                "DCM": "#134f5c",
                                "HPAL": "#2f9a7f",
                                "ONC": "#31681a",
                                "Lainnya": "#ff9900"
                            }
                        )

                        fig.update_traces(
                            textposition="outside",
                            textfont_size=14
                        )

                        fig.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            bargap=0.35,
                            xaxis_title="Business Unit",
                            yaxis_title="Jumlah",
                            margin=dict(t=20, b=20, l=20, r=20)
                        )

                        st.plotly_chart(
                            fig,
                            use_container_width=True,
                            key="bar_vio"
                        )                                
                possible_bus = ["DCM", "HPAL", "ONC"]
                available_bus = [bu for bu in possible_bus if bu in pivot_vio.columns]

                if len(available_bus) == 0:
                    pivot_vio["Total"] = 0
                else:
                    pivot_vio["Total"] = pivot_vio[available_bus].sum(axis=1)

                pivot_vio = pivot_vio.sort_values("Total", ascending=False, ignore_index=True)
                pivot_vio.index = pivot_vio.index + 1
                with st.expander("Data Refresh Violation Weekly"):
                    st.dataframe(pivot_vio, height=300)
        with col3:
            header2=st.container()
            violation_chart = violation.pivot_table(
                index="Departement",
                columns="BU",
                values="Nama Karyawan",
                aggfunc="count",
                fill_value=0).reset_index()
            with header2:
                st.markdown(f"""<h3 style='text-align: center;'>Data 2026</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bu_vio_2026 = [bu for bu in possible_bus if bu in violation_chart.columns]

                if len(available_bu_vio_2026) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_totals_vio_2026 = {bu: violation_chart[bu].sum() for bu in available_bu_vio_2026}

                    labels2 = list(bu_totals_vio_2026.keys())
                    sizes2  = list(bu_totals_vio_2026.values())

                    if sum(sizes2) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_2= pd.DataFrame({
                            "BU": labels2,
                            "Total": sizes2
                        })

                        fig2 = px.pie(
                            chart_2,
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
                        total_all = sum(sizes2)
                        fig2.update_layout(
                            height=400,
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
                            fig2,
                            use_container_width=True,
                            key="pie_violaton_2026"
                        )
                if len(available_bu_vio_2026) == 0:
                    violation_chart["Total"] = 0
                else:
                    violation_chart["Total"] = violation_chart[available_bu_vio_2026].sum(axis=1)

                violation_chart = violation_chart.sort_values("Total", ascending=False, ignore_index=True)
                violation_chart.index = violation_chart.index + 1
                with st.expander("Data Refresh Violation 2026"):
                    st.dataframe(violation_chart, height=300)
    st.divider()

    #non violation
    with st.container(border=True):
        st.markdown(
            f"""
            <style>
                .judul h2 {{
                    font-size: 35px;
                    font-weight: 500;
                    color: black;
                    text-align: center;
                }}
            </style>
            <div class="judul">
                <h2>Refresh Non Violation</h2>
            </div>
            """,
            unsafe_allow_html=True)
        a, b, c= st.columns([4,0.1,3])
        with a: 
            headerA=st.container()
            pivot_refresh = filtered_refresh.pivot_table(
                index="Area",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerA:
                st.markdown(f"""<h3 style='text-align: center;'>Data {judul_week}</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_refresh = [bu for bu in possible_bus if bu in pivot_refresh.columns]

                if len(available_refresh) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_refresh = {bu: pivot_refresh[bu].sum() for bu in available_refresh}

                    labelsA= list(bu_total_refresh.keys())
                    sizesA  = list(bu_total_refresh.values())

                    if sum(sizesA) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_refresh = pd.DataFrame({
                            "BU": labelsA,
                            "Total": sizesA
                        })

                        figA = px.bar(
                            chart_refresh,
                            x="BU",
                            y="Total",
                            text="Total",
                            color="BU",
                            color_discrete_map={
                                "DCM": "#134f5c",
                                "HPAL": "#2f9a7f",
                                "ONC": "#31681a",
                                "Lainnya": "#ff9900"
                            }
                        )

                        figA.update_traces(
                            textposition="outside",
                            textfont_size=14
                        )

                        figA.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            bargap=0.35,
                            xaxis_title="Business Unit",
                            yaxis_title="Jumlah",
                            margin=dict(t=20, b=20, l=20, r=20)
                        )

                        st.plotly_chart(
                            figA,
                            use_container_width=True,
                            key="bar_refresh"
                        )                                
                if len(available_refresh) == 0:
                    pivot_refresh["Total"] = 0
                else:
                    pivot_refresh["Total"] = pivot_refresh[available_refresh].sum(axis=1)

                pivot_refresh = pivot_refresh.sort_values("Total", ascending=False, ignore_index=True)
                pivot_refresh.index = pivot_refresh.index + 1
                with st.expander("Data Refresh Non Violation Weekly"):
                    st.dataframe(pivot_refresh, height=300)
        with c:
            headerC=st.container()
            refresh_chart = refresh.pivot_table(
                index="Area",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerC:
                st.markdown(f"""<h3 style='text-align: center;'>Data 2026</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bu_refresh = [bu for bu in possible_bus if bu in refresh_chart.columns]

                if len(available_bu_refresh) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_refresh_2026 = {bu: refresh_chart[bu].sum() for bu in available_bu_refresh}

                    labelsC = list(bu_total_refresh_2026.keys())
                    sizesC  = list(bu_total_refresh_2026.values())

                    if sum(sizesC) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        refresh_2026= pd.DataFrame({
                            "BU": labelsC,
                            "Total": sizesC
                        })

                        figC = px.pie(
                            refresh_2026,
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

                        figC.update_traces(
                            textinfo="label+value",
                            textfont_size=14
                        )
                        total_all_refresh = sum(sizesC)
                        figC.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(t=20, b=20, l=20, r=20),
                            annotations=[
                                dict(
                                    text=f"<b>{total_all_refresh}</b>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=26, color="black"),
                                    showarrow=False
                                )
                            ]
                        )

                        st.plotly_chart(
                            figC,
                            use_container_width=True,
                            key="pie_refresh_2026"
                        )
                if len(available_bu_refresh) == 0:
                    refresh_chart["Total"] = 0
                else:
                    refresh_chart["Total"] = refresh_chart[available_bu_refresh].sum(axis=1)

                refresh_chart = refresh_chart.sort_values("Total", ascending=False, ignore_index=True)
                refresh_chart.index = refresh_chart.index + 1
                with st.expander("Refresh Non Violation 2026"):
                    st.dataframe(refresh_chart, height=300)
    st.divider()

    #pembekalan
    with st.container(border=True):
        st.markdown(
            f"""
            <style>
                .judul h2 {{
                    font-size: 35px;
                    font-weight: 500;
                    color: black;
                    text-align: center;
                }}
            </style>
            <div class="judul">
                <h2>Pembekalan Orientasi</h2>
            </div>
            """,
            unsafe_allow_html=True)
        d,e,f= st.columns([4,0.1,3])
        with d: 
            headerD=st.container()
            pivot_orientasi = filtered_ori.pivot_table(
                index="Departement",
                columns="BU",
                values="Nama Karyawan",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerD:
                st.markdown(f"""<h3 style='text-align: center;'>Data {judul_week}</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_orientasi = [bu for bu in possible_bus if bu in pivot_orientasi.columns]

                if len(available_orientasi) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_ori = {bu: pivot_orientasi[bu].sum() for bu in available_orientasi}

                    labelsD= list(bu_total_ori.keys())
                    sizesD  = list(bu_total_ori.values())

                    if sum(sizesD) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_orientasi = pd.DataFrame({
                            "BU": labelsD,
                            "Total": sizesD
                        })

                        figD = px.bar(
                            chart_orientasi,
                            x="BU",
                            y="Total",
                            text="Total",
                            color="BU",
                            color_discrete_map={
                                "DCM": "#134f5c",
                                "HPAL": "#2f9a7f",
                                "ONC": "#31681a",
                                "Lainnya": "#ff9900"
                            }
                        )

                        figD.update_traces(
                            textposition="outside",
                            textfont_size=14
                        )

                        figD.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            bargap=0.35,
                            xaxis_title="Business Unit",
                            yaxis_title="Jumlah",
                            margin=dict(t=20, b=20, l=20, r=20)
                        )

                        st.plotly_chart(
                            figD,
                            use_container_width=True,
                            key="bar_orientasi"
                        )                                
                if len(available_orientasi) == 0:
                    pivot_orientasi["Total"] = 0
                else:
                    pivot_orientasi["Total"] = pivot_orientasi[available_orientasi].sum(axis=1)

                pivot_orientasi = pivot_orientasi.sort_values("Total", ascending=False, ignore_index=True)
                pivot_orientasi.index = pivot_orientasi.index + 1
                with st.expander("Data Pembekalan Orientasi Weekly"):
                    st.dataframe(pivot_orientasi, height=300)
        with f:
            headerF=st.container()
            pivot_orientasi2026 = pembekalan.pivot_table(
                index="Departement",
                columns="BU",
                values="Nama Karyawan",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerF:
                st.markdown(f"""<h3 style='text-align: center;'>Data 2026</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bu_orientasi2026 = [bu for bu in possible_bus if bu in pivot_orientasi2026.columns]

                if len(available_bu_orientasi2026) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_orientasi_2026 = {bu: pivot_orientasi2026[bu].sum() for bu in available_bu_orientasi2026}

                    labelsF = list(bu_total_orientasi_2026.keys())
                    sizesF  = list(bu_total_orientasi_2026.values())

                    if sum(sizesF) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        orientasi_2026= pd.DataFrame({
                            "BU": labelsF,
                            "Total": sizesF
                        })

                        figF = px.pie(
                            orientasi_2026,
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

                        figF.update_traces(
                            textinfo="label+value",
                            textfont_size=14
                        )
                        total_all_orientasi = sum(sizesF)
                        figF.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(t=20, b=20, l=20, r=20),
                            annotations=[
                                dict(
                                    text=f"<b>{total_all_orientasi}</b>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=26, color="black"),
                                    showarrow=False
                                )
                            ]
                        )

                        st.plotly_chart(
                            figF,
                            use_container_width=True,
                            key="pie_orientasi_2026"
                        )
                if len(available_bu_orientasi2026) == 0:
                    pivot_orientasi2026["Total"] = 0
                else:
                    pivot_orientasi2026["Total"] = pivot_orientasi2026[available_bu_orientasi2026].sum(axis=1)

                pivot_orientasi2026 = pivot_orientasi2026.sort_values("Total", ascending=False, ignore_index=True)
                pivot_orientasi2026.index = pivot_orientasi2026.index + 1
                with st.expander("Data Pembekalan Orientasi 2026"):
                    st.dataframe(pivot_orientasi2026, height=300)
    st.divider()

    #ddc
    with st.container(border=True):
        st.markdown(
            f"""
            <style>
                .judul h2 {{
                    font-size: 35px;
                    font-weight: 500;
                    color: black;
                    text-align: center;
                }}
            </style>
            <div class="judul">
                <h2>Defensive Driving</h2>
            </div>
            """,
            unsafe_allow_html=True)
        g,h,i= st.columns([4,0.1,3])
        with g: 
            headerG=st.container()
            pivot_ddc = filtered_ddc.pivot_table(
                index="Departemen",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerG:
                st.markdown(f"""<h3 style='text-align: center;'>Data {judul_week}</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_ddc = [bu for bu in possible_bus if bu in pivot_ddc.columns]

                if len(available_ddc) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_ddc = {bu: pivot_ddc[bu].sum() for bu in available_ddc}

                    labelsG= list(bu_total_ddc.keys())
                    sizesG  = list(bu_total_ddc.values())

                    if sum(sizesG) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        chart_ddc = pd.DataFrame({
                            "BU": labelsG,
                            "Total": sizesG
                        })

                        figG = px.bar(
                            chart_ddc,
                            x="BU",
                            y="Total",
                            text="Total",
                            color="BU",
                            color_discrete_map={
                                "DCM": "#134f5c",
                                "HPAL": "#2f9a7f",
                                "ONC": "#31681a",
                                "Lainnya": "#ff9900"
                            }
                        )

                        figG.update_traces(
                            textposition="outside",
                            textfont_size=14
                        )

                        figG.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            bargap=0.35,
                            xaxis_title="Business Unit",
                            yaxis_title="Jumlah",
                            margin=dict(t=20, b=20, l=20, r=20)
                        )

                        st.plotly_chart(
                            figG,
                            use_container_width=True,
                            key="bar_ddc"
                        )                                
                if len(available_ddc) == 0:
                    pivot_ddc["Total"] = 0
                else:
                    pivot_ddc["Total"] = pivot_ddc[available_ddc].sum(axis=1)

                pivot_ddc = pivot_ddc.sort_values("Total", ascending=False, ignore_index=True)
                pivot_ddc.index = pivot_ddc.index + 1
                with st.expander("Data Defensive Driving Weekly"):
                    st.dataframe(pivot_ddc, height=300)
        with i:
            headerI=st.container()
            pivot_ddc2026 = ddc.pivot_table(
                index="Departemen",
                columns="BU",
                values="Nama",
                aggfunc="count",
                fill_value=0).reset_index()
            with headerI:
                st.markdown(f"""<h3 style='text-align: center;'>Data 2026</h3>""", unsafe_allow_html=True)
                possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]

                available_bu_ddc2026 = [bu for bu in possible_bus if bu in pivot_ddc2026.columns]

                if len(available_bu_ddc2026) == 0:
                    st.info("Tidak ada data untuk ditampilkan.")
                else:
                    bu_total_ddc_2026 = {bu: pivot_ddc2026[bu].sum() for bu in available_bu_ddc2026}

                    labelsI = list(bu_total_ddc_2026.keys())
                    sizesI  = list(bu_total_ddc_2026.values())

                    if sum(sizesI) == 0:
                        st.info("Data kosong, tidak bisa menampilkan chart.")
                    else:
                        ddc_2026= pd.DataFrame({
                            "BU": labelsI,
                            "Total": sizesI
                        })

                        figI = px.pie(
                            ddc_2026,
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

                        figI.update_traces(
                            textinfo="label+value",
                            textfont_size=14
                        )
                        total_all_ddc = sum(sizesI)
                        figI.update_layout(
                            height=400,
                            showlegend=False,
                            plot_bgcolor="white",
                            paper_bgcolor="white",
                            margin=dict(t=20, b=20, l=20, r=20),
                            annotations=[
                                dict(
                                    text=f"<b>{total_all_ddc}</b>",
                                    x=0.5,
                                    y=0.5,
                                    font=dict(size=26, color="black"),
                                    showarrow=False
                                )
                            ]
                        )

                        st.plotly_chart(
                            figI,
                            use_container_width=True,
                            key="pie_ddc_2026"
                        )
                if len(available_bu_ddc2026) == 0:
                    pivot_ddc2026["Total"] = 0
                else:
                    pivot_ddc2026["Total"] = pivot_ddc2026[available_bu_ddc2026].sum(axis=1)

                pivot_ddc2026 = pivot_ddc2026.sort_values("Total", ascending=False, ignore_index=True)
                pivot_ddc2026.index = pivot_ddc2026.index + 1
                with st.expander("Data Defensive Driving 2026"):
                    st.dataframe(pivot_ddc2026, height=300)
    st.divider()
  
    st.divider()
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
    import streamlit as st

    import base64
    import requests

    cols_per_row = 3
    rows = refresh_dokumentasi.to_dict("records")

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


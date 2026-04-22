def app():
    import streamlit as st
    import pandas as pd
    import altair as alt
    from streamlit_navigation_bar import st_navbar
    from datetime import date
    from streamlit.components.v1 import html
    from streamlit_option_menu import option_menu
    import matplotlib
    import gspread
    from google.oauth2 import service_account
    from google.oauth2.service_account import Credentials
    import base64

    SHEET_ID = "1aL4qAd8MPbKXnvN_2aeQcmwHVhrApBgEwGxbY8vB7gI"
    SHEET_DATA = "Data"
    SHEET_MANPOWER = "Manpower"

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )

    client = gspread.authorize(creds)

    @st.cache_resource
    def get_gspread_client():
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=SCOPES
        )
        return gspread.authorize(creds)

    gc = get_gspread_client()

    @st.cache_data(ttl=300) 
    def load_sheet(sheet_name: str) -> pd.DataFrame:
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet(sheet_name)

        values = ws.get_all_values()
        headers = values[0]
        rows = values[1:]

        seen = {}
        unique_headers = []
        for h in headers:
            if h in seen:
                seen[h] += 1
                unique_headers.append(f"{h}_{seen[h]}")
            else:
                seen[h] = 0
                unique_headers.append(h)

        df = pd.DataFrame(rows, columns=unique_headers)
        return df
    
    def fig_to_base64(fig):
        if fig is None:
            return ""
        return base64.b64encode(fig.to_image(format="png")).decode()
    
    def create_bar_chart(df, color_map):
        import plotly.express as px

        fig = px.bar(
            df,
            x="BU",
            y="Total",
            text="Total",
            color="BU",
            color_discrete_map=color_map,
            template="seaborn"
        )

        fig.update_traces(textposition="outside")

        fig.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            margin=dict(l=20, r=20, t=20, b=20)
        )
        return fig

    def create_pie_chart(pivot, color_map):
        import plotly.express as px
        import pandas as pd

        possible_bus = ["DCM", "HPAL", "ONC", "Lainnya"]
        available_bus = [bu for bu in possible_bus if bu in pivot.columns]
        bu_total = {bu: pivot[bu].sum() for bu in available_bus}

        bu_total = {k: v for k, v in bu_total.items() if v > 0}
        if len(bu_total) == 0:
            return None

        df_chart = pd.DataFrame({
            "BU": list(bu_total.keys()),
            "Total": list(bu_total.values())
        })

        fig = px.pie(
            df_chart,
            names="BU",
            values="Total",
            color="BU",
            color_discrete_map=color_map,
            hole=0.4
        )

        fig.update_traces(
            textinfo="label+value",
            textfont_size=14
        )

        total_all = df_chart["Total"].sum()

        fig.update_layout(
            height=500,
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

        return fig

    ssDokumentasi = "DOKUMENTASI"
    sheetDokumentasi = client.open(ssDokumentasi)
    wsDok = sheetDokumentasi.worksheet("Dokumentasi")
    valuesDok = wsDok.get("D:L")
    dokumentasi = pd.DataFrame(valuesDok[1:],columns=valuesDok[0])

    df = load_sheet(SHEET_DATA)
    df_manpower = load_sheet(SHEET_MANPOWER)

    data = df.iloc[:, 3:7]

    weeks = sorted(data["Week"].dropna().unique())
    week1_start = date(2026, 1, 2)
    today = date.today()
    days_since = (today - week1_start).days
    current_week = (days_since // 7) + 1
    default_week = f"Week {current_week}"
    default_week = [default_week] if default_week in weeks else []
    week_filter = st.multiselect("Pilih Week", weeks, default=default_week,width=250)
    filtered_df = data[
    data["Week"].isin(week_filter)]
    sum_kegiatan = (
        filtered_df.groupby(["BU", "Kegiatan"])["Jumlah"]
        .sum()
        .reset_index()
    )
    sum_kegiatan["Jumlah"] = (
        pd.to_numeric(
            sum_kegiatan["Jumlah"],
            errors="coerce"
        )
        .fillna(0)
        .astype(int)
    )
    df_ok=filtered_df[filtered_df["BU"]!="Lainnya"]
    pivot=df_ok.pivot_table(
        index="Kegiatan",
        columns="BU",
        values="Jumlah",
        aggfunc="sum",
        fill_value=0
    ).reset_index()
    possible_bus = ["DCM", "HPAL", "ONC"]
    available_bus = [bu for bu in possible_bus if bu in pivot.columns]
    if len(available_bus) == 0:
        pivot["Total"] = 0
    else:
        pivot[available_bus] = pivot[available_bus].apply(pd.to_numeric, errors="coerce")
        pivot["Total"] = pivot[available_bus].sum(axis=1)
    pivot_tabel = pivot.sort_values("Total", ascending=False, ignore_index=True)
    pivot_tabel.index = pivot_tabel.index + 1

    st.title("📥 Download Report")
    
    ##Overview
    def generate_overview_html(df_manpower, sum_kegiatan, pivot_tabel, week_filter):

        total_manpower = len(df_manpower)
        onsite = (df_manpower["On/Off Site"] == "On-Site").sum()
        offsite = (df_manpower["On/Off Site"] == "Off-Site").sum()
        total_kegiatan = sum(sum_kegiatan["Jumlah"])

        #chart kegiatan
        import plotly.express as px
        import base64

        def fig_to_base64(fig):
            return base64.b64encode(fig.to_image(format="png")).decode()
        
        sum_kegiatan["label"] = sum_kegiatan["Jumlah"].apply(lambda x: x if x != 0 else "")

        fig = px.bar(
            sum_kegiatan,
            x="Jumlah",
            y="Kegiatan",
            color="BU",
            orientation="h",
            text="label",
            color_discrete_map={
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a",
                "Lainnya": "#ff9900"
            }
        )

        fig.update_layout(
            yaxis=dict(
                categoryorder="total ascending"  # biar kebalik jadi terbesar di atas
            )
        )

        fig.update_traces(
            textposition="outside",
            textfont_size=14
        )

        fig.update_layout(
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis_title=None,
            yaxis_title=None,
            showlegend=True,
            margin=dict(l=150, r=20, t=20, b=20)
        )

        chart_base64 = fig_to_base64(fig)

        html = f"""
        <html>
        <head>
        <style>
            body {{
                font-family: Figtree;
                padding:20px;
                background: linear-gradient(135deg, #eef2ff, #f8fafc);
            }}
            .kpi-container {{display:flex; gap:20px;padding: 10px;}}
            .kpi {{
                border:1px solid #ddd;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                padding:10px;
                border-radius:10px;
                width:200px;
                text-align:center;
                background-color: #fff}}
            .kpi2 {{
                padding:15px;
                border-radius:10px;
                border:1px solid #ddd;
                text-align:center;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                max-width:900px;background-color: #fff
            }}
            .kpi-simper {{
                border:1px solid #ddd;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                padding:10px;
                border-radius:10px;
                max-width:600px;
                text-align:center;
                background-color: #fff}}
            .container {{
                padding:15px;
                border:1px solid #ddd;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                border-radius:10px;
                text-align:center;
                max-width:1400px;background-color: #fff
            }}
            .square-img {{
                width: 100%;
                aspect-ratio: 1 / 1;
                overflow: hidden;
                border-radius: 14px;
                box-shadow: 0 4px 12px rgba(0,0,0,0.08);
            }}

            .square-img img {{
                width: 100%;
                height: 100%;
                object-fit: cover;
                transition: 0.3s ease-in-out;
            }}

            .square-img img:hover {{
                transform: scale(1.05);
            }}

            .img-caption {{
                text-align: center;
                font-size: 14px;
                margin-top: 6px;
                color: #444;
            }}
        </style>
        <style>
            table {{
                border-collapse: collapse;
                width: 100%;
            }}
            th {{
                text-align: center;
                padding: 8px;
            }}
            td {{
                text-align: center;
                padding: 6px;
            }}
        </style>
        </head>
        <body>


        <h1>📊 Weekly Report Asset Management</h1>
        <h3>Week: {", ".join(week_filter)}</h3>

        <div class="kpi-container">
            <div class="kpi"><h4>Total Manpower</h4><h2>{total_manpower}</h2></div>
            <div class="kpi"><h4>On Site</h4><h2>{onsite}</h2></div>
            <div class="kpi"><h4>Off Site</h4><h2>{offsite}</h2></div>
            <div class="kpi"><h4>Total Kegiatan</h4><h2>{total_kegiatan}</h2></div>
        </div>

        <div class="kpi-container">
            <div class="kpi2"><h2>📊 Chart Kegiatan</h2>
                <img src="data:image/png;base64,{chart_base64}" width="600">
                <details>
                <summary style="padding-top:10px">Lihat Detail</summary>
                {pivot_tabel.to_html(index=False)}
                </details></div>

            <div class="kpi2"><h2>👷 Manpower</h2>
            {df_manpower.to_html(index=False)}</div>
        </div>
        </body>
        </html>
        """


        return html

    #INDUKSI
    spreadsheetinduksi = "8. INDUKSI ALL BU 2026"
    sheetinduksi = client.open(spreadsheetinduksi)
    worksheetInduksi = sheetinduksi.worksheet("2026")
    valuesInduksi = worksheetInduksi.get("D:N")
    dfInduksi = pd.DataFrame(valuesInduksi[1:],columns=valuesInduksi[0])
    filtered_induksi = dfInduksi[dfInduksi["Week"].isin(week_filter)]
    induksi_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Induksi") &
        (dokumentasi["Year"]=="2026")
    ]
    pivot_induksi = filtered_induksi.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_induksi = (
        filtered_induksi["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_induksi.columns = ["BU", "Total"]

    pivot_induksi2026 = dfInduksi.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_induksi_html(week_filter, filtered_induksi, pivot_induksi2026, induksi_dokumentasi):
        import requests
        import base64
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"
        total_induksi = len(filtered_induksi)
        total_perusahaan = filtered_induksi["Perusahaan"].nunique()
        total_dept = filtered_induksi["Department"].nunique()

        pivot_induksi = filtered_induksi.pivot_table(
            index="Perusahaan",
            columns="BU",
            values="Nama",
            aggfunc="count",
            fill_value=0
        ).reset_index()

        bu_counts = (
            filtered_induksi["BU"]
            .value_counts()
            .reset_index()
        )
        bu_counts.columns = ["BU", "Total"]

        fig_induksi = create_bar_chart(
            bu_counts,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        chart_induksi = fig_to_base64(fig_induksi)

        pie_induksi = create_pie_chart(
            pivot_induksi2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        piechart_induksi = fig_to_base64(pie_induksi)

        html = f"""
        <h2 style="text-align:center;">🛡️ Induksi {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi"><h4>Total Induksi</h4><h2>{total_induksi}</h2></div>
            <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
            <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
        </div>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Induksi</h2>
                <img src="data:image/png;base64,{chart_induksi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_induksi.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Induksi 2026</h2>
                <img src="data:image/png;base64,{piechart_induksi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_induksi2026.to_html(index=False)}
                </details>
            </div>
        </div>

        <h3>📸 Dokumentasi Induksi</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in induksi_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue

        html += "</div>"  # 🔥 tutup grid

        return html
    

    #INSPEKSI
    spreadsheetinspeksi = "INSPEKSI ALL BU 2026"
    sheetinspeksi = client.open(spreadsheetinspeksi)
    worksheetInspeksi = sheetinspeksi.worksheet("2026")
    valuesInspeksi= worksheetInspeksi.get("E:O")
    dfInspeksi = pd.DataFrame(valuesInspeksi[1:],columns=valuesInspeksi[0])
    df_inspeksi = dfInspeksi[dfInspeksi["BU"].notna() & (dfInspeksi["BU"].str.strip() != "")]
    filtered_inspeksi = df_inspeksi[df_inspeksi["Week"].isin(week_filter)]
    inspeksi_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Inspeksi & Observasi") &
        (dokumentasi["Year"]=="2026")]
    pivot_inspeksi =filtered_inspeksi.pivot_table(
        index="Departement",
        columns="BU",
        values="Nomor Lambung",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_inspeksi = (
        filtered_inspeksi["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_inspeksi.columns = ["BU", "Total"]

    pivot_inspeksi2026 =df_inspeksi.pivot_table(
        index="Departement",
        columns="BU",
        values="Nomor Lambung",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    #Temuan Inspeksi
    ws_temuan = sheetinspeksi.worksheet("Olah Temuan")
    values_temuan = ws_temuan.get("A:H")
    df_temuan = pd.DataFrame(values_temuan[1:],columns=values_temuan[0])
    filtered_temuan = df_temuan[df_temuan["Week"].isin(week_filter)]

    #Observasi
    spreadsheetobservasi = "OBSERVASI ALL BU 2026"
    sheetobservasi = client.open(spreadsheetobservasi)
    worksheetObservasi = sheetobservasi.worksheet("Observasi")
    valuesObservasi= worksheetObservasi.get("A:O")
    dfObservasi = pd.DataFrame(valuesObservasi[1:],columns=valuesObservasi[0])
    df_observasi = dfObservasi[dfObservasi["BU"].notna() & (dfObservasi["BU"].str.strip() != "")]
    filtered_observasi = df_observasi[df_observasi["Week"].isin(week_filter)]
    pivot_observasi =filtered_observasi.pivot_table(
        index="Departemen",
        columns="BU",
        values="Driver/Operator",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_observasi= (
        filtered_observasi["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_observasi.columns = ["BU", "Total"]

    pivot_observasi2026 = df_observasi.pivot_table(
        index="Departemen",
        columns="BU",
        values="Driver/Operator",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_inspeksi_html(week_filter, filtered_inspeksi,pivot_inspeksi,filtered_observasi, pivot_observasi, pivot_inspeksi2026, pivot_observasi2026, filtered_temuan, inspeksi_dokumentasi):
        import requests
        import plotly.express as px
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        total_inspeksi = len(filtered_inspeksi)
        total_observasi = len(filtered_observasi)
        total_perusahaan_inspeksi = filtered_inspeksi["Perusahaan"].nunique()
        total_dept_inspeksi = filtered_inspeksi["Departement"].nunique()
        bu_counts_inspeksi = (
            filtered_inspeksi["BU"]
            .value_counts()
            .reset_index()
        )
        bu_counts_inspeksi.columns = ["BU", "Total"]
        fig_inspeksi = create_bar_chart(
            bu_counts_inspeksi,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_inspeksi = fig_to_base64(fig_inspeksi)

        pie_inspeksi = create_pie_chart(
            pivot_inspeksi2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        piechart_inspeksi = fig_to_base64(pie_inspeksi)

        bu_counts_observasi = (
            filtered_observasi["BU"]
            .value_counts()
            .reset_index()
        )
        bu_counts_observasi.columns = ["BU", "Total"]
        fig_observasi = create_bar_chart(
            bu_counts_observasi,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_observasi = fig_to_base64(fig_observasi)

        pie_observasi = create_pie_chart(
            pivot_observasi2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        piechart_observasi= fig_to_base64(pie_observasi)

        top_temuan = (
            filtered_temuan
            .groupby("Klasifikasi Temuan", as_index=False)
            .size()
            .rename(columns={"size": "Jumlah"})
            .sort_values("Jumlah", ascending=False)
            .head(10)  
        )

        fig_temuan = px.bar(
            top_temuan,
            y="Klasifikasi Temuan",
            x="Jumlah",               
            text="Jumlah"
        )

        fig_temuan.update_traces(textposition="outside", width= 0.4)
        fig_temuan.update_layout(
            bargap=0.3,
            yaxis_title=None,
            xaxis_title=None,
            xaxis=dict(categoryorder="total ascending",
                       tickfont=dict(size=10)),
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(
                tickfont=dict(size=10),
                categoryorder="total ascending"
            ),
            margin=dict(l=250, r=20, t=20, b=50)
        )
        chart_temuan_weekly= fig_to_base64(fig_temuan)
        html = f"""
            <html>
            <head>
            <style>
                .kpi-container {{display:flex; gap:20px;padding: 10px;}}
                .kpi {{
                    border:1px solid #ddd;
                    box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                    padding:10px;
                    border-radius:10px;
                    width:200px;
                    text-align:center;
                    background-color: #fff}}
                .kpi2 {{
                    padding:15px;
                    border:1px solid #ddd;
                    box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                    border-radius:10px;
                    text-align:center;
                    width:700px;background-color: #fff
                }}
                .container {{
                    padding:15px;
                    border:1px solid #ddd;
                    box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                    border-radius:10px;
                    text-align:center;
                    width:1400px;background-color: #fff
                }}
            </style>
            </head>

            <body>
            <h2 style="text-align:center;">🔎 Inspeksi & Observasi {judul_week}</h2>
            <div class="kpi-container">
                <div class="kpi"><h4>Total Inspeksi</h4><h2>{total_inspeksi}</h2></div>
                <div class="kpi"><h4>Total Observasi</h4><h2>{total_observasi}</h2></div>
                <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan_inspeksi}</h2></div>
                <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept_inspeksi}</h2></div>
            </div>

            <div class="kpi-container">
                <div class="kpi2"><h2>📊 Chart Inspeksi</h2>
                <img src="data:image/png;base64,{chart_inspeksi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data Inspeksi</summary>
                    {pivot_inspeksi.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2"><h2>📊 Chart Observasi</h2>
                <img src="data:image/png;base64,{chart_observasi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data Observasi</summary>
                    {pivot_observasi.to_html(index=False)}
                </details>
            </div>
            </div>
            </body>            
            </html>"""
        html += f"""
            <div class="kpi-container">
                <div class="kpi2"><h2>📊 Chart Inspeksi 2026</h2>
                <img src="data:image/png;base64,{piechart_inspeksi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data Inspeksi</summary>
                    {pivot_inspeksi2026.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2"><h2>📊 Chart Observasi 2026</h2>
                <img src="data:image/png;base64,{piechart_observasi}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data Observasi</summary>
                    {pivot_observasi2026.to_html(index=False)}
                </details>
            </div>
            </div>"""
        html += f"""
        <div class="kpi-container">
            <div class="container"><h2>Temuan Inspeksi Weekly</h2>
                <img src="data:image/png;base64,{chart_temuan_weekly}" width="1250">
            </div>
        </div>"""
        html += """    
            <h3>📸 Dokumentasi Inspeksi & Observasi</h3>

            <div style="
                display: grid;
                grid-template-columns: repeat(3, 1fr);
                gap: 20px;
            ">
            """

        for _, row in inspeksi_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="width:100%; 
                        height:300px; 
                        object-fit:cover; 
                        border-radius:10px;
                        border:1px solid #ddd;
                        box-shadow: 0 3px 8px rgba(0,0,0,0.08);">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue
        html += "</div>" 

        return html
    

    #Recommissioning
    spreadsheetrecom = "COMMISSIONING & RECOMMISSIONING ALL BU 2026"
    sheetrecom = client.open(spreadsheetrecom)
    worksheetRecom = sheetrecom.worksheet("2026")
    valuesRecom = worksheetRecom.get("D:U")
    dfRecom = pd.DataFrame(valuesRecom[1:],columns=valuesRecom[0])
    filtered_recom = dfRecom[dfRecom["Week"].isin(week_filter)]
    ws_temuan_recom=sheetrecom.worksheet("Olah Temuan")
    values_temuan_recom=ws_temuan_recom.get("A:H")
    df_temuan_recom = pd.DataFrame(values_temuan_recom[1:],columns=values_temuan_recom[0])
    filtered_temuan_recom = df_temuan_recom[df_temuan_recom["Week"].isin(week_filter)]

    com = dfRecom[dfRecom["Keterangan"] == "COMMISSIONING"]
    recom = dfRecom[dfRecom["Keterangan"] == "RECOMMISSIONING"]

    filtered_commissioning = filtered_recom[filtered_recom["Keterangan"] == "COMMISSIONING"]
    filtered_recommissioning = filtered_recom[filtered_recom["Keterangan"] == "RECOMMISSIONING"]
    
    recom_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"].isin(["Recommissioning", "Commissining"])) &
        (dokumentasi["Year"]=="2026")
    ]

    #commissioning
    df_lulus_com=filtered_commissioning[filtered_commissioning["LULUS/TIDAK"] == "LULUS UJI KENDARAAN"]
    pivot_lulus_com = df_lulus_com.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()
    
    bu_counts_lulus_com= (
        df_lulus_com["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_lulus_com.columns = ["BU", "Total"]
    df_tidaklulus_com=filtered_commissioning[filtered_commissioning["LULUS/TIDAK"] == "BELUM LULUS UJI KENDARAAN"]
    pivot_tidaklulus_com = df_tidaklulus_com.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()

    bu_counts_tidaklulus_com= (
        df_tidaklulus_com["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_tidaklulus_com.columns = ["BU", "Total"]

    #recommissioning
    df_lulus_recom=filtered_recommissioning[filtered_recommissioning["LULUS/TIDAK"] == "LULUS UJI KENDARAAN"]
    pivot_lulus_recom = df_lulus_recom.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()
    
    bu_counts_lulus_recom= (
        df_lulus_recom["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_lulus_recom.columns = ["BU", "Total"]
    df_tidaklulus_recom=filtered_recommissioning[filtered_recommissioning["LULUS/TIDAK"] == "BELUM LULUS UJI KENDARAAN"]
    pivot_tidaklulus_recom = df_tidaklulus_recom.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()

    bu_counts_tidaklulus_recom= (
        df_tidaklulus_recom["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_tidaklulus_recom.columns = ["BU", "Total"]

    #2026
    pivot_recom = recom.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()
    
    bu_counts_recom= (
        recom["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_recom.columns = ["BU", "Total"]

    pivot_com = com.pivot_table(
            index="Departemen",
            columns="BU",
            values="No. Unit",
            aggfunc="count",
            fill_value=0
        ).rename_axis(None, axis=1).reset_index()
    
    bu_counts_com= (
        com["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_com.columns = ["BU", "Total"]

    def generate_recom_html(week_filter, filtered_recom,
                            filtered_commissioning,
                            filtered_recommissioning,
                            pivot_lulus_com,
                            pivot_tidaklulus_com, 
                            bu_counts_lulus_com, 
                            bu_counts_tidaklulus_com,
                            pivot_lulus_recom,
                            pivot_tidaklulus_recom,
                            bu_counts_lulus_recom, 
                            bu_counts_tidaklulus_recom,
                            pivot_recom,
                            pivot_com,
                            bu_counts_recom, 
                            bu_counts_com,
                            filtered_temuan_recom, 
                            recom_dokumentasi):
        import requests
        import base64
        import pandas as pd
        import plotly.express as px
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        total_commissioning = len(filtered_commissioning)
        total_recommissioning = len(filtered_recommissioning)
        total_perusahaan = filtered_recom["Perusahaan"].nunique()
        total_dept = filtered_recom["Departemen"].nunique()

        #lulus Commissioning
        fig_lulus_com = create_bar_chart(
            bu_counts_lulus_com,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_lulus_com = fig_to_base64(fig_lulus_com)

        #tidak lulus com
        fig_tidaklulus_com = create_bar_chart(
            bu_counts_tidaklulus_com,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_tidaklulus_com = fig_to_base64(fig_tidaklulus_com)

         #lulus Recommissioning
        fig_lulus_recom = create_bar_chart(
            bu_counts_lulus_recom,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_lulus_recom = fig_to_base64(fig_lulus_recom)

        #tidak lulus recom
        fig_tidaklulus_recom = create_bar_chart(
            bu_counts_tidaklulus_recom,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_tidaklulus_recom = fig_to_base64(fig_tidaklulus_recom)

        #com
        fig_com = create_bar_chart(
            bu_counts_com,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_com = fig_to_base64(fig_com)

        #recom
        fig_recom = create_bar_chart(
            bu_counts_recom,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_recom = fig_to_base64(fig_recom)

        top_temuan_recom =(
        filtered_temuan_recom
        .groupby("Klasifikasi Temuan", as_index=False)
        .size()
        .rename(columns={"size": "Jumlah"})
        .sort_values("Jumlah", ascending=False)
        .head(10)
        )

        fig_temuan_recom = px.bar(
            top_temuan_recom,
            y="Klasifikasi Temuan",
            x="Jumlah",               
            text="Jumlah"
        )
        fig_temuan_recom.update_traces(textposition="outside", width= 0.4)
        fig_temuan_recom.update_layout(
            bargap=0.3,
            yaxis_title=None,
            xaxis_title=None,
            xaxis=dict(categoryorder="total ascending",
                       tickfont=dict(size=10)),
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(
                tickfont=dict(size=10),
                categoryorder="total ascending"
            ),
            margin=dict(l=250, r=20, t=20, b=50)
        )
        chart_temuan_recom= fig_to_base64(fig_temuan_recom)

        html = f"""
        <h2 style="text-align:center;">🚜 Commissioning & Recommissioning {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi"><h4>Total Commissioning</h4><h2>{total_commissioning}</h2></div>
            <div class="kpi"><h4>Total Recommissioning</h4><h2>{total_recommissioning}</h2></div>            
            <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
            <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
        </div>


        <h3>Commissioning</h3>
        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Lulus Uji Kelayakan</h2>
                <img src="data:image/png;base64,{chart_lulus_com}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_lulus_com.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Tidak Lulus Uji Kelayakan</h2>
                <img src="data:image/png;base64,{chart_tidaklulus_com}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_tidaklulus_com.to_html(index=False)}
                </details>
            </div>
        </div>"""

        html += f"""
        <h3>Recommissioning</h3>
        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Lulus Uji Kelayakan</h2>
                <img src="data:image/png;base64,{chart_lulus_recom}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_lulus_recom.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Tidak Lulus Uji Kelayakan</h2>
                <img src="data:image/png;base64,{chart_tidaklulus_recom}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_tidaklulus_recom.to_html(index=False)}
                </details>
            </div>
        </div>"""

        html += f"""
        <h2>Commissioning & Recommissioning 2026</h2>
        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Commissioning 2026</h2>
                <img src="data:image/png;base64,{chart_com}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_com.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Recommissioning 2026</h2>
                <img src="data:image/png;base64,{chart_recom}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_recom.to_html(index=False)}
                </details>
            </div>
        </div>"""
        
        html += f"""
        <div class="kpi-container">
            <div class="container"><h2>Temuan Commissioning & Recommissioning Weekly</h2>
                <img src="data:image/png;base64,{chart_temuan_recom}" width="1250">
            </div>
        </div>"""
        html+=f"""
        <h3>📸 Dokumentasi Commissioning & Recommissioning</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in recom_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue

        html += "</div>" 

        return html
    
    #refresh
    spreadsheetviolation= "REFRESH VIOLATION 2026"
    sheetviolation = client.open(spreadsheetviolation)
    worksheetViolation = sheetviolation.worksheet("2026")
    valuesViolation = worksheetViolation.get("D:U")
    dfViolation= pd.DataFrame(valuesViolation[1:],columns=valuesViolation[0])
    filtered_violation = dfViolation[dfViolation["Week"].isin(week_filter)]
    refresh_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"].isin(["Refresh", "Pembekalan"])) &
        (dokumentasi["Year"]=="2026")
    ]
    pivot_dfVio = dfViolation.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama Karyawan",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_vio = filtered_violation.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama Karyawan",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_vio = (
        filtered_violation["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_vio.columns = ["BU", "Total"]

    #refresh non violation
    sp_refresh = "TRAINING 2026"
    sheet_refresh = client.open(sp_refresh)
    ws_refresh = sheet_refresh.worksheet("Teknik Pengoperasian")
    values_refresh = ws_refresh.get("A:O")
    refresh = pd.DataFrame(values_refresh[1:], columns=values_refresh[0])
    filtered_refresh= refresh[refresh["Week"].isin(week_filter)]
    pivot_fullRefresh= refresh.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_refresh = filtered_refresh.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_refresh = (
        filtered_refresh["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_refresh.columns = ["BU", "Total"]

    #pembekalan orientasi
    sp_pembekalan = "PEMBEKALAN ORIENTASI 2026"
    sheet_pembekalan = client.open(sp_pembekalan)
    ws_pembekalan = sheet_pembekalan.worksheet("2026")
    values_pembekalan = ws_pembekalan.get("A:M")
    pembekalan = pd.DataFrame(values_pembekalan[1:], columns=values_pembekalan[0])
    filtered_pembekalan= pembekalan[pembekalan["Week"].isin(week_filter)]
    pivot_fullPembekalan= pembekalan.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama Karyawan",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_pembekalan = filtered_pembekalan.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama Karyawan",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_pembekalan = (
        filtered_pembekalan["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_pembekalan.columns = ["BU", "Total"]

    #defensive driving
    sp_ddc = "DEFENSIVE DRIVING ALL BU 2026"
    sheet_ddc = client.open(sp_ddc)
    ws_ddc = sheet_ddc.worksheet("2026")
    values_ddc = ws_ddc.get("A:M")
    ddc = pd.DataFrame(values_ddc[1:], columns=values_ddc[0])
    filtered_ddc= ddc[ddc["Week"].isin(week_filter)]
    pivot_fullddc= ddc.pivot_table(
        index="Departemen",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_ddc = filtered_ddc.pivot_table(
        index="Departemen",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_ddc = (
        filtered_ddc["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_ddc.columns = ["BU", "Total"]
    
    def generate_refresh_html(week_filter, pivot_vio, 
                              bu_counts_vio, 
                              pivot_dfVio,
                              pivot_fullRefresh,
                              pivot_refresh,
                              bu_counts_refresh,
                              pivot_fullPembekalan,
                              pivot_pembekalan,
                              bu_counts_pembekalan,
                              pivot_fullddc,
                              pivot_ddc,
                              bu_counts_ddc,
                              refresh_dokumentasi):
        import requests
        import base64
        import pandas as pd
        import plotly.express as px
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        #refresh violation
        fig_vio = create_bar_chart(
            bu_counts_vio,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_vio = fig_to_base64(fig_vio)
        
        pie_vio = create_pie_chart(
            pivot_dfVio,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_vio= fig_to_base64(pie_vio)

        #refresh non violation
        fig_refresh = create_bar_chart(
            bu_counts_refresh,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_refresh = fig_to_base64(fig_refresh)
        
        pie_refresh = create_pie_chart(
            pivot_fullRefresh,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_refresh= fig_to_base64(pie_refresh)

        #pembekalan orientasi
        fig_pembekalan = create_bar_chart(
            bu_counts_pembekalan,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_pembekalan = fig_to_base64(fig_pembekalan)
        
        pie_pembekalan = create_pie_chart(
            pivot_fullPembekalan,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_pembekalan= fig_to_base64(pie_pembekalan)

        #defensive driving
        fig_ddc = create_bar_chart(
            bu_counts_ddc,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )
        chart_ddc = fig_to_base64(fig_ddc)
        
        pie_ddc = create_pie_chart(
            pivot_fullddc,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_ddc= fig_to_base64(pie_ddc)
        
        #html violation
        html = f"""
        <h2 style="text-align:center;">⚠️ Refresh Violation {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Refresh Violation (Coaching)</h2>
                <img src="data:image/png;base64,{chart_vio}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_vio.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Refresh Violation (Coaching) 2026</h2>
                <img src="data:image/png;base64,{piechart_vio}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_dfVio.to_html(index=False)}
                </details>
            </div>
        </div>"""

        #html non violation
        html += f"""
        <h2 style="text-align:center;">🔄 Refresh Non Violation {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Refresh Non-Violation</h2>
                <img src="data:image/png;base64,{chart_refresh}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_refresh.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Refresh Non-Violation 2026</h2>
                <img src="data:image/png;base64,{piechart_refresh}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_fullRefresh.to_html(index=False)}
                </details>
            </div>
        </div>"""

        #html pembekalan
        html += f"""
        <h2 style="text-align:center;">📚 Pembekalan Orientasi {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Pembekalan Orientasi</h2>
                <img src="data:image/png;base64,{chart_pembekalan}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_pembekalan.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Pembekalan Orientasi 2026</h2>
                <img src="data:image/png;base64,{piechart_pembekalan}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_fullPembekalan.to_html(index=False)}
                </details>
            </div>
        </div>"""

        #html ddc
        html += f"""
        <h2 style="text-align:center;">🚘 Defensive Driving {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Defensive Driving</h2>
                <img src="data:image/png;base64,{chart_ddc}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_ddc.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Defensive Driving 2026</h2>
                <img src="data:image/png;base64,{piechart_ddc}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_fullddc.to_html(index=False)}
                </details>
            </div>
        </div>"""

        html+=f"""
        <h3>📸 Dokumentasi Refresh</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in refresh_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue

        html += "</div>" 

        return html
    
    #SIMPER
    ss_simper = "7. DATA SIMPER ALL BU 2026"
    sheet_simper= client.open(ss_simper)
    ws_simper = sheet_simper.worksheet("2026")
    values_simper = ws_simper.get("A:J")
    df_simper = pd.DataFrame(values_simper[1:], columns=values_simper[0])
    filtered_simper = df_simper[df_simper["Week 2026"].isin(week_filter)]
    #full
    filtered_full=filtered_simper[filtered_simper["Status SIMPER"] == "F"]
    full=df_simper[df_simper["Status SIMPER"] == "F"]
    pivot_fullSIMPER= full.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_full = filtered_full.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    #probation
    filtered_prob=filtered_simper[filtered_simper["Status SIMPER"] == "P"]
    probation=df_simper[df_simper["Status SIMPER"] == "P"]
    pivot_probSIMPER= probation.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_prob = filtered_prob.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    #sementara
    filtered_sememntara=filtered_simper[filtered_simper["Status SIMPER"] == "T"]
    sementara=df_simper[df_simper["Status SIMPER"] == "T"]
    pivot_semSIMPER= sementara.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_sementara = filtered_sememntara.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_simper_html(week_filter,filtered_simper, pivot_full, pivot_prob, pivot_sementara,pivot_fullSIMPER, pivot_probSIMPER, pivot_semSIMPER):
        import requests
        import base64
        import pandas as pd
        import plotly.express as px
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        total_simper = len(filtered_simper)
        total_perusahaan = filtered_simper["Perusahaan"].nunique()
        total_dept = filtered_simper["Departemen"].nunique()

        #full
        pie_filtered_full = create_pie_chart(
            pivot_full,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_full= fig_to_base64(pie_filtered_full)

        #probation
        pie_filtered_prob = create_pie_chart(
            pivot_prob,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_prob= fig_to_base64(pie_filtered_prob)

        #sementara
        pie_filtered_sem = create_pie_chart(
            pivot_sementara,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_sem= fig_to_base64(pie_filtered_sem)


        html = f"""
            <h2 style="text-align:center;">🪪 SIMPER {judul_week}</h2>


            <div class="kpi-container">
                <div class="kpi"><h4>Total SIMPER</h4><h2>{total_simper}</h2></div>
                <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
                <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
            </div>

            <div class="kpi-container">
                <div class="kpi-simper">
                    <h2>📊SIMPER Full</h2>
                    <img src="data:image/png;base64,{piechart_filtered_full}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_full.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊SIMPER Probation</h2>
                    <img src="data:image/png;base64,{piechart_filtered_prob}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_prob.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊SIMPER Sementara</h2>
                    <img src="data:image/png;base64,{piechart_filtered_sem}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_sementara.to_html(index=False)}
                    </details>
                </div>
            </div>"""
        
        #full
        pie_filtered_full2026 = create_pie_chart(
            pivot_fullSIMPER,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_full2026= fig_to_base64(pie_filtered_full2026)

        #probation
        pie_filtered_prob2026 = create_pie_chart(
            pivot_probSIMPER,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_prob2026= fig_to_base64(pie_filtered_prob2026)

        #sementara
        pie_filtered_sem2026 = create_pie_chart(
            pivot_semSIMPER,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_sem2026= fig_to_base64(pie_filtered_sem2026)

        html += f"""
            <div class="kpi-container">
                <div class="kpi-simper">
                    <h2>📊SIMPER Full 2026</h2>
                    <img src="data:image/png;base64,{piechart_filtered_full2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_fullSIMPER.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊SIMPER Probation 2026</h2>
                    <img src="data:image/png;base64,{piechart_filtered_prob2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_probSIMPER.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊SIMPER Sementara 2026</h2>
                    <img src="data:image/png;base64,{piechart_filtered_sem2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_semSIMPER.to_html(index=False)}
                    </details>
                </div>
            </div>"""

        html += "</div>" 

        return html

    #TES PRAKTIK
    
    ss_praktik = "TES PRAKTIK ALL BU 2026"
    sheet_praktik= client.open(ss_praktik)
    ws_praktik = sheet_praktik.worksheet("2026")
    values_praktik = ws_praktik.get("A:Q")
    df_praktik = pd.DataFrame(values_praktik[1:], columns=values_praktik[0])
    filtered_praktik = df_praktik[df_praktik["Week"].isin(week_filter)]

    praktik_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Tes Praktik") &(dokumentasi["Year"]=="2026")
    ]

    #kandidat
    filtered_kandidat=filtered_praktik[filtered_praktik["Kategori"] == "TES KANDIDAT"]
    kandidat=df_praktik[df_praktik["Kategori"] == "TES KANDIDAT"]
    pivot_kandidat2026= kandidat.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_kandidat = filtered_kandidat.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    #praktik
    filtered_tespraktik=filtered_praktik[filtered_praktik["Kategori"] == "TES PRAKTIK"]
    praktik=df_praktik[df_praktik["Kategori"] == "TES PRAKTIK"]
    pivot_praktik2026= praktik.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_praktik= filtered_tespraktik.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    #penambahan
    filtered_penambahan=filtered_praktik[filtered_praktik["Kategori"] == "Penambahan Versatility"]
    penambahan=df_praktik[df_praktik["Kategori"] == "Penambahan Versatility"]
    pivot_penambahan2026= penambahan.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    pivot_penambahan = filtered_penambahan.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_praktik_html(week_filter, filtered_praktik, pivot_kandidat, pivot_kandidat2026,
                              pivot_praktik, pivot_praktik2026,
                              pivot_penambahan, pivot_penambahan2026, praktik_dokumentasi):
        import requests
        import base64
        import pandas as pd
        import plotly.express as px
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"
        total_lulus = filtered_praktik[filtered_praktik["KET"] == "LULUS"]
        total_tdklulus = filtered_praktik[filtered_praktik["KET"] == "TIDAK LULUS"]

        total_lulus = len(total_lulus)
        total_tdklulus = len(total_tdklulus)
        total_perusahaan = filtered_praktik["Perusahaan"].nunique()
        total_dept = filtered_praktik["Department"].nunique()

        #praktik
        pie_filtered_praktik = create_pie_chart(
            pivot_praktik,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_praktik= fig_to_base64(pie_filtered_praktik)

        #kandidat
        pie_filtered_kandidat = create_pie_chart(
            pivot_kandidat,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_kandidat= fig_to_base64(pie_filtered_kandidat)

        #penambahan
        pie_filtered_penambahan = create_pie_chart(
            pivot_penambahan,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_penambahan= fig_to_base64(pie_filtered_penambahan)

        html = f"""
            <h2 style="text-align:center;">🏗️ Tes Praktik {judul_week}</h2>

            <div class="kpi-container">
                <div class="kpi"><h4>Total Lulus</h4><h2>{total_lulus}</h2></div>
                <div class="kpi"><h4>Total Tidak Lulus</h4><h2>{total_tdklulus}</h2></div>
                <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
                <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
            </div>

            <div class="kpi-container">
                <div class="kpi-simper">
                    <h2>📊 Tes Kandidat</h2>
                    <img src="data:image/png;base64,{piechart_filtered_kandidat}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_kandidat.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊 Tes Praktik</h2>
                    <img src="data:image/png;base64,{piechart_filtered_praktik}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_praktik.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊 Penambahan Versatility</h2>
                    <img src="data:image/png;base64,{piechart_filtered_penambahan}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_penambahan.to_html(index=False)}
                    </details>
                </div>
            </div>"""
        
        #praktik
        pie_filtered_praktik2026 = create_pie_chart(
            pivot_praktik2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_praktik2026= fig_to_base64(pie_filtered_praktik2026)

        #kandidat
        pie_filtered_kandidat2026 = create_pie_chart(
            pivot_kandidat2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_kandidat2026= fig_to_base64(pie_filtered_kandidat2026)

        #penambahan
        pie_filtered_penambahan2026 = create_pie_chart(
            pivot_penambahan2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_filtered_penambahan2026= fig_to_base64(pie_filtered_penambahan2026)

        html += f"""
            <div class="kpi-container">
                <div class="kpi-simper">
                    <h2>📊 Tes Kandidat</h2>
                    <img src="data:image/png;base64,{piechart_filtered_kandidat2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_kandidat2026.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊 Tes Praktik</h2>
                    <img src="data:image/png;base64,{piechart_filtered_praktik2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_praktik2026.to_html(index=False)}
                    </details>
                </div>

                <div class="kpi-simper">
                    <h2>📊 Penambahan Versatility</h2>
                    <img src="data:image/png;base64,{piechart_filtered_penambahan2026}" width="450">
                    <details>
                        <summary style="padding-top:10px">Detail Data</summary>
                        {pivot_penambahan2026.to_html(index=False)}
                    </details>
                </div>
            </div>"""

        html+=f"""
        <h3>📸 Dokumentasi Tes Praktik</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in praktik_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue
        
        html += "</div>" 

        return html

    #TRAINING
    spreadsheetTraining = "TRAINING 2026"
    sheettraining = client.open(spreadsheetTraining)
    ws_training = sheettraining.worksheet("Training")
    values_training = ws_training.get("E:Q")
    df_training = pd.DataFrame(values_training[1:], columns=values_training[0])
    filtered_training = df_training[df_training["Week"].isin(week_filter)]
    training_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Training") &
        (dokumentasi["Year"]=="2026")
    ]
    bu_counts_training = (
        filtered_training["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_training.columns = ["BU", "Total"]

    pivot_training2026 = df_training.pivot_table(
        index="Perusahaan",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_training_html(week_filter, filtered_training, pivot_training2026, training_dokumentasi):
        import requests
        import base64
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        total_training = len(filtered_training)
        total_perusahaan = filtered_training["Perusahaan"].nunique()
        total_dept = filtered_training["Departement"].nunique()

        pivot_training = filtered_training.pivot_table(
            index="Perusahaan",
            columns="BU",
            values="Nama",
            aggfunc="count",
            fill_value=0
        ).reset_index()

        bu_counts_training = (
            filtered_training["BU"]
            .value_counts()
            .reset_index()
        )
        bu_counts_training.columns = ["BU", "Total"]

        fig_training = create_bar_chart(
            bu_counts_training,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        chart_training = fig_to_base64(fig_training)

        piechart_training = create_pie_chart(pivot_training2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            })
        piechart_training = fig_to_base64(piechart_training)

        html = f"""
        <h2 style="text-align:center;">👷🏽 Training {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi"><h4>Total Training</h4><h2>{total_training}</h2></div>
            <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
            <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
        </div>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Training</h2>
                <img src="data:image/png;base64,{chart_training}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_training.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Training 2026</h2>
                <img src="data:image/png;base64,{piechart_training}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_training2026.to_html(index=False)}
                </details>
            </div>
        </div>

        <h3>📸 Dokumentasi Training</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in training_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue

        html += "</div>"

        return html
    
        #Complience Rate
    sscompliencerate = "TRAINING 2026"
    sheetCR = client.open(sscompliencerate)
    worksheetCR = sheetCR.worksheet("Training Complience Rate")
    valuesCR = worksheetCR.get("E:R")
    dfCR = pd.DataFrame(valuesCR[1:],columns=valuesCR[0])
    filtered_CR = dfCR[dfCR["Week"].isin(week_filter)]
    CR_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Complience Rate") &
        (dokumentasi["Year"]=="2026")
    ]
    pivot_CR = filtered_CR.pivot_table(
        index="Departement",
        columns="Judul Training",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()
    bu_counts_CR = (
        filtered_CR["BU"]
        .value_counts()
        .reindex(possible_bus, fill_value=0)
        .reset_index()
    )
    bu_counts_CR.columns = ["BU", "Total"]

    pivot_CR2026 = dfCR.pivot_table(
        index="Departement",
        columns="BU",
        values="Nama",
        aggfunc="count",
        fill_value=0
    ).rename_axis(None, axis=1).reset_index()

    def generate_CR_html(week_filter, filtered_CR, pivot_CR, pivot_CR2026, CR_dokumentasi):
        import requests
        import base64
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"
        total_CR = len(filtered_CR)
        total_perusahaan = filtered_CR["Perusahaan"].nunique()
        total_dept = filtered_CR["Departement"].nunique()

        bu_counts_CR = (
            filtered_CR["BU"]
            .value_counts()
            .reset_index()
        )
        bu_counts_CR.columns = ["BU", "Total"]

        fig_CR = create_bar_chart(
            bu_counts_CR,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        chart_CR = fig_to_base64(fig_CR)

        pie_CR = create_pie_chart(
            pivot_CR2026,
            {
                "DCM": "#134f5c",
                "HPAL": "#2f9a7f",
                "ONC": "#31681a"
            }
        )

        piechart_CR = fig_to_base64(pie_CR)

        html = f"""
        <h2 style="text-align:center;">👷🏻‍♂️ Complience Rate {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi"><h4>Total Peserta</h4><h2>{total_CR}</h2></div>
            <div class="kpi"><h4>Total Perusahaan</h4><h2>{total_perusahaan}</h2></div>
            <div class="kpi"><h4>Total Departemen</h4><h2>{total_dept}</h2></div>
        </div>

        <div class="kpi-container">
            <div class="kpi2">
                <h2>📊 Chart Complience Rate</h2>
                <img src="data:image/png;base64,{chart_CR}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_CR.to_html(index=False)}
                </details>
            </div>

            <div class="kpi2">
                <h2>📊 Chart Complience Rate 2026</h2>
                <img src="data:image/png;base64,{piechart_CR}" width="600">
                <details>
                    <summary style="padding-top:10px">Detail Data</summary>
                    {pivot_CR2026.to_html(index=False)}
                </details>
            </div>
        </div>

        <h3>📸 Dokumentasi Complience Rate</h3>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in CR_dokumentasi.iterrows():
            url = row.get("url_clean", "")

            try:
                res = requests.get(url)
                img_base64 = base64.b64encode(res.content).decode()

                html += f"""
                <div style="text-align:center;">
                    <img src="data:image/jpeg;base64,{img_base64}" 
                        style="
                            width:100%;
                            height:300px;
                            object-fit:cover;
                            border-radius:10px;
                            border:1px solid #ddd;
                            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
                        ">
                    <small>{row.get("Keterangan","")}</small>
                </div>
                """
            except:
                continue

        html += "</div>"

        return html

    #P5M
    ssp5m = "6. BRIEFING P5M 2026"
    sheetp5m = client.open(ssp5m)
    worksheetp5m = sheetp5m.worksheet("2026")
    valuesp5m = worksheetp5m.get("E:V")
    dfp5m = pd.DataFrame(valuesp5m[1:],columns=valuesp5m[0])
    filtered_p5m = dfp5m[dfp5m["Week"].isin(week_filter)]

    def generate_p5m_html(week_filter, filtered_p5m):
        import requests
        import base64
        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"
        html = f"""
        <h2 style="text-align:center;">📢 Pembicaraan 5 Menit (P5M) {judul_week}</h2>

        <div style="
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 20px;
        ">
        """

        for _, row in filtered_p5m.iterrows():
            url = row.get("url_clean", "")

            if not url:
                continue

            try:
                res = requests.get(url, timeout=10)
                content_type = res.headers.get("Content-Type", "")

                materi = row.get("Materi", "")
                lokasi = row.get("Lokasi", "")

                import re
                items = re.split(r'\d+\.\s*', materi)
                items = [i.strip() for i in items if i.strip()]

                list_html = "".join([f"<li>{item}</li>" for item in items])

                if "image" in content_type:
                    img_base64 = base64.b64encode(res.content).decode()

                    html += f"""
                    <div>
                        <div class="square-img">
                            <img src="data:image/jpeg;base64,{img_base64}">
                        </div>

                        <div class="img-caption">
                            <b>{lokasi}</b>
                            <ol style="text-align:left; padding-left:18px;">
                                {list_html}
                            </ol>
                        </div>
                    </div>
                    """
                else:
                    html += f"""
                    <div>
                        <p style="color:red;">Link bukan gambar</p>
                        <small>{url}</small>
                    </div>
                    """

            except Exception as e:
                html += f"""
                <div>
                    <p style="color:red;">Gagal load gambar</p>
                    <small>{str(e)}</small>
                </div>
                """
        html += "</div>"
        return html
    
        #Issues
    ssissues = "ASSET MANAGEMENT 2026"
    sheetissues = client.open(ssissues)
    worksheetissues = sheetissues.worksheet("Feedback Issues")
    valuesissuses = worksheetissues.get("A:I")
    dfissues = pd.DataFrame(valuesissuses[1:],columns=valuesissuses[0])
    filtered_issues= dfissues[dfissues["Week"].isin(week_filter)]

    def generate_issues_html(dfissues, filtered_issues, week_filter):
        df_open = dfissues[
            (dfissues["Status"] == "Open") &
            (~dfissues["Week"].isin(week_filter))
        ]
    
        total_issue = dfissues["Status"].count()
        total_close = (dfissues["Status"] == "Close").sum()

        closing_rate = (total_close / total_issue * 100) if total_issue > 0 else 0

        judul_week = ", ".join(week_filter) if week_filter else "Semua Week"

        html = f"""
        <h2 style="text-align:center;">🚧 Issues {judul_week}</h2>

        <div class="kpi-container">
            <div class="kpi">
                <h4>Closing Rate</h4>
                <h2>{closing_rate:.1f}%</h2>
            </div>
        </div>

        <div style="
            display:grid;
            grid-template-columns: repeat(3, 1fr);
            gap:20px;
            margin-top:20px;
        ">
        """
        for _, row in filtered_issues.iterrows():

            issue = row.get("Issue", "-")
            solusi = row.get("Solusi", "-")
            pic = row.get("PIC", "-")
            status = row.get("Status", "-")
            due = row.get("Due Date", "-")

            html += f"""
            <div style="
                background:white;
                padding:15px;
                border-radius:12px;
                box-shadow: 0 3px 8px rgba(0,0,0,0.08);
            ">
                <h4>🚧 {issue}</h4>

                <p><b>Action:</b> {solusi}</p>
                <p><b>PIC:</b> {pic}</p>
                <p><b>Status:</b> {status}</p>
                <p><b>Due Date:</b> {due}</p>
            </div>
            """
        html += "</div>"

        html += "<h3>📌 Open Issues from Previous Weeks</h3>"

        if df_open.empty:
            html += "<p>✅ Tidak ada issues open dari week sebelumnya</p>"
        else:
            html += df_open[
                ["Week", "Date", "Issue", "PIC", "Status", "Due Date"]
            ].to_html(index=False)
        return html

    
 
    
    overview_html = generate_overview_html(
        df_manpower,
        sum_kegiatan,
        pivot_tabel,
        week_filter)

    induksi_html = generate_induksi_html(
        week_filter, filtered_induksi, pivot_induksi2026,
        induksi_dokumentasi
    )

    inspeksi_html = generate_inspeksi_html(
        week_filter, filtered_inspeksi,
        pivot_inspeksi,
        filtered_observasi, 
        pivot_observasi, 
        pivot_inspeksi2026, 
        pivot_observasi2026, 
        filtered_temuan, 
        inspeksi_dokumentasi
    )

    recom_html = generate_recom_html(
        week_filter, filtered_recom,
        filtered_commissioning,
        filtered_recommissioning,
        pivot_lulus_com,
        pivot_tidaklulus_com, 
        bu_counts_lulus_com, 
        bu_counts_tidaklulus_com,
        pivot_lulus_recom,
        pivot_tidaklulus_recom,
        bu_counts_lulus_recom, 
        bu_counts_tidaklulus_recom,
        pivot_recom,
        pivot_com,
        bu_counts_recom, 
        bu_counts_com,
        filtered_temuan_recom, 
        recom_dokumentasi
    )

    refresh_html = generate_refresh_html(
        week_filter, pivot_vio, 
        bu_counts_vio, 
        pivot_dfVio,
        pivot_fullRefresh,
        pivot_refresh,
        bu_counts_refresh,
        pivot_fullPembekalan,
        pivot_pembekalan,
        bu_counts_pembekalan,
        pivot_fullddc,
        pivot_ddc,
        bu_counts_ddc,
        refresh_dokumentasi
    )

    simper_html = generate_simper_html(
        week_filter, filtered_simper, pivot_full, pivot_prob, pivot_sementara,
        pivot_fullSIMPER, pivot_probSIMPER, pivot_semSIMPER
    )

    praktik_html = generate_praktik_html(
        week_filter, filtered_praktik, pivot_kandidat, pivot_kandidat2026,
        pivot_praktik, pivot_praktik2026,
        pivot_penambahan, pivot_penambahan2026, praktik_dokumentasi
    )

    training_html = generate_training_html(week_filter, filtered_training,
                                           pivot_training2026, 
                                           training_dokumentasi)
    
    CR_html = generate_CR_html(
        week_filter, filtered_CR, pivot_CR, pivot_CR2026, CR_dokumentasi    
    )

    p5m_html = generate_p5m_html(
        week_filter, filtered_p5m    
    )

    issues_html = generate_issues_html(
        dfissues, filtered_issues, week_filter   
    )

    final_html = (
        overview_html.replace("</body></html>", "")
        + induksi_html
        + inspeksi_html
        + recom_html
        + praktik_html
        + simper_html
        + refresh_html        
        + training_html
        + CR_html
        + p5m_html
        + issues_html
        + "</body></html>"
)
    st.download_button(
        label="📥 Download HTML Report",
        data=final_html,
        file_name="weekly_report.html",
        mime="text/html"
    )

    st.components.v1.html(final_html, height=600, scrolling=True)
   
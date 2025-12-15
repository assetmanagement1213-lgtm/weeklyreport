
def app():
    import streamlit as st
    import pandas as pd
    import gspread
    from google.oauth2.service_account import Credentials
    from datetime import date
    import requests
    import matplotlib.pyplot as plt

    q,w =st.columns([4,1])
    with q:
        st.header("Induksi")
    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )

    client = gspread.authorize(creds)

    spreadsheet_name = "INDUKSI ALL BU 2025"
    sheet = client.open(spreadsheet_name)

    # === ambil sheet induksi ===
    naemSheetDok = "DOKUMENTASI"
    sheet_dok = client.open(naemSheetDok)
    ws_induksi = sheet.worksheet("2025")
    values_induksi = ws_induksi.get("E:N")
    df_induksi = pd.DataFrame(values_induksi[1:], columns=values_induksi[0])

    # === ambil sheet dokumentasi ===
    ws_dok = sheet_dok.worksheet("Dokumentasi")
    values_dok = ws_dok.get("E:L")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    # ================= WEEK FILTER ==================
    weeks = sorted(df_induksi["Week"].dropna().unique())

    week1_start = date(2024, 12, 27)
    today = date.today()
    days_since = (today - week1_start).days
    current_week = (days_since // 7) + 1
    default_week = f"Week {current_week}"

    with w:
        week_filter = st.multiselect(
            "Pilih Week",
            weeks,
            default="Week 48",
            max_selections=None
        )

    # ================= FILTER KE 2 DATA ==================
    filtered_induksi = df_induksi[df_induksi["Week"].isin(week_filter)]
    induksi_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter)) &
        (dokumentasi["Kegiatan"] == "Induksi")
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
                <h4>Total induksi</h4>
                <div class="value">{len(filtered_induksi)}</div>
            </div>
        """, unsafe_allow_html=True)
    with x2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Perusahaan</h4>
                <div class="value">{filtered_induksi["Perusahaan"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)
    with x3:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Departemen</h4>
                <div class="value">{filtered_induksi["Department"].nunique()}</div>
            </div>
            
        """, unsafe_allow_html=True)

    st.divider()
    col1, col2 = st.columns([5,3])
    with col1 :
        pivot = filtered_induksi.pivot_table(
            index="Perusahaan",
            columns="BU",
            values="Nama",
            aggfunc="count",
            fill_value=0
        ).reset_index()

        # --- Kolom BU yang mungkin ---
        possible_bus = ["DCM", "HPAL", "ONC"]

        # --- Kolom BU yang benar-benar ada di pivot ---
        available_bus = [bu for bu in possible_bus if bu in pivot.columns]

        # --- Jika tidak ada sama sekali, buat kolom Total = 0 ---
        if len(available_bus) == 0:
            pivot["Total"] = 0
        else:
            pivot["Total"] = pivot[available_bus].sum(axis=1)

        # --- Urutkan pivot ---
        pivot = pivot.sort_values("Total", ascending=False, ignore_index=True)
        pivot.index = pivot.index + 1
        st.dataframe(pivot, height=400)
    with col2:
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
                fig1, ax1 = plt.subplots()

                # warna custom
                colors = {
                    "DCM": "#134f5c",
                    "HPAL": "#2f9a7f",
                    "ONC": "#31681a",
                    "Lainnya": "#ff9900"
                }
                color_list = [colors[bu] for bu in labels]

                # =====================
                # PIE CHART
                # =====================
                if sum(sizes) == 0:
                    st.info("Data kosong, tidak bisa menampilkan chart.")
                else:
                    fig, ax = plt.subplots(figsize=(6, 6))
                    ax.pie(
                        sizes,
                        labels=labels,
                        autopct="%1.1f%%",
                        startangle=90,
                        colors=color_list
                    )
                    ax.axis("equal")

                    st.pyplot(fig)

    st.divider()

    st.subheader("📸 Dokumentasi Kegiatan")
    import requests
    import streamlit as st

    cols_per_row = 3
    rows = induksi_dokumentasi.to_dict("records")

    for i in range(0, len(rows), cols_per_row):
        cols = st.columns(cols_per_row)

        for col, row in zip(cols, rows[i:i+cols_per_row]):
            url = row["url_clean"]

            if url:
                try:
                    response = requests.get(url, stream=True)

                    content_type = response.headers.get("Content-Type", "")

                    with col:
                        # Jika benar-benar image
                        if "image" in content_type:
                            st.image(
                                response.content,
                                caption=row.get("Keterangan", "")
                            )
                        else:
                            st.warning("Link tidak mengembalikan file gambar.")
                            st.write(row["url_clean"])

                except Exception as e:
                    with col:
                        st.error(f"Error load: {e}")

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
                <h1>Briefing Pembicaraan 5 Menit (P5M)</h1>
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

    spreadsheet_name = "6. BRIEFING P5M 2026"
    sheet = client.open(spreadsheet_name)

    # === ambil sheet dokumentasi ===
    nameSheetDok = "6. BRIEFING P5M 2026"
    sheet_dok = client.open(nameSheetDok)
    ws_recom = sheet.worksheet("2026")
    values_recom = ws_recom.get("E:V")
    df_recom = pd.DataFrame(values_recom[1:], columns=values_recom[0])

    # === ambil sheet dokumentasi ===
    ws_dok = sheet_dok.worksheet("2026")
    values_dok = ws_dok.get("D:N")
    dokumentasi = pd.DataFrame(values_dok[1:], columns=values_dok[0])

    # ================= WEEK FILTER ==================
    weeks = sorted(df_recom["Week"].dropna().unique())

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

        return f"{week_sorted[0]} – {week_sorted[-1]}"

    recom_dokumentasi = dokumentasi[
        (dokumentasi["Week"].isin(week_filter))&(dokumentasi["Year"]=="2026")]


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
    
    if recom_dokumentasi.empty:
        st.info(f"Tidak ada data yang dapat ditampilkan.")
        st.stop()

    cols_per_row = 3
    rows = recom_dokumentasi.to_dict("records")

    for i in range(0, len(rows), cols_per_row):
        cols = st.columns(cols_per_row)

        for col, row in zip(cols, rows[i:i+cols_per_row]):

            url = row.get("url_clean", "")

            with col:
                if not url:
                    continue

                try:
                    response = requests.get(url, stream=True, timeout=10)
                    content_type = response.headers.get("Content-Type", "")

                    lokasi = row.get("Lokasi", "")
                    materi = row.get("Materi", "")

                    if "image" in content_type:
                        img_base64 = base64.b64encode(response.content).decode()
                        st.markdown(f"""
                        <div class="square-img">
                            <img src="data:image/jpeg;base64,{img_base64}">
                        </div>
                        """, unsafe_allow_html=True)

                        caption_text = f"**{lokasi}**\n\n{materi}"

                        st.markdown(caption_text)

                    else:
                        st.warning("Link tidak mengembalikan file gambar.")
                        st.write(url)

                except Exception as e:

                    st.error(f"Gagal load gambar: {e}")


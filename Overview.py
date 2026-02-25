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
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

USERS = {
    "asset": "asset1213",
}
st.markdown("""
    <style>
    [data-testid="stHeader"] {
            background-color: rgba(0, 0, 0, 0);}
    </style>
""", unsafe_allow_html=True)
def show_login():

    st.markdown("""
        <style>
        .stApp {
            background: linear-gradient(135deg, #134f5c, #2f9a7f);
        }
        .login-title {
            text-align: center;
            font-size: 28px;
            font-weight: 600;
            margin-bottom: 30px;
            color: #134f5c;
        }
        [class="stVerticalBlock st-key-form st-emotion-cache-1rf3gxw e12zf7d53"] {
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);

        }
        </style>
    """, unsafe_allow_html=True)
    main_container=st.container(
        height=400,
        width=500,
        key="form",
        horizontal_alignment="center",
        vertical_alignment="center", border=True
    )
    with main_container:
        content_container = st.container(
            width=300, 
            gap="medium"
        )
    with content_container :
        st.markdown('<div class="login-title">🔐 Dashboard Login</div>', unsafe_allow_html=True)

        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        if st.button("Login", use_container_width=True):
            if username in USERS and USERS[username] == password:
                st.session_state.logged_in = True
                st.session_state.username = username
                st.rerun()
            else:
                st.error("Username atau Password salah")

        st.markdown('</div>', unsafe_allow_html=True)

if not st.session_state.logged_in:
    show_login()
    st.stop()

if "page" not in st.session_state:
    st.session_state.page = "Overview"


st.set_page_config(
    page_title="Weekly Report Asset Management 2026",
    page_icon="👷‍♂️",
    layout="wide"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Poppins', sans-serif;
}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<style>
h1 {
    margin-bottom: 0px;
}
h2 {
    margin-top: 0px;
    margin-bottom: 8px;
}
</style>
""", unsafe_allow_html=True)
st.markdown("""
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.11.1/font/bootstrap-icons.css">
""", unsafe_allow_html=True)

#sidebar
with st.sidebar:
    if st.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.username = None
        st.rerun()
    st.divider()
    st.markdown("""
    <div style="
        display: flex;
        align-items: center;
        text-align: center;
        padding-left: 32px;
        gap: 8px;
        font-family: 'Poppins', sans-serif;
        font-size: 16px;
        font-weight: 600;
        color: black;
    ">
        <i class="bi bi-grid-fill"></i>
        <span>Activity</span>
    </div>
    """, unsafe_allow_html=True)
    

    selected = option_menu(
        menu_title=None,
        options=[
            "Overview","Induksi","Training","Complience Rate",
            "Com/Re-com","Inspeksi & Observasi",
            "SIMPER","Tes Praktik","Refresh","Briefing P5M","Lainnya"
        ],
        icons=[
        "speedometer2",      # Overview
        "person-badge",      # Induksi
        "mortarboard",       # Training
        "chat-square-text",  # Sharing Knowledge
        "tools",             # Commissioning
        "clipboard-check",   # Inspeksi & Observasi
        "credit-card",       # SIMPER
        "check2-square",     # Tes Praktik
        "arrow-clockwise",   # Refresh
        "megaphone",         # Briefing P5M
        "three-dots"
        ],
        styles={
            "icon": {
                "font-size": "15px"
            },
            "container": {
                "padding": "8px",
                "background-color": "#F2f2f3f3"
            },
            "menu-title": {
                "font-size": "16px",
                "font-weight": "600",
                "color": "#111",
                "padding-bottom": "8px"
            },
            "nav-link": {
                "font-size": "13px",
                "font-weight": "500",
                "text-align": "left",
                "margin": "2px 0px",
                "--hover-color": "#eee",
            },
            "nav-link-selected": {
                "background-color": "#000000",
                "color": "white",
                "font-weight": "600"
            }
        }
    )
import Com_Recom
import Induksi
import Inspeksi
import Tes_praktik
import Sharing
import SIMPER
import TrainingA2B
import Refresh
import P5M
import Lainnya

# PAGES = {
#     "Overview": None,
#     "Induksi": Induksi.app,
#     "Training": TrainingA2B.app,
#     "Sharing Knowledge": Sharing.app,
#     "Com/Re-com": Com_Recom.app,
#     "Inspeksi & Observasi": Observasi.app,
#     "SIMPER": SIMPER.app,
# }
# page_func = PAGES.get(selected)

# if page_func:
#     page_func()


import base64

def img_to_base64(path):
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode()

logo_base64 = img_to_base64("asset.png")  # ganti sesuai nama file kamu
st.markdown("""
    <style>
    .header-box {
        background: white;
        border-radius: 20px;
        padding: 28px 40px;
        margin-top: -120px;
        margin-bottom: 16px;
        box-shadow: 0 6px 14px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        gap: 24px;  
    }

    /* teks */
    .header-text h1 {
        margin: 0;
        font-size: 45px;
        font-weight: 700;
        color: #1f2937;
    }

    .header-text h2 {
        margin-top: -25px;
        font-size: 20px;
        font-weight: 500;
        color: #374151;
    }

    /* logo bulat */
    .header-logo img {
        width: 150px;
        height: 150px;
        border-radius: 50%;
        object-fit: cover;
        box-shadow: 0 6px 14px rgba(0,0,0,0.15);
    }
    </style>
    """, unsafe_allow_html=True)
st.markdown(f"""
<div class="header-box">
    <div class="header-logo">
        <img src="data:image/png;base64,{logo_base64}">
    </div>
    <div class="header-text">
        <h1>Weekly Report</h1>
        <h2>Asset Management 2026</h2>
    </div>
</div>
""", unsafe_allow_html=True)


import base64

def set_bg(image_file):
    with open(image_file, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url(data:image/png;base64,{encoded});
            background-size: cover;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

set_bg("bg.jpg")

if selected == "Overview":
    SHEET_ID = "1aL4qAd8MPbKXnvN_2aeQcmwHVhrApBgEwGxbY8vB7gI"
    SHEET_DATA = "Data"
    SHEET_MANPOWER = "Manpower"

    SCOPES = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    # ======================
    # AUTH GOOGLE
    # ======================
    @st.cache_resource
    def get_gspread_client():
        creds = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=SCOPES
        )
        return gspread.authorize(creds)

    gc = get_gspread_client()

    # ======================
    # HELPER: LOAD SHEET + FIX HEADER
    # ======================
    @st.cache_data(ttl=300)  # auto refresh tiap 5 menit
    def load_sheet(sheet_name: str) -> pd.DataFrame:
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet(sheet_name)

        values = ws.get_all_values()
        headers = values[0]
        rows = values[1:]

        # ---- FIX HEADER DUPLICATE ----
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

    # ======================
    # LOAD DATA
    # ======================
    df = load_sheet(SHEET_DATA)
    df_manpower = load_sheet(SHEET_MANPOWER)

    # contoh slicing kayak kode awal kamu
    data = df.iloc[:, 3:7]

    weeks = sorted(data["Week"].dropna().unique())
    st.markdown("""
    <style>
    div[data-baseweb="tag"] {
        background-color: black !important; 
        color: black !important;
        border: 1px solid black!important;
    }
    div[data-baseweb="tag"] svg {
        color: black !important;
    }
    div[data-baseweb="select"] > div {
        border: 1px solid black !important;
        border-radius: 8px !important;
    }
    </style>
    """, unsafe_allow_html=True)

    week1_start = date(2026, 1, 2)
    today = date.today()
    days_since = (today - week1_start).days
    current_week = (days_since // 7) + 1
    default_week = f"Week {current_week}"
    default_week = [default_week] if default_week in weeks else []


    a,b = st.columns([4,1])
    with b:
        week_filter = st.multiselect("Pilih Week", weeks, default=default_week,width=250)
        filtered_df = data[
        data["Week"].isin(week_filter)]
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
        [class="stColumn st-emotion-cache-1lq70ut e12zf7d52"] {
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
            text-align: center;
            diplay: inline-block;
        }
        [class="stColumn st-emotion-cache-1qlw70m e12zf7d52"] {
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
            text-align: center;
            diplay: inline-block;
        }
        .st-emotion-cache-l64bj3 {
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
        }
        .marks {
            background-color: white;
        }
        .st-emotion-cache-7czcpc > img {
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
        }
        </style>
        """, unsafe_allow_html=True)

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
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Manpower</h4>
                <div class="value">{len(df_manpower)}</div>
            </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Manpower On-Site</h4>
                <div class="value">{(df_manpower["On/Off Site"] == "On-Site").sum()}</div>
            </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Manpower Off-Site</h4>
                <div class="value">{(df_manpower["On/Off Site"] == "Off-Site").sum()}</div>
            </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
            <div class="metric-card">
                <h4>Total Kegiatan</h4>
                <div class="value">{sum(sum_kegiatan["Jumlah"])}</div>
            </div>
        """, unsafe_allow_html=True)

    # PREPARE CHART

    total_kegiatan = (
        sum_kegiatan.groupby("Kegiatan", as_index=False)["Jumlah"]
        .sum()
    )

    chart = (
        alt.Chart(sum_kegiatan)
        .mark_bar()
        .encode(
            y=alt.Y("Kegiatan:N", sort="-x", axis=alt.Axis(title=None, labelFontSize = 12)),
            x=alt.X("Jumlah:Q", axis=alt.Axis(title=None, labels=False)),
            color=alt.Color(
                "BU:N",
                scale=alt.Scale(
                    domain=["DCM", "HPAL", "ONC", "Lainnya"],
                    range=["#134f5c", "#2f9a7f", "#31681a", "#ff9900"]
                )
            )))

    

    labels = (
        alt.Chart(total_kegiatan)
        .mark_text(align="left", baseline="middle", dx=3)
        .encode(
            y="Kegiatan:N",
            x="Jumlah:Q",
            text=alt.Text("Jumlah:Q", format="0")
        )
    )
    final_chart = (
        (chart + labels)
        .properties(height=400)
        .configure_view(
            fill="white"          # area plot
        )
        .configure(
            background="white"    # seluruh canvas
        )
        .configure_legend(
            fillColor="white"
        )
    )

    # CONDITIONAL FORMATTING
    def highlight_offsite(row):
        if row["On/Off Site"].strip().lower() == "off-site":
            return ["background-color: #CECECE"] * len(row)
        return ["background-color: white"] * len(row)

    df_manpower = df_manpower.copy()
    df_manpower = df_manpower.reset_index(drop=True)   
    styled_df = (
        df_manpower.style
        .apply(highlight_offsite, axis=1)
        .hide(axis="index")
        .set_properties(**{
            "border": "1px solid #ddd",
            "text-align": "center",
            "padding": "6px",
            "font-size": "12px",
        })
    )

    st.divider()
    col1, col2 = st.columns([3,5], border=True)

    with col1:
        st.markdown("<h4 style='text-align: center;'>Manpowers</h4>", unsafe_allow_html=True)
        st.dataframe(styled_df, hide_index=True)
    with col2:
        container2 = st.container()
        container2.markdown("<div class='card-container'>", unsafe_allow_html=True)

        st.markdown("<h4 style='text-align: center;'>📈 Jumlah per Kegiatan</h4>", unsafe_allow_html=True)
        st.altair_chart(final_chart, use_container_width=True)
        container2.markdown("</div>", unsafe_allow_html=True)
        with st.expander("Overview"):
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
            st.dataframe(pivot_tabel, height=400)

elif selected == "Induksi":
    Induksi.app()

elif selected == "Training":
    TrainingA2B.app()

elif selected == "Complience Rate":
    Sharing.app()

elif selected == "Com/Re-com":
    Com_Recom.app()

elif selected == "Inspeksi & Observasi":
    Inspeksi.app()

elif selected == "SIMPER":
    SIMPER.app()

elif selected == "Tes Praktik":
    Tes_praktik.app()

elif selected == "Refresh":
    Refresh.app()

elif selected == "Briefing P5M":
    P5M.app()

elif selected == "Lainnya":
    Lainnya.app()



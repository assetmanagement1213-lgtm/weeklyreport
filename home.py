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


st.set_page_config(
    page_title="Dashboard Asset Management",
    page_icon="👷‍♂️",
    layout="wide"
)

USERNAME = st.secrets["auth"]["username"]
PASSWORD = st.secrets["auth"]["password"]

def login_page():
    st.markdown("""
    <style>
    .st-emotion-cache-tn0cau {
        max-width: 350px;
        margin: auto;
        padding: 30px;
        border-radius: 12px;
        border: 1px solid #ddd;
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        text-align: center;
        background-color : black;
        color : white;
    }
    .st-emotion-cache-y7k0q2 {
        color: black;}
    .st-emotion-cache-1o7eotc {
        color:white;}
    </style>
    """, unsafe_allow_html=True)

    st.markdown("## 🔐 Login")

    user = st.text_input("Username")
    pw = st.text_input("Password", type="password")

    if st.button("Login"):
        if user == USERNAME and pw == PASSWORD:
            st.session_state["logged_in"] = True
            st.rerun()
        else:
            st.error("❌ Username atau password salah")

    st.markdown("</div>", unsafe_allow_html=True)
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False

if not st.session_state["logged_in"]:
    login_page()
    st.stop()

st.markdown("""
    <style>
        section[data-testid="stSidebar"] {display: none;}
        div[data-testid="stSidebar"] {display: none;}

        .main .block-container {
            padding-left: 2rem;
            padding-right: 2rem;
        }
    </style>
""", unsafe_allow_html=True)

st.image("abc.png")

SHEET_ID = "1UCyov9SZzwCzruemj7eUCFpc_ONV9du3fio00K_JHtI"
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
/* Tag item terpilih (chip) */
div[data-baseweb="tag"] {
    background-color: black !important; /* ganti sesuai warna kamu */
    color: black !important;
    border: 1px solid black!important;
}

/* Text X (remove button) */
div[data-baseweb="tag"] svg {
    color: black !important;
}

/* Dropdown box border */
div[data-baseweb="select"] > div {
    border: 1px solid black !important;
    border-radius: 8px !important;
}
</style>
""", unsafe_allow_html=True)

week1_start = date(2024, 12, 27)  # Jumat
week1_end = date(2025, 1, 2)      # Kamis
today = date.today()

days_since_week1 = (today - week1_start).days
current_week_number = (days_since_week1 // 7) + 1

default_week = f"Week {current_week_number}"

tab1, tab2 = st.tabs(["📑 Overview", "👷‍♂️ Activity"])

##OVERVIEW
with tab1:
    a,b = st.columns([4,1])
    with a:
        st.subheader("📊 Weekly Overview")
    with b:
        week_filter = st.multiselect("Pilih Week", weeks, default=[default_week],width=250)
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
        .st-emotion-cache-codo9b{
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
            text-align: center;
            diplay: inline-block;
        }
        .st-emotion-cache-1o7eotc{
            background-color: white;
            box-shadow: 0 3px 8px rgba(0,0,0,0.08);
        }
        .marks {
            background-color: white;
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
            y=alt.Y("Kegiatan:N", sort="-x", axis=alt.Axis(title=None)),
            x=alt.X("Jumlah:Q", axis=alt.Axis(title=None, labels=False, grid=False)),
            color=alt.Color(
                "BU:N",
                scale=alt.Scale(
                    domain=["DCM", "HPAL", "ONC", "Lainnya"],
                    range=["#134f5c", "#2f9a7f", "#31681a", "#ff9900"]
                )
            )
        )
    )

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
        st.markdown("<h3 style='text-align: center;'>👬🏼 Manpowers</h3>", unsafe_allow_html=True)
        st.dataframe(styled_df, hide_index=True)
    with col2:
        container2 = st.container()
        container2.markdown("<div class='card-container'>", unsafe_allow_html=True)

        st.markdown("<h3 style='text-align: center;'>📈 Jumlah per Kegiatan</h3>", unsafe_allow_html=True)
        st.altair_chart(final_chart, use_container_width=True)
        container2.markdown("</div>", unsafe_allow_html=True)
with tab2:
    st.markdown("""
    <style>
    .st-key-styledradio .stRadio p {
        font-size: 14px;
        padding : 5px;
    }
    .st-key-styledradio {
    margin-top: 20px;
    }
    .column {
        background: white;
        padding: 15px;
        border-radius: 12px;
        border: 1px solid #dfdfdf;
        box-shadow: 0px 3px 8px rgba(0,0,0,0.12);
        margin-bottom: 10px;
    }
    </style>
    """, unsafe_allow_html=True)
    
    q, w = st.columns([1, 6], border=True)
    import training
    import induksi
    
    with q:
        menu = st.radio(
            "Pilih Activity",
            ["Induksi", "Training", "Sharing Knowledge","Commissining/Recommissioning", "Inspeksi", "Observasi", "SIMPER", "Tes Praktik", "Refresh", "Pembekalan", "Briefing P5M"],
            index=0,
            key="styledradio"
        )
        st.markdown(
            """<style>
        div[class*="stRadio"] > label > div[data-testid="stMarkdownContainer"] > p {
            font-size: 20px; font-weight: bold;
        }
            </style>
            """, unsafe_allow_html=True)   

    with w:
        if menu == "Induksi":
            induksi.app()

        elif menu == "Training":
            training.app()
    st.markdown('</div>', unsafe_allow_html=True)

if st.button("🔄 Refresh Data"):
    st.cache_data.clear()
st.divider()
st.markdown(
    """
    <p style="text-align: left; color: gray; font-size: 13px;">
        Asset Management 2025
    </p>
    """,
    unsafe_allow_html=True)

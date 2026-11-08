import streamlit as st
import pandas as pd
import openpyxl
import requests
import re
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
import io

# Page Config
st.set_page_config(
    page_title="ERP Route Planner Engine",
    page_icon="🚚",
    layout="wide"
)

# Base Depot Settings
BASE_DEPOT = "Paras Nutritions Pvt Ltd, Road, Moga, Punjab, India"
BASE_LAT = 30.8165
BASE_LON = 75.1717

# -----------------------------------------------------------------------------
# DYNAMIC REGEX & LIVE API RESOLUTION ENGINES
# -----------------------------------------------------------------------------

def clean_vpo_name(party_name):
    """Clean ERP party names dynamically using RegEx."""
    if not isinstance(party_name, str) or not party_name.strip():
        return ""
    
    txt = party_name.strip()
    # Strip noise patterns: LF-9707, AG 473, R-12, RT15, /142044, etc.
    txt = re.sub(r'[\,\s\/]+(LF[-\s]*\d+|AG\s*\d+|R[-]*\d+|RT\d+|\d{4,}).*$', '', txt, flags=re.IGNORECASE)
    
    pos_vpo = re.search(r'VPO[-.\s]*', txt, flags=re.IGNORECASE)
    if pos_vpo:
        raw_vpo = txt[pos_vpo.end():].strip('.')
    elif ',' in txt:
        raw_vpo = txt.split(',')[0]
    elif '/' in txt:
        raw_vpo = txt.split('/')[0]
    else:
        raw_vpo = txt
        
    words = raw_vpo.split()
    clean_words = []
    for w in words:
        w_upper = w.upper().strip()
        if w_upper in ["AG", "LF", "R", "NO"] or w_upper.startswith("R-") or w_upper.isdigit():
            break
        clean_words.append(w)
        
    res = " ".join(clean_words).strip()
    return res.title() if res else party_name.title()


def get_live_geocoding_and_km(vpo_name, party_text, api_key=""):
    """
    Fetch exact District Name and Driving KM using Google Maps API (if provided)
    or Free OpenStreetMap/OSRM Engine (Fallback).
    """
    if not vpo_name:
        return "Unknown", "N/A", ""

    district_name = "Punjab"
    driving_km = "N/A"
    
    # Check text explicit district after comma
    if ',' in party_text:
        parts = party_text.split(',')
        if len(parts) > 1:
            cand = parts[1].strip()
            if "LF" not in cand.upper() and "R-" not in cand.upper() and len(cand) > 2:
                district_name = cand.split()[0].title()

    # METHOD A: Google Maps API Execution (If API Key provided)
    if api_key and api_key.strip():
        try:
            # 1. Geocode District
            geo_url = f"https://maps.googleapis.com/maps/api/geocode/json?address={urllib.parse.quote(vpo_name + ', Punjab, India')}&key={api_key}"
            g_resp = requests.get(geo_url, timeout=3).json()
            if g_resp.get("status") == "OK" and g_resp.get("results"):
                for component in g_resp["results"][0].get("address_components", []):
                    if "administrative_area_level_2" in component.get("types", []):
                        district_name = component["long_name"].replace("District", "").strip().title()
                        break
            
            full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
            
            # 2. Distance Matrix
            dist_url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={urllib.parse.quote(BASE_DEPOT)}&destinations={urllib.parse.quote(full_addr)}&mode=driving&key={api_key}"
            d_resp = requests.get(dist_url, timeout=3).json()
            if d_resp.get("status") == "OK" and d_resp.get("rows"):
                elements = d_resp["rows"][0].get("elements", [])
                if elements and elements[0].get("status") == "OK":
                    driving_km = elements[0]["distance"]["text"].upper()
                    
            maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
            return district_name, driving_km, maps_link
        except Exception:
            pass

    # METHOD B: Free OpenStreetMap & OSRM Engine (High-Speed Multi-Threaded Engine)
    try:
        osm_url = f"https://nominatim.openstreetmap.org/search?q={urllib.parse.quote(vpo_name + ', Punjab, India')}&format=json&addressdetails=1&limit=1"
        headers = {'User-Agent': 'StreamlitRoutePlanner/1.0'}
        res = requests.get(osm_url, headers=headers, timeout=2).json()
        
        if res and len(res) > 0:
            addr = res[0].get("address", {})
            district_name = addr.get("county", addr.get("state_district", district_name)).replace("District", "").strip().title()
            
            dest_lat = res[0]["lat"]
            dest_lon = res[0]["lon"]
            
            osrm_url = f"http://router.project-osrm.org/route/v1/driving/{BASE_LON},{BASE_LAT};{dest_lon},{dest_lat}?overview=false"
            osrm_res = requests.get(osrm_url, timeout=2).json()
            if osrm_res.get("routes"):
                meters = osrm_res["routes"][0]["distance"]
                driving_km = f"{round(meters / 1000.0, 1)} KM"
    except Exception:
        pass

    full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
    maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
    return district_name, driving_km, maps_link


def process_single_sheet(sheet_name, df, api_key=""):
    """Fast Parallel Processing for Individual Route Sheet."""
    # Find Party Name Column
    party_col = None
    header_idx = None
    
    for idx, row in df.head(10).iterrows():
        for col in df.columns:
            val = str(row[col]).upper()
            if "SOLD TO PARTY NAME" in val or "FARMER NAME" in val or "SOLD-TO PARTY NAME" in val:
                party_col = col
                header_idx = idx
                break
        if party_col is not None:
            break

    if party_col is None:
        return sheet_name, df, None

    # Filter Valid Data Rows
    data_df = df.iloc[header_idx + 1:].copy() if header_idx is not None else df.copy()
    
    unique_parties = data_df[party_col].dropna().unique()
    cache = {}

    def fetch_task(party):
        if not party or str(party).upper() == "TOTAL":
            return
        clean_vpo = clean_vpo_name(str(party))
        dist, km, link = get_live_geocoding_and_km(clean_vpo, str(party), api_key)
        cache[party] = {
            "clean_vpo": clean_vpo,
            "district": dist,
            "km": km,
            "link": link
        }

    # Parallel Execution Threads (Fast 10x Processing)
    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(fetch_task, unique_parties)

    # Append Output Columns
    vpo_list, dist_list, addr_list, km_list, link_list = [], [], [], [], []

    for _, row in data_df.iterrows():
        party = row[party_col]
        if party in cache:
            c = cache[party]
            vpo_list.append(c["clean_vpo"])
            dist_list.append(c["district"])
            addr_list.append(f"{c['clean_vpo']}, District {c['district']}, Punjab, India")
            km_list.append(c["km"])
            link_list.append(c["link"])
        else:
            vpo_list.append("")
            dist_list.append("")
            addr_list.append("")
            km_list.append("")
            link_list.append("")

    data_df["Extracted Village (VPO)"] = vpo_list
    data_df["District Name"] = dist_list
    data_df["Exact Google Maps Address"] = addr_list
    data_df["Dist from Base (KM)"] = km_list
    data_df["Google Maps Directions Link"] = link_list

    return sheet_name, data_df, cache

# -----------------------------------------------------------------------------
# STREAMLIT UI DASHBOARD
# -----------------------------------------------------------------------------

st.title("🚚 ERP Route Sheet Highway Planner Engine")
st.caption("Powered by Paras Nutritions Pvt Ltd | Multi-Threaded Dynamic Processing Engine")

st.sidebar.header("⚙️ Configuration Settings")
api_key_input = st.sidebar.text_input("Google Maps API Key (Optional):", type="password", help="Leave blank to use Free OpenStreetMap/OSRM Routing Engine")

uploaded_file = st.file_uploader("Upload ERP Excel File (.xlsx or .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    st.success("File Uploaded Successfully! Click Process to start.")
    
    if st.button("🚀 Process All Route Sheets Now", type="primary"):
        with st.spinner("Processing Route Sheets dynamically across live Highway Routing API..."):
            excel_file = pd.ExcelFile(uploaded_file)
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_count = 0
                
                for sheet in excel_file.sheet_names:
                    sheet_upper = sheet.strip().upper()
                    if "ROUTE" in sheet_upper and sheet_upper not in ["ROUTE INDEX", "ROUTE #"]:
                        df = pd.read_excel(uploaded_file, sheet_name=sheet)
                        s_name, processed_df, _ = process_single_sheet(sheet, df, api_key_input)
                        processed_df.to_excel(writer, sheet_name=s_name, index=False)
                        processed_count += 1
                    else:
                        # Copy original non-route sheets as is
                        df = pd.read_excel(uploaded_file, sheet_name=sheet)
                        df.to_excel(writer, sheet_name=sheet, index=False)

            output.seek(0)
            
            st.balloons()
            st.success(f"✅ Processed {processed_count} Route Sheets successfully in seconds!")
            
            st.download_button(
                label="📥 Download Updated ERP Excel Report",
                data=output,
                file_name="Updated_Route_Plan_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

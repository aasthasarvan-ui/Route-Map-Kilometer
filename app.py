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
    page_title="ERP Route Sheet Highway Planner",
    page_icon="🚚",
    layout="wide"
)

# Base Origin Depot (Paras Nutritions Pvt Ltd, Moga)
BASE_DEPOT = "Paras Nutritions Pvt Ltd, Road, Moga, Punjab, India"
BASE_LAT = 30.8165
BASE_LON = 75.1717

# -----------------------------------------------------------------------------
# PURE DYNAMIC LOGIC & LIVE GEOCODING ENGINES (ZERO HARDCODED DATA)
# -----------------------------------------------------------------------------

def clean_vpo_name_dynamic(party_name):
    """
    Pure Dynamic RegEx Parser to extract clean VPO name from complex ERP strings.
    Strips noise like LF-9707, AG 473, R-12, RT15, farmer names, etc.
    """
    if not isinstance(party_name, str) or not party_name.strip():
        return ""
    
    txt = party_name.strip()
    
    # 1. Remove ERP System Noise codes
    txt = re.sub(r'[\,\s\/]+(LF[-\s]*\d+|AG\s*\d+|R[-]*\d+|RT\d+|\d{4,}).*$', '', txt, flags=re.IGNORECASE)
    
    # 2. Extract after VPO / V.P.O. if present
    pos_vpo = re.search(r'VPO[-.\s]*', txt, flags=re.IGNORECASE)
    if pos_vpo:
        raw_vpo = txt[pos_vpo.end():].strip('.')
    elif ',' in txt:
        raw_vpo = txt.split(',')[0]
    elif '/' in txt:
        raw_vpo = txt.split('/')[0]
    else:
        raw_vpo = txt
        
    # 3. Dynamic word filtering (Removes code suffixes)
    words = raw_vpo.split()
    clean_words = []
    for w in words:
        w_upper = w.upper().strip()
        if w_upper in ["AG", "LF", "R", "NO"] or w_upper.startswith("R-") or w_upper.isdigit():
            break
        clean_words.append(w)
        
    res = " ".join(clean_words).strip()
    return res.title() if res else party_name.title()


def dynamic_live_geocoding_engine(vpo_name, party_text, api_key=""):
    """
    100% Dynamic Engine (No prefilled arrays).
    Fetches real-time District Name and exact driving KM using Google Maps API 
    or OpenStreetMap + OSRM Live Highway Router.
    """
    if not vpo_name:
        return "Unknown", "N/A", ""

    district_name = "Punjab"
    driving_km = "N/A"

    # METHOD A: Live Google Maps API (If API Key provided in sidebar)
    if api_key and api_key.strip():
        try:
            geo_url = f"https://maps.googleapis.com/maps/api/geocode/json?address={urllib.parse.quote(vpo_name + ', Punjab, India')}&key={api_key}"
            g_resp = requests.get(geo_url, timeout=4).json()
            
            if g_resp.get("status") == "OK" and g_resp.get("results"):
                for component in g_resp["results"][0].get("address_components", []):
                    types = component.get("types", [])
                    if "administrative_area_level_2" in types:
                        district_name = component["long_name"].replace("District", "").strip().title()
                        break
            
            full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
            
            dist_url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={urllib.parse.quote(BASE_DEPOT)}&destinations={urllib.parse.quote(full_addr)}&mode=driving&key={api_key}"
            d_resp = requests.get(dist_url, timeout=4).json()
            if d_resp.get("status") == "OK" and d_resp.get("rows"):
                elements = d_resp["rows"][0].get("elements", [])
                if elements and elements[0].get("status") == "OK":
                    driving_km = elements[0]["distance"]["text"].upper()
                    
            maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
            return district_name, driving_km, maps_link
        except Exception:
            pass

    # METHOD B: Live OpenStreetMap Geocoding + Live OSRM Highway Engine (Pure Free API)
    try:
        osm_url = f"https://nominatim.openstreetmap.org/search?q={urllib.parse.quote(vpo_name + ', Punjab, India')}&format=json&addressdetails=1&limit=1"
        headers = {'User-Agent': 'PureDynamicERPRoutePlanner/4.0'}
        res = requests.get(osm_url, headers=headers, timeout=3).json()
        
        if res and len(res) > 0:
            addr = res[0].get("address", {})
            
            # Dynamic District Property Extraction Logic
            raw_d = addr.get("county", addr.get("state_district", addr.get("district", "")))
            if raw_d:
                district_name = raw_d.replace("District", "").replace("Tehsil", "").strip().title()
            
            dest_lat = float(res[0]["lat"])
            dest_lon = float(res[0]["lon"])
            
            # Real Highway Driving Distance calculation via OSRM Web Engine
            osrm_url = f"http://router.project-osrm.org/route/v1/driving/{BASE_LON},{BASE_LAT};{dest_lon},{dest_lat}?overview=false"
            osrm_res = requests.get(osrm_url, timeout=3).json()
            if osrm_res.get("routes"):
                meters = osrm_res["routes"][0]["distance"]
                driving_km = f"{round(meters / 1000.0, 1)} KM"
    except Exception:
        pass

    # Dynamic fallback string format if OSM search fails for very small village
    full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
    maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
    return district_name, driving_km, maps_link


def process_single_sheet(sheet_name, df, api_key=""):
    """Multi-Threaded Dynamic Processor for Route Sheets."""
    party_col = None
    header_idx = None
    
    # Dynamic Column Auto-Detection Logic
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
        return sheet_name, df

    data_df = df.iloc[header_idx + 1:].copy() if header_idx is not None else df.copy()
    unique_parties = data_df[party_col].dropna().unique()
    cache = {}

    def fetch_task(party):
        if not party or str(party).upper() == "TOTAL":
            return
        clean_vpo = clean_vpo_name_dynamic(str(party))
        dist, km, link = dynamic_live_geocoding_and_km(clean_vpo, str(party), api_key)
        cache[party] = {
            "clean_vpo": clean_vpo,
            "district": dist,
            "km": km,
            "link": link
        }

    # Parallel Processing Threads for Ultra-Fast Batching
    with ThreadPoolExecutor(max_workers=10) as executor:
        executor.map(fetch_task, unique_parties)

    vpo_list, dist_list, addr_list, km_list, link_list = [], [], [], [], []

    for _, row in data_df.iterrows():
        party = row[party_col]
        if party in cache:
            c = cache[party]
            vpo_list.append(c["clean_vpo"])
            dist_list.append(c["district"])
            
            full_address_text = f"{c['clean_vpo']}, District {c['district']}, Punjab, India"
            addr_list.append(full_address_text)
            km_list.append(c["km"])
            
            # Excel Native Hyperlink Formula
            excel_hyperlink = f'=HYPERLINK("{c["link"]}", "📍 Open Route Map ({c["km"]})")'
            link_list.append(excel_hyperlink)
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

    return sheet_name, data_df

# -----------------------------------------------------------------------------
# STREAMLIT UI DASHBOARD
# -----------------------------------------------------------------------------

st.title("🚚 ERP Route Sheet Highway Planner Engine")
st.caption("Pure Dynamic Geocoding Architecture — 0% Hardcoded Data")

st.sidebar.header("Settings")
api_key_input = st.sidebar.text_input("Google Maps API Key (Optional):", type="password", help="Enter key for Google Maps driving routes or leave blank for Free OpenStreetMap Routing Engine")

uploaded_file = st.file_uploader("Upload ERP Excel File (.xlsx or .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    if st.button("🚀 Process All Route Sheets Now", type="primary"):
        with st.spinner("Processing Route Sheets dynamically across Live Geocoding API..."):
            excel_file = pd.ExcelFile(uploaded_file)
            output = io.BytesIO()
            
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            for sheet in excel_file.sheet_names:
                sheet_upper = sheet.strip().upper()
                df = pd.read_excel(uploaded_file, sheet_name=sheet)
                
                if "ROUTE" in sheet_upper and sheet_upper not in ["ROUTE INDEX", "ROUTE #"]:
                    s_name, processed_df = process_single_sheet(sheet, df, api_key_input)
                else:
                    s_name, processed_df = sheet, df
                    
                ws = wb.create_sheet(title=s_name)
                ws.append(list(processed_df.columns))
                
                for row in processed_df.itertuples(index=False):
                    ws.append(list(row))

            wb.save(output)
            output.seek(0)
            
            st.success("✅ Complete! All ERP sheets dynamically processed.")
            st.download_button(
                label="📥 Download Updated Excel Sheet",
                data=output,
                file_name="Updated_Route_Plan_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

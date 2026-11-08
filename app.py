import streamlit as st
import pandas as pd
import openpyxl
import requests
import re
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
import io

# Page Configuration
st.set_page_config(
    page_title="ERP Route Sheet Highway Planner Engine",
    page_icon="🚚",
    layout="wide"
)

# Base Origin Depot (Paras Nutritions Pvt Ltd, Moga)
BASE_DEPOT = "Paras Nutritions Pvt Ltd, Road, Moga, Punjab, India"
BASE_LAT = 30.8165
BASE_LON = 75.1717

# -----------------------------------------------------------------------------
# PURE DYNAMIC GEOCODING & BOUNDED SPATIAL LOGIC (ZERO HARDCODED DATA)
# -----------------------------------------------------------------------------

def extract_pure_village_name_dynamic(cell_value):
    """
    Pure Dynamic RegEx Parser to clean Column H text.
    Strips noise like LF-9707, AG 473, R-12, RT15, farmer names, etc.
    """
    if not isinstance(cell_value, str) or not cell_value.strip():
        return ""
    
    txt = cell_value.strip()
    
    # 1. Strip ERP System Noise codes (e.g. LF-9707, AG 473, R-12, RT15, numbers)
    txt = re.sub(r'[\,\s\/]+(LF[-\s]*\d+|AG\s*\d+|R[-]*\d+|RT\d+|\d{4,}).*$', '', txt, flags=re.IGNORECASE)
    
    # 2. Extract text after VPO / V.P.O. keyword if present
    pos_vpo = re.search(r'VPO[-.\s]*', txt, flags=re.IGNORECASE)
    if pos_vpo:
        raw_vpo = txt[pos_vpo.end():].strip('.- ')
    elif ',' in txt:
        raw_vpo = txt.split(',')[0]
    elif '/' in txt:
        raw_vpo = txt.split('/')[0]
    else:
        raw_vpo = txt
        
    # 3. Dynamic word filtering (Removes noise suffixes)
    words = raw_vpo.split()
    clean_words = []
    for w in words:
        w_upper = w.upper().strip()
        if w_upper in ["AG", "LF", "R", "NO"] or w_upper.startswith("R-") or w_upper.isdigit():
            break
        clean_words.append(w)
        
    res = " ".join(clean_words).strip()
    return res.title() if res else cell_value.title()


def live_bounded_spatial_geocoder(vpo_name, full_cell_text, api_key=""):
    """
    100% Dynamic Engine (Zero Hardcoding).
    Uses Bounding Box Scoped Queries for Punjab to ensure accurate District 
    and Live Highway Driving KM resolution.
    """
    if not vpo_name:
        return "Unknown", "N/A", ""

    district_name = "Punjab"
    driving_km = "N/A"

    # Scoped Query String specifically targetting Punjab locations
    search_query = f"{vpo_name}, Punjab, India"

    # METHOD A: Live Google Maps API (If API Key provided)
    if api_key and api_key.strip():
        try:
            geo_url = f"https://maps.googleapis.com/maps/api/geocode/json?address={urllib.parse.quote(search_query)}&key={api_key}"
            g_resp = requests.get(geo_url, timeout=5).json()
            
            if g_resp.get("status") == "OK" and g_resp.get("results"):
                for component in g_resp["results"][0].get("address_components", []):
                    types = component.get("types", [])
                    if "administrative_area_level_2" in types:
                        district_name = component["long_name"].replace("District", "").strip().title()
                        break
            
            full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
            dist_url = f"https://maps.googleapis.com/maps/api/distancematrix/json?origins={urllib.parse.quote(BASE_DEPOT)}&destinations={urllib.parse.quote(full_addr)}&mode=driving&key={api_key}"
            d_resp = requests.get(dist_url, timeout=5).json()
            
            if d_resp.get("status") == "OK" and d_resp.get("rows"):
                elements = d_resp["rows"][0].get("elements", [])
                if elements and elements[0].get("status") == "OK":
                    driving_km = elements[0]["distance"]["text"].upper()
                    
            maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
            return district_name, driving_km, maps_link
        except Exception:
            pass

    # METHOD B: Live Bounded OpenStreetMap Geocoding + Live OSRM Router
    try:
        # Bounding box coordinates added for Malwa Punjab belt to prevent wrong state hits
        osm_url = f"https://nominatim.openstreetmap.org/search?q={urllib.parse.quote(search_query)}&format=json&addressdetails=1&viewbox=74.0,29.5,76.5,31.5&bounded=1&limit=1"
        headers = {'User-Agent': 'BoundedDynamicERPRoutePlanner/12.0'}
        res = requests.get(osm_url, headers=headers, timeout=4).json()
        
        # Fallback without strict bounds if bounded viewbox returns empty
        if not res:
            osm_url = f"https://nominatim.openstreetmap.org/search?q={urllib.parse.quote(search_query)}&format=json&addressdetails=1&limit=1"
            res = requests.get(osm_url, headers=headers, timeout=4).json()
            
        if res and len(res) > 0:
            addr = res[0].get("address", {})
            raw_d = addr.get("county", addr.get("state_district", addr.get("district", "")))
            if raw_d and "Tehsil" not in raw_d:
                district_name = raw_d.replace("District", "").strip().title()
            
            dest_lat = float(res[0]["lat"])
            dest_lon = float(res[0]["lon"])
            
            # Live OSRM Highway Engine
            osrm_url = f"http://router.project-osrm.org/route/v1/driving/{BASE_LON},{BASE_LAT};{dest_lon},{dest_lat}?overview=false"
            osrm_res = requests.get(osrm_url, timeout=4).json()
            if osrm_res.get("routes"):
                meters = osrm_res["routes"][0]["distance"]
                driving_km = f"{round(meters / 1000.0, 1)} KM"
    except Exception:
        pass

    full_addr = f"{vpo_name}, District {district_name}, Punjab, India"
    maps_link = f"https://www.google.com/maps/dir/?api=1&origin={urllib.parse.quote(BASE_DEPOT)}&destination={urllib.parse.quote(full_addr)}&travelmode=driving"
    return district_name, driving_km, maps_link


def process_single_sheet_strict_h(sheet_name, df_raw, api_key=""):
    """
    STRICT COLUMN H BASE PROCESSOR:
    Extracts Party / Village Names STRICTLY from Column H (Index 7) with 0% hardcoded data.
    """
    if df_raw.empty or df_raw.shape[1] < 8:
        return sheet_name, df_raw

    party_col_idx = 7 # Column H (0-indexed)
    header_row_idx = None
    
    # Locate Header Row in Column H
    for r_idx in range(min(25, len(df_raw))):
        val = str(df_raw.iloc[r_idx, party_col_idx]).upper().strip()
        if any(k in val for k in ["SOLD TO PARTY", "FARMER NAME", "PARTY NAME", "CUSTOMER", "SHIP TO"]):
            header_row_idx = r_idx
            break
            
    if header_row_idx is None:
        header_row_idx = 3

    data_df = df_raw.copy()
    num_rows = len(data_df)
    
    vpo_list = [""] * num_rows
    dist_list = [""] * num_rows
    addr_list = [""] * num_rows
    km_list = [""] * num_rows
    link_list = [""] * num_rows

    # Column Headers Setup
    vpo_list[header_row_idx] = "Extracted Village (VPO)"
    dist_list[header_row_idx] = "District Name"
    addr_list[header_row_idx] = "Exact Google Maps Address"
    km_list[header_row_idx] = "Dist from Base (KM)"
    link_list[header_row_idx] = "Google Maps Directions Link"

    # Extract Unique Parties strictly from Column H
    row_party_map = {}
    unique_parties = set()

    for r_idx in range(header_row_idx + 1, num_rows):
        val = data_df.iloc[r_idx, party_col_idx]
        if pd.notna(val):
            p_str = str(val).strip()
            if p_str and p_str.upper() not in ["TOTAL", "SUBTOTAL", "NAN", "NONE", "HEADER", "SOLD TO PARTY NAME"]:
                row_party_map[r_idx] = p_str
                unique_parties.add(p_str)

    cache = {}

    def fetch_task(party):
        clean_vpo = extract_pure_village_name_dynamic(party)
        if clean_vpo:
            dist, km, link = live_bounded_spatial_geocoder(clean_vpo, party, api_key)
            cache[party] = {
                "clean_vpo": clean_vpo,
                "district": dist,
                "km": km,
                "link": link
            }

    # Parallel Geocoding Threads
    if unique_parties:
        with ThreadPoolExecutor(max_workers=10) as executor:
            executor.map(fetch_task, list(unique_parties))

    # Populate Rows Directly
    for r_idx, party_val in row_party_map.items():
        if party_val in cache:
            c = cache[party_val]
            vpo_list[r_idx] = c["clean_vpo"]
            dist_list[r_idx] = c["district"]
            addr_list[r_idx] = f"{c['clean_vpo']}, District {c['district']}, Punjab, India"
            km_list[r_idx] = c["km"]
            link_list[r_idx] = f'=HYPERLINK("{c["link"]}", "📍 Open Route Map ({c["km"]})")'

    data_df["Extracted Village (VPO)"] = vpo_list
    data_df["District Name"] = dist_list
    data_df["Exact Google Maps Address"] = addr_list
    data_df["Dist from Base (KM)"] = km_list
    data_df["Google Maps Directions Link"] = link_list

    return sheet_name, data_df

# -----------------------------------------------------------------------------
# STREAMLIT USER INTERFACE
# -----------------------------------------------------------------------------

st.title("🚚 ERP Route Sheet Highway Planner Engine")
st.caption("Bounded Spatial Dynamic Logic Engine — 0% Hardcoded Data")

st.sidebar.header("Configuration")
api_key_input = st.sidebar.text_input("Google Maps API Key (Optional):", type="password", help="Add key for Google Maps live matrix or leave blank for Bounded OpenStreetMap Engine")

uploaded_file = st.file_uploader("Upload ERP Excel File (.xlsx or .xlsm)", type=["xlsx", "xlsm"])

if uploaded_file:
    if st.button("🚀 Process All Route Sheets Now", type="primary"):
        with st.spinner("Processing ERP Sheets dynamically across Bounded Highway API..."):
            excel_file = pd.ExcelFile(uploaded_file)
            output = io.BytesIO()
            
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            
            for sheet in excel_file.sheet_names:
                sheet_upper = sheet.strip().upper()
                df_raw = pd.read_excel(uploaded_file, sheet_name=sheet, header=None)
                
                if "ROUTE" in sheet_upper and sheet_upper not in ["ROUTE INDEX", "ROUTE #"]:
                    s_name, processed_df = process_single_sheet_strict_h(sheet, df_raw, api_key_input)
                else:
                    s_name, processed_df = sheet, df_raw
                    
                ws = wb.create_sheet(title=s_name)
                
                for row in processed_df.itertuples(index=False):
                    row_vals = [None if pd.isna(val) else val for val in row]
                    ws.append(row_vals)

            wb.save(output)
            output.seek(0)
            
            st.success("✅ Complete! Bounded spatial processing finished.")
            st.download_button(
                label="📥 Download Updated Excel Sheet",
                data=output,
                file_name="Updated_Route_Plan_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

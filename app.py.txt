import streamlit as st
import pandas as pd
import folium
from streamlit_folium import st_folium
from streamlit_js_eval import get_geolocation
import os
from datetime import datetime, date
from PIL import Image
import numpy as np

# -----------------------------------------------------------------------------
# Configuration & Setup
# -----------------------------------------------------------------------------
st.set_page_config(page_title="FMCG Live Survey Portal", layout="wide")

MASTER_FILE_PATH = "master_questions.xlsx"
SURVEY_FILE = "survey_responses.xlsx"
IMAGE_FOLDER = "survey_images"

LOGO_FILES = ["drishti_logo.png", "logo.png", "drishti.png", "drishti_logo.jpg", "logo.jpg"]
LOGO_PATH = None
for f in LOGO_FILES:
    if os.path.exists(f):
        LOGO_PATH = f
        break

if not os.path.exists(IMAGE_FOLDER):
    os.makedirs(IMAGE_FOLDER)

if "surveys_completed" not in st.session_state:
    st.session_state.surveys_completed = set()

if "selected_outlet" not in st.session_state:
    st.session_state.selected_outlet = None

if "form_step" not in st.session_state:
    st.session_state.form_step = 1

if "temp_form_data" not in st.session_state:
    st.session_state.temp_form_data = {}

if "order_item_count" not in st.session_state:
    st.session_state.order_item_count = 1

if "just_submitted" not in st.session_state:
    st.session_state.just_submitted = False

if "last_submitted_outlet" not in st.session_state:
    st.session_state.last_submitted_outlet = ""

# -----------------------------------------------------------------------------
# Fast Caching & Data Helpers
# -----------------------------------------------------------------------------
@st.cache_data(ttl=2)
def load_survey_data():
    if os.path.exists(SURVEY_FILE):
        try:
            return pd.read_excel(SURVEY_FILE)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()

@st.cache_data
def load_excel_sheets(file):
    excel_file = pd.ExcelFile(file)
    sheets_dict = {}
    for sheet in excel_file.sheet_names:
        sheets_dict[sheet] = pd.read_excel(excel_file, sheet_name=sheet)
    return excel_file.sheet_names, sheets_dict

saved_responses_df = load_survey_data()
if not saved_responses_df.empty:
    col_out = next((c for c in saved_responses_df.columns if "OUTLET" in c.upper()), None)
    if col_out:
        st.session_state.surveys_completed = set(saved_responses_df[col_out].dropna().astype(str).tolist())

def save_survey_data(data_dict):
    df_new = pd.DataFrame([data_dict])
    if os.path.exists(SURVEY_FILE):
        try:
            df_old = pd.read_excel(SURVEY_FILE)
            df_final = pd.concat([df_old, df_new], ignore_index=True)
        except Exception:
            df_final = df_new
    else:
        df_final = df_new
    df_final.to_excel(SURVEY_FILE, index=False)

def calculate_distance_km(lat1, lon1, lat2, lon2):
    try:
        R = 6371.0
        lat1, lon1, lat2, lon2 = map(np.radians, [float(lat1), float(lon1), float(lat2), float(lon2)])
        dlat = lat2 - lat1
        dlon = lon2 - lon1
        a = np.sin(dlat / 2.0)**2 + np.cos(lat1) * np.cos(lat2) * np.sin(dlon / 2.0)**2
        c = 2 * np.arcsin(np.sqrt(a))
        return R * c
    except Exception:
        return 9999.0

st.markdown("""
    <style>
        .block-container { padding-top: 1.2rem !important; padding-bottom: 1rem !important; }
        [data-testid="stSidebar"] [data-testid="stVerticalBlock"] { gap: 0.8rem; }
    </style>
""", unsafe_allow_html=True)

try:
    params = st.query_params
    mode_val = params.get("mode", None)
except Exception:
    params = st.experimental_get_query_params()
    mode_val = params.get("mode", [None])[0]

is_surveyor_mode = (mode_val == "surveyor")
uploaded_master = None

if is_surveyor_mode:
    st.markdown("""
        <style>
            [data-testid="stSidebar"] {display: none !important;}
            [data-testid="collapsedControl"] {display: none !important;}
            header {visibility: hidden !important;}
            .block-container {padding-top: 0.8rem !important;}
        </style>
    """, unsafe_allow_html=True)
    app_mode = "📱 Surveyor Mode (Mobile)"
else:
    with st.sidebar:
        if LOGO_PATH:
            st.image(LOGO_PATH, width=70)
        else:
            st.title("👁️ Drishti")
            
        st.title("🎛️ Admin Control Panel")
        app_mode = st.radio(
            "Select Mode:",
            ["🖥️ Admin Dashboard (Laptop Live)", "📱 Surveyor Mode Preview"]
        )
        uploaded_master = st.file_uploader(
            "📂 Update Master Excel File (.xlsx)", 
            type=["xlsx"], 
            key="admin_master_uploader"
        )

loc = get_geolocation()
surveyor_lat, surveyor_lon = None, None

if loc and 'coords' in loc:
    surveyor_lat = float(loc['coords']['latitude'])
    surveyor_lon = float(loc['coords']['longitude'])
    if not is_surveyor_mode:
        st.sidebar.success(f"GPS Active: {surveyor_lat:.4f}, {surveyor_lon:.4f}")

active_file = None
if uploaded_master:
    active_file = uploaded_master
elif os.path.exists(MASTER_FILE_PATH):
    active_file = MASTER_FILE_PATH

if active_file:
    try:
        sheet_names, sheets_dict = load_excel_sheets(active_file)

        he_sheet_name = next((s for s in sheet_names if "HE" in s or "Hunter" in s), sheet_names[0])
        outlets_sheet_name = next((s for s in sheet_names if "Group" in s or "Outlet" in s), sheet_names[1] if len(sheet_names)>1 else sheet_names[0])

        df_he = sheets_dict[he_sheet_name].copy()
        df_outlets = sheets_dict[outlets_sheet_name].copy()

        df_he_columns = [str(c).strip() for c in df_he.columns]
        df_outlets.columns = df_outlets.columns.str.strip()

        master_dropdowns = {}
        for s_name in sheet_names:
            s_df = sheets_dict[s_name].copy()
            s_df.columns = s_df.columns.str.strip()
            for col in s_df.columns:
                clean_col = col.strip()
                vals = s_df[col].dropna().astype(str).str.strip().tolist()
                if vals:
                    master_dropdowns[clean_col] = vals
                    master_dropdowns[s_name.strip()] = vals

        hunter_col = next((c for c in df_outlets.columns if "HUNTER" in c.upper() or "SURVEYOR" in c.upper()), "Hunter Name")
        hunter_list = df_outlets[hunter_col].dropna().astype(str).unique().tolist() if hunter_col in df_outlets.columns else []

        def extract_coords(row):
            loc_val = row.get("EDS Id Location", "")
            if pd.notna(loc_val):
                try:
                    parts = str(loc_val).split(",")
                    return float(parts[0].strip()), float(parts[1].strip())
                except:
                    pass
            return None, None

        df_outlets['out_lat'], df_outlets['out_lon'] = zip(*df_outlets.apply(extract_coords, axis=1))
        outlet_col = next((c for c in df_outlets.columns if "OUTLET" in c.upper()), df_outlets.columns[0])

        header_title = "FMCG Field Survey Portal" if (is_surveyor_mode or app_mode == "📱 Surveyor Mode Preview") else "Live Field Survey Admin Dashboard"
        st.markdown(f"<h1 style='margin:0; padding-top: 0px; margin-bottom: 10px; font-size: 2.1rem;'>{header_title}</h1>", unsafe_allow_html=True)
        st.markdown("---")

        # =====================================================================
        # MODE 1: SURVEYOR MOBILE MODE
        # =====================================================================
        if is_surveyor_mode or app_mode == "📱 Surveyor Mode Preview":

            selected_hunter = st.selectbox("👤 Select Your Name (Surveyor)", options=hunter_list)

            if selected_hunter:
                df_filtered_outlets = df_outlets[df_outlets[hunter_col].astype(str) == selected_hunter].copy()
            else:
                df_filtered_outlets = df_outlets.copy()

            if surveyor_lat and surveyor_lon:
                df_filtered_outlets['distance_km'] = df_filtered_outlets.apply(
                    lambda r: calculate_distance_km(surveyor_lat, surveyor_lon, r['out_lat'], r['out_lon']), axis=1
                )
                df_filtered_outlets = df_filtered_outlets.sort_values(by='distance_km').reset_index(drop=True)

            outlet_list = df_filtered_outlets[outlet_col].dropna().astype(str).unique().tolist()

            pending_outlets_df = df_filtered_outlets[~df_filtered_outlets[outlet_col].astype(str).isin(st.session_state.surveys_completed)]
            suggested_outlet_name = None
            suggested_distance = None

            if not pending_outlets_df.empty:
                next_row = pending_outlets_df.iloc[0]
                suggested_outlet_name = str(next_row[outlet_col])
                if 'distance_km' in next_row and next_row['distance_km'] < 9000:
                    suggested_distance = f"{next_row['distance_km']:.2f} km"

            if suggested_outlet_name:
                dist_str = f" ({suggested_distance})" if suggested_distance else ""
                st.info(f"💡 **Suggested Next Outlet:** `{suggested_outlet_name}`{dist_str}")

            if not st.session_state.selected_outlet or st.session_state.selected_outlet not in outlet_list:
                st.session_state.selected_outlet = suggested_outlet_name if suggested_outlet_name else (outlet_list[0] if outlet_list else None)

            def sync_outlet_selection():
                st.session_state.selected_outlet = st.session_state.outlet_selectbox_widget
                st.session_state.form_step = 1
                st.session_state.order_item_count = 1
                st.session_state.just_submitted = False

            col_left, col_right = st.columns([1.1, 1.2])

            with col_left:
                st.subheader("🗺️ Outlet Map & Navigation")
                
                cur_match = df_filtered_outlets[df_filtered_outlets[outlet_col].astype(str) == st.session_state.selected_outlet]
                if not cur_match.empty:
                    cur_row = cur_match.iloc[0]
                    sel_lat, sel_lon = cur_row.get('out_lat', None), cur_row.get('out_lon', None)
                else:
                    sel_lat, sel_lon = None, None

                if pd.notna(sel_lat) and pd.notna(sel_lon):
                    map_center = [sel_lat, sel_lon]
                    map_zoom = 16
                elif surveyor_lat and surveyor_lon:
                    map_center = [surveyor_lat, surveyor_lon]
                    map_zoom = 14
                else:
                    map_center = [26.8941, 80.9584]
                    map_zoom = 14

                m = folium.Map(location=map_center, zoom_start=map_zoom)

                if surveyor_lat and surveyor_lon:
                    folium.Marker(
                        [surveyor_lat, surveyor_lon],
                        popup=f"<b>📍 Your Live Location ({selected_hunter})</b>",
                        tooltip="Your Live Position",
                        icon=folium.Icon(color="blue", icon="user", prefix="fa")
                    ).add_to(m)

                for idx, row in df_filtered_outlets.iterrows():
                    out_name = str(row.get(outlet_col, f"Outlet_{idx}"))
                    lat, lon = row['out_lat'], row['out_lon']

                    if pd.notna(lat) and pd.notna(lon):
                        is_selected = (out_name == st.session_state.selected_outlet)
                        is_done = out_name in st.session_state.surveys_completed
                        is_suggested = (out_name == suggested_outlet_name)

                        if surveyor_lat and surveyor_lon:
                            point_gmaps_url = f"https://www.google.com/maps/dir/?api=1&origin={surveyor_lat},{surveyor_lon}&destination={lat},{lon}&travelmode=driving"
                        else:
                            point_gmaps_url = f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}"

                        dist_text = f"{row['distance_km']:.2f} km away" if 'distance_km' in row and row['distance_km'] < 9000 else ""
                        
                        popup_html = f"""
                            <b>{out_name}</b><br>
                            <small>{dist_text}</small><br><br>
                            <a href="{point_gmaps_url}" target="_blank" style="background-color:#007bff; color:white; padding:4px 8px; text-decoration:none; border-radius:4px; font-size:11px;">🧭 Navigate Here</a>
                        """

                        if is_done:
                            point_color = "green"
                        elif is_selected:
                            point_color = "red"
                        elif is_suggested:
                            point_color = "purple"
                        else:
                            point_color = "orange"

                        if is_selected:
                            folium.Marker(
                                [lat, lon],
                                popup=folium.Popup(popup_html, max_width=250),
                                tooltip=out_name,
                                icon=folium.Icon(color="red", icon="star", prefix="fa")
                            ).add_to(m)
                        elif is_suggested:
                            folium.Marker(
                                [lat, lon],
                                popup=folium.Popup(popup_html, max_width=250),
                                tooltip=out_name,
                                icon=folium.Icon(color="purple", icon="info-sign")
                            ).add_to(m)
                        else:
                            folium.CircleMarker(
                                location=[lat, lon],
                                radius=8,
                                popup=folium.Popup(popup_html, max_width=250),
                                tooltip=out_name,
                                color=point_color,
                                fill=True,
                                fill_color=point_color,
                                fill_opacity=0.85
                            ).add_to(m)

                map_output = st_folium(m, width=550, height=480, key="surveyor_interactive_map")

                clicked_point_name = None
                if map_output:
                    if map_output.get("last_object_clicked_tooltip"):
                        clicked_point_name = str(map_output["last_object_clicked_tooltip"]).strip()
                    elif map_output.get("last_object_clicked_popup"):
                        raw_pop = str(map_output["last_object_clicked_popup"])
                        for name_opt in outlet_list:
                            if name_opt in raw_pop:
                                clicked_point_name = name_opt
                                break

                if clicked_point_name and clicked_point_name in outlet_list:
                    if clicked_point_name != st.session_state.selected_outlet:
                        st.session_state.selected_outlet = clicked_point_name
                        st.session_state.form_step = 1
                        st.session_state.order_item_count = 1
                        st.session_state.just_submitted = False
                        st.rerun()

            with col_right:
                st.subheader(f"📝 Dynamic Survey Form ({selected_hunter})")

                # =============================================================
                # CONFIRMATION SCREEN AFTER SUBMISSION
                # =============================================================
                if st.session_state.just_submitted:
                    st.success(f"🎉 **Success!** Survey for `{st.session_state.last_submitted_outlet}` successfully submitted and saved.")
                    st.info("🗺️ Map and Form have automatically moved to the next suggested outlet.")
                    
                    if st.button("👉 Go to Next Outlet Survey", use_container_width=True):
                        st.session_state.just_submitted = False
                        st.session_state.selected_outlet = suggested_outlet_name if suggested_outlet_name else (outlet_list[0] if outlet_list else None)
                        st.session_state.form_step = 1
                        st.session_state.order_item_count = 1
                        st.rerun()
                else:
                    try:
                        curr_idx = outlet_list.index(st.session_state.selected_outlet)
                    except ValueError:
                        curr_idx = 0

                    selected_outlet = st.selectbox(
                        "Select Target Outlet",
                        options=outlet_list,
                        index=curr_idx,
                        key="outlet_selectbox_widget",
                        on_change=sync_outlet_selection
                    )

                    matched_rows = df_filtered_outlets[df_filtered_outlets[outlet_col].astype(str) == st.session_state.selected_outlet]
                    if not matched_rows.empty:
                        out_data = matched_rows.iloc[0]
                    else:
                        out_data = df_filtered_outlets.iloc[0]

                    sel_lat = out_data.get('out_lat', None)
                    sel_lon = out_data.get('out_lon', None)
                    
                    if pd.notna(sel_lat) and pd.notna(sel_lon):
                        if surveyor_lat and surveyor_lon:
                            gmaps_url = f"https://www.google.com/maps/dir/?api=1&origin={surveyor_lat},{surveyor_lon}&destination={sel_lat},{sel_lon}&travelmode=driving"
                        else:
                            gmaps_url = f"https://www.google.com/maps/dir/?api=1&destination={sel_lat},{sel_lon}"
                        
                        st.link_button("🗺️ Open Google Maps Route Navigation", gmaps_url, use_container_width=True)

                    if 'distance_km' in out_data and out_data['distance_km'] < 9000:
                        st.caption(f"📏 Live Distance to Selected Outlet: **{out_data['distance_km']:.2f} km**")

                    if st.session_state.form_step == 1:
                        with st.form(key=f"dynamic_form_{st.session_state.selected_outlet}", clear_on_submit=False):
                            form_values = {}

                            channel_hierarchy = {}
                            if "Channel Sub Channel" in sheet_names:
                                csc_df = sheets_dict["Channel Sub Channel"].copy()
                                csc_df.columns = csc_df.columns.str.strip()
                                if "Channel Name" in csc_df.columns and "Sub Channel Name" in csc_df.columns:
                                    for c_val, group in csc_df.groupby("Channel Name"):
                                        channel_hierarchy[str(c_val).strip()] = group["Sub Channel Name"].dropna().astype(str).str.strip().tolist()

                            products_list = master_dropdowns.get("Products", master_dropdowns.get("Product ", []))

                            form_values[hunter_col] = selected_hunter

                            order_processed = False
                            quantity_processed = False
                            amount_processed = False

                            for col_heading in df_he_columns:
                                head_clean = col_heading.strip()
                                head_lower = head_clean.lower()

                                # HIDE EDS ID LOCATION FIELD COMPLETELY FROM FORM VIEW (Automatic background save)
                                if "eds id location" in head_lower:
                                    form_values[head_clean] = f"{surveyor_lat},{surveyor_lon}" if surveyor_lat else str(out_data.get("EDS Id Location", "GPS Not Locked"))
                                    continue

                                if head_lower in ["id"]:
                                    form_values[head_clean] = f"SURVEY_{datetime.now().strftime('%Y%m%d%H%M%S')}"
                                elif head_lower in ["date"]:
                                    form_values[head_clean] = datetime.now().strftime("%Y-%m-%d")
                                elif head_lower in ["time"]:
                                    form_values[head_clean] = datetime.now().strftime("%H:%M:%S")
                                elif head_lower in ["location"]:
                                    form_values[head_clean] = f"{surveyor_lat},{surveyor_lon}" if surveyor_lat else "GPS Not Locked"

                                elif head_clean in df_filtered_outlets.columns or head_clean.upper() in [c.upper() for c in df_filtered_outlets.columns]:
                                    matched_col = next((c for c in df_filtered_outlets.columns if c.upper() == head_clean.upper()), None)
                                    default_val = str(out_data.get(matched_col, "")) if matched_col else ""
                                    
                                    # MAKE MASTER COLUMNS SELECTABLE / FILLABLE (Instead of rigid Auto-filled text box)
                                    col_options = df_outlets[matched_col].dropna().astype(str).unique().tolist() if matched_col else []
                                    if col_options:
                                        if default_val in col_options:
                                            default_idx = col_options.index(default_val)
                                        else:
                                            default_idx = 0
                                        form_values[head_clean] = st.selectbox(head_clean, options=col_options, index=default_idx, key=f"sel_master_{head_clean}_{st.session_state.selected_outlet}")
                                    else:
                                        form_values[head_clean] = st.text_input(head_clean, value=default_val, key=f"inp_{head_clean}_{st.session_state.selected_outlet}")

                                elif "outlet name" in head_lower:
                                    form_values[head_clean] = st.text_input(head_clean, value=st.session_state.selected_outlet, key=f"inp_outname_{st.session_state.selected_outlet}")

                                elif head_lower == "channel name":
                                    c_options = list(channel_hierarchy.keys()) if channel_hierarchy else master_dropdowns.get("Channel Name", ["General"])
                                    form_values[head_clean] = st.selectbox(head_clean, options=c_options, key=f"chan_{st.session_state.selected_outlet}")

                                elif head_lower == "sub channel name":
                                    sel_chan = st.session_state.get(f"chan_{st.session_state.selected_outlet}", None)
                                    sub_opts = channel_hierarchy.get(sel_chan, []) if sel_chan else master_dropdowns.get("Sub Channel Name", ["General"])
                                    if not sub_opts:
                                        sub_opts = master_dropdowns.get("Sub Channel Name", ["General"])
                                    form_values[head_clean] = st.selectbox(head_clean, options=sub_opts, key=f"subchan_{st.session_state.selected_outlet}")

                                elif any(k in head_lower for k in ["multi", "brands", "available products", "issues", "services"]):
                                    dropdown_opts = master_dropdowns.get(head_clean, products_list if products_list else ["Option 1", "Option 2"])
                                    selected_items = st.multiselect(f"{head_clean} (Multi-select)", options=dropdown_opts, key=f"multi_{head_clean}_{st.session_state.selected_outlet}")
                                    form_values[head_clean] = ", ".join(selected_items) if selected_items else "None"

                                elif "order" in head_lower or "product" in head_lower:
                                    if not order_processed:
                                        order_processed = True
                                        order_items_collected = []
                                        for i in range(st.session_state.order_item_count):
                                            if products_list:
                                                item_val = st.selectbox(f"Order Item #{i+1}", options=["None"] + products_list, key=f"prod_{head_clean}_{i}_{st.session_state.selected_outlet}")
                                            else:
                                                item_val = st.text_input(f"Order Item #{i+1}", key=f"txt_{head_clean}_{i}_{st.session_state.selected_outlet}")
                                            if item_val and item_val != "None":
                                                order_items_collected.append(item_val)
                                        
                                        form_values[head_clean] = ", ".join(order_items_collected) if order_items_collected else "None"
                                    else:
                                        continue

                                elif "quantity" in head_lower or "qty" in head_lower:
                                    if not quantity_processed:
                                        quantity_processed = True
                                        form_values[head_clean] = st.number_input(head_clean, min_value=0, step=1, key=f"num_{head_clean}_{st.session_state.selected_outlet}")
                                    else:
                                        continue

                                elif "amount" in head_lower or "price" in head_lower or "value" in head_lower:
                                    if not amount_processed:
                                        amount_processed = True
                                        form_values[head_clean] = st.number_input(head_clean, min_value=0.0, step=10.0, key=f"amt_{head_clean}_{st.session_state.selected_outlet}")
                                    else:
                                        continue

                                elif "image" in head_lower or "photo" in head_lower or "shop image" in head_lower:
                                    pass

                                else:
                                    dropdown_opts = None
                                    for key, opt_list in master_dropdowns.items():
                                        if key.lower().replace(" ", "").replace("_", "") in head_lower.replace(" ", "").replace("_", ""):
                                            dropdown_opts = opt_list
                                            break

                                    if dropdown_opts:
                                        form_values[head_clean] = st.selectbox(head_clean, options=dropdown_opts, key=f"opt_{head_clean}_{st.session_state.selected_outlet}")
                                    else:
                                        form_values[head_clean] = st.text_input(head_clean, key=f"txtin_{head_clean}_{st.session_state.selected_outlet}")

                            proceed_btn = st.form_submit_button("📸 Submit Form & Open Camera", use_container_width=True)

                            if proceed_btn:
                                st.session_state.temp_form_data = form_values
                                st.session_state.form_step = 2
                                st.rerun()

                        if st.button("➕ Add Another Order Item"):
                            st.session_state.order_item_count += 1
                            st.rerun()

                    elif st.session_state.form_step == 2:
                        st.success("✅ Form data recorded! Capture store photo to finalize the survey.")
                        captured_img = st.camera_input("📷 Take Photo Now to Finalize", key=f"camera_step2_{st.session_state.selected_outlet}")

                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            if st.button("⬅️ Back to Form Edit", use_container_width=True):
                                st.session_state.form_step = 1
                                st.rerun()

                        if captured_img:
                            final_data = st.session_state.temp_form_data.copy()
                            img_col = next((c for c in df_he_columns if "image" in c.lower() or "photo" in c.lower()), "Shop Image")

                            img_filename = f"{st.session_state.selected_outlet}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
                            img_path = os.path.join(IMAGE_FOLDER, img_filename)
                            image = Image.open(captured_img)
                            image.save(img_path)
                            final_data[img_col] = img_filename

                            save_survey_data(final_data)
                            
                            completed_outlet_name = st.session_state.selected_outlet
                            st.session_state.surveys_completed.add(completed_outlet_name)
                            
                            temp_pending_df = df_filtered_outlets[~df_filtered_outlets[outlet_col].astype(str).isin(st.session_state.surveys_completed)]
                            next_suggested = str(temp_pending_df.iloc[0][outlet_col]) if not temp_pending_df.empty else outlet_list[0]

                            st.session_state.last_submitted_outlet = completed_outlet_name
                            st.session_state.selected_outlet = next_suggested
                            st.session_state.form_step = 1
                            st.session_state.order_item_count = 1
                            st.session_state.just_submitted = True
                            st.rerun()

        # =====================================================================
        # MODE 2: ADMIN LIVE DASHBOARD
        # =====================================================================
        else:
            total_outlets = len(df_outlets)
            current_completed = len(st.session_state.surveys_completed)
            current_pending = total_outlets - current_completed
            completion_rate = (current_completed / total_outlets * 100) if total_outlets > 0 else 0

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("📍 Total Outlets", f"{total_outlets}")
            m2.metric("✅ Completed Surveys", f"{current_completed}", delta=f"{completion_rate:.1f}% Done")
            m3.metric("⏳ Pending Outlets", f"{current_pending}")
            m4.metric("👥 Total Surveyors", f"{len(hunter_list)}")

            st.markdown("---")

            st.subheader("🔗 Quick Application Access Links")
            base_url = st.text_input("🌐 App Base URL (Streamlit Cloud / Localhost)", value="http://localhost:8501", key="base_url_input")

            admin_url = f"{base_url.strip('/')}/"
            surveyor_url = f"{base_url.strip('/')}/?mode=surveyor"

            col_link_1, col_link_2 = st.columns(2)
            with col_link_1:
                st.markdown("**🖥️ Admin Portal Link:**")
                st.code(admin_url, language="text")
            with col_link_2:
                st.markdown("**📱 Surveyor Mobile Link:**")
                st.code(surveyor_url, language="text")

            st.markdown("---")
            st.subheader("🗺️ Live Map Tracking & Filter Options")
            
            c_filter1, c_filter2 = st.columns(2)
            with c_filter1:
                selected_admin_hunter = st.selectbox("👤 Filter by Surveyor", options=["All Surveyors"] + hunter_list)
            with c_filter2:
                map_status_filter = st.selectbox("📌 Filter Outlets Status on Map", options=["Show All Points", "✅ Completed Points Only", "⏳ Pending Points Only", "📍 Surveyor Live Location Only"])

            if selected_admin_hunter != "All Surveyors":
                adm_outlets = df_outlets[df_outlets[hunter_col].astype(str) == selected_admin_hunter]
            else:
                adm_outlets = df_outlets

            adm_map = folium.Map(location=[26.8941, 80.9584], zoom_start=13)

            if surveyor_lat and surveyor_lon:
                folium.Marker(
                    [surveyor_lat, surveyor_lon],
                    popup="<b>📍 Active Surveyor Current GPS Location</b>",
                    tooltip="Surveyor Live Position",
                    icon=folium.Icon(color="blue", icon="user", prefix="fa")
                ).add_to(adm_map)

            if map_status_filter != "📍 Surveyor Live Location Only":
                for idx, row in adm_outlets.iterrows():
                    out_name = str(row.get(outlet_col, f"Outlet_{idx}"))
                    h_owner = str(row.get(hunter_col, "Unknown"))
                    lat, lon = row['out_lat'], row['out_lon']

                    if pd.notna(lat) and pd.notna(lon):
                        is_done = out_name in st.session_state.surveys_completed
                        
                        if map_status_filter == "✅ Completed Points Only" and not is_done:
                            continue
                        if map_status_filter == "⏳ Pending Points Only" and is_done:
                            continue

                        color = "green" if is_done else "orange"
                        status_lbl = "Completed ✅" if is_done else "Pending ⏳"

                        folium.CircleMarker(
                            location=[lat, lon],
                            radius=7,
                            popup=f"<b>{out_name}</b><br>Surveyor: {h_owner}<br>Status: <b>{status_lbl}</b>",
                            tooltip=f"{out_name} ({status_lbl})",
                            color=color,
                            fill=True,
                            fill_color=color,
                            fill_opacity=0.85
                        ).add_to(adm_map)

            st_folium(adm_map, width=1200, height=500, key="admin_master_map")

            st.markdown("---")
            col_rep_head, col_rep_date = st.columns([1.2, 1])

            with col_rep_head:
                st.subheader("📊 Surveyor-wise Progress Report")

            resp_df = load_survey_data()

            with col_rep_date:
                today = date.today()
                date_range = st.date_input("📅 Select Date Range Filter", value=(today, today), key="progress_date_range")

            filtered_resp_df = resp_df.copy()

            if not resp_df.empty:
                date_col = next((c for c in resp_df.columns if "DATE" in c.upper()), None)
                if date_col:
                    filtered_resp_df[date_col] = pd.to_datetime(filtered_resp_df[date_col], errors='coerce').dt.date
                    if isinstance(date_range, tuple) and len(date_range) == 2:
                        start_d, end_d = date_range
                        filtered_resp_df = filtered_resp_df[(filtered_resp_df[date_col] >= start_d) & (filtered_resp_df[date_col] <= end_d)]

            if not filtered_resp_df.empty:
                resp_outlet_col = next((c for c in filtered_resp_df.columns if "OUTLET" in c.upper()), None)
                filtered_completed_set = set(filtered_resp_df[resp_outlet_col].dropna().astype(str).tolist()) if resp_outlet_col else set()
            else:
                filtered_completed_set = set()

            progress_data = []
            for h_name in hunter_list:
                h_outlets = df_outlets[df_outlets[hunter_col].astype(str) == h_name]
                h_total = len(h_outlets)
                h_outlet_names = set(h_outlets[outlet_col].dropna().astype(str).tolist())
                h_done = len(h_outlet_names.intersection(filtered_completed_set))
                h_pending = h_total - h_done
                h_pct = (h_done / h_total * 100) if h_total > 0 else 0

                progress_data.append({
                    "Surveyor Name": h_name,
                    "Total Outlets Assigned": h_total,
                    "Completed Points ✅": h_done,
                    "Pending Points ⏳": h_pending,
                    "Completion %": f"{h_pct:.1f}%"
                })

            df_progress = pd.DataFrame(progress_data)
            st.dataframe(df_progress, use_container_width=True)

            st.markdown("---")
            st.subheader("📥 Download Survey Data")
            
            ot_resp = load_survey_data()
            if not ot_resp.empty:
                import io
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    ot_resp.to_excel(writer, index=False, sheet_name='Survey_Data')
                buffer.seek(0)

                st.download_button(
                    "Download Survey Responses (Excel)",
                    data=buffer.getvalue(),
                    file_name=f"Survey_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("Abhi tak koi survey response save nahi hua hai.")

    except Exception as e:
        st.error(f"Error reading Master Excel file: {str(e)}")

else:
    st.info("👈 Please upload Master Excel File or place `master_questions.xlsx` on server.")
import streamlit as st
import pandas as pd
import re
import os
import requests
from urllib.parse import urlparse
import json
from io import BytesIO

# Configure page
st.set_page_config(
    page_title="🎓 Academic Attendance Calculator",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

def load_module_mappings():
    """Load existing module mappings from JSON file"""
    try:
        if os.path.exists('module_mappings.json'):
            with open('module_mappings.json', 'r') as f:
                mappings = json.load(f)
                # Also update session state
                st.session_state.mappings = mappings
                return mappings
        else:
            return {}
    except Exception as e:
        st.error(f"Error loading mappings: {e}")
        return {}
    
if 'mappings' not in st.session_state:
    st.session_state.mappings = load_module_mappings()  # Load from file on startup
if 'df' not in st.session_state:
    st.session_state.df = None
if 'modules' not in st.session_state:
    st.session_state.modules = []
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False

def is_exam_session(text):
    """Check if session is exam-related (these don't count for attendance)"""
    if pd.isna(text):
        return False
    
    text = str(text).lower()
    exam_keywords = ["examination", "exam", "coursework", "viva", "cw", "course work"]
    
    return any(keyword in text for keyword in exam_keywords)

def normalize_session_text(text):
    """Clean up session text for mapping"""
    if pd.isna(text):
        return ""
    
    text = str(text).strip()
    
    # Remove instructor names (improved pattern)
    text = re.sub(r'\s*-\s*(Ms|Mr|Dr|Prof|Professor)\.?\s+[A-Z][a-z]+.*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+(Ms|Mr|Dr|Prof|Professor)\.?\s+[A-Z][a-z]+.*$', '', text, flags=re.IGNORECASE)
    
    # Remove tutorial/practical indicators
    text = re.sub(r'\s*-\s*(tutorial|practical|tute|prac|lab)\s*$', '', text, flags=re.IGNORECASE)
    text = re.sub(r'\s+(tutorial|practical|tute|prac|lab)\s*$', '', text, flags=re.IGNORECASE)
    
    # Clean up extra spaces and dashes
    text = re.sub(r'\s*-\s*$', '', text)  # Remove trailing dash
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text

def get_module_code_for_session(session_text, mappings):
    """Get module code for a session"""
    # First check if it's an exam session
    if is_exam_session(session_text):
        return "EXAM"
    
    normalized = normalize_session_text(session_text).lower()
    
    # Skip empty or common non-academic sessions
    if not normalized or normalized in ["inauguration", "holiday", "break", "lunch", "nan"]:
        return None
    
    # Check if we already have a mapping for this normalized text
    if normalized in mappings:
        return mappings[normalized]
    
    # Check if any existing mapping key is contained in this session (both ways)
    for key, code in mappings.items():
        key_lower = key.lower()
        if key_lower in normalized or normalized in key_lower:
            return code
        
        # Also check if the key matches closely (handling small variations)
        if key_lower.replace(" ", "") == normalized.replace(" ", ""):
            return code
    
    # Return None if not found (will be handled by UI)
    return None

def is_for_module(cell, module_name, mappings):
    """Check if the session is related to the selected module"""
    if pd.isna(cell): 
        return False
    
    # Get the module code for this session
    session_module = get_module_code_for_session(cell, mappings)
    return session_module == module_name

def is_tutorial_or_practical(cell):
    """Check if the session is a tutorial or practical"""
    if pd.isna(cell): 
        return False
    cell = str(cell).lower()
    return any(keyword in cell for keyword in ["tutorial", "practical", "tute", "prac", "lab"])

def load_schedule(file_data):
    """Load and format the Excel schedule file"""
    try:
        xls = pd.ExcelFile(file_data)
        df = xls.parse(xls.sheet_names[0])
        
        # Handle different column structures
        if len(df.columns) < 3:
            raise ValueError("Excel file must have at least 3 columns (Date, Morning, Afternoon)")
        
        df.columns = ["Date", "Morning", "Afternoon"] + [f"Extra_{i}" for i in range(len(df.columns) - 3)]
        df = df.dropna(subset=["Date"])
        
        # Skip header row if it exists
        if df.iloc[0]["Date"] and str(df.iloc[0]["Date"]).lower() in ["date", "day"]:
            df = df[1:]
        
        return df
        
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return None

def get_all_session_columns(df):
    """Get all columns that contain sessions (Morning, Afternoon, Extra_0, etc.)"""
    session_columns = []
    for col in df.columns:
        if col in ["Morning", "Afternoon"] or col.startswith("Extra_"):
            session_columns.append(col)
    return session_columns

def get_unmapped_sessions(df, mappings):
    """Get all sessions that need mapping"""
    session_columns = get_all_session_columns(df)
    all_sessions = pd.concat([df[col] for col in session_columns]).dropna()
    unmapped = []
    
    print("Debug: Checking sessions against mappings...")  # Debug
    
    for session in all_sessions.astype(str):
        session = session.strip()
        if session and session.lower() not in ["inauguration", "holiday", "break", "lunch", "nan"]:
            normalized = normalize_session_text(session).lower()
            
            # Skip if it's an exam session
            if is_exam_session(session):
                continue
                
            # Check if we have a mapping
            found_mapping = False
            
            # Direct match
            if normalized in mappings:
                found_mapping = True
            else:
                # Check partial matches (both ways)
                for key in mappings.keys():
                    key_lower = key.lower()
                    if key_lower in normalized or normalized in key_lower:
                        found_mapping = True
                        break
                    # Also check without spaces
                    if key_lower.replace(" ", "") == normalized.replace(" ", ""):
                        found_mapping = True
                        break
            
            if not found_mapping:
                print(f"Debug: No mapping found for '{session}' -> '{normalized}'")  # Debug
                unmapped.append((session, normalized))
    
    # Remove duplicates while preserving order
    seen = set()
    unique_unmapped = []
    for session, normalized in unmapped:
        if normalized not in seen:
            seen.add(normalized)
            unique_unmapped.append((session, normalized))
    
    return unique_unmapped

def build_module_list(df, mappings):
    """Build list of all modules by processing all sessions"""
    session_columns = get_all_session_columns(df)
    all_sessions = pd.concat([df[col] for col in session_columns]).dropna()
    modules = set()
    
    for session in all_sessions.astype(str):
        session = session.strip()
        if session and session.lower() not in ["inauguration", "holiday", "break", "lunch", "nan"]:
            module_code = get_module_code_for_session(session, mappings)
            if module_code and module_code != "EXAM":
                modules.add(module_code)
    
    return sorted(list(modules))

def count_lectures_for_module(df, module_name, mappings):
    """Count total sessions for a specific module"""
    total = 0
    session_details = []
    session_columns = get_all_session_columns(df)
    
    for idx, row in df.iterrows():
        date = row["Date"]
        day_sessions = []
        
        # Check all session columns
        for col in session_columns:
            session_text = str(row[col]) if not pd.isna(row[col]) else ""
            if session_text:
                session_match = is_for_module(session_text, module_name, mappings)
                if session_match:
                    is_tut = is_tutorial_or_practical(session_text)
                    session_type = "Tutorial/Practical" if is_tut else "Lecture"
                    
                    # Map column names to readable format
                    if col == "Morning":
                        time_slot = "🌅 Morning"
                    elif col == "Afternoon":
                        time_slot = "🌆 Afternoon"
                    else:
                        extra_num = col.split('_')[1] if '_' in col else "0"
                        time_slot = f"📚 Extra_{extra_num}"
                    
                    day_sessions.append({
                        'time_slot': time_slot,
                        'session_text': session_text,
                        'session_type': session_type
                    })
                    total += 1
        
        # Add the day if there are sessions
        if day_sessions:
            session_details.append({
                'date': date,
                'sessions': day_sessions
            })

    return total, session_details

def calculate_holiday_allowance(total_sessions, current_missed, min_percentage=80):
    """Calculate how many more sessions can be missed while maintaining minimum attendance"""
    min_sessions_needed = int(total_sessions * (min_percentage / 100))
    sessions_attended = total_sessions - current_missed
    max_total_missed = total_sessions - min_sessions_needed
    additional_misses_allowed = max_total_missed - current_missed
    
    return max(0, additional_misses_allowed), min_sessions_needed

def download_from_sharepoint():
    """Download Excel from SharePoint public link"""
    shared_link = "https://nibm-my.sharepoint.com/:x:/g/personal/chandula_nibm_lk/EWelzFX_1ipFhyHvho9Su2oBBR0L2UAnqPzQHFhROycGiQ?e=aMcKjo"
    
    try:
        # Extract file id token from URL
        file_id = urlparse(shared_link).path.split('/')[-1]
        direct_download_url = f"https://nibm-my.sharepoint.com/personal/chandula_nibm_lk/_layouts/15/download.aspx?share={file_id}"
        
        response = requests.get(direct_download_url, stream=True)
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            st.error(f"Failed to download file (HTTP {response.status_code})")
            return None
    except Exception as e:
        st.error(f"Error downloading file: {e}")
        return None
    
def download_from_sharepoint2():
    """Download Excel from SharePoint public link"""
    shared_link = "https://nibm-my.sharepoint.com/:x:/g/personal/amilau_nibm_lk/EdzoomT5ACJCjUpohkuSd78B6FMiyqSt2LeF4a3nZuXFkw?e=phnvms"
    
    try:
        # Extract file id token from URL
        file_id = urlparse(shared_link).path.split('/')[-1]
        direct_download_url = f"https://nibm-my.sharepoint.com/personal/chandula_nibm_lk/_layouts/15/download.aspx?share={file_id}"
        
        response = requests.get(direct_download_url, stream=True)
        if response.status_code == 200:
            return BytesIO(response.content)
        else:
            st.error(f"Failed to download file (HTTP {response.status_code})")
            return None
    except Exception as e:
        st.error(f"Error downloading file: {e}")
        return None

# Main Streamlit App
def main():
    st.title("🎓 Academic Attendance Calculator")
    st.markdown("---")
    
    # Sidebar for file upload and settings
    with st.sidebar:
        st.header("📁 File Upload")
        
        # Option to download from SharePoint or upload file
        data_source = st.radio("Choose data source:", 
                              ["📥 Upload Excel File", "🌐 DSE", "🌐 DCSD"])
        
        if data_source == "🌐 DSE":
            if st.button("📥 Download Schedule"):
                with st.spinner("Downloading schedule..."):
                    file_data = download_from_sharepoint()
                    if file_data:
                        st.session_state.df = load_schedule(file_data)
                        if st.session_state.df is not None:
                            st.success("✅ Schedule downloaded successfully!")
                            st.session_state.file_processed = True
                        else:
                            st.error("❌ Failed to process downloaded file")
        elif data_source == "🌐 DCSD":
            if st.button("📥 Download Schedule"):
                with st.spinner("Downloading schedule..."):
                    file_data = download_from_sharepoint2()
                    if file_data:
                        st.session_state.df = load_schedule(file_data)
                        if st.session_state.df is not None:
                            st.success("✅ Schedule downloaded successfully!")
                            st.session_state.file_processed = True
                        else:
                            st.error("❌ Failed to process downloaded file")
        else:
            uploaded_file = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'])
            if uploaded_file is not None:
                st.session_state.df = load_schedule(uploaded_file)
                if st.session_state.df is not None:
                    st.success("✅ File uploaded successfully!")
                    st.session_state.file_processed = True
                else:
                    st.error("❌ Failed to process uploaded file")
    
    # Main content area
    if st.session_state.df is not None:
        # Show file info
        st.success(f"📊 Loaded {len(st.session_state.df)} rows of schedule data")
        
        # Load mappings from file (in case they were updated outside the app)
        mappings = load_module_mappings()
        
        # Show existing mappings at startup if available
        if mappings:
            st.info(f"📋 Loaded {len(mappings)} existing session mappings from module_mappings.json")
            
            # Debug section to show what mappings are loaded
            with st.expander("🔍 Debug: Show Raw Mappings"):
                st.json(mappings)
            
            with st.expander("View All Current Mappings"):
                # Group by module code for better display
                module_groups = {}
                for session, code in mappings.items():
                    if code not in module_groups:
                        module_groups[code] = []
                    module_groups[code].append(session)
                
                for code in sorted(module_groups.keys()):
                    st.write(f"**📘 {code}:**")
                    for session in sorted(module_groups[code]):
                        st.write(f"  • {session}")
        
        # Check for unmapped sessions
        unmapped_sessions = get_unmapped_sessions(st.session_state.df, mappings)
        
        # Debug: Show what sessions are being processed
        with st.expander("🔍 Debug: Session Processing"):
            session_columns = get_all_session_columns(st.session_state.df)
            all_sessions = pd.concat([st.session_state.df[col] for col in session_columns]).dropna()
            
            st.write("**Sample original sessions:**")
            for session in list(all_sessions.astype(str))[:10]:  # Show first 10
                normalized = normalize_session_text(session).lower()
                found_in_mappings = normalized in mappings
                st.write(f"- Original: `{session}`")
                st.write(f"  Normalized: `{normalized}`")
                st.write(f"  Found in mappings: {found_in_mappings}")
                if found_in_mappings:
                    st.write(f"  Maps to: `{mappings[normalized]}`")
                st.write("---")
        
        if unmapped_sessions:
            st.warning(f"⚠️ Found {len(unmapped_sessions)} unmapped sessions. Let @silknye know:")
            
            # Create mapping interface
            st.subheader("🗂️ Session Mapping")
            
            # Show existing mappings
            if mappings:
                with st.expander("View Current Mappings"):
                    mapping_df = pd.DataFrame([
                        {"Session": session, "Module Code": code} 
                        for session, code in mappings.items()
                    ])
                    st.dataframe(mapping_df)
            
            # Map unmapped sessions
        
        # Build module list
        st.session_state.modules = build_module_list(st.session_state.df, mappings)
        
        if st.session_state.modules:
            st.subheader("📘 Available Modules")
            selected_module = st.selectbox(
                "Choose a module:",
                st.session_state.modules,
                key="module_selector"
            )
            
            if selected_module:
                # Count sessions for selected module
                total_sessions, session_details = count_lectures_for_module(
                    st.session_state.df, selected_module, mappings
                )
                
                if total_sessions > 0:
                    # Display session details
                    st.subheader(f"📋 Sessions for {selected_module}")
                    
                    # Show session breakdown
                    lecture_count = sum(1 for day in session_details 
                                      for session in day['sessions'] 
                                      if session['session_type'] == 'Lecture')
                    tutorial_count = sum(1 for day in session_details 
                                       for session in day['sessions'] 
                                       if session['session_type'] == 'Tutorial/Practical')
                    
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📚 Regular Lectures", lecture_count)
                    with col2:
                        st.metric("🛠️ Tutorials/Practicals", tutorial_count)
                    with col3:
                        st.metric("📊 Total Sessions", total_sessions)
                    
                    # Show detailed session list
                    with st.expander("View All Sessions"):
                        for day in session_details:
                            st.write(f"**📅 {day['date']}:**")
                            for session in day['sessions']:
                                st.write(f"  {session['time_slot']}: {session['session_text']} ({session['session_type']})")
                    
                    # Attendance calculation
                    st.subheader("📊 Attendance Calculation")
                    
                    missed_sessions = st.number_input(
                        "❌ How many sessions did you miss?",
                        min_value=0,
                        max_value=total_sessions,
                        value=0,
                        step=1
                    )
                    
                    if st.button("📈 Calculate Attendance"):
                        # Calculate results
                        attended = total_sessions - missed_sessions
                        attendance_percentage = (attended / total_sessions) * 100
                        holiday_allowance, min_sessions_needed = calculate_holiday_allowance(
                            total_sessions, missed_sessions
                        )
                        
                        # Display results
                        st.subheader("📊 Attendance Report")
                        
                        # Metrics
                        col1, col2, col3, col4 = st.columns(4)
                        with col1:
                            st.metric("📚 Total Sessions", total_sessions)
                        with col2:
                            st.metric("✅ Attended", attended)
                        with col3:
                            st.metric("❌ Missed", missed_sessions)
                        with col4:
                            st.metric("📈 Attendance %", f"{attendance_percentage:.1f}%")
                        
                        # Status indicator
                        if attendance_percentage >= 80:
                            st.success("🟢 **Status: GOOD** - Meeting attendance requirements")
                            if holiday_allowance > 0:
                                st.info(f"💡 You can miss {holiday_allowance} more session(s) and still maintain 80%")
                        elif attendance_percentage >= 75:
                            st.warning("🟡 **Status: WARNING** - Close to minimum requirement")
                            if holiday_allowance > 0:
                                st.warning(f"⚠️ You can only miss {holiday_allowance} more session(s) to maintain 80%")
                            else:
                                st.error("🚨 Cannot miss any more sessions to maintain 80%")
                        else:
                            st.error("🔴 **Status: CRITICAL** - Below minimum attendance requirement")
                            sessions_needed = min_sessions_needed - attended
                            st.error(f"📉 You need {sessions_needed} more attended sessions to reach 80%")
                        
                        # Holiday allowance details
                        st.subheader("🏖️ Holiday Allowance (80% minimum)")
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("📋 Minimum sessions needed", min_sessions_needed)
                        with col2:
                            st.metric("🔢 Max total misses allowed", total_sessions - min_sessions_needed)
                        
                        if holiday_allowance > 0:
                            st.success(f"🎉 Sessions you can still miss: **{holiday_allowance}**")
                        else:
                            sessions_over_limit = missed_sessions - (total_sessions - min_sessions_needed)
                            st.error(f"⚠️ Already at/over limit! Sessions over: **{sessions_over_limit}**")
                
                else:
                    st.error(f"❌ No sessions found for module: {selected_module}")
        else:
            st.info("📝 Please map the sessions above to see available modules")
    
    else:
        st.info("📁 Please upload an Excel file or download from SharePoint to get started")
        
        # Show instructions
        st.subheader("📋 Instructions")
        st.markdown("""
        1. **Upload your Excel file** or **download from SharePoint**
        2. **Map sessions** to their respective module codes
        3. **Select a module** to analyze
        4. **Enter missed sessions** to calculate attendance
        5. **View results** and holiday allowance
        """)

if __name__ == "__main__":
    main()
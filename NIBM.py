import pandas as pd
import re
import os
import requests
from urllib.parse import urlparse
import json

def load_module_mappings():
    """Load existing module mappings from file"""
    try:
        with open('module_mappings.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {}

def save_module_mappings(mappings):
    """Save module mappings to file"""
    with open('module_mappings.json', 'w') as f:
        json.dump(mappings, f, indent=2)

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
    """Get module code for a session, asking user if not found"""
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
    
    # Check if any existing mapping key is contained in this session
    for key, code in mappings.items():
        if key.lower() in normalized or normalized in key.lower():
            return code
    
    # Not found, ask user
    print(f"\n🤔 New session found: '{session_text}'")
    print(f"   Normalized to: '{normalized}'")
    print("What module code should this be mapped to?")
    
    # Show existing modules for reference
    if mappings:
        existing_codes = sorted(set(v for v in mappings.values() if v != "EXAM"))
        print(f"📚 Existing modules: {', '.join(existing_codes)}")
    
    while True:
        module_code = input("Enter module code (e.g., DLO, ECS, DSA, OOP): ").strip().upper()
        if module_code:
            # Add this mapping
            mappings[normalized] = module_code
            save_module_mappings(mappings)
            print(f"✅ Mapped '{session_text}' → {module_code}")
            return module_code
        else:
            print("❌ Please enter a valid module code")

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

def load_schedule(file_path):
    """Load and format the Excel schedule file"""
    try:
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found: {file_path}")
        
        xls = pd.ExcelFile(file_path)
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
        print(f"❌ Error loading file: {e}")
        return None

def get_all_session_columns(df):
    """Get all columns that contain sessions (Morning, Afternoon, Extra_0, etc.)"""
    session_columns = []
    for col in df.columns:
        if col in ["Morning", "Afternoon"] or col.startswith("Extra_"):
            session_columns.append(col)
    return session_columns

def build_module_list(df, mappings):
    """Build list of all modules by processing all sessions"""
    session_columns = get_all_session_columns(df)
    all_sessions = pd.concat([df[col] for col in session_columns]).dropna()
    modules = set()
    
    print("\n🔄 Processing sessions to build module list...")
    print("(Exam sessions will be auto-mapped to EXAM and excluded from attendance)")
    
    for session in all_sessions.astype(str):
        session = session.strip()
        if session and session.lower() not in ["inauguration", "holiday", "break", "lunch", "nan"]:
            # This will ask for mapping if not found
            module_code = get_module_code_for_session(session, mappings)
            if module_code and module_code != "EXAM":  # Don't include EXAM in selectable modules
                modules.add(module_code)
    
    return sorted(list(modules))

def count_lectures_for_module(df, module_name, mappings):
    """Count total sessions for a specific module"""
    total = 0
    session_columns = get_all_session_columns(df)
    
    print(f"\n🔍 Sessions found for module '{module_name}':")
    print("-" * 50)
    
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
                        # For Extra columns, use a different emoji
                        extra_num = col.split('_')[1] if '_' in col else "0"
                        time_slot = f"📚 Extra_{extra_num}"
                    
                    day_sessions.append(f"  {time_slot}: {session_text} ({session_type})")
                    total += 1
        
        # Print the day if there are sessions
        if day_sessions:
            print(f"📅 {date}:")
            for session in day_sessions:
                print(session)

    return total

def calculate_holiday_allowance(total_sessions, current_missed, min_percentage=80):
    """Calculate how many more sessions can be missed while maintaining minimum attendance"""
    # Calculate minimum sessions needed to attend
    min_sessions_needed = int(total_sessions * (min_percentage / 100))
    
    # Calculate sessions already attended
    sessions_attended = total_sessions - current_missed
    
    # Calculate how many more sessions can be missed
    max_total_missed = total_sessions - min_sessions_needed
    additional_misses_allowed = max_total_missed - current_missed
    
    return max(0, additional_misses_allowed), min_sessions_needed

def get_valid_file_path():
    """Download Excel from SharePoint public link or ask user if fails"""
    shared_link = "https://nibm-my.sharepoint.com/:x:/g/personal/amilau_nibm_lk/EdzoomT5ACJCjUpohkuSd78B6FMiyqSt2LeF4a3nZuXFkw?e=phnvms"

    # Extract file id token from URL
    file_id = urlparse(shared_link).path.split('/')[-1]

    # Construct direct download URL
    direct_download_url = f"https://nibm-my.sharepoint.com/personal/amilau_nibm_lk/_layouts/15/download.aspx?share={file_id}"

    print(f"🌐 Downloading schedule from SharePoint link...")
    try:
        response = requests.get(direct_download_url, stream=True)
        if response.status_code == 200:
            local_file = "excel_data.xlsx"
            with open(local_file, "wb") as f:
                for chunk in response.iter_content(1024):
                    f.write(chunk)
            print(f"✅ File downloaded successfully as '{local_file}'")
            return local_file
        else:
            print(f"❌ Failed to download file (HTTP {response.status_code}).")
    except Exception as e:
        print(f"❌ Error downloading file: {e}")

    # Fallback: manual input if download fails
    while True:
        file_path = input("📂 Enter Excel file path manually: ").strip().strip('"\'')
        if os.path.exists(file_path):
            print(f"✅ Using file: {file_path}")
            return file_path
        else:
            print(f"❌ File not found: {file_path}")
            retry = input("Try again? (y/n): ").lower()
            if retry != 'y':
                return None

def get_valid_module_choice(modules):
    """Get and validate module choice from user"""
    while True:
        try:
            choice = int(input(f"\n❓ Choose module number (1-{len(modules)}): "))
            if 1 <= choice <= len(modules):
                return choice - 1  # Convert to 0-based index
            else:
                print(f"❌ Please enter a number between 1 and {len(modules)}")
        except ValueError:
            print("❌ Please enter a valid number")

def get_valid_missed_sessions(total_sessions):
    """Get and validate number of missed sessions"""
    while True:
        try:
            missed = int(input(f"❌ How many sessions did you miss? (0-{total_sessions}): "))
            if 0 <= missed <= total_sessions:
                return missed
            else:
                print(f"❌ Please enter a number between 0 and {total_sessions}")
        except ValueError:
            print("❌ Please enter a valid number")

def show_session_breakdown(df, module_name, mappings):
    """Show detailed breakdown of sessions for verification"""
    print(f"\n📋 Session Breakdown for {module_name}:")
    print("=" * 50)
    
    session_columns = get_all_session_columns(df)
    lecture_count = 0
    tutorial_count = 0
    
    for idx, row in df.iterrows():
        for col in session_columns:
            session_text = str(row[col]) if not pd.isna(row[col]) else ""
            if session_text:
                session_match = is_for_module(session_text, module_name, mappings)
                if session_match:
                    is_tut = is_tutorial_or_practical(session_text)
                    if is_tut:
                        tutorial_count += 1
                    else:
                        lecture_count += 1
    
    print(f"📚 Regular Lectures: {lecture_count}")
    print(f"🛠️  Tutorials/Practicals: {tutorial_count}")
    print(f"📊 Total Sessions: {lecture_count + tutorial_count}")

def show_current_mappings(mappings):
    """Show current module mappings"""
    if mappings:
        print("\n🗂️  Current Module Mappings:")
        print("-" * 40)
        
        # Group by module code
        module_groups = {}
        for session, code in mappings.items():
            if code not in module_groups:
                module_groups[code] = []
            module_groups[code].append(session)
        
        for code in sorted(module_groups.keys()):
            print(f"  📘 {code}:")
            for session in sorted(module_groups[code]):
                print(f"    • {session}")

def show_exam_sessions(df, mappings):
    """Show all exam sessions found (for information)"""
    exam_sessions = []
    session_columns = get_all_session_columns(df)
    
    for idx, row in df.iterrows():
        for col in session_columns:
            session_text = str(row[col]) if not pd.isna(row[col]) else ""
            if session_text and is_exam_session(session_text):
                col_name = col if col in ["Morning", "Afternoon"] else f"Extra Column ({col})"
                exam_sessions.append(f"{row['Date']} {col_name}: {session_text}")
    
    if exam_sessions:
        print(f"\n📝 Exam Sessions Found (excluded from attendance):")
        print("-" * 50)
        for session in exam_sessions:
            print(f"  🎯 {session}")

def attendance_for_module():
    """Main function to calculate attendance"""
    print("🎓 Academic Attendance Calculator")
    print("=" * 35)
    
    # Load existing mappings
    mappings = load_module_mappings()
    
    # Show current mappings if any
    if mappings:
        show_current_mappings(mappings)
    
    # Get file path
    file_path = get_valid_file_path()
    if not file_path:
        print("❌ Exiting program")
        return
    
    # Load schedule
    df = load_schedule(file_path)
    if df is None:
        return
    
    # Show what columns we're processing
    session_columns = get_all_session_columns(df)
    print(f"\n📊 Processing columns: {', '.join(session_columns)}")
    
    # Show exam sessions
    show_exam_sessions(df, mappings)
    
    # Build module list (this will ask for mappings for new sessions)
    modules = build_module_list(df, mappings)
    
    if not modules:
        print("❌ No modules found in the schedule")
        return
    
    # Display available modules
    print("\n📘 Available Modules:")
    print("-" * 20)
    for i, mod in enumerate(modules, 1):
        print(f"{i:2d}. {mod}")
    
    # Get module choice
    choice_index = get_valid_module_choice(modules)
    selected_module = modules[choice_index]
    
    # Count sessions with debugging
    total_lectures = count_lectures_for_module(df, selected_module, mappings)
    
    if total_lectures == 0:
        print(f"❌ No sessions found for module: {selected_module}")
        return
    
    # Show session breakdown
    show_session_breakdown(df, selected_module, mappings)
    
    # Get missed sessions
    missed = get_valid_missed_sessions(total_lectures)
    
    # Calculate attendance
    attended = total_lectures - missed
    attendance_percentage = (attended / total_lectures) * 100
    
    # Calculate holiday allowance
    holiday_allowance, min_sessions_needed = calculate_holiday_allowance(total_lectures, missed)
    
    # Display results
    print("\n📊 Attendance Report")
    print("=" * 35)
    print(f"🧠 Module: {selected_module}")
    print(f"📚 Total Sessions: {total_lectures}")
    print(f"✅ Attended: {attended}")
    print(f"❌ Missed: {missed}")
    print(f"📈 Attendance: {attendance_percentage:.2f}%")
    
    # Holiday allowance information
    print(f"\n🏖️  Holiday Allowance (80% minimum):")
    print(f"📋 Minimum sessions needed: {min_sessions_needed}")
    if holiday_allowance > 0:
        print(f"🎉 Sessions you can still miss: {holiday_allowance}")
        print(f"🔢 Maximum total misses allowed: {total_lectures - min_sessions_needed}")
    else:
        sessions_over_limit = missed - (total_lectures - min_sessions_needed)
        print(f"⚠️  Already at/over limit! Sessions over: {sessions_over_limit}")
    
    # Attendance status
    if attendance_percentage >= 80:
        print(f"\n🟢 Status: GOOD - Meeting attendance requirements")
        if holiday_allowance > 0:
            print(f"   💡 You can miss {holiday_allowance} more session(s) and still maintain 80%")
    elif attendance_percentage >= 75:
        print(f"\n🟡 Status: WARNING - Close to minimum requirement")
        if holiday_allowance > 0:
            print(f"   ⚠️  You can only miss {holiday_allowance} more session(s) to maintain 80%")
        else:
            print(f"   🚨 Cannot miss any more sessions to maintain 80%")
    else:
        print(f"\n🔴 Status: CRITICAL - Below minimum attendance requirement")
        print(f"   📉 You need {min_sessions_needed - attended} more attended sessions to reach 80%")

def main():
    """Main program entry point"""
    try:
        attendance_for_module()
    except KeyboardInterrupt:
        print("\n\n👋 Program interrupted by user")
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        print("Please check your Excel file format and try again")

# Run the program
if __name__ == "__main__":
    main()
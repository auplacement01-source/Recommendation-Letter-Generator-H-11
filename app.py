from __future__ import annotations

import hmac
import io
import numbers
import os
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from supabase import Client, create_client


# ------------------------------ App configuration ------------------------------
st.set_page_config(
    page_title="Air University | Letter Desk",
    page_icon="✉️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    :root { --ink:#172b4d; --blue:#0c5c91; --pale:#eff6fb; --line:#dce5ed; }
    .stApp { background: linear-gradient(180deg,#f8fbfd 0%,#ffffff 360px); }
    [data-testid="stSidebar"] { background:#f1f6fa; border-right:1px solid #e2eaf0; }
    .hero { padding:1.35rem 1.55rem; border:1px solid #dce8f1; border-radius:18px;
            background:linear-gradient(120deg,#ffffff 0%,#eef6fb 100%); margin-bottom:1rem; }
    .hero h1 { color:#163b5b; margin:0 0 .3rem 0; font-size:2rem; }
    .hero p { color:#52677a; margin:0; }
    .subtle { color:#63788a; font-size:.92rem; }
    .step-card { padding:.8rem 1rem; border:1px solid #e0e9f0; border-radius:12px; background:#fff; }
    div[data-testid="stMetric"] { background:#fff; border:1px solid #e3ebf1; border-radius:12px; padding:.65rem .85rem; }
    .stButton > button[kind="primary"] { background:#0c5c91; border-color:#0c5c91; }
    </style>
    """,
    unsafe_allow_html=True,
)

# Program names are based on the three AU department pages selected for this app.
# The selected full name is stored and printed in the letter body; the code is used
# only in the letter number. Codes are explicit so BSAF can never become just "B".
PROGRAMS: dict[str, str] = {
    # Undergraduate programs
    "Bachelor of Science in International Relations": "BS-IR",
    "Bachelor of Business Administration (BBA-Honors)": "BBA",
    "Bachelor of Science in Aviation Management": "BSAvM",
    "Bachelor of Electrical Engineering": "BEE",
    "Bachelor of Mechatronics Engineering": "BMECH",
    "Bachelor of Mechanical Engineering": "BME",
    "Bachelor of Computer Engineering": "BCE",
    "Bachelor of Science in Computer Sciences": "BSCS",
    "Bachelor of Science in Information Technology": "BSIT",
    "Bachelor of Science in Cyber Security": "BSCybSec",
    "Bachelor of Science in Accounting and Finance": "BSAF",
    "Bachelor of Science in English": "BSENG",
    "Bachelor of Science in Physics": "BSPHY",
    "Bachelor of Science in Mathematics": "BSMATH",
    "Bachelor of Science in Psychology": "BS Psychology",
    "Bachelor of Biomedical Engineering": "BE Biomedical Engineering",
    "Bachelor of Science in Healthcare Management": "BSHCM",
    "Bachelor of Science in Tourism and Hospitality Management": "BSTHM",
    "Bachelor of Science in Gaming and Multimedia": "BSGM",
    "Bachelor of Science in Software Engineering": "BSSE",
    "Bachelor of Science in Artificial Intelligence": "BSAI",
    # Master's programs
    "MS / M.Phil in Linguistics and Literature": "MSLL",
    "M.Phil in Education": "MPhil-EDU",
    "Master of Business Administration": "MBA",
    "Master of Business Administration (Executive)": "MBA-EXEC",
    "Master of Science in Applied Physics": "MSAP",
    "Master of Science in Computer Sciences": "MSCS",
    "Master of Science in Aerospace Engineering": "MSAE",
    "Master of Science in Avionics Engineering": "MSAvE",
    "Master of Science in Information Security": "MSIS",
    "Master of Science in Electrical Engineering": "MSEE",
    "Master of Science in Mechatronics Engineering": "MSME",
    "Master of Science in Management Sciences": "MS-MGT",
    "Master of Science in Business Analytics": "MSBA",
    "Master of Science in Mathematics": "MSMATH",
    "Master of Science in Mechanical Engineering": "MSMECH",
    "Master of Science in Project Management": "MSPM",
    "Master of Science in Strategic Studies": "MSSS",
    "Master of Science in Data Science": "MSDS",
    "Master of Science in Gaming and Multimedia": "MSGM",
    "Master of Science in Artificial Intelligence": "MSAI",
    "Master of Science in Systems Security": "MSSYSSEC",
    "Master of Science in Clinical Psychology": "MSCP",
    "Master of Science in Cyber Security": "MSCY",
    "Master of Science in Biomedical Engineering": "MSBME",
    "Bachelor of Science in Clinical Psychology": "BSCP",
    "M.Phil in Professional Psychology": "MPhil-PP",
    "Bachelor of Communication and Media Studies (Morning)": "BCMS",
    "Master of Science in Teaching English to Speakers of Other Languages": "MS-TESOL",
    "PhD in Linguistics": "PhD-LING",
    "Master of Science in Counseling Psychology": "MS-CPSY",
    "Master of Science in Cyberpsychology": "MS-CYPSY",
    "PhD in Psychology": "PhD-PSY",
    "Bachelor of Science in Cyberpsychology": "BS-CYPSY",
    "Bachelor of Science in Financial Technology": "BSFinTech",
    "Bachelor of Science in Business and Information Technology": "BSBIT",
    "Bachelor of Science in Business Analytics": "BSBA",
    "Master of Science in Operations and Supply Chain Management": "MSOSCM",
    "Master of Science in Finance": "MSFIN",
    "PhD in Project Management": "PhD-PM",
    "Master of Science in International Relations and Strategic Studies": "MS IR & SS",
    "PhD in Strategic Studies": "PhD-SS",
    # PhD programs
    "PhD in Engineering": "PhD-ENG",
    "PhD in Management Sciences": "PhD-MS",
    "PhD in Linguistics and Literature": "PhD-LL",
    "PhD in Mathematics": "PhD-MATH",
    "PhD in Physics": "PhD-PHY",
    "PhD in Computer Science": "PhD-CS",
    "PhD in Information Security": "PhD-IS",
    "PhD in Cyber Security": "PhD-CY",
    "Master of Science in Management Sciences (HR/Finance/Marketing)": "MS-MGT",
    "PhD in Management Sciences (HR/Finance/Marketing)": "PhD-MGT",
    "Other / enter below": "OTHER",
}

# The user-provided department pages define the intended catalog for this app.
# Keep these canonical body names even where AU's page gives only a shorter title.
DEPARTMENT_PROGRAMS: dict[str, list[str]] = {
    "Department of Management Studies (DMS)": [
        "Bachelor of Science in Tourism and Hospitality Management",
        "Bachelor of Science in Healthcare Management",
        "Bachelor of Science in Accounting and Finance",
        "Bachelor of Science in Aviation Management",
        "Bachelor of Science in Financial Technology",
    ],
    "Department of Business Studies (DBS)": [
        "Bachelor of Business Administration (BBA-Honors)",
        "Master of Business Administration",
        "Bachelor of Science in Business and Information Technology",
        "Bachelor of Science in Business Analytics",
    ],
    "Department of Graduate Studies (DGS)": [
        "Master of Science in Management Sciences (HR/Finance/Marketing)",
        "Master of Science in Project Management",
        "Master of Science in Business Analytics",
        "Master of Science in Operations and Supply Chain Management",
        "Master of Science in Finance",
        "PhD in Project Management",
        "PhD in Management Sciences (HR/Finance/Marketing)",
    ],
    "Humanities": [
        "Bachelor of Science in English",
        "Bachelor of Communication and Media Studies (Morning)",
        "Master of Science in Teaching English to Speakers of Other Languages",
        "MS / M.Phil in Linguistics and Literature",
        "PhD in Linguistics",
        "PhD in Linguistics and Literature",
    ],
    "Psychology": [
        "Master of Science in Counseling Psychology",
        "Master of Science in Cyberpsychology",
        "Master of Science in Clinical Psychology",
        "PhD in Psychology",
        "Bachelor of Science in Psychology",
        "Bachelor of Science in Cyberpsychology",
        "Bachelor of Science in Clinical Psychology",
        "M.Phil in Professional Psychology",
    ],
    "International Relations and Strategic Studies": [
        "Bachelor of Science in International Relations",
        "Master of Science in International Relations and Strategic Studies",
        "PhD in Strategic Studies",
    ],
}

# Filter the wider AU catalog to the departments and programs supplied by the user.
PROGRAMS = {
    name: PROGRAMS[name]
    for department_programs in DEPARTMENT_PROGRAMS.values()
    for name in department_programs
}
PROGRAMS["Other / enter below"] = "OTHER"

PROGRAM_DEPARTMENTS: dict[str, str] = {
    name: department
    for department, department_programs in DEPARTMENT_PROGRAMS.items()
    for name in department_programs
}

PROGRAM_LEVELS: dict[str, str] = {
    name: (
        "Undergraduate" if name.startswith(("Bachelor",))
        else "Master's" if name.startswith(("MS ", "MS /", "M.Phil", "Master"))
        else "PhD" if name.startswith("PhD")
        else "Other"
    )
    for name in PROGRAMS
}

# Accepted forms in bulk spreadsheets, including legacy AU page wording and
# common abbreviations. Every alias resolves to one canonical full body name.
PROGRAM_ALIASES: dict[str, str] = {
    "BS IR": "Bachelor of Science in International Relations",
    "Bachelor of Science International Relations": "Bachelor of Science in International Relations",
    "BBA": "Bachelor of Business Administration (BBA-Honors)",
    "Bachelor of Business Administration": "Bachelor of Business Administration (BBA-Honors)",
    "BS AviMgt": "Bachelor of Science in Aviation Management",
    "BS Aviation Management": "Bachelor of Science in Aviation Management",
    "BEE": "Bachelor of Electrical Engineering",
    "BMECH": "Bachelor of Mechatronics Engineering",
    "BME": "Bachelor of Mechanical Engineering",
    "BCE": "Bachelor of Computer Engineering",
    "BSCS": "Bachelor of Science in Computer Sciences",
    "Bachelor of Science in Computer Science": "Bachelor of Science in Computer Sciences",
    "BSIT": "Bachelor of Science in Information Technology",
    "BSCybSec": "Bachelor of Science in Cyber Security",
    "BSAF": "Bachelor of Science in Accounting and Finance",
    "Bachelor Studies in Accounting and Finance": "Bachelor of Science in Accounting and Finance",
    "Bachelor of Science in Accounting & Finance": "Bachelor of Science in Accounting and Finance",
    "BS Accounting and Finance": "Bachelor of Science in Accounting and Finance",
    "BS Accounting & Finance": "Bachelor of Science in Accounting and Finance",
    "BSENG": "Bachelor of Science in English",
    "BSPHY": "Bachelor of Science in Physics",
    "BSMATH": "Bachelor of Science in Mathematics",
    "BS Psychology": "Bachelor of Science in Psychology",
    "BSHM": "Bachelor of Science in Healthcare Management",
    "BS Healthcare Management": "Bachelor of Science in Healthcare Management",
    "BSTHM": "Bachelor of Science in Tourism and Hospitality Management",
    "BS Tourism and Hospitality Management": "Bachelor of Science in Tourism and Hospitality Management",
    "BSGM": "Bachelor of Science in Gaming and Multimedia",
    "Bachelor of Science in Gaming & Multimedia": "Bachelor of Science in Gaming and Multimedia",
    "BSSE": "Bachelor of Science in Software Engineering",
    "Bachelor of Science in Software Engineering as HEC Approved Non PEC Program": "Bachelor of Science in Software Engineering",
    "BSAI": "Bachelor of Science in Artificial Intelligence",
    "Bachelor Studies in Accounting and Finance (BSAF)": "Bachelor of Science in Accounting and Finance",
    "Bachelor of Communication & Media Studies (Morning)": "Bachelor of Communication and Media Studies (Morning)",
    "MS in TESOL": "Master of Science in Teaching English to Speakers of Other Languages",
    "MS /M.Phil in Linguistics and Literature": "MS / M.Phil in Linguistics and Literature",
    "PhD Linguistics": "PhD in Linguistics",
    "MS Counseling Psychology": "Master of Science in Counseling Psychology",
    "MS Cyberpsychology": "Master of Science in Cyberpsychology",
    "MS Clinical Psychology": "Master of Science in Clinical Psychology",
    "PhD Psychology": "PhD in Psychology",
    "BS Cyberpsychology": "Bachelor of Science in Cyberpsychology",
    "BSFinTech": "Bachelor of Science in Financial Technology",
    "BSBIT": "Bachelor of Science in Business and Information Technology",
    "BSBA": "Bachelor of Science in Business Analytics",
    "MSOSCM": "Master of Science in Operations and Supply Chain Management",
    "MSFIN": "Master of Science in Finance",
    "PhD-PM": "PhD in Project Management",
    "PhD-MGT": "PhD in Management Sciences (HR/Finance/Marketing)",
    "MS IR & SS": "Master of Science in International Relations and Strategic Studies",
    "MS in IR & SS": "Master of Science in International Relations and Strategic Studies",
    "PhD Strategic Studies": "PhD in Strategic Studies",
    "MSCP": "Master of Science in Clinical Psychology",
    "MS Clinical Psychology": "Master of Science in Clinical Psychology",
    "MBA": "Master of Business Administration",
    "Master of Business Administration 2 Years": "Master of Business Administration",
    "MS Business Analytics": "Master of Science in Business Analytics",
    "MS in Applied Physics": "Master of Science in Applied Physics",
    "MS in Strategic Studies": "Master of Science in Strategic Studies",
    "MS Strategic Studies (MSSS)": "Master of Science in Strategic Studies",
    "MS Cyber Security": "Master of Science in Cyber Security",
    "MS Bio-Medical Engineering": "Master of Science in Biomedical Engineering",
    "Masters of Science in Systems Security": "Master of Science in Systems Security",
    "Master of Science in Management Sciences (18 Months)": "Master of Science in Management Sciences",
    "PhD Engineering": "PhD in Engineering",
    "PhD Management Sciences": "PhD in Management Sciences",
    "PhD Linguistics and Literature": "PhD in Linguistics and Literature",
    "PhD Mathematics": "PhD in Mathematics",
    "PhD Physics": "PhD in Physics",
    "PhD Computer Science": "PhD in Computer Science",
    "PhD Information Security": "PhD in Information Security",
    "PhD Cyber Security": "PhD in Cyber Security",
    "MS Management (HR/Finance/Marketing)": "Master of Science in Management Sciences (HR/Finance/Marketing)",
    "MS-Management (HR/Finance/Marketing)": "Master of Science in Management Sciences (HR/Finance/Marketing)",
    "MS Management Sciences": "Master of Science in Management Sciences (HR/Finance/Marketing)",
    "PhD Management Sciences (HR/Finance/Marketing)": "PhD in Management Sciences (HR/Finance/Marketing)",
    "BS English": "Bachelor of Science in English",
    "BS Clinical Psychology": "Bachelor of Science in Clinical Psychology",
    "MPhil Professional Psychology": "M.Phil in Professional Psychology",
}


def program_key(value: Any) -> str:
    """Normalize labels and abbreviations for reliable spreadsheet matching."""
    text = "" if value is None else str(value).strip()
    return re.sub(r"[^a-z0-9]+", "", text.lower())


PROGRAM_LOOKUP: dict[str, str] = {}
for _program_name, _program_code in PROGRAMS.items():
    if _program_name != "Other / enter below":
        PROGRAM_LOOKUP[program_key(_program_name)] = _program_name
        PROGRAM_LOOKUP[program_key(_program_code)] = _program_name
for _alias, _canonical_name in PROGRAM_ALIASES.items():
    if _canonical_name in PROGRAMS:
        PROGRAM_LOOKUP[program_key(_alias)] = _canonical_name

SEMESTERS = list(range(1, 13))
DEPARTMENT_PRESETS = [
    "IBD/AU/DSA/PLAC/H-11",
    "IBD/AU/PLAC/H-11",
    "Custom prefix",
]

# This intentionally small, editable offline dictionary makes a best-effort guess.
# Names can be unisex or differently gendered by family; always allow a human override.
FEMALE_NAMES = set("""
Ayesha Fatima Zainab Maryam Mariam Amina Aaminah Amna Sana Hira Iqra Mahnoor Noor
Zubda Zubda Fatima Kazam Imaan Iman Amnah Amna Siddiqui Laiba Eman Emaan Anum
Anam Hania Haniya Dua Aleena Alina Aliza Areeba Arooba Arooj Mehak Mahak Saira
Sadia Nadia Rabia Rabiya Sidra Saba Kiran Kiran Bushra Beenish Sehrish Zoya Zara
Hafsa Hafsa Anaya Anayah Minahil Minal Nimra Nimrah Momina Mominah Muskan Rida
Ridaa Sehr Hoorain Aiman Ayra Aiza Aqsa Aqsa Rameen Rimsha Saniya Sania Komal
Maham Maira Misha Huda Hiba Hifza Esha Ayesha Anisha Areej Erum Irum Sumbal
Shazia Farah Farwa Fiza Fariha Faryal Arisha Hareem Hareem Alishba Ayesha Noor
Aqsa Javeria Jaweria Khadija Khadijah Aaliyah Aliya Sehrish Muneeba Muneebah
Nawal Nida Noreen Naila Naima Naeema Samina Samra Samreen Shumaila Sumaira
Tahira Tehreem Tuba Tooba Urooj Yasmin Yusra Zunaira Zunairah Zainab Zahra
Abeer Arooj Afsheen Afreen Alishba Aqila Bushra Faiza Fiza Gul Hina Iffat
Jannat Kanwal Lubna Mahwish Mehwish Nargis Parisa Rabail Rukhsana Saman Shabana
Shagufta Sharmeen Shazia Tehmina Uzma Wajiha Warda Wania Yumna
""".split())
MALE_NAMES = set("""
Muhammad Mohammad Mohammed Ahmed Ahmad Ali Hassan Hussain Husain Abdullah Abdul
Abdulrahman Usman Uthman Umar Omar Bilal Hamza Ibrahim Ismail Ismail Imran Imran
Ahsan Asad Salman Zain Zayn Saad Shahzaib Shehryar Talha Taha Tayyab Tayab Yasir
Yasir Farhan Fahad Faizan Faisal Fahim Fahad Muneeb Muneeb Daniyal Danish Dawood
Dawood Awais Owais Arslan Arsalan Haris Harris Huzaifa Huzayfa Junaid Junayd
Kamran Kashif Khurram Luqman Moiz Moeed Mustafa Musab Noman Nouman Qasim Qasim
Raza Rehan Rizwan Saif Sameer Samir Shahbaz Shahid Sohaib Suhaib Usama Waleed
Walid Yousuf Yusuf Zubair Zain Ali Abbas Adnan Adeel Aftab Akbar Amir Ameer
Anas Aqib Arham Asif Atif Azlan Azhar Babar Basit Burhan Ehtesham Ehsan
Faisal Ghazi Ghulam Haider Hamid Haroon Hashim Hammad Haseeb Ilyas Irfan Jibran
Khalid Majid Malik Mansoor Masood Mehmood Nabeel Nabil Naveed Nawaz Niaz
Parvez Pervez Raheel Rashid Rauf Shayan Sohail Suhail Tariq Waqar Waqas
""".split())

LOGO_PATH = Path(__file__).with_name("air_university_logo.png")


def secret_value(section: str, key: str, env_name: str | None = None) -> str:
    """Read a value from Streamlit secrets, falling back to an environment variable."""
    try:
        section_data = st.secrets.get(section, {})
        value = section_data.get(key, "") if hasattr(section_data, "get") else ""
        if value:
            return str(value).strip()
    except Exception:
        pass
    if env_name:
        return os.getenv(env_name, "").strip()
    return ""


@st.cache_resource(show_spinner=False)
def get_supabase(url: str, key: str) -> Client:
    return create_client(url, key)


def require_password() -> None:
    configured = secret_value("app", "password", "APP_PASSWORD")
    if not configured:
        st.error("**Setup required:** add a strong shared password as `APP_PASSWORD` in Streamlit Secrets before using this app.")
        st.code('[app]\npassword = "replace-with-a-long-random-password"', language="toml")
        st.stop()
    if st.session_state.get("authenticated"):
        return
    st.markdown(
        '<div class="hero"><h1>Air University · Letter Desk</h1>'
        '<p>Secure recommendation and internship letter workspace</p></div>',
        unsafe_allow_html=True,
    )
    with st.form("login_form"):
        entered = st.text_input("Office password", type="password")
        submitted = st.form_submit_button("Unlock workspace", type="primary", use_container_width=True)
    if submitted:
        if hmac.compare_digest(entered.encode("utf-8"), configured.encode("utf-8")):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("That password does not match. Please try again.")
    st.caption("This is a simple shared-password gate for a small office workflow. Keep the app private and do not share the Supabase secret key.")
    st.stop()


def get_client_or_stop() -> Client:
    url = secret_value("supabase", "url", "SUPABASE_URL")
    key = (
        secret_value("supabase", "key", "SUPABASE_SECRET_KEY")
        or secret_value("supabase", "service_role_key", "SUPABASE_SERVICE_ROLE_KEY")
    )
    if not url or not key:
        st.error("Supabase is not configured yet. Add your project URL and server-side secret key in Streamlit Secrets, then restart the app.")
        st.code('[supabase]\nurl = "https://YOUR-PROJECT-REF.supabase.co"\nkey = "YOUR-SUPABASE-SERVER-ONLY-SECRET-KEY"', language="toml")
        st.info("Do not put the key in GitHub, in app.py, or in a browser-side script. Follow the README setup steps.")
        st.stop()
    try:
        return get_supabase(url, key)
    except Exception as exc:
        st.error(f"Could not connect to Supabase: {exc}")
        st.stop()


def ordinal(n: int) -> str:
    if 10 <= n % 100 <= 20:
        suffix = "th"
    else:
        suffix = {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th")
    return f"{n}{suffix}"


def infer_gender(full_name: str) -> tuple[str, str]:
    """Return a best-effort value from the first name; unknown names remain neutral."""
    tokens = re.findall(r"[A-Za-zÀ-ÿ'-]+", (full_name or "").strip())
    if not tokens:
        return "They", "No name to analyze"
    # Pakistani full names may start with a given name, title, or compound name.
    first = tokens[0].strip("'-").title()
    normalized = first.replace("'", "")
    if normalized in FEMALE_NAMES:
        return "She", "Best-effort name match"
    if normalized in MALE_NAMES:
        return "He", "Best-effort name match"
    if len(tokens) > 1:
        second = tokens[1].strip("'-").title().replace("'", "")
        if second in FEMALE_NAMES:
            return "She", "Best-effort match on second name"
        if second in MALE_NAMES:
            return "He", "Best-effort match on second name"
    return "They", "Name not recognized—please choose pronouns"


def resolve_gender(name: str, explicit: Any = None) -> tuple[str, str]:
    options = {"female": "She", "she": "She", "f": "She", "male": "He", "he": "He", "m": "He", "they": "They", "other": "They"}
    if explicit is not None and not pd.isna(explicit) and str(explicit).strip():
        given = str(explicit).strip().lower()
        if given in options:
            return options[given], "Excel override"
        # Also accept a full selector label, e.g. "Female (She/Her)".
        if "she" in given or "female" in given:
            return "She", "Excel override"
        if "he" in given or "male" in given:
            return "He", "Excel override"
        if "they" in given:
            return "They", "Excel override"
    return infer_gender(name)


def safe_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, numbers.Real) and not isinstance(value, bool):
        if float(value).is_integer():
            return str(int(value))
    return str(value).strip()


ROSTER_HEADER_ALIASES: dict[str, set[str]] = {
    "student_name": {
        "studentname", "name", "fullname", "studentfullname", "nameofstudent",
        "studentsname", "studentsfullname", "candidatename", "candidate", "applicantname",
        "scholarname", "traineename",
    },
    "registration_id": {
        "registrationid", "registrationnumber", "registrationno", "registration",
        "reg", "regid", "regno", "regnumber", "registrationidno", "studentregistrationid",
        "studentregistrationnumber", "studentregistrationno", "studentid", "studentnumber",
        "studentno", "roll", "rollno",
        "rollnumber", "enrollmentno", "enrolmentno", "enrollmentnumber",
        "enrolmentnumber", "universityid", "registrationnumberid",
    },
    "program": {
        "program", "programme", "programname", "programtitle", "degree", "degreeprogram",
        "degreeprogramme", "degreename", "degreetitle", "course", "coursename",
        "courseofstudy", "qualification", "major", "discipline", "academicprogram",
    },
    "semester": {
        "semester", "sem", "semesterno", "semesternumber", "currentsemester", "term",
    },
    "gender": {"gender", "sex", "pronoun", "pronouns"},
    "company_name": {
        "company", "companyname", "organization", "organizationname", "organisation",
        "organisationname", "employer", "employername", "firm", "internshipcompany",
        "hostorganization", "hostorganisation",
    },
    "city": {"city", "cityname", "location"},
    "recipient_mode": {
        "recipientmode", "recipienttype", "lettermode", "lettertype", "companyspecific",
        "generalletter",
    },
}
ROSTER_HEADER_LOOKUP = {
    re.sub(r"[^a-z0-9]", "", alias.casefold()): field
    for field, aliases in ROSTER_HEADER_ALIASES.items()
    for alias in aliases
}


def roster_header_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", safe_text(value).casefold())


def roster_field_for_header(value: Any) -> str | None:
    token = roster_header_token(value)
    direct = ROSTER_HEADER_LOOKUP.get(token)
    if direct:
        return direct
    if "student" in token and "name" in token:
        return "student_name"
    if ("registration" in token or token.startswith("reg")) and any(x in token for x in ("id", "no", "num", "number")):
        return "registration_id"
    if ("enrollment" in token or "enrolment" in token) and any(x in token for x in ("id", "no", "num", "number")):
        return "registration_id"
    if "roll" in token and any(x in token for x in ("id", "no", "num", "number")):
        return "registration_id"
    if any(x in token for x in ("program", "programme", "degree", "course", "qualification")):
        return "program"
    if "semester" in token or token == "sem":
        return "semester"
    if "gender" in token or "pronoun" in token or token == "sex":
        return "gender"
    if any(x in token for x in ("company", "organization", "organisation", "employer", "firm")):
        return "company_name"
    if "city" in token or token == "location":
        return "city"
    return None


def detect_roster_header(raw: pd.DataFrame, scan_rows: int = 20) -> tuple[int, dict[str, list[int]], int] | None:
    """Find the strongest likely header row, even when a sheet starts with a title."""
    best: tuple[int, dict[str, list[int]], int] | None = None
    for row_index in range(min(scan_rows, len(raw))):
        mapping: dict[str, list[int]] = {}
        for column_index, value in enumerate(raw.iloc[row_index].tolist()):
            field = roster_field_for_header(value)
            if field:
                mapping.setdefault(field, []).append(column_index)
        if not {"student_name", "registration_id"}.issubset(mapping):
            continue
        score = len(mapping) * 10 + len(mapping.get("student_name", [])) + len(mapping.get("registration_id", []))
        candidate = (row_index, mapping, score)
        if best is None or score > best[2]:
            best = candidate
    return best


def _canonicalize_roster_sheet(raw: pd.DataFrame, sheet_name: str) -> tuple[pd.DataFrame, dict[str, Any]] | None:
    detected = detect_roster_header(raw)
    if detected is None:
        return None
    header_row, mapping, score = detected
    output: list[dict[str, Any]] = []
    for row_index in range(header_row + 1, len(raw)):
        values = raw.iloc[row_index].tolist()
        record: dict[str, Any] = {}
        for field, column_indexes in mapping.items():
            record[field] = next(
                (safe_text(values[column]) for column in column_indexes
                 if column < len(values) and safe_text(values[column])),
                "",
            )
        if not any(record.values()):
            continue
        record["_source_sheet"] = sheet_name
        record["_source_row"] = row_index + 1
        output.append(record)
    columns = list(ROSTER_HEADER_ALIASES) + ["_source_sheet", "_source_row"]
    frame = pd.DataFrame(output, columns=columns).fillna("")
    mapped_headers = {
        field: [safe_text(raw.iat[header_row, col]) for col in indexes]
        for field, indexes in mapping.items()
    }
    return frame, {"sheet": sheet_name, "header_row": header_row + 1, "mapped_headers": mapped_headers, "score": score}


def parse_uploaded_roster(filename: str, content: bytes) -> tuple[pd.DataFrame, list[dict[str, Any]]]:
    """Read CSV/XLSX rosters with varied headers and optional title rows/sheets."""
    if filename.casefold().endswith(".csv"):
        try:
            raw = pd.read_csv(io.BytesIO(content), header=None, dtype=object, encoding="utf-8-sig")
        except UnicodeDecodeError:
            raw = pd.read_csv(io.BytesIO(content), header=None, dtype=object, encoding="latin-1")
        sheets = {"CSV": raw}
    else:
        sheets = pd.read_excel(io.BytesIO(content), sheet_name=None, header=None, dtype=object)
    parsed: list[pd.DataFrame] = []
    details: list[dict[str, Any]] = []
    for sheet_name, raw in sheets.items():
        if raw is None or raw.empty:
            continue
        result = _canonicalize_roster_sheet(raw, str(sheet_name))
        if result:
            frame, detail = result
            if not frame.empty:
                parsed.append(frame)
                details.append(detail)
    if not parsed:
        raise ValueError(
            "Could not find a roster header row with both a student-name column and a registration-ID column. "
            "Common accepted headings include Name / Student Name and Registration ID / Reg No / Roll No."
        )
    combined = pd.concat(parsed, ignore_index=True).fillna("")
    return combined, details


def make_row(
    *,
    student_name: str,
    registration_id: str,
    gender: str,
    program: str,
    program_code: str,
    semester: int,
    department_prefix: str,
    letter_date: date,
    recipient_mode: str,
    company_name: str = "",
    city: str = "",
) -> dict[str, Any]:
    return {
        "student_name": student_name.strip(),
        "registration_id": registration_id.strip(),
        "gender": gender,
        "program": program,
        "program_code": program_code,
        "semester": int(semester),
        "department_prefix": department_prefix.strip().strip("/"),
        "letter_date": letter_date.isoformat(),
        "recipient_mode": recipient_mode,
        "company_name": company_name.strip() if recipient_mode == "Company / organization" else "",
        "city": city.strip() if recipient_mode == "Company / organization" else "",
    }


class PartialSaveError(RuntimeError):
    def __init__(self, message: str, saved_records: list[dict[str, Any]]):
        super().__init__(message)
        self.saved_records = saved_records


def is_unique_number_conflict(exc: Exception) -> bool:
    code = safe_text(getattr(exc, "code", ""))
    message = str(exc).casefold()
    return code == "23505" or "23505" in message or "duplicate key" in message or "unique constraint" in message


def allocate_and_save(client: Client, rows: list[dict[str, Any]], include_semester: bool) -> list[dict[str, Any]]:
    """Allocate per-office/program/year serials atomically and never insert a duplicate number."""
    saved: list[dict[str, Any]] = []
    for row in rows:
        yr = date.fromisoformat(row["letter_date"]).strftime("%y")
        try:
            for attempt in range(25):
                rpc_result = client.rpc(
                    "next_letter_serial",
                    {
                        "p_department_prefix": row["department_prefix"],
                        "p_program_code": row["program_code"],
                        "p_year_2d": yr,
                    },
                ).execute()
                sequence = rpc_result.data
                if isinstance(sequence, list):
                    sequence = sequence[0]
                if isinstance(sequence, dict):
                    sequence = next(iter(sequence.values()))
                sequence = int(sequence)
                parts = [row["department_prefix"], row["program_code"], yr]
                if include_semester:
                    parts.append(f"S{row['semester']}")
                parts.append(f"{sequence:03d}")
                record = {
                    **row,
                    "year_2d": yr,
                    "sequence_no": sequence,
                    "letter_number": "/".join(parts),
                }
                try:
                    result = client.table("letters").insert(record).execute()
                except Exception as insert_error:
                    if is_unique_number_conflict(insert_error) and attempt < 24:
                        continue
                    raise
                if not result.data:
                    raise RuntimeError("Supabase did not return the inserted letter record.")
                saved.append(result.data[0])
                break
            else:
                raise RuntimeError("Could not find an unused letter number after 25 allocation attempts.")
        except Exception as exc:
            raise PartialSaveError(
                f"Could not save {row.get('student_name', 'student')} ({row.get('registration_id', '')}): {exc}",
                saved,
            ) from exc
    return saved


def format_letter_date(value: Any) -> str:
    if isinstance(value, date):
        d = value
    else:
        d = date.fromisoformat(str(value)[:10])
    return f"{ordinal(d.day)} {d.strftime('%B, %Y')}"


def remove_table_borders(table) -> None:
    tbl_pr = table._tbl.tblPr
    borders = tbl_pr.first_child_found_in("w:tblBorders")
    if borders is None:
        from docx.oxml import OxmlElement
        borders = OxmlElement("w:tblBorders")
        tbl_pr.append(borders)
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = "w:" + edge
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "nil")


def set_cell_margins(cell, top=0, start=0, bottom=0, end=0) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def style_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Inches(0.72)
    section.bottom_margin = Inches(0.72)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = RGBColor(25, 35, 45)
    normal.paragraph_format.space_after = Pt(8)
    normal.paragraph_format.line_spacing = 1.12


def add_letter(doc: Document, record: dict[str, Any], final_page: bool = False) -> None:
    header = doc.add_table(rows=1, cols=2)
    header.autofit = False
    header.columns[0].width = Inches(4.55)
    header.columns[1].width = Inches(2.0)
    remove_table_borders(header)
    left, right = header.rows[0].cells
    left.width = Inches(4.55)
    right.width = Inches(2.0)
    left.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
    right.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_margins(left, 0, 0, 0, 0)
    set_cell_margins(right, 0, 0, 0, 0)
    left_p = left.paragraphs[0]
    left_p.paragraph_format.space_after = Pt(2)
    if LOGO_PATH.exists():
        left_p.add_run().add_picture(str(LOGO_PATH), width=Inches(1.0))
    else:
        run = left_p.add_run("AIR UNIVERSITY")
        run.bold = True
        run.font.color.rgb = RGBColor(0, 91, 145)
        run.font.size = Pt(12)
    number_p = left.add_paragraph()
    number_p.paragraph_format.space_before = Pt(0)
    number_p.paragraph_format.space_after = Pt(0)
    nr = number_p.add_run(safe_text(record.get("letter_number")))
    nr.font.size = Pt(8.5)
    date_p = right.paragraphs[0]
    date_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    date_p.paragraph_format.space_after = Pt(0)
    dr = date_p.add_run(format_letter_date(record.get("letter_date", date.today().isoformat())))
    dr.font.size = Pt(9.5)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(5)
    mode = safe_text(record.get("recipient_mode"))
    company = safe_text(record.get("company_name"))
    city = safe_text(record.get("city"))
    if mode == "Company / organization" and company:
        to = doc.add_paragraph()
        to.paragraph_format.space_after = Pt(0)
        r = to.add_run("HR Manager")
        r.bold = True
        org = doc.add_paragraph()
        org.paragraph_format.space_after = Pt(0)
        org.add_run(company)
        if city:
            loc = doc.add_paragraph()
            loc.paragraph_format.space_after = Pt(0)
            loc.add_run(city)
    salutation = doc.add_paragraph()
    salutation.alignment = WD_ALIGN_PARAGRAPH.CENTER
    salutation.paragraph_format.space_before = Pt(8)
    salutation.paragraph_format.space_after = Pt(15)
    s = salutation.add_run("To Whom it May Concern")
    s.bold = True

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(20)
    t = title.add_run("REQUEST FOR INTERNSHIP OPPORTUNITY FOR AIR UNIVERSITY STUDENT")
    t.bold = True
    t.underline = True
    t.font.size = Pt(10.5)

    name = safe_text(record.get("student_name"))
    reg = safe_text(record.get("registration_id"))
    gender = safe_text(record.get("gender")) or "They"
    # Sentence pronoun in the source template is subject case: He / She (or neutral They).
    program = safe_text(record.get("program"))
    semester = int(record.get("semester") or 1)
    body = doc.add_paragraph()
    body.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    bold_prefix = body.add_run(f"{name}, Registration No. {reg},")
    bold_prefix.bold = True
    body.add_run(
        f" is a student at Air University, a federally chartered public sector university, recognized by HEC. "
        f"{gender} is currently enrolled in the {ordinal(semester)} Semester of {program}."
    )

    organization_phrase = f"at {company}" if mode == "Company / organization" and company else "at your esteemed organization"
    paragraph = doc.add_paragraph(
        f"It is requested that the student may please be provided with an internship opportunity {organization_phrase}. "
        "The student will benefit a great deal by working at an organization with a good reputation, which would go a long way in professional development."
    )
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_paragraph("If you have any internship/ job opportunity in future, please share with us.")

    sign_table = doc.add_table(rows=1, cols=2)
    sign_table.autofit = False
    sign_table.columns[0].width = Inches(3.65)
    sign_table.columns[1].width = Inches(2.9)
    remove_table_borders(sign_table)
    blank, sign = sign_table.rows[0].cells
    blank.width = Inches(3.65)
    sign.width = Inches(2.9)
    for cell in (blank, sign):
        set_cell_margins(cell, 0, 0, 0, 0)
    sign_p = sign.paragraphs[0]
    sign_p.paragraph_format.space_before = Pt(78)
    sign_p.paragraph_format.space_after = Pt(0)
    sign_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    officer = safe_text(record.get("officer_name")) or "MOBEEN JAMSHED"
    role = safe_text(record.get("officer_title")) or "Assistant Director"
    office = safe_text(record.get("office_name")) or "Placement & Alumni Affairs"
    rr = sign_p.add_run(f"({officer})")
    rr.bold = True
    sign_p.add_run(f"\n{role}\n{office}")


def build_docx(records: list[dict[str, Any]]) -> bytes:
    doc = Document()
    style_document(doc)
    for index, record in enumerate(records):
        if index:
            doc.add_page_break()
        add_letter(doc, record)
    stream = io.BytesIO()
    doc.save(stream)
    return stream.getvalue()


def get_year(date_text: str) -> str:
    return date.fromisoformat(date_text).strftime("%y")


def normalize_program(raw: Any, default_program: str) -> tuple[str, str]:
    text = safe_text(raw)
    if not text:
        text = safe_text(default_program)
    canonical_name = PROGRAM_LOOKUP.get(program_key(text))
    if canonical_name:
        return canonical_name, PROGRAMS[canonical_name]
    if program_key(text) == program_key("Other / enter below"):
        return "Other / enter below", "OTHER"
    # For a custom course name, create a stable short code from its words.
    code = re.sub(r"[^A-Z0-9]", "", "".join(part[:1] for part in re.findall(r"[A-Za-z0-9]+", text.upper())))[:8]
    return text, code or "OTHER"


def normalize_semester(raw: Any, default_semester: int) -> int:
    text = safe_text(raw)
    if not text:
        return int(default_semester)
    match = re.search(r"\d+", text)
    if not match:
        raise ValueError(f"Invalid semester value: {text}")
    value = int(match.group())
    if not 1 <= value <= 12:
        raise ValueError(f"Semester must be between 1 and 12: {text}")
    return value


def template_excel() -> bytes:
    example = pd.DataFrame(columns=[
        "student_name", "registration_id", "program", "semester", "gender",
        "company_name", "city", "recipient_mode",
    ])
    info = pd.DataFrame({
        "Column": ["student_name", "registration_id", "program", "semester", "gender", "company_name", "city", "recipient_mode"],
        "Required?": ["Yes", "Yes", "No (uses selected default)", "No (uses selected default)", "No (auto-detects from name)", "No", "No", "No (uses selected default)"],
        "Example / accepted values": ["Ayesha Khan", "2504332", "Master of Science in Clinical Psychology or MSCP", "3 or 3rd Semester", "Female, Male, or They; blank = best-effort", "PIMS", "Islamabad", "Company / organization or General"],
    })
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        example.to_excel(writer, index=False, sheet_name="Students")
        info.to_excel(writer, index=False, sheet_name="Read me")
    return out.getvalue()


def excel_bytes(records: list[dict[str, Any]]) -> bytes:
    df = pd.DataFrame(records)
    preferred = [
        "letter_number", "letter_date", "student_name", "registration_id", "gender",
        "program", "program_code", "semester", "recipient_mode", "company_name", "city",
        "department_prefix", "year_2d", "sequence_no", "created_at",
    ]
    columns = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df[columns].to_excel(writer, index=False, sheet_name="Letter records")
        ws = writer.sheets["Letter records"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_length = min(max(len(str(cell.value or "")) for cell in column_cells) + 2, 45)
            ws.column_dimensions[column_cells[0].column_letter].width = max(12, max_length)
    return out.getvalue()


def fetch_all_records(client: Client, limit: int = 1000) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    start = 0
    while True:
        batch = client.table("letters").select("*").order("created_at", desc=True).range(start, start + limit - 1).execute().data or []
        rows.extend(batch)
        if len(batch) < limit:
            break
        start += limit
    return rows


def date_from_record(record: dict[str, Any]) -> str:
    value = record.get("letter_date")
    if isinstance(value, str):
        return value[:10]
    return date.today().isoformat()


def app_main() -> None:
    require_password()
    client = get_client_or_stop()

    st.markdown(
        '<div class="hero"><h1>Air University · Letter Desk</h1>'
        '<p>Create, store, find and reprint internship recommendation letters — one at a time or from an Excel roster.</p></div>',
        unsafe_allow_html=True,
    )

    with st.sidebar:
        st.subheader("Letter settings")
        prefix_choice = st.selectbox("Department / office prefix", DEPARTMENT_PRESETS, index=0)
        if prefix_choice == "Custom prefix":
            department_prefix = st.text_input("Enter department prefix", value="IBD/AU/DSA/PLAC/H-11")
        else:
            department_prefix = prefix_choice
        include_semester = st.checkbox("Include semester in letter number", value=False, help="Off matches your example: PREFIX/PROGRAM/YY/001")
        st.caption("Example: `IBD/AU/DSA/PLAC/H-11/BBA/26/001`. The year follows the letter date; serials increase separately for each office, program, and year.")
        with st.expander("Signature block", expanded=False):
            officer_name = st.text_input("Officer name", value="MOBEEN JAMSHED")
            officer_title = st.text_input("Officer title", value="Assistant Director")
            office_name = st.text_input("Office", value="Placement & Alumni Affairs")
        st.divider()
        if st.button("Lock workspace", use_container_width=True):
            st.session_state["authenticated"] = False
            st.rerun()
        st.caption("Tip: Keep this workspace password-protected and limit access to authorized office staff.")

    common = {
        "department_prefix": department_prefix,
        "officer_name": officer_name,
        "officer_title": officer_title,
        "office_name": office_name,
    }
    tabs = st.tabs(["Create one", "Bulk from Excel", "Find / reprint", "Records & export"])

    # ------------------------------ Single letter ------------------------------
    with tabs[0]:
        st.markdown("### Create an individual letter")
        st.caption("Enter the student details. Name-based pronoun matching is an offline best-effort guess; review the pronoun selector before generating.")
        col1, col2 = st.columns(2)
        with col1:
            student_name = st.text_input(
                "Student full name *",
                placeholder="e.g. Imaan Shah",
                key="single_name",
                on_change=lambda: st.session_state.update(
                    single_gender=infer_gender(st.session_state.get("single_name", ""))[0]
                ),
            )
            registration_id = st.text_input("Registration ID *", placeholder="e.g. 2504269", key="single_reg")
            selectable_programs = list(PROGRAMS.keys())
            default_program_index = selectable_programs.index("Master of Science in Clinical Psychology")
            program_label = st.selectbox(
                "Program *",
                selectable_programs,
                index=default_program_index,
                format_func=lambda name: (
                    name if name == "Other / enter below"
                    else f"{PROGRAM_DEPARTMENTS[name]} · {name} ({PROGRAMS[name]})"
                ),
                key="single_program",
            )
            custom_program = ""
            if program_label == "Other / enter below":
                custom_program = st.text_input("Program name", placeholder="Enter full program name")
            else:
                st.caption(f"Letter-number code: **{PROGRAMS[program_label]}** · The full program name is used in the letter body.")
            semester = st.selectbox("Semester *", SEMESTERS, index=2, format_func=lambda n: f"{ordinal(n)} Semester", key="single_semester")
        with col2:
            detected_gender, gender_reason = infer_gender(student_name)
            pronoun_choice = st.selectbox(
                "Pronoun for the letter *",
                ["She", "He", "They"],
                index=["She", "He", "They"].index(detected_gender),
                format_func=lambda v: {"She": "She / her", "He": "He / him", "They": "They / them"}[v],
                key="single_gender",
            )
            if student_name:
                if detected_gender == "They":
                    st.warning("We could not confidently match this name. Please select the correct pronoun above.")
                else:
                    st.info(f"Suggested from the name: **{detected_gender}** · {gender_reason}. You can change it.")
            recipient_mode = st.radio("Letter recipient", ["General (any organization)", "Company / organization"], horizontal=True, key="single_recipient")
            company_name = ""
            city = ""
            if recipient_mode == "Company / organization":
                company_name = st.text_input("Company / organization name *", placeholder="e.g. PIMS")
                city = st.text_input("City (optional)", placeholder="e.g. Islamabad")
            letter_date = st.date_input("Letter date", value=date.today(), format="DD/MM/YYYY", key="single_date")
        single_program = custom_program.strip() if program_label == "Other / enter below" else program_label
        if program_label == "Other / enter below":
            _, single_program_code = normalize_program(single_program, single_program)
        else:
            single_program_code = PROGRAMS[program_label]
        generate_one = st.button("Generate and save letter", type="primary", use_container_width=True, key="generate_single")
        if generate_one:
            issues = []
            if not student_name.strip(): issues.append("Enter the student's full name.")
            if not registration_id.strip(): issues.append("Enter the registration ID.")
            if not department_prefix.strip(): issues.append("Enter a department prefix in the sidebar.")
            if program_label == "Other / enter below" and not single_program: issues.append("Enter the custom program name.")
            if recipient_mode == "Company / organization" and not company_name.strip(): issues.append("Enter the company or organization name, or choose General.")
            if issues:
                for item in issues: st.error(item)
            else:
                row = make_row(
                    student_name=student_name, registration_id=registration_id, gender=pronoun_choice,
                    program=single_program, program_code=single_program_code, semester=semester,
                    department_prefix=department_prefix, letter_date=letter_date,
                    recipient_mode=recipient_mode, company_name=company_name, city=city,
                )
                try:
                    with st.spinner("Allocating the next letter number and saving to Supabase…"):
                        saved = allocate_and_save(client, [{**row, **common}], include_semester)[0]
                    st.session_state["last_single_record"] = saved
                    st.success(f"Saved as **{saved['letter_number']}**. You can print it now or find it later by registration ID.")
                except Exception as exc:
                    st.error(str(exc))
        last_single = st.session_state.get("last_single_record")
        if last_single:
            st.download_button(
                "Download / print latest letter (.docx)", build_docx([last_single]),
                file_name=f"{last_single['letter_number'].replace('/', '-')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                key=f"download_single_{last_single['id']}",
            )
            with st.expander("Latest saved record details"):
                st.json({k: last_single.get(k) for k in ("letter_number", "student_name", "registration_id", "gender", "program", "semester", "company_name", "city")})

    # ------------------------------ Bulk workbook ------------------------------
    with tabs[1]:
        st.markdown("### Generate a batch from an Excel roster")
        st.write("Upload `.xlsx` or `.csv`. Each student gets a separate letter number; all letters are saved in Supabase and combined into one downloadable Word document.")
        st.download_button("Download blank Excel template", template_excel(), "student_letter_template.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="template_download")
        bulk1, bulk2, bulk3 = st.columns(3)
        with bulk1:
            default_bulk_program = st.selectbox(
                "Default program for blank cells",
                [program for program in PROGRAMS if program != "Other / enter below"],
                index=list(PROGRAMS).index("Master of Science in Clinical Psychology"),
                format_func=lambda name: f"{PROGRAM_DEPARTMENTS[name]} · {name} ({PROGRAMS[name]})",
                key="bulk_program",
            )
        with bulk2:
            default_bulk_semester = st.selectbox("Default semester for blank cells", SEMESTERS, index=2, format_func=lambda n: f"{ordinal(n)} Semester", key="bulk_semester")
        with bulk3:
            bulk_default_mode = st.selectbox("Default recipient mode", ["General (any organization)", "Company / organization"], key="bulk_mode")
        bulk_date = st.date_input("Date for this batch", value=date.today(), format="DD/MM/YYYY", key="bulk_date")
        uploaded = st.file_uploader("Choose student roster", type=["xlsx", "csv"], key="bulk_upload")
        if uploaded:
            try:
                df, import_details = parse_uploaded_roster(uploaded.name, uploaded.getvalue())
                source_summary = ", ".join(
                    f"{detail['sheet']} (header row {detail['header_row']})" for detail in import_details
                )
                st.caption(
                    f"Detected {len(df)} student row(s) from {len(import_details)} roster sheet(s): {source_summary}. "
                    "Extra columns are ignored; blank program/semester cells use the selected defaults."
                )
                mapped_text = []
                for detail in import_details:
                    for field, headers in detail["mapped_headers"].items():
                        mapped_text.append(f"`{field}` ← {', '.join(headers)}")
                st.info("Detected columns: " + " · ".join(dict.fromkeys(mapped_text)))
                st.dataframe(df.head(15), use_container_width=True, hide_index=True)
                if not {"student_name", "registration_id"}.issubset(df.columns):
                    st.error("The spreadsheet needs a student name and registration ID. Download the blank template above or rename your columns to common headings such as `Name` and `Reg No`.")
                else:
                    if st.button("Review and generate all letters", type="primary", key="bulk_generate"):
                        default_program, default_code = normalize_program(default_bulk_program, default_bulk_program)
                        prepared: list[dict[str, Any]] = []
                        problems: list[str] = []
                        seen: set[tuple[str, str, str]] = set()
                        for idx, rowdata in df.iterrows():
                            excel_row = safe_text(rowdata.get("_source_row")) or str(idx + 2)
                            excel_sheet = safe_text(rowdata.get("_source_sheet"))
                            source_location = f"{excel_sheet}, row {excel_row}" if excel_sheet else f"row {excel_row}"
                            name = safe_text(rowdata.get("student_name"))
                            reg = safe_text(rowdata.get("registration_id"))
                            if not name or not reg:
                                problems.append(f"{source_location}: missing student name or registration ID; skipped.")
                                continue
                            key = (reg.casefold(), safe_text(rowdata.get("company_name")).casefold(), safe_text(rowdata.get("program")).casefold())
                            if key in seen:
                                problems.append(f"{source_location}: duplicate registration/company/program in this upload; skipped.")
                                continue
                            seen.add(key)
                            try:
                                program, code = normalize_program(rowdata.get("program"), default_program)
                                if program == "Other / enter below":
                                    problems.append(f"{source_location}: the 'Other' program option needs a full program name in the spreadsheet; skipped.")
                                    continue
                                sem = normalize_semester(rowdata.get("semester"), default_bulk_semester)
                                gender, _ = resolve_gender(name, rowdata.get("gender"))
                                mode_raw = safe_text(rowdata.get("recipient_mode")).lower()
                                company = safe_text(rowdata.get("company_name"))
                                city_value = safe_text(rowdata.get("city"))
                                if mode_raw in {"company", "organization", "company / organization", "specific"}:
                                    mode = "Company / organization"
                                elif mode_raw in {"general", "general (any organization)", "any"}:
                                    mode = "General (any organization)"
                                elif company:
                                    mode = "Company / organization"
                                else:
                                    mode = bulk_default_mode
                                if mode == "Company / organization" and not company:
                                    problems.append(f"{source_location}: company mode was selected but no company name was found; skipped.")
                                    continue
                                prepared.append(make_row(
                                    student_name=name, registration_id=reg, gender=gender,
                                    program=program, program_code=code, semester=sem,
                                    department_prefix=department_prefix, letter_date=bulk_date,
                                    recipient_mode=mode, company_name=company, city=city_value,
                                ))
                            except Exception as exc:
                                problems.append(f"{source_location}: {exc}; skipped.")
                        if problems:
                            with st.expander(f"Import notes ({len(problems)})", expanded=True):
                                for problem in problems: st.write(f"- {problem}")
                        if not prepared:
                            st.error("No valid student rows to generate.")
                        else:
                            st.info(f"Ready to generate **{len(prepared)}** letter(s). Rows with issues above will not be processed.")
                            try:
                                with st.spinner(f"Saving {len(prepared)} letters to Supabase…"):
                                    saved_batch = allocate_and_save(client, [{**row, **common} for row in prepared], include_semester)
                                st.session_state["last_bulk_records"] = saved_batch
                                st.success(f"Successfully saved {len(saved_batch)} letter(s). The combined document is ready below.")
                            except PartialSaveError as exc:
                                st.session_state["last_bulk_records"] = exc.saved_records
                                if exc.saved_records:
                                    st.error(
                                        f"Batch stopped after {len(exc.saved_records)} letter(s) were saved. "
                                        f"Those letters are available below and remain searchable for reprint. "
                                        f"Fix the failed row before uploading the remaining students. Details: {exc}"
                                    )
                                else:
                                    st.error(f"No letters in this batch were saved. Details: {exc}")
                            except Exception as exc:
                                st.error(f"Could not finish the batch. Check Records & export before retrying. Details: {exc}")
                    last_batch = st.session_state.get("last_bulk_records", [])
                    if last_batch:
                        st.dataframe(pd.DataFrame(last_batch)[[c for c in ["letter_number", "student_name", "registration_id", "gender", "program", "semester", "company_name"] if c in last_batch[0]]], use_container_width=True, hide_index=True)
                        combined_bytes = build_docx(last_batch)
                        st.download_button(
                            f"Download all {len(last_batch)} letters in one Word document",
                            combined_bytes, "air_university_internship_letters.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary", key=f"bulk_docx_{last_batch[0].get('id')}"
                        )
                        st.download_button(
                            "Download this batch record as Excel", excel_bytes(last_batch), "generated_letter_batch.xlsx",
                            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"bulk_xlsx_{last_batch[0].get('id')}"
                        )
            except Exception as exc:
                st.error(f"Could not read this file: {exc}")

    # ------------------------------ Search / reprint ------------------------------
    with tabs[2]:
        st.markdown("### Find a previous letter")
        st.caption("Search using the student's exact registration ID. Every matching saved letter can be downloaded again for printing.")
        with st.form("search_form"):
            search_reg = st.text_input("Registration ID", placeholder="Enter registration ID")
            search_button = st.form_submit_button("Search saved letters", type="primary")
        if search_button:
            if not search_reg.strip():
                st.warning("Enter a registration ID first.")
            else:
                try:
                    results = client.table("letters").select("*").eq("registration_id", search_reg.strip()).order("created_at", desc=True).execute().data or []
                    st.session_state["search_results"] = results
                    st.session_state["search_reg"] = search_reg.strip()
                except Exception as exc:
                    st.error(f"Search failed: {exc}")
        results = st.session_state.get("search_results", [])
        if results:
            st.success(f"Found {len(results)} saved letter(s) for registration ID {st.session_state.get('search_reg')}.")
            for record in results:
                with st.container(border=True):
                    left, right = st.columns([4, 1])
                    with left:
                        st.markdown(f"**{record.get('letter_number', '')}** · {record.get('student_name', '')}")
                        st.caption(f"{safe_text(record.get('program'))} · Semester {record.get('semester', '')} · {record.get('company_name') or 'General letter'} · {record.get('letter_date', '')}")
                    with right:
                        st.download_button(
                            "Download / print",
                            build_docx([record]),
                            f"{str(record.get('letter_number', 'letter')).replace('/', '-')}.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"reprint_{record.get('id')}",
                            use_container_width=True,
                        )
        elif search_button:
            st.info("No saved letters were found for that registration ID.")

    # ------------------------------ Records / export ------------------------------
    with tabs[3]:
        st.markdown("### Letter register")
        st.write("Download a copy of the saved record data for office filing. The database remains the primary record.")
        if st.button("Load saved records", key="load_records"):
            try:
                with st.spinner("Reading records from Supabase…"):
                    all_records = fetch_all_records(client)
                st.session_state["all_records"] = all_records
            except Exception as exc:
                st.error(f"Could not load records: {exc}")
        all_records = st.session_state.get("all_records", [])
        if all_records:
            records_df = pd.DataFrame(all_records)
            st.metric("Saved letters", len(records_df))
            query = st.text_input("Filter by name, registration ID, or letter number", key="record_filter")
            if query.strip():
                q = query.strip().lower()
                mask = records_df.astype(str).apply(lambda col: col.str.lower().str.contains(re.escape(q), na=False)).any(axis=1)
                view_df = records_df[mask]
            else:
                view_df = records_df
            visible_cols = [c for c in ["letter_number", "letter_date", "student_name", "registration_id", "gender", "program", "semester", "company_name", "city"] if c in view_df.columns]
            st.dataframe(view_df[visible_cols], use_container_width=True, hide_index=True)
            st.download_button(
                "Download all matching records as Excel",
                excel_bytes(view_df.to_dict(orient="records")),
                "air_university_letter_register.xlsx",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", key="records_export",
            )
        else:
            st.info("Click **Load saved records** to show and export the Supabase register.")

    st.divider()
    st.caption("Letter wording follows the supplied sample. Verify student details, pronouns, recipient, and letter number before printing or issuing an official letter.")


if __name__ == "__main__":
    app_main()

import streamlit as st
import os
import json
import uuid

from datetime import datetime
from zoneinfo import ZoneInfo
from threading import Lock

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

# =========================================================
# PAGE SETTINGS
# =========================================================
st.set_page_config(
    page_title="GenAI and E-Commerce Business Performance Survey",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# =========================================================
# AUTOSAVE AND RESPONSE STORAGE
# =========================================================

DRAFT_DIR = "survey_drafts"

# All final submitted responses will be stored here
RESPONSE_FILE = "survey_responses.xlsx"

os.makedirs(DRAFT_DIR, exist_ok=True)

# Prevent two respondents writing to the Excel file
# at exactly the same moment
excel_lock = Lock()


# All questionnaire fields that should be autosaved
AUTOSAVE_KEYS = [
    "consent",

    # Section 2
    "A1",
    "A2",
    "A3",
    "A3_other",
    "A3_final",
    "A4",
    "A4_other",
    "A4_final",
    "A5",
    "A6",
    "A7",
    "A8",
    "A9",
    "A10",
    "A10_other",
    "A10_final",
    "A12",

    # GenAI Tool Quality
    "GAITQ1",
    "GAITQ2",
    "GAITQ3",
    "GAITQ4",
    "GAITQ5",
    "GAITQ6",
    "GAITQ7",
    "GAITQ8",

    # GenAI Adoption
    "GAIA1",
    "GAIA2",
    "GAIA3",
    "GAIA4",
    "GAIA5",

    # Business Performance
    "ECBP1",
    "ECBP2",
    "ECBP3",
    "ECBP4",
    "ECBP5",
    "ECBP6",
    "ECBP7"
]

# =========================================================
# FINAL EXCEL RESPONSE COLUMNS
# =========================================================

RESPONSE_HEADERS = [
    "Response_No",
    "Timestamp",
    "Response_ID",
    "Consent",

    # Demographic Information
    "A1_Age_Group",
    "A2_Gender",
    "A3_Education",
    "A4_Job_Role",
    "A6_Work_Experience",
    "A5_Industry",
    "A7_Organisation_Size",
    "A8_GenAI_Experience",
    "A9_GenAI_Frequency",
    "A10_GenAI_Purpose",
    "A12_GenAI_Tools",

    # GenAI Tool Quality
    "GAITQ1",
    "GAITQ2",
    "GAITQ3",
    "GAITQ4",
    "GAITQ5",
    "GAITQ6",
    "GAITQ7",
    "GAITQ8",

    # GenAI Adoption
    "GAIA1",
    "GAIA2",
    "GAIA3",
    "GAIA4",
    "GAIA5",

    # Business Performance
    "ECBP1",
    "ECBP2",
    "ECBP3",
    "ECBP4",
    "ECBP5",
    "ECBP6",
    "ECBP7"
]

# =========================================================
# CREATE OR RECOVER DRAFT ID
# =========================================================

if "draft_id" not in st.session_state:

    existing_draft_id = st.query_params.get("draft", None)

    # Accept only a valid 32-character hexadecimal ID
    if (
        existing_draft_id
        and len(str(existing_draft_id)) == 32
        and all(
            character in "0123456789abcdefABCDEF"
            for character in str(existing_draft_id)
        )
    ):
        st.session_state["draft_id"] = str(existing_draft_id)

    else:
        st.session_state["draft_id"] = uuid.uuid4().hex
        st.query_params["draft"] = st.session_state["draft_id"]


def get_draft_path():
    return os.path.join(
        DRAFT_DIR,
        f"{st.session_state['draft_id']}.json"
    )


def get_response_path():
    return RESPONSE_FILE


def load_draft():

    draft_path = get_draft_path()

    if not os.path.exists(draft_path):
        return

    try:
        with open(draft_path, "r", encoding="utf-8") as file:
            saved_data = json.load(file)

        saved_answers = saved_data.get("answers", {})

        for key, value in saved_answers.items():

            if key not in st.session_state:
                st.session_state[key] = value

        if "page" in saved_data:
            st.session_state["page"] = saved_data["page"]

    except (json.JSONDecodeError, OSError):
        pass

def restore_saved_answers(keys):

    draft_path = get_draft_path()

    if not os.path.exists(draft_path):
        return

    try:
        with open(draft_path, "r", encoding="utf-8") as file:
            saved_data = json.load(file)

        saved_answers = saved_data.get("answers", {})

        for key in keys:

            # Restore only when Streamlit no longer has
            # the widget value in the current session
            if key in saved_answers and key not in st.session_state:
                st.session_state[key] = saved_answers[key]

    except (json.JSONDecodeError, OSError):
        pass

def save_draft():

    draft_path = get_draft_path()

    # Keep answers already saved from previous sections
    previous_data = {}

    if os.path.exists(draft_path):

        try:

            with open(draft_path, "r", encoding="utf-8") as file:
                previous_data = json.load(file)

        except (json.JSONDecodeError, OSError):
            previous_data = {}


    saved_answers = previous_data.get("answers", {})


    # Update values currently available in session state
    for key in AUTOSAVE_KEYS:

        if key in st.session_state:
            saved_answers[key] = st.session_state[key]


    draft_data = {
        "draft_id": st.session_state["draft_id"],
        "page": st.session_state.get("page", 1),
        "last_saved": datetime.now().isoformat(
            timespec="seconds"
        ),
        "answers": saved_answers
    }


    # Create a unique temporary file for every autosave
    temporary_path = (
        draft_path
        + "."
        + uuid.uuid4().hex
        + ".tmp"
    )


    try:

        with open(
            temporary_path,
            "w",
            encoding="utf-8"
        ) as file:

            json.dump(
                draft_data,
                file,
                ensure_ascii=False,
                indent=4
            )


        os.replace(
            temporary_path,
            draft_path
        )


    except PermissionError:

        # Windows may briefly lock a temporary file.
        # The questionnaire should continue running.
        pass


    finally:

        # Clean up temporary file if necessary
        if os.path.exists(temporary_path):

            try:
                os.remove(temporary_path)

            except PermissionError:
                pass


def delete_draft():

    draft_path = get_draft_path()

    if os.path.exists(draft_path):

        try:
            os.remove(draft_path)
        except OSError:
            pass


def save_final_response():

    # Make sure the latest answers are saved
    save_draft()

    draft_path = get_draft_path()

    saved_answers = {}

    if os.path.exists(draft_path):

        try:
            with open(draft_path, "r", encoding="utf-8") as file:
                draft_data = json.load(file)

            saved_answers = draft_data.get("answers", {})

        except (json.JSONDecodeError, OSError):
            saved_answers = {}


    # =====================================================
    # HELPER FOR LIKERT NUMBERS
    # =====================================================

    def numeric_answer(key):

        value = saved_answers.get(key)

        if value is None:
            return None

        try:
            return int(value)

        except (ValueError, TypeError):
            return value


    # =====================================================
    # SRI LANKA SUBMISSION TIME
    # =====================================================

    submitted_at = datetime.now(
        ZoneInfo("Asia/Colombo")
    ).replace(tzinfo=None)


    # =====================================================
    # WRITE TO ONE EXCEL WORKBOOK
    # =====================================================

    with excel_lock:

        # ---------------------------------------------
        # Open existing workbook
        # ---------------------------------------------
        if os.path.exists(RESPONSE_FILE):

            workbook = load_workbook(RESPONSE_FILE)

            if "Responses" in workbook.sheetnames:
                worksheet = workbook["Responses"]

            else:
                worksheet = workbook.active
                worksheet.title = "Responses"

        # ---------------------------------------------
        # Create workbook for first response
        # ---------------------------------------------
        else:

            workbook = Workbook()

            worksheet = workbook.active
            worksheet.title = "Responses"

            worksheet.append(RESPONSE_HEADERS)

            # Header formatting
            header_fill = PatternFill(
                fill_type="solid",
                fgColor="087F5B"
            )

            header_font = Font(
                color="FFFFFF",
                bold=True
            )

            for cell in worksheet[1]:

                cell.fill = header_fill
                cell.font = header_font

                cell.alignment = Alignment(
                    horizontal="center",
                    vertical="center"
                )

            worksheet.freeze_panes = "A2"


        # =================================================
        # PREVENT DUPLICATE RESPONSE
        # =================================================

        current_response_id = st.session_state["draft_id"]

        response_id_column = 3

        existing_ids = [
            str(row[0])
            for row in worksheet.iter_rows(
                min_row=2,
                min_col=response_id_column,
                max_col=response_id_column,
                values_only=True
            )
            if row[0] is not None
        ]

        if current_response_id in existing_ids:

            workbook.close()

            delete_draft()

            return


        # =================================================
        # RESPONSE NUMBER
        # =================================================

        response_number = worksheet.max_row


        # =================================================
        # BUILD ONE RESPONSE ROW
        # =================================================

        response_row = [
            response_number,
            submitted_at,
            current_response_id,

            saved_answers.get("consent"),

            # Demographic Information
            saved_answers.get("A1"),
            saved_answers.get("A2"),

            saved_answers.get(
                "A3_final",
                saved_answers.get("A3")
            ),

            saved_answers.get(
                "A4_final",
                saved_answers.get("A4")
            ),

            saved_answers.get("A6"),
            saved_answers.get("A5"),
            saved_answers.get("A7"),
            saved_answers.get("A8"),
            saved_answers.get("A9"),

            saved_answers.get(
                "A10_final",
                saved_answers.get("A10")
            ),

            saved_answers.get("A12"),

            # GenAI Tool Quality
            numeric_answer("GAITQ1"),
            numeric_answer("GAITQ2"),
            numeric_answer("GAITQ3"),
            numeric_answer("GAITQ4"),
            numeric_answer("GAITQ5"),
            numeric_answer("GAITQ6"),
            numeric_answer("GAITQ7"),
            numeric_answer("GAITQ8"),

            # GenAI Adoption
            numeric_answer("GAIA1"),
            numeric_answer("GAIA2"),
            numeric_answer("GAIA3"),
            numeric_answer("GAIA4"),
            numeric_answer("GAIA5"),

            # Business Performance
            numeric_answer("ECBP1"),
            numeric_answer("ECBP2"),
            numeric_answer("ECBP3"),
            numeric_answer("ECBP4"),
            numeric_answer("ECBP5"),
            numeric_answer("ECBP6"),
            numeric_answer("ECBP7")
        ]


        # =================================================
        # ADD RESPONSE TO NEXT EXCEL ROW
        # =================================================

        worksheet.append(response_row)


        # Timestamp formatting
        timestamp_cell = worksheet.cell(
            row=worksheet.max_row,
            column=2
        )

        timestamp_cell.number_format = "yyyy-mm-dd hh:mm:ss"


        # Filter across all responses
        worksheet.auto_filter.ref = worksheet.dimensions


        # =================================================
        # COLUMN WIDTHS
        # =================================================

        column_widths = {
            "A": 13,
            "B": 21,
            "C": 36,
            "D": 35,
            "E": 18,
            "F": 15,
            "G": 25,
            "H": 22,
            "I": 22,
            "J": 28,
            "K": 24,
            "L": 22,
            "M": 22,
            "N": 25,
            "O": 30
        }

        for column, width in column_widths.items():
            worksheet.column_dimensions[column].width = width


        # Research variable columns
        for column_number in range(16, len(RESPONSE_HEADERS) + 1):

            worksheet.column_dimensions[
                worksheet.cell(
                    row=1,
                    column=column_number
                ).column_letter
            ].width = 12


        # =================================================
        # SAVE EXCEL WORKBOOK
        # =================================================

        workbook.save(RESPONSE_FILE)

        workbook.close()


    # =====================================================
    # DELETE TEMPORARY AUTOSAVE DRAFT
    # =====================================================

    delete_draft()

def response_already_submitted():

    if not os.path.exists(RESPONSE_FILE):
        return False

    try:

        with excel_lock:

            workbook = load_workbook(
                RESPONSE_FILE,
                read_only=True,
                data_only=True
            )

            if "Responses" not in workbook.sheetnames:
                workbook.close()
                return False

            worksheet = workbook["Responses"]

            current_response_id = st.session_state["draft_id"]

            for row in worksheet.iter_rows(
                min_row=2,
                min_col=3,
                max_col=3,
                values_only=True
            ):

                if row[0] is not None and str(row[0]) == current_response_id:

                    workbook.close()

                    return True

            workbook.close()

    except (OSError, PermissionError):
        return False

    return False

# =========================================================
# RESTORE PREVIOUS DRAFT
# =========================================================

if "draft_loaded" not in st.session_state:

    load_draft()

    st.session_state["draft_loaded"] = True


# =========================================================
# SESSION STATE
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = 1

if "survey_ended" not in st.session_state:
    st.session_state.survey_ended = False

if "survey_completed" not in st.session_state:
    st.session_state.survey_completed = False

if "scroll_to_top" not in st.session_state:
    st.session_state.scroll_to_top = False

# =========================================================
# AUTO SCROLL HELPER FUNCTION
# =========================================================

def scroll_page_to_top():

    st.html(
        """
        <script>
        (function() {

            function goToTop() {

                const main = document.querySelector(
                    'section[data-testid="stMain"]'
                );

                if (main) {
                    main.scrollTop = 0;
                    main.scrollTo({
                        top: 0,
                        left: 0,
                        behavior: 'auto'
                    });
                }

                document.documentElement.scrollTop = 0;
                document.body.scrollTop = 0;

                window.scrollTo(0, 0);
            }

            requestAnimationFrame(function() {

                goToTop();

                setTimeout(goToTop, 100);
                setTimeout(goToTop, 300);

            });

        })();
        </script>
        """,
        unsafe_allow_javascript=True
    )

# =========================================================
# ADMIN RESPONSE DOWNLOAD
# =========================================================

if st.query_params.get("admin") == "responses":

    st.title("Survey Response Administration")

    admin_password = st.text_input(
        "Admin Password",
        type="password"
    )

    if admin_password == st.secrets["admin_password"]:

        if os.path.exists(RESPONSE_FILE):

            with open(RESPONSE_FILE, "rb") as excel_file:

                st.download_button(
                    label="Download Survey Responses",
                    data=excel_file,
                    file_name="survey_responses.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )

        else:
            st.warning("No response file currently exists on the cloud server.")

    elif admin_password:
        st.error("Incorrect password.")

    st.stop()

# =========================================================
# CSS
# =========================================================
st.markdown("""
<style>

/* Outside background */
.stApp {
    background-color: #E9F7F0;
    font-family: Arial, sans-serif;
    color: #202124;
}

/* Main white questionnaire */
.block-container {
    max-width: 850px;
    background-color: #FFFFFF;
    padding: 35px 42px 50px 42px;
    margin-top: 25px;
    margin-bottom: 30px;
    border-radius: 8px;
    border: 1px solid #DDE7E1;
    box-shadow: 0 3px 12px rgba(0,0,0,0.10);
}

/* Main title */
.main-title {
    background: linear-gradient(
        135deg,
        #056B4F 0%,
        #07875F 50%,
        #045A42 100%
    );
    color: white;
    font-family: Arial, sans-serif;
    font-size: 34pt;
    font-weight: 700;
    line-height: 1.15;
    text-align: center;
    text-transform: uppercase;
    padding: 28px 24px;
    border-radius: 7px;
    margin-bottom: 32px;
    box-shadow: 0 6px 14px rgba(4, 90, 66, 0.30);
}

/* =========================================================
   SECTION HEADING - CLEAN BOX STYLE
========================================================= */

.section-heading {
    font-family: Arial, sans-serif;

    font-size: 16pt;

    font-weight: 840;
    
    border: 2px solid #8FD8BE;

    text-transform: uppercase;

    text-align: center;

    color: #087F5B;

    background-color: #EAF6EF;

    border: 1px solid #C9CED1;

    border-radius: 6px;

    padding: 12px 18px;

    margin-top: 28px;

    margin-bottom: 20px;

    box-shadow:
        0 4px 7px rgba(120, 120, 120, 0.18);
}

/* Sub headings */
.sub-heading {
    font-family: Arial, sans-serif;
    font-size: 14pt;
    font-weight: 700;
    color: #202124;
    margin-top: 20px;
    margin-bottom: 12px;
}

/* Dear Sir/Madam */
.dear-heading {
    font-family: Arial, sans-serif;
    font-size: 12pt;
    font-weight: 400;
    color: #202124;
    margin-bottom: 15px;
}

/* Paragraphs */
.content-text {
    font-family: Arial, sans-serif;
    font-size: 12pt;
    line-height: 1.65;
    color: #202124;
    text-align: justify;
    margin-bottom: 15px;
}

/* Researcher */
.researcher {
    font-family: Arial, sans-serif;
    font-size: 12pt;
    font-weight: 600;
    color: #202124;
    margin-top: 18px;
    margin-bottom: 25px;
}

/* Instruction box */
.instruction-box {
    background-color: #F5FAF7;
    border-left: 4px solid #087F5B;
    border-radius: 5px;
    padding: 16px 18px;
    font-family: Arial, sans-serif;
    font-size: 13pt;
    line-height: 1.8;
    color: #202124;
    margin-bottom: 20px;
}

/* Consent box */
.consent-box {
    background-color: #F5FAF7;
    border: 1px solid #D4E5DB;
    border-radius: 6px;
    padding: 20px;
    font-family: Arial, sans-serif;
    font-size: 12pt;
    line-height: 1.65;
    color: #202124;
    text-align: justify;
    margin-bottom: 18px;
}

/* Radio text */
div[data-testid="stRadio"] label p {
    font-family: Arial, sans-serif !important;
    font-size: 12.5pt !important;
    color: #202124 !important;
    opacity: 1 !important;
}

/* General button */
.stButton button {
    font-family: Arial, sans-serif !important;
    font-size: 16pt !important;
    font-weight: 700 !important;
    min-height: 42px;
    padding: 8px 28px;
    border-radius: 5px;
}

/* Hide Streamlit items */
#MainMenu {
    visibility: hidden;
}

footer {
    visibility: hidden;
}

header {
    visibility: hidden;
}

/* Mobile */
@media (max-width: 700px) {

    .block-container {
        width: 96%;
        padding: 22px 20px 35px 20px;
        margin-top: 10px;
    }

    .main-title {
        font-size: 23pt;
        padding: 22px 15px;
    }

    .section-heading {
        font-size: 15pt;
    }

    .content-text {
        font-size: 11.5pt;
        text-align: left;
    }
}

/* =========================================================
   DEMOGRAPHIC SUBSECTION HEADING
========================================================= */
.demo-subheading {
    font-family: Arial, sans-serif;
    font-size: 15pt;
    font-weight: 800;
    color: #087F5B;
    margin-top: 10px;
    margin-bottom: 8px;
}


/* =========================================================
   DEMOGRAPHIC DESCRIPTION
========================================================= */
.demo-description {
    font-family: Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.5;
    color: #202124;
    margin-bottom: 16px;
}


/* =========================================================
   DEMOGRAPHIC QUESTION TITLE
========================================================= */
.demo-question {
    font-family: Arial, sans-serif;
    font-size: 14pt;
    font-weight: 750;
    color: #202124;
    margin-top: 14px;
    margin-bottom: 5px;
}

/* =========================================================
   RESEARCH QUESTION SEGMENTS
========================================================= */

.research-segment {
    font-family: Arial, sans-serif;
    font-size: 14pt;
    font-weight: 750;
    color: #087F5B;

    margin-top: 24px;
    margin-bottom: 7px;
}


/* Segment instructions */
.research-description {
    font-family: Arial, sans-serif;
    font-size: 11pt;
    line-height: 1.55;
    color: #404040;

    margin-top: 2px;
    margin-bottom: 8px;
}


/* Individual research statement */
.research-question {
    font-family: Arial, sans-serif;
    font-size: 12.5pt;
    font-weight: 600;
    line-height: 1.5;
    color: #202124;
    text-align: justify;

    margin-top: 15px;
    margin-bottom: 4px;
}


/* Likert guide above each segment */
.scale-guide {
    background-color: #F2FAF6;

    border: 1px solid #A9DFC9;
    border-radius: 5px;

    font-family: Arial, sans-serif;
    font-size: 12pt;
    font-weight: 650;

    color: #202124;
    text-align: center;

    width: 100%;
    box-sizing: border-box;

    padding: 10px 8px;

    margin-top: 8px;
    margin-bottom: 12px;
}

/* =========================================================
   RESEARCH CONSTRUCT SUBHEADING
========================================================= */
.construct-heading {
    font-family: Arial, sans-serif;

    font-size: 16pt;

    font-weight: 840;

    color: #087F5B;

    margin-top: 24px;

    margin-bottom: 8px;
}


/* Divider between research constructs */
.research-divider {
    width: 100%;
    height: 2px;

    background-color: #74BFA5;

    margin-top: 18px;
    margin-bottom: 10px;
}

/* Likert buttons - unselected */
.st-key-research_likert button {
    background-color: #EAF4E3 !important;
    color: #000000 !important;
    border: 2px solid #228B22 !important;
    font-weight: 700 !important;
}

/* Keep unselected number black */
.st-key-research_likert button p {
    color: #000000 !important;
    font-weight: 700 !important;
}

/* Likert button - selected */
.st-key-research_likert button[kind="primary"],
.st-key-research_likert button[data-testid="stBaseButton-primary"] {
    background-color: #228B22 !important;
    color: #FFFFFF !important;
    border: 2px solid #228B22 !important;
}

/* Selected tick and number white */
.st-key-research_likert button[kind="primary"] p,
.st-key-research_likert button[data-testid="stBaseButton-primary"] p {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}


/* Selected hover */
.st-key-research_likert div[data-testid="stButton"] button[kind="primary"]:hover {

    background-color: #1B6F1B !important;

    border-color: #1B6F1B !important;

    color: #FFFFFF !important;
}


/* SECTION NAVIGATION BUTTONS */

div[data-testid="stButton"] button {
    font-family: Arial, sans-serif !important;
    font-size: 14pt !important;
    font-weight: 800 !important;
}


/* Previous / secondary button */
div[data-testid="stButton"] button[kind="secondary"] {
    background-color: #19437D !important;
    color: #FFFFFF !important;
    border: 3px solid #19437D !important;
}

div[data-testid="stButton"] button[kind="secondary"] p {
    color: #FFFFFF !important;
    font-weight: 700 !important;
}


/* =========================================================
   SECTION 3 LIKERT BUTTONS
========================================================= */

/* Unselected Likert option */
.st-key-research_likert div[data-testid="stButton"] button[kind="secondary"] {
    background-color: #EAF4E3 !important;
    color: #000000 !important;
    border: 2px solid #228B22 !important;
    
    min-height: 32px !important;
    height: 32px !important;
    padding: 2px 8px !important;
}

/* Larger unselected numbers */
.st-key-research_likert div[data-testid="stButton"] button[kind="secondary"] p {
    color: #000000 !important;
    font-size: 13pt !important;
    font-weight: 600 !important;
}

/* Selected Likert option */
.st-key-research_likert div[data-testid="stButton"] button[kind="primary"] {
    background-color: #228B22 !important;
    color: #FFFFFF !important;
    border: 2px solid #228B22 !important;
    min-height: 32px !important;
    height: 32px !important;
    padding: 2px 8px !important;
}

/* Larger selected number */
.st-key-research_likert div[data-testid="stButton"] button[kind="primary"] p {
    color: #FFFFFF !important;
    font-size: 15pt !important;
    font-weight: 900 !important;
}

/* Selected option hover */
.st-key-research_likert div[data-testid="stButton"] button[kind="primary"]:hover {
    background-color: #1B6F1B !important;
    border-color: #EAF4E3 !important;
}

/* =========================================================
   NEXT AND SUBMIT BUTTON TEXT - EXTRA BOLD
========================================================= */

.st-key-section2_next button,
.st-key-section3_submit button {
    font-weight: 900 !important;
}

.st-key-section2_next button p,
.st-key-section3_submit button p {
    font-weight: 900 !important;
    color: #FFFFFF !important;
}

/* Reduce space before text boxes */
div[data-testid="stTextInput"] {
    margin-top: 6px !important;
    margin-bottom: 5px !important;
}

/* Mobile */
@media (max-width: 700px) {

    .block-container {
        width: 96%;
        padding: 22px 20px 35px 20px;
        margin-top: 10px;
    }

    .main-title {
        font-size: 23pt;
        padding: 22px 15px;
    }

    .section-heading {
        font-size: 15pt;
    }

    .content-text {
        font-size: 11.5pt;
        text-align: left;
    }
}

</style>
""", unsafe_allow_html=True)

# =========================================================
# QUESTIONNAIRE PROGRESS BAR
# =========================================================

if (
    not st.session_state.get("survey_completed", False)
    and not st.session_state.get("survey_ended", False)
):

    current_page = st.session_state.get("page", 1)

    if current_page == 1:
        progress_value = 15
        progress_text = "Questionnaire Progress — Section 1 of 3"

    elif current_page == 2:
        progress_value = 45
        progress_text = "Questionnaire Progress — Section 2 of 3"

    else:

        research_keys = [
            "GAITQ1", "GAITQ2", "GAITQ3", "GAITQ4",
            "GAITQ5", "GAITQ6", "GAITQ7", "GAITQ8",
            "GAIA1", "GAIA2", "GAIA3", "GAIA4", "GAIA5",
            "ECBP1", "ECBP2", "ECBP3", "ECBP4",
            "ECBP5", "ECBP6", "ECBP7"
        ]

        completed_answers = sum(
            1
            for key in research_keys
            if st.session_state.get(key) is not None
        )

        # Section 3 progresses from 65% to 95%
        progress_value = 65 + int(
            (completed_answers / len(research_keys)) * 30
        )

        progress_text = (
            f"Questionnaire Progress — "
            f"{completed_answers} of {len(research_keys)} "
            f"research statements completed"
        )

    st.progress(
        progress_value,
        text=progress_text
    )

# =========================================================
# COMPLETED SURVEY PAGE
# =========================================================

if st.session_state.get("survey_completed", False):

    st.balloons()

    completion_html = (
        '<div style="max-width:700px;margin:50px auto 30px auto;'
        'background:#FFFFFF;border:1px solid #D7E8DF;'
        'border-top:7px solid #087F5B;border-radius:12px;'
        'padding:50px 40px;text-align:center;'
        'box-shadow:0 6px 20px rgba(0,0,0,0.10);">'

        '<div style="width:75px;height:75px;background:#E8F6EF;'
        'border-radius:50%;margin:0 auto 22px auto;'
        'display:flex;align-items:center;justify-content:center;'
        'font-family:Arial,sans-serif;font-size:36px;'
        'font-weight:900;color:#087F5B;">'
        '✓'
        '</div>'

        '<div style="font-family:Arial,sans-serif;'
        'font-size:23pt;font-weight:800;color:#087F5B;'
        'margin-bottom:14px;">'
        'Response Recorded'
        '</div>'

        '<div style="font-family:Arial,sans-serif;'
        'font-size:16pt;font-weight:700;color:#555555;'
        'margin-bottom:26px;">'
        'Your questionnaire has been successfully submitted.'
        '</div>'

        '<div style="font-family:Arial,sans-serif;'
        'font-size:12.5pt;line-height:1.8;color:#202124;'
        'max-width:590px;margin:0 auto;">'
        'Thank you very much for participating in this research. '
        'Your valuable time, experiences, and information are sincerely appreciated. '
        'Your contribution will provide important insights and support the successful '
        'completion of this study.'
        '</div>'

        '<div style="width:90px;height:2px;background:#74BFA5;'
        'margin:32px auto 25px auto;"></div>'

        '<div style="font-family:Arial,sans-serif;'
        'font-size:13pt;font-weight:800;color:#202124;'
        'margin-bottom:5px;">'
        'Imali Abeysinghe'
        '</div>'

        '<div style="font-family:Arial,sans-serif;'
        'font-size:12pt;line-height:1.7;color:#555555;">'
        'MBA in Business Analytics<br>'
        'Queen Margaret University, UK'
        '</div>'

        '</div>'
    )

    st.markdown(
        completion_html,
        unsafe_allow_html=True
    )

    st.stop()

# =========================================================
# CUSTOM LIKERT SCALE
# =========================================================
def likert_scale(key):

    if key not in st.session_state:
        st.session_state[key] = None

    cols = st.columns(5)

    for i, col in enumerate(cols, start=1):

        with col:

            selected = st.session_state[key] == str(i)

            button_text = f"✓  {i}" if selected else str(i)

            if st.button(
                button_text,
                key=f"{key}_option_{i}",
                type="primary" if selected else "secondary",
                use_container_width=True
            ):
                st.session_state[key] = str(i)
                save_draft()
                st.rerun()

    return st.session_state[key]

# =========================================================
# SECTION 1
# =========================================================
if st.session_state.page == 1:

    # Restore previously saved Section 1 response
    restore_saved_answers([
        "consent"
    ])


    # Main research title
    st.markdown(
        '<div class="main-title">THE IMPACT OF GENERATIVE AI TOOLS ON THE BUSINESS PERFORMANCE OF E-COMMERCE BUSINESSES IN SRI LANKA</div>',
        unsafe_allow_html=True
    )

    # Introduction
    st.markdown(
        '<div class="section-heading">Section A - Introduction</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="sub-heading">Overview to the Survey</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="dear-heading">Dear Sir / Madam,</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">I am currently pursuing a Master of Business Administration (MBA) in Business Analytics at Queen Margaret University, UK. As part of my degree programme, I am conducting a research study to examine the impact of Generative AI (GenAI) tools on the business performance of e-commerce firms in Sri Lanka, which are rarely studied in Sri Lanka.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">The research is intended to examine the impact of GenAI tool quality on business performance, GenAI tool quality on GenAI adoption, GenAI adoption on business performance and finally the extent to which GenAI adoption mediates the relationship between GenAI tool quality and business performance of e-commerce firms in Sri Lanka. Therefore, I invite all founders, owners, management and employees of e-commerce organisations in Sri Lanka to participate in this survey.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">This survey involves completing an online questionnaire, which will take only a few minutes of your time. You are kindly requested to complete this questionnaire based on your actual experience with Generative AI at work. Your responses will provide valuable insights for this research.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">It is completely voluntary to participate in this survey; you can withdraw at any time from the survey before submitting the questionnaire. No questions that can uniquely identify you or your organisation are included in this survey, and all the data gathered will be used for academic purposes and will not be shared with any e-commerce organisation, regulatory body or market research organisation for commercial purposes.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">Findings obtained by analysing these survey data will be reported in aggregate form. Meanwhile, all the gathered data will be stored securely in a password-protected online drive and will be disposed of five years after the completion of the research.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="content-text">Your participation is highly appreciated and will significantly contribute to the success of this research. Thank you for your time and support.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="researcher">Imali Abeysinghe</div>',
        unsafe_allow_html=True
    )


    # =====================================================
    # INSTRUCTIONS
    # =====================================================
    st.markdown(
        '<div class="sub-heading">Instructions</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="instruction-box">📌 Please provide one response for each survey item.<br><br>📌 Where response options are provided, select the most appropriate answer.<br><br>📌 There are no right or wrong answers; respond according to your own experience.<br><br>📌 From Section 3 onwards, use the 5-point scale shown below to provide responses.</div>',
        unsafe_allow_html=True
    )

    # =====================================================
    # LIKERT GUIDE
    # =====================================================
    cols = st.columns(5)

    labels = [
        ("1", "Strongly Disagree"),
        ("2", "Disagree"),
        ("3", "Neutral"),
        ("4", "Agree"),
        ("5", "Strongly Agree")
    ]

    for col, (number, label) in zip(cols, labels):
        with col:
            st.markdown(
                f'<div style="background:#F5FAF7;border:1px solid #D4E5DB;border-radius:5px;text-align:center;padding:12px 4px;min-height:90px;"><div style="font-size:13pt;font-weight:700;color:#087F5B;">{number}</div><div style="font-size:12pt;font-weight:700;color:#202124;margin-top:8px;line-height:1.2;">{label}</div></div>',
                unsafe_allow_html=True
            )


    # =====================================================
    # CONSENT
    # =====================================================
    st.markdown(
        '<div class="sub-heading">Consent Statement</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="consent-box">I confirm that I have read and understood the information provided about this research study. I understand that my participation is voluntary and that I may withdraw before submitting the questionnaire. I understand that my responses will be kept confidential, used only for academic research purposes, and reported without identifying me personally.<br><br>By selecting <b>“I agree”</b> below, I voluntarily consent to participate in this study.</div>',
        unsafe_allow_html=True
    )


    consent = st.radio(
        "Consent",
        [
            "I agree to participate in this study",
            "I do not agree to participate in this study"
        ],
        index=None,
        label_visibility="collapsed",
        key="consent"
    )

    # =====================================================
    # AUTOSAVE CONSENT RESPONSE
    # =====================================================

    save_draft()

    # =====================================================
    # AGREE
    # =====================================================
    if consent == "I agree to participate in this study":

        st.markdown("""
        <style>
        div[data-testid="stButton"] button {
            background-color: #087F5B !important;
            border-color: #087F5B !important;
            color: white !important;
        }

        div[data-testid="stButton"] button p {
            color: white !important;
        }

        div[data-testid="stButton"] button:hover {
            background-color: #056B4F !important;
            border-color: #056B4F !important;
        }
        </style>
        """, unsafe_allow_html=True)

        if st.button("NEXT", key="next_button"):
            st.session_state.page = 2
            save_draft()
            st.session_state.scroll_to_top = True
            st.rerun()


    # =====================================================
    # DISAGREE
    # =====================================================
    elif consent == "I do not agree to participate in this study":

        st.markdown("""
        <style>
        div[data-testid="stButton"] button {
            background-color: #C62828 !important;
            border-color: #C62828 !important;
            color: white !important;
        }

        div[data-testid="stButton"] button p {
            color: white !important;
        }

        div[data-testid="stButton"] button:hover {
            background-color: #A51D1D !important;
            border-color: #A51D1D !important;
        }
        </style>
        """, unsafe_allow_html=True)

        if st.button("END SURVEY", key="end_button"):
            delete_draft()
            st.session_state.survey_ended = True
            st.rerun()


# =========================================================
# END SURVEY
# =========================================================
if st.session_state.survey_ended:

    st.markdown(
        '<div style="background:#FFFFFF;border:1px solid #D6E1D8;border-radius:8px;padding:35px;text-align:center;margin-top:40px;"><div style="font-family:Arial;font-size:16pt;font-weight:700;color:#087F5B;margin-bottom:15px;">THANK YOU</div><div style="font-family:Arial;font-size:12pt;line-height:1.7;color:#202124;">Thank you for your time. As you have chosen not to consent to participate in this study, the questionnaire has now ended. No survey responses have been collected.</div></div>',
        unsafe_allow_html=True
    )

    st.stop()

# =========================================================
# SECTION 2 - DEMOGRAPHIC INFORMATION
# =========================================================
if st.session_state.page == 2:

    # =====================================================
    # RESTORE AUTOSAVED SECTION 2 RESPONSES
    # =====================================================

    restore_saved_answers([
        "A1",
        "A2",
        "A3",
        "A3_other",
        "A3_final",
        "A4",
        "A4_other",
        "A4_final",
        "A5",
        "A6",
        "A7",
        "A8",
        "A9",
        "A10",
        "A10_other",
        "A10_final",
        "A12"
    ])


    # -----------------------------------------------------
    # RESEARCH TOPIC - SAME AS SECTION 1
    # -----------------------------------------------------
    st.markdown(
        '<div class="main-title">THE IMPACT OF GENERATIVE AI TOOLS ON THE BUSINESS PERFORMANCE OF E-COMMERCE BUSINESSES IN SRI LANKA</div>',
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # MAIN SECTION HEADING
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-heading">Section B - Demographic Information</div>',
        unsafe_allow_html=True
    )


    # =====================================================
    # SUBSECTION 1 - PERSONAL CHARACTERISTICS
    # =====================================================
    st.markdown(
        '<div class="demo-subheading">Personal Characteristics</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="demo-description">Please provide the most appropriate response that describes your personal and professional background.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div style="font-family:Arial;font-size:10.5pt;color:#666;margin-top:-6px;margin-bottom:-5strpx;text-align:right;"><span style="color:#C62828;font-weight:700;">*</span> Required question</div>',
        unsafe_allow_html=True
    )

    # -----------------------------------------------------
    # AGE GROUP
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Age Group <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    age = st.radio(
        "Age Group",
        [
            "Under 25",
            "25–34",
            "35–44",
            "45–54",
            "55 or above"
        ],
        index=None,
        key="A1",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # GENDER
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Gender <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    gender = st.radio(
        "Gender",
        [
            "Male",
            "Female"
        ],
        index=None,
        key="A2",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # EDUCATIONAL QUALIFICATION
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Highest Educational Qualification <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    education = st.radio(
        "Highest Educational Qualification",
        [
            "G.C.E. (O/L) or below",
            "G.C.E. A/L",
            "Diploma",
            "Bachelor's",
            "Master's",
            "Doctorate",
            "Other"
        ],
        index=None,
        key="A3",
        label_visibility="collapsed"
    )

    education_other = ""

    if education == "Other":
        st.markdown(
            '<div style="font-family:Arial;font-size:10.5pt;color:#555;margin-top:-8px;margin-bottom:-12px;">Please specify your educational qualification:</div>',
            unsafe_allow_html=True
        )

        education_other = st.text_input(
            "Educational qualification",
            key="A3_other",
            label_visibility="collapsed"
        )


    # -----------------------------------------------------
    # CURRENT JOB ROLE
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Current Job Role <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    job_role = st.radio(
        "Current Job Role",
        [
            "Owner/Founder",
            "Senior Management",
            "Middle Management",
            "Executive / Officer",
            "Support Staff",
            "Other"
        ],
        index=None,
        key="A4",
        label_visibility="collapsed"
    )

    job_role_other = ""

    if job_role == "Other":
        st.markdown(
            '<div style="font-family:Arial;font-size:10.5pt;color:#555;margin-top:-8px;margin-bottom:-12px;">Please specify your current job role:</div>',
            unsafe_allow_html=True
        )

        job_role_other = st.text_input(
            "Current job role",
            key="A4_other",
            label_visibility="collapsed"
        )


    # -----------------------------------------------------
    # YEARS OF WORK EXPERIENCE
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Years of Work Experience <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    work_experience = st.radio(
        "Years of Work Experience",
        [
            "Below 5 years",
            "5–10 years",
            "More than 10 years"
        ],
        index=None,
        key="A6",
        label_visibility="collapsed"
    )


    # VISIBLE SEPARATOR LINE
    st.markdown(
    "<hr style='border:none; border-top:5px solid #74BFA5; margin:18px 0 8px 0;'>",
    unsafe_allow_html = True
)


    # =====================================================
    # SUBSECTION 2 - ORGANISATIONAL CHARACTERISTICS
    # =====================================================
    st.markdown(
        '<div class="demo-subheading">Organisational Characteristics</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="demo-description">Please provide the most appropriate response based on the e-commerce organisation in which you currently work or operate.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div style="font-family:Arial;font-size:10.5pt;color:#666;margin-top:-16px;margin-bottom:10px;text-align:right;"><span style="color:#C62828;font-weight:700;">*</span> Required question</div>',
        unsafe_allow_html=True
    )

    # -----------------------------------------------------
    # INDUSTRY / SECTOR
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Industry / Sector <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    industry = st.text_input(
        "Industry / Sector",
        placeholder="For Example: Fashion & Apparel, Jewellery & Accessories, Food & Grocery, Electronics, Educational Services",
        key="A5",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # ORGANISATION SIZE
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Organisation Size (Number of Employees) <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    organisation_size = st.radio(
        "Organisation Size",
        [
            "Micro (1–9)",
            "Small (10–49)",
            "Medium (50–249)",
            "Large (250+)"
        ],
        index=None,
        key="A7",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # EXPERIENCE USING GENAI
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Experience Using GenAI <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    genai_experience = st.radio(
        "Experience Using GenAI",
        [
            "Below 6 months",
            "6–12 months",
            "1–2 years",
            "More than 2 years"
        ],
        index=None,
        key="A8",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # FREQUENCY OF GENAI USE
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Frequency of GenAI Use <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    genai_frequency = st.radio(
        "Frequency of GenAI Use",
        [
            "Rarely",
            "Monthly",
            "Weekly",
            "Several times a week",
            "Daily"
        ],
        index=None,
        key="A9",
        label_visibility="collapsed"
    )


    # -----------------------------------------------------
    # MAIN PURPOSE OF USING GENAI
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Main Purpose of Using GenAI <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    genai_purpose = st.radio(
        "Main Purpose of Using GenAI",
        [
            "Content generation",
            "Information search",
            "Decision support",
            "Customer service",
            "Data analysis",
            "Other"
        ],
        index=None,
        key="A10",
        label_visibility="collapsed"
    )

    purpose_other = ""

    if genai_purpose == "Other":
        st.markdown(
            '<div style="font-family:Arial;font-size:10.5pt;color:#555;margin-top:-8px;margin-bottom:-12px;">Please specify the main purpose:</div>',
            unsafe_allow_html=True
        )

        purpose_other = st.text_input(
            "Main purpose",
            key="A10_other",
            label_visibility="collapsed"
        )


    # -----------------------------------------------------
    # MAIN GENAI TOOLS USED
    # -----------------------------------------------------
    st.markdown(
        '<div class="demo-question">Main GenAI Tool(s) Used <span style="color:#C62828;">*</span></div>',
        unsafe_allow_html=True
    )

    genai_tools = st.text_input(
        "Main GenAI Tool(s) Used",
        placeholder="For example: ChatGPT, Gemini, Microsoft Copilot",
        key="A12",
        label_visibility="collapsed"
    )

    # =====================================================
    # AUTOSAVE SECTION 2 RESPONSES
    # =====================================================

    save_draft()

    # =====================================================
    # NAVIGATION
    # =====================================================
    st.markdown("<br>", unsafe_allow_html=True)

    previous_col, spacer_col, next_col = st.columns([1.4, 4, 1.4])


    # PREVIOUS BUTTON
    with previous_col:

        if st.button(
            "PREVIOUS",
            key="section2_previous",
            use_container_width=True
        ):
            st.session_state.page = 1
            save_draft()
            st.session_state.scroll_to_top = True
            st.rerun()


    # NEXT BUTTON
    with next_col:

        if st.button(
            "NEXT",
            key="section2_next",
            type="primary",
            use_container_width=True
        ):

            missing_fields = []

            if age is None:
                missing_fields.append("Age Group")

            if gender is None:
                missing_fields.append("Gender")

            if education is None:
                missing_fields.append("Highest Educational Qualification")

            if education == "Other" and not education_other.strip():
                missing_fields.append("Educational Qualification - Other")

            if job_role is None:
                missing_fields.append("Current Job Role")

            if job_role == "Other" and not job_role_other.strip():
                missing_fields.append("Current Job Role - Other")

            if work_experience is None:
                missing_fields.append("Years of Work Experience")

            if not industry.strip():
                missing_fields.append("Industry / Sector")

            if organisation_size is None:
                missing_fields.append("Organisation Size")

            if genai_experience is None:
                missing_fields.append("Experience Using GenAI")

            if genai_frequency is None:
                missing_fields.append("Frequency of GenAI Use")

            if genai_purpose is None:
                missing_fields.append("Main Purpose of Using GenAI")

            if genai_purpose == "Other" and not purpose_other.strip():
                missing_fields.append("Main Purpose - Other")

            if not genai_tools.strip():
                missing_fields.append("Main GenAI Tool(s) Used")


            if missing_fields:

                st.error(
                    "Please complete all required questions before proceeding."
                )

            else:

                # Store cleaned values for later saving
                st.session_state["A3_final"] = (
                    education_other.strip()
                    if education == "Other"
                    else education
                )

                st.session_state["A4_final"] = (
                    job_role_other.strip()
                    if job_role == "Other"
                    else job_role
                )

                st.session_state["A10_final"] = (
                    purpose_other.strip()
                    if genai_purpose == "Other"
                    else genai_purpose
                )

                st.session_state.page = 3

                save_draft()

                st.session_state.scroll_to_top = True

                st.rerun()


# =========================================================
# SECTION 3 - RESEARCH INFORMATION
# =========================================================
if st.session_state.page == 3:

    # -----------------------------------------------------
    # RESEARCH TOPIC
    # -----------------------------------------------------
    st.markdown(
        '<div class="main-title">THE IMPACT OF GENERATIVE AI TOOLS ON THE BUSINESS PERFORMANCE OF E-COMMERCE BUSINESSES IN SRI LANKA</div>',
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # MAIN SECTION HEADING
    # -----------------------------------------------------
    st.markdown(
        '<div class="section-heading">Section C - Research Information</div>',
        unsafe_allow_html=True
    )


    # -----------------------------------------------------
    # GENERAL INSTRUCTION
    # -----------------------------------------------------
    st.markdown(
        '<div class="research-description">Please select one response for each statement using the 5-point agreement scale.</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div style="font-family:Arial;font-size:10.5pt;color:#666;text-align:right;margin-top:-5px;margin-bottom:10px;"><span style="color:#C62828;font-weight:700;">*</span> Required question</div>',
        unsafe_allow_html=True
    )


    # =====================================================
    # LIKERT QUESTION CONTAINER
    # =====================================================
    with st.container(key="research_likert"):

        # =================================================
        # GENAI TOOL QUALITY
        # =================================================
        st.markdown(
            '<div class="construct-heading">GenAI Tool Quality</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="research-description">Please indicate the extent to which you agree with the following statements about the quality of GenAI tools used in your organisation.</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="scale-guide">1 = Strongly Disagree &nbsp;&nbsp;&nbsp;&nbsp; 2 = Disagree &nbsp;&nbsp;&nbsp;&nbsp; 3 = Neutral &nbsp;&nbsp;&nbsp;&nbsp; 4 = Agree &nbsp;&nbsp;&nbsp;&nbsp; 5 = Strongly Agree</div>',
            unsafe_allow_html=True
        )


        # GAITQ1
        st.markdown(
            '<div class="research-question">The GenAI tools used in our organisation operate reliably and consistently. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq1 = likert_scale("GAITQ1")


        # GAITQ2
        st.markdown(
            '<div class="research-question">The GenAI tools are easy to access and provide responses within an appropriate time. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq2 = likert_scale("GAITQ2")


        # GAITQ3
        st.markdown(
            '<div class="research-question">Information produced by GenAI tools is accurate, relevant, and useful for our business activities. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq3 = likert_scale("GAITQ3")


        # GAITQ4
        st.markdown(
            '<div class="research-question">GenAI tools respond effectively to the needs and requests of users in our organisation. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq4 = likert_scale("GAITQ4")


        # GAITQ5
        st.markdown(
            '<div class="research-question">GenAI tools provide appropriate and personalised assistance for completing business-related tasks. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq5 = likert_scale("GAITQ5")


        # GAITQ6
        st.markdown(
            '<div class="research-question">Explanations produced by GenAI tools are clear, logical, and well structured. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq6 = likert_scale("GAITQ6")


        # GAITQ7
        st.markdown(
            '<div class="research-question">Recommendations generated by GenAI tools are convincing and supported by sufficient reasoning. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq7 = likert_scale("GAITQ7")


        # GAITQ8
        st.markdown(
            '<div class="research-question">The information and sources provided by GenAI tools are perceived as trustworthy and credible. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaitq8 = likert_scale("GAITQ8")


        # =================================================
        # SEPARATOR
        # =================================================
        st.markdown(
            "<hr style='border:none;border-top:2px solid #74BFA5;margin:28px 0 10px 0;'>",
            unsafe_allow_html=True
        )


        # =================================================
        # GENAI ADOPTION
        # =================================================
        st.markdown(
            '<div class="construct-heading">GenAI Adoption</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="research-description">Please indicate the extent to which the following statements describe the actual use of GenAI within your organisation.</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="scale-guide">1 = Strongly Disagree &nbsp;&nbsp;&nbsp;&nbsp; 2 = Disagree &nbsp;&nbsp;&nbsp;&nbsp; 3 = Neutral &nbsp;&nbsp;&nbsp;&nbsp; 4 = Agree &nbsp;&nbsp;&nbsp;&nbsp; 5 = Strongly Agree</div>',
            unsafe_allow_html=True
        )


        # GAIA1
        st.markdown(
            '<div class="research-question">Our organisation has adopted GenAI tools as part of its regular business activities. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaia1 = likert_scale("GAIA1")


        # GAIA2
        st.markdown(
            '<div class="research-question">GenAI tools are frequently used for relevant business tasks within our organisation. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaia2 = likert_scale("GAIA2")


        # GAIA3
        st.markdown(
            '<div class="research-question">GenAI tools have been integrated into existing work processes and operations. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaia3 = likert_scale("GAIA3")


        # GAIA4
        st.markdown(
            '<div class="research-question">The use of GenAI tools has expanded across different business activities or functions in our organisation. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaia4 = likert_scale("GAIA4")


        # GAIA5
        st.markdown(
            '<div class="research-question">GenAI use has become an established and continuing practice within our organisation. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        gaia5 = likert_scale("GAIA5")


        # =================================================
        # SEPARATOR
        # =================================================
        st.markdown(
            "<hr style='border:none;border-top:2px solid #74BFA5;margin:28px 0 10px 0;'>",
            unsafe_allow_html=True
        )


        # =================================================
        # BUSINESS PERFORMANCE
        # =================================================
        st.markdown(
            '<div class="construct-heading">Business Performance</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="research-description">Considering the performance of your organisation over the past five years, please indicate the extent to which you agree with the following statements.</div>',
            unsafe_allow_html=True
        )

        st.markdown(
            '<div class="scale-guide">1 = Strongly Disagree &nbsp;&nbsp;&nbsp;&nbsp; 2 = Disagree &nbsp;&nbsp;&nbsp;&nbsp; 3 = Neutral &nbsp;&nbsp;&nbsp;&nbsp; 4 = Agree &nbsp;&nbsp;&nbsp;&nbsp; 5 = Strongly Agree</div>',
            unsafe_allow_html=True
        )


        # ECBP1
        st.markdown(
            '<div class="research-question">Our organisation has generated a satisfactory return on investment. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp1 = likert_scale("ECBP1")


        # ECBP2
        st.markdown(
            '<div class="research-question">Our sales performance has improved. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp2 = likert_scale("ECBP2")


        # ECBP3
        st.markdown(
            '<div class="research-question">The organisation has maintained a strong level of profitability. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp3 = likert_scale("ECBP3")


        # ECBP4
        st.markdown(
            '<div class="research-question">Our business has experienced noticeable overall growth. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp4 = likert_scale("ECBP4")


        # ECBP5
        st.markdown(
            '<div class="research-question">The organisation has strengthened its market position and market share. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp5 = likert_scale("ECBP5")


        # ECBP6
        st.markdown(
            '<div class="research-question">Overall, our organisation has been successful in achieving its business goals. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp6 = likert_scale("ECBP6")


        # ECBP7
        st.markdown(
            '<div class="research-question">Customer satisfaction with our products and services has improved. <span style="color:#C62828;">*</span></div>',
            unsafe_allow_html=True
        )

        ecbp7 = likert_scale("ECBP7")


    # =====================================================
    # NAVIGATION
    # Outside the Likert container
    # =====================================================
    st.markdown("<br>", unsafe_allow_html=True)

    previous_col, spacer_col, submit_col = st.columns(
        [1.5, 3.5, 1.7]
    )


    # -----------------------------------------------------
    # PREVIOUS BUTTON
    # -----------------------------------------------------
    with previous_col:

        if st.button(
            "PREVIOUS",
            key="section3_previous",
            use_container_width=True
        ):

            st.session_state.page = 2
            save_draft()
            st.session_state.scroll_to_top = True
            st.rerun()


    # -----------------------------------------------------
    # SUBMIT BUTTON
    # -----------------------------------------------------
    with submit_col:

        if st.button(
            "SUBMIT",
            key="section3_submit",
            type="primary",
            use_container_width=True
        ):

            research_answers = [
                gaitq1,
                gaitq2,
                gaitq3,
                gaitq4,
                gaitq5,
                gaitq6,
                gaitq7,
                gaitq8,

                gaia1,
                gaia2,
                gaia3,
                gaia4,
                gaia5,

                ecbp1,
                ecbp2,
                ecbp3,
                ecbp4,
                ecbp5,
                ecbp6,
                ecbp7
            ]


            # -------------------------------------------------
            # VALIDATION
            # -------------------------------------------------
            if any(answer is None for answer in research_answers):

                st.error(
                    "Please provide a response to all research statements before submitting the questionnaire."
                )



            else:

                # Permanently store the completed questionnaire

                save_final_response()

                # Open the final thank-you page

                st.session_state["survey_completed"] = True

                st.rerun()

# =========================================================
# SCROLL TO TOP AFTER SECTION NAVIGATION
# =========================================================

if st.session_state.get("scroll_to_top", False):

    scroll_page_to_top()

    st.session_state.scroll_to_top = False

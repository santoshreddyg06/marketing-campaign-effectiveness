import uuid
import random
import threading

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# --------------------------------------------------
# PAGE CONFIGURATION
# --------------------------------------------------

st.set_page_config(
    page_title="Marketing Campaign Effectiveness",
    page_icon="📊",
    layout="centered"
)

st.title("Marketing Campaign Effectiveness Analysis")
st.write(
    "This study examines consumer behaviour and the effectiveness "
    "of different promotional strategies. Please answer honestly."
)


# --------------------------------------------------
# GOOGLE SHEETS CONNECTION
# --------------------------------------------------

SPREADSHEET_ID = "1jDAnymdh0bID9wa6B8e1T1OuDTganGvIxb7QfSaohYo"

HEADERS = [
    "respondent_id",
    "age_group",
    "platform",
    "ad_frequency",
    "promotion_preference",
    "clicked_ad",
    "ad_influence",
    "purchased_after_ad",
    "ignore_reason",
    "condition",
    "purchase_intention",
    "offer_attractiveness",
    "status"
]

AGE_TARGETS = {
    "Below 18": 5,
    "18–22": 15,
    "23–30": 15,
    "Above 30": 15
}

PROMOTION_TARGETS = {
    "Discount": 17,
    "Cashback": 17,
    "No Promotion": 16
}

# Helps coordinate quota checks among sessions on this one app process.
# It does not provide a cross-server/database transaction.
quota_lock = threading.Lock()


@st.cache_resource
def connect_to_sheets():
    credentials = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=[
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
    )

    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    try:
        response_sheet = spreadsheet.worksheet("Experiment Responses")
    except gspread.WorksheetNotFound:
        response_sheet = spreadsheet.add_worksheet(
            title="Experiment Responses",
            rows=1000,
            cols=len(HEADERS)
        )

    if not response_sheet.row_values(1):
        response_sheet.append_row(HEADERS)
    elif response_sheet.row_values(1) != HEADERS:
        raise ValueError(
            "The headers in Experiment Responses do not match "
            "the expected headers. Please check row 1."
        )

    return response_sheet


try:
    sheet = connect_to_sheets()
except Exception as e:
    st.error("Could not connect to Google Sheets.")
    st.error(str(e))
    st.stop()


# --------------------------------------------------
# QUOTA HELPERS
# --------------------------------------------------

def get_valid_records():
    """Count both in-progress and completed responses."""
    records = sheet.get_all_records()
    return [
        record for record in records
        if record.get("status") in ("In progress", "Completed")
    ]


def get_counts(records):
    age_counts = {group: 0 for group in AGE_TARGETS}
    promotion_counts = {group: 0 for group in PROMOTION_TARGETS}

    for record in records:
        age = record.get("age_group")
        condition = record.get("condition")

        if age in age_counts:
            age_counts[age] += 1

        if condition in promotion_counts:
            promotion_counts[condition] += 1

    return age_counts, promotion_counts


def choose_available_promotion(promotion_counts):
    available = [
        group for group, target in PROMOTION_TARGETS.items()
        if promotion_counts[group] < target
    ]

    if not available:
        return None

    # Balance groups while respecting their maximum targets.
    minimum_count = min(promotion_counts[group] for group in available)
    least_filled = [
        group for group in available
        if promotion_counts[group] == minimum_count
    ]
    return random.choice(least_filled)


def get_response_row(respondent_id):
    ids = sheet.col_values(1)
    matching_rows = [
        index + 1
        for index, value in enumerate(ids)
        if value == respondent_id
    ]
    return matching_rows[-1] if matching_rows else None


# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

if "phase" not in st.session_state:
    st.session_state.phase = 1

if "respondent_id" not in st.session_state:
    st.session_state.respondent_id = str(uuid.uuid4())

if "response_row" not in st.session_state:
    st.session_state.response_row = None

if "condition" not in st.session_state:
    st.session_state.condition = None


# --------------------------------------------------
# PHASE 1: CONSUMER BEHAVIOUR SURVEY
# --------------------------------------------------

if st.session_state.phase == 1:
    st.header("Part 1: Consumer Behaviour Survey")
    st.caption("Please answer all questions before continuing.")

    with st.form("survey_form"):
        age_group = st.selectbox(
            "1. What is your age group?",
            ["Select an option", "Below 18", "18–22", "23–30", "Above 30"],
            index=0
        )

        platform = st.selectbox(
            "2. Which platform do you commonly use to discover products?",
            ["Select an option", "Instagram", "YouTube", "Google/Search", "Email", "Other"],
            index=0
        )

        ad_frequency = st.selectbox(
            "3. How frequently do you encounter online advertisements?",
            ["Select an option", "Rarely", "Sometimes", "Often", "Very often"],
            index=0
        )

        promotion_preference = st.selectbox(
            "4. Which promotional offer attracts you the most?",
            ["Select an option", "Discount", "Cashback", "Coupons", "Free delivery", "Buy-one-get-one"],
            index=0
        )

        clicked_ad = st.radio(
            "5. Have you ever clicked on an online advertisement?",
            ["Select an option", "Yes", "No"],
            index=0
        )

        ad_influence = st.selectbox(
            "6. How frequently do advertisements influence your shopping decisions?",
            ["Select an option", "Never", "Rarely", "Sometimes", "Often", "Always"],
            index=0
        )

        purchased_after_ad = st.radio(
            "7. Have you ever purchased a product after seeing an advertisement?",
            ["Select an option", "Yes", "No"],
            index=0
        )

        ignore_reason = st.selectbox(
            "8. What is the main reason you ignore advertisements?",
            [
                "Select an option",
                "Irrelevant products",
                "Lack of trust",
                "No attractive offers",
                "Too many advertisements",
                "Other"
            ],
            index=0
        )

        survey_submitted = st.form_submit_button(
            "Continue to Promotional Experiment"
        )

    if survey_submitted:
        answers = {
            "Age group": age_group,
            "Platform": platform,
            "Advertisement frequency": ad_frequency,
            "Promotion preference": promotion_preference,
            "Clicked advertisement": clicked_ad,
            "Advertisement influence": ad_influence,
            "Purchased after advertisement": purchased_after_ad,
            "Reason for ignoring advertisements": ignore_reason
        }

        unanswered = [
            question for question, answer in answers.items()
            if answer == "Select an option"
        ]

        if unanswered:
            st.error("Please answer all questions before continuing.")
            for question in unanswered:
                st.write(f"- {question}")

        else:
            try:
                # Quota check and row creation are coordinated within this
                # running app process. In-progress rows reserve quota slots.
                with quota_lock:
                    records = get_valid_records()
                    age_counts, promotion_counts = get_counts(records)

                    total_count = len(records)
                    if total_count >= sum(AGE_TARGETS.values()):
                        st.error(
                            "The study has reached its target of 50 responses. "
                            "Thank you for your interest."
                        )
                        st.stop()

                    if age_counts[age_group] >= AGE_TARGETS[age_group]:
                        st.error(
                            f"The target for the '{age_group}' age group "
                            "has been reached. Please do not select a different "
                            "age group unless it accurately describes your age."
                        )
                        st.stop()

                    condition = choose_available_promotion(promotion_counts)
                    if condition is None:
                        st.error(
                            "All promotional experiment groups have reached "
                            "their targets. Response collection is closed."
                        )
                        st.stop()

                    row = [
                        st.session_state.respondent_id,
                        age_group,
                        platform,
                        ad_frequency,
                        promotion_preference,
                        clicked_ad,
                        ad_influence,
                        purchased_after_ad,
                        ignore_reason,
                        condition,
                        "",
                        "",
                        "In progress"
                    ]

                    sheet.append_row(row, value_input_option="USER_ENTERED")
                    response_row = get_response_row(
                        st.session_state.respondent_id
                    )

                    if response_row is None:
                        st.error(
                            "Your survey was saved, but the response row "
                            "could not be located. Please contact the research "
                            "team instead of submitting again."
                        )
                        st.stop()

                    st.session_state.response_row = response_row
                    st.session_state.condition = condition
                    st.session_state.phase = 2

                st.rerun()

            except Exception as e:
                st.error("Could not save your survey responses.")
                st.error(str(e))


# --------------------------------------------------
# PHASE 2: PROMOTIONAL EXPERIMENT
# --------------------------------------------------

elif st.session_state.phase == 2:
    st.header("Part 2: Promotional Experiment")
    condition = st.session_state.condition

    st.write("Imagine you are shopping online and see the following offer.")

    if condition == "Discount":
        st.subheader("Special Offer: ₹400 Discount")
        st.info("You receive a ₹400 discount on your purchase.")
    elif condition == "Cashback":
        st.subheader("Special Offer: ₹400 Cashback")
        st.info("You receive ₹400 cashback after your purchase.")
    else:
        st.subheader("Online Shopping")
        st.info(
            "You see the product advertised, but no promotional "
            "offer is available."
        )

    st.write(
        "Assume the product is relevant to your needs and its price is affordable."
    )

    with st.form("experiment_form"):
        purchase_intention = st.radio(
            "1. How likely are you to purchase the product?",
            [
                "Select an option",
                "1 - Very unlikely",
                "2 - Unlikely",
                "3 - Neutral",
                "4 - Likely",
                "5 - Very likely"
            ],
            index=0
        )

        offer_attractiveness = st.radio(
            "2. How attractive do you find this offer?",
            [
                "Select an option",
                "1 - Not attractive at all",
                "2 - Slightly attractive",
                "3 - Moderately attractive",
                "4 - Very attractive",
                "5 - Extremely attractive"
            ],
            index=0
        )

        experiment_submitted = st.form_submit_button("Submit My Responses")

    if experiment_submitted:
        if (
            purchase_intention == "Select an option"
            or offer_attractiveness == "Select an option"
        ):
            st.error("Please answer both experiment questions.")
        else:
            try:
                response_row = st.session_state.response_row

                if response_row is None:
                    st.error(
                        "Your survey response could not be located. "
                        "Please contact the research team."
                    )
                    st.stop()

                purchase_score = int(purchase_intention[0])
                attractiveness_score = int(offer_attractiveness[0])

                with quota_lock:
                    # Avoid updating an already completed response.
                    current_status = sheet.cell(response_row, 13).value
                    if current_status == "Completed":
                        st.session_state.phase = 3
                        st.rerun()

                    sheet.update(
                        range_name=f"K{response_row}:M{response_row}",
                        values=[[
                            purchase_score,
                            attractiveness_score,
                            "Completed"
                        ]]
                    )

                st.session_state.phase = 3
                st.rerun()

            except Exception as e:
                st.error("Could not save your experiment responses.")
                st.error(str(e))


# --------------------------------------------------
# PHASE 3: COMPLETION
# --------------------------------------------------

elif st.session_state.phase == 3:
    st.success("Thank you for participating!")
    st.write("Your responses have been recorded successfully.")
    st.write("You may now close this page.")

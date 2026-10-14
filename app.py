import streamlit as st
import gspread
import random
import uuid

st.set_page_config(
    page_title="Marketing Campaign Analysis",
    layout="centered"
)

SHEET_ID = "1jDAnymdh0bID9wa6B8e1T1OuDTganGvIxb7QfSaohYo"

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


@st.cache_resource
def connect_sheet():
    gc = gspread.service_account(
        filename="service_account.json"
    )

    spreadsheet = gc.open_by_key(SHEET_ID)
    return spreadsheet.worksheet("Experiment Responses")


try:
    sheet = connect_sheet()

    if sheet.row_values(1) != HEADERS:
        if len(sheet.get_all_values()) > 1:
            st.error(
                "The worksheet headers do not match the application. "
                "Please resolve this before collecting responses."
            )
            st.stop()

        sheet.update(
            range_name="A1:M1",
            values=[HEADERS]
        )

except Exception as e:
    st.error(f"Google Sheets connection failed: {e}")
    st.stop()


def get_records():
    records = sheet.get_all_records()
    return records


def choose_promotion():
    records = get_records()

    counts = {
        group: 0 for group in PROMOTION_TARGETS
    }

    for record in records:
        group = record.get("condition")

        if group in counts:
            counts[group] += 1

    available = [
        group
        for group, target in PROMOTION_TARGETS.items()
        if counts[group] < target
    ]

    if not available:
        return None

    minimum = min(counts[group] for group in available)

    candidates = [
        group for group in available
        if counts[group] == minimum
    ]

    return random.choice(candidates)


def find_row(respondent_id):
    ids = sheet.col_values(1)

    try:
        return ids.index(respondent_id) + 1
    except ValueError:
        return None


if "respondent_id" not in st.session_state:
    st.session_state.respondent_id = str(uuid.uuid4())

if "phase" not in st.session_state:
    st.session_state.phase = 1

if "submitted" not in st.session_state:
    st.session_state.submitted = False


st.title("Marketing Campaign Effectiveness Analysis")

st.write(
    "This academic study explores digital advertising, "
    "consumer behavior, and promotional offers."
)

st.info(
    "Participation is voluntary. The shopping scenario is hypothetical. "
    "Please do not provide unnecessary personal information."
)


if st.session_state.submitted:
    st.success("Your response has been recorded. Thank you!")
    st.stop()


# --------------------------------------------------
# PHASE 1: CONSUMER SURVEY
# --------------------------------------------------

if st.session_state.phase == 1:

    st.header("Part A: Consumer Behavior")

    with st.form("consumer_survey"):

        age = st.selectbox(
            "1. What is your age group?",
            ["Select an option"] + list(AGE_TARGETS),
            index=0
        )

        platform = st.selectbox(
            "2. Where do you encounter advertisements most frequently?",
            [
                "Select an option",
                "Instagram",
                "YouTube",
                "Google/Search",
                "Email",
                "Other"
            ],
            index=0
        )

        platform_other = ""

        if platform == "Other":
            platform_other = st.text_input(
                "Specify the platform"
            )

        frequency = st.selectbox(
            "3. How frequently do you encounter digital advertisements?",
            [
                "Select an option",
                "Rarely",
                "Sometimes",
                "Often",
                "Very often"
            ],
            index=0
        )

        preference = st.selectbox(
            "4. Which promotional offer attracts you most?",
            [
                "Select an option",
                "Discount",
                "Cashback",
                "Coupon",
                "Free delivery",
                "Buy-one-get-one",
                "Other"
            ],
            index=0
        )

        preference_other = ""

        if preference == "Other":
            preference_other = st.text_input(
                "Specify the promotional offer"
            )

        clicked = st.selectbox(
            "5. Have you clicked a digital advertisement in the past three months?",
            ["Select an option", "Yes", "No"],
            index=0
        )

        influence = st.selectbox(
            "6. How often do advertisements influence your purchases?",
            [
                "Select an option",
                "Never",
                "Rarely",
                "Sometimes",
                "Often",
                "Always"
            ],
            index=0
        )

        purchased = st.selectbox(
            "7. Have you purchased a product after seeing a digital advertisement in the past three months?",
            ["Select an option", "Yes", "No"],
            index=0
        )

        reason = st.selectbox(
            "8. What is the main reason you ignore advertisements?",
            [
                "Select an option",
                "Irrelevant products",
                "Lack of trust",
                "Unattractive offers",
                "Too many ads",
                "Other"
            ],
            index=0
        )

        reason_other = ""

        if reason == "Other":
            reason_other = st.text_input(
                "Specify the reason"
            )

        start_experiment = st.form_submit_button(
            "Continue to Experiment"
        )

    if start_experiment:

        required = [
            age != "Select an option",
            platform != "Select an option",
            frequency != "Select an option",
            preference != "Select an option",
            clicked != "Select an option",
            influence != "Select an option",
            purchased != "Select an option",
            reason != "Select an option"
        ]

        if not all(required):
            st.error("Please answer all eight questions.")

        elif platform == "Other" and not platform_other.strip():
            st.error("Please specify the platform.")

        elif preference == "Other" and not preference_other.strip():
            st.error("Please specify the promotional offer.")

        elif reason == "Other" and not reason_other.strip():
            st.error("Please specify the reason.")

        else:
            records = get_records()

            age_count = sum(
                1 for record in records
                if record.get("age_group") == age
            )

            if age_count >= AGE_TARGETS[age]:
                st.error(
                    "This age-group target has been reached. "
                    "Please contact the study organizer."
                )

            else:
                condition = choose_promotion()

                if condition is None:
                    st.error(
                        "All experimental group targets have been reached."
                    )

                else:
                    row = [
                        st.session_state.respondent_id,
                        age,
                        platform_other.strip()
                        if platform == "Other" else platform,
                        frequency,
                        preference_other.strip()
                        if preference == "Other" else preference,
                        clicked,
                        influence,
                        purchased,
                        reason_other.strip()
                        if reason == "Other" else reason,
                        condition,
                        "",
                        "",
                        "In progress"
                    ]

                    try:
                        sheet.append_row(row)

                        st.session_state.condition = condition
                        st.session_state.phase = 2

                        st.rerun()

                    except Exception as e:
                        st.error(f"Could not save your response: {e}")


# --------------------------------------------------
# PHASE 2: PROMOTIONAL EXPERIMENT
# --------------------------------------------------

if st.session_state.phase == 2:

    st.header("Part B: Promotional Experiment")

    condition = st.session_state.condition

    st.write(
        "Imagine you are considering wireless headphones "
        "originally priced at ₹2,000."
    )

    if condition == "Discount":
        st.success(
            "₹400 discount. You pay ₹1,600 at checkout."
        )

    elif condition == "Cashback":
        st.success(
            "₹400 cashback within 7 days after purchase. "
            "You pay ₹2,000 at checkout."
        )

    else:
        st.info(
            "No promotion is available. You pay ₹2,000."
        )

    with st.form("experiment_form"):

        purchase_intention = st.select_slider(
            "9. How likely are you to purchase this product?",
            options=[1, 2, 3, 4, 5],
            value=None,
            format_func=lambda x: {
                1: "1 - Very unlikely",
                2: "2 - Unlikely",
                3: "3 - Neither likely nor unlikely",
                4: "4 - Likely",
                5: "5 - Very likely"
            }[x]
        )

        attractiveness = st.select_slider(
            "10. How attractive is this offer?",
            options=[1, 2, 3, 4, 5],
            value=None,
            format_func=lambda x: {
                1: "1 - Not at all attractive",
                2: "2 - Slightly attractive",
                3: "3 - Moderately attractive",
                4: "4 - Very attractive",
                5: "5 - Extremely attractive"
            }[x]
        )

        finish = st.form_submit_button("Submit Final Response")

    if finish:

        if purchase_intention is None or attractiveness is None:
            st.error("Please answer both experimental questions.")

        else:
            try:
                row_number = find_row(
                    st.session_state.respondent_id
                )

                if row_number is None:
                    st.error(
                        "Your survey record could not be found. "
                        "Please contact the study organizer."
                    )

                else:
                    sheet.update(
                        range_name=f"K{row_number}:M{row_number}",
                        values=[[
                            purchase_intention,
                            attractiveness,
                            "Completed"
                        ]]
                    )

                    st.session_state.submitted = True

                    st.success(
                        "Thank you! Your responses have been recorded."
                    )

            except Exception as e:
                st.error(f"Could not save your response: {e}")
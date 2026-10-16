import random
import uuid
from datetime import datetime, timezone

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Marketing Campaign Effectiveness", page_icon="📊", layout="centered")

SPREADSHEET_ID = "1jDAnymdh0bID9wa6B8e1T1OuDTganGvIxb7QfSaohYo"

AGE_QUOTAS = {
    "Under 18": 5,
    "18–22": 15,
    "23–30": 15,
    "31–45": 10,
    "Above 45": 5,
}

SURVEY_HEADERS = [
    "respondent_id", "timestamp_utc", "age_group", "discovery_platform",
    "ad_frequency", "online_product_category", "clicked_ad_3m",
    "purchased_after_ad_3m", "purchase_ad_platform", "ad_trust",
    "purchase_influence", "ad_ignore_reason", "discount_importance",
    "preferred_promotion", "choose_promoted_product_likelihood",
    "assigned_condition", "status"
]

EXPERIMENT_HEADERS = [
    "respondent_id", "assigned_condition", "purchase_intention",
    "offer_attractiveness", "value_for_money_confidence", "timestamp_utc"
]

PLATFORMS = [
    "Instagram", "YouTube", "Google/Search",
    "E-commerce apps (Amazon, Flipkart, etc.)", "Facebook", "Other"
]

CONDITIONS = {
    "Discount": {
        "title": "₹400 Immediate Discount",
        "description": (
            "The headphones have an original price of ₹2,000. You can buy them "
            "now for ₹1,600 after an immediate ₹400 discount."
        ),
    },
    "Cashback": {
        "title": "₹400 Cashback",
        "description": (
            "The headphones cost ₹2,000. You pay ₹2,000 at checkout and receive "
            "₹400 cashback to your original payment method within 7 days. "
            "No additional purchase or minimum-spend condition applies."
        ),
    },
    "No Promotion": {
        "title": "No Promotional Offer",
        "description": (
            "The headphones cost ₹2,000. No discount, cashback, coupon, or other "
            "promotional offer is available."
        ),
    },
}

def connect_to_sheets():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    if "gcp_service_account" in st.secrets:
        credentials = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes
        )
    else:
        credentials = Credentials.from_service_account_file(
            "service_account.json", scopes=scopes
        )

    spreadsheet = gspread.authorize(credentials).open_by_key(SPREADSHEET_ID)

    def get_or_create(title, headers, rows=1000):
        try:
            worksheet = spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(
                title=title, rows=rows, cols=len(headers)
            )

        first_row = worksheet.row_values(1)
        if not first_row:
            worksheet.append_row(headers)
        elif first_row != headers:
            raise ValueError(
                f"Headers in worksheet '{title}' do not match this app's expected "
                "columns. Back up your existing data before changing row 1. "
                f"Expected: {headers}"
            )
        return worksheet

    survey_ws = get_or_create("Survey Responses", SURVEY_HEADERS, 1000)
    experiment_ws = get_or_create("Experiment Responses", EXPERIMENT_HEADERS, 1000)
    return survey_ws, experiment_ws

st.title("Marketing Campaign Effectiveness Analysis")
st.write(
    "This voluntary study examines online advertising behaviour, consumer trust, "
    "price sensitivity, and responses to promotional offers."
)
st.caption(
    "Please do not enter your name, email address, phone number, or other direct "
    "identifying information. Responses are for academic research."
)

try:
    survey_ws, experiment_ws = connect_to_sheets()
except Exception as exc:
    st.error("Could not connect to Google Sheets. Check credentials, sharing permissions, and worksheet headers.")
    st.code(str(exc))
    st.stop()

def records_for_age_group(group):
    try:
        return sum(
            1 for row in survey_ws.get_all_records()
            if row.get("age_group") == group and row.get("status") == "Completed"
        )
    except Exception:
        return 0

def quota_status():
    return {group: (records_for_age_group(group), target) for group, target in AGE_QUOTAS.items()}

if "page" not in st.session_state:
    st.session_state.page = "survey"
if "respondent_id" not in st.session_state:
    st.session_state.respondent_id = str(uuid.uuid4())
if "survey_data" not in st.session_state:
    st.session_state.survey_data = {}
if "assigned_condition" not in st.session_state:
    st.session_state.assigned_condition = None

if st.session_state.page == "survey":
    st.header("Part 1 — Consumer Behaviour Survey")
    st.write("There are 12 core questions and one conditional follow-up. Most people finish in 3–5 minutes.")

    with st.expander("Participant targets by age group"):
        for group, (count, target) in quota_status().items():
            st.write(f"**{group}:** {count}/{target}")

    st.info(
        "Participation is voluntary. If you are under 18, participate only if the "
        "research team's consent and institutional requirements permit it."
    )

    with st.form("consumer_survey"):
        age_group = st.radio(
            "1. What is your age group?",
            ["Select an option", *AGE_QUOTAS.keys()],
        )
        discovery_platform = st.selectbox(
            "2. Which platform do you most often use to discover or learn about products?",
            ["Select an option", *PLATFORMS],
        )
        ad_frequency = st.radio(
            "3. How frequently do you encounter online advertisements?",
            ["Select an option", "Never", "Rarely", "Sometimes", "Often", "Very often"],
        )
        product_category = st.selectbox(
            "4. Which type of product do you most frequently purchase online?",
            [
                "Select an option", "Electronics", "Clothing and fashion",
                "Food and groceries", "Beauty and personal care",
                "Home and lifestyle", "Other / I rarely purchase online",
            ],
        )
        clicked_ad = st.radio(
            "5. In the past three months, have you clicked on an online advertisement to learn more about a product?",
            ["Select an option", "Yes", "No", "Not sure"],
        )
        purchased_after_ad = st.radio(
            "6. In the past three months, have you purchased a product after seeing an online advertisement?",
            ["Select an option", "Yes", "No", "Not sure"],
        )
        purchase_ad_platform = st.selectbox(
            "6a. If yes, which platform did you see the advertisement on before purchasing?",
            ["Select an option", *PLATFORMS, "Cannot remember"],
            disabled=(purchased_after_ad != "Yes"),
        )
        ad_trust = st.radio(
            "7. How much do you trust product information shown in online advertisements?",
            [
                "Select an option", "1 — Do not trust at all", "2 — Trust very little",
                "3 — Neither trust nor distrust", "4 — Trust somewhat",
                "5 — Trust completely",
            ],
        )
        purchase_influence = st.selectbox(
            "8. What most influences your decision to purchase a product after seeing an advertisement?",
            [
                "Select an option", "Price or discount", "Product quality and features",
                "Customer reviews and ratings", "Brand reputation",
                "Recommendation from friends or family", "The advertisement itself", "Other",
            ],
        )
        ignore_reason = st.selectbox(
            "9. What is the main reason you ignore an online advertisement?",
            [
                "Select an option", "The product is irrelevant to me",
                "I do not trust the advertisement", "The price or offer is unattractive",
                "I see too many advertisements", "I prefer researching products independently",
                "Other",
            ],
        )
        discount_importance = st.radio(
            "10. How important is a discount when deciding whether to purchase a product?",
            [
                "Select an option", "1 — Not at all important",
                "2 — Slightly important", "3 — Moderately important",
                "4 — Very important", "5 — Extremely important",
            ],
        )
        preferred_promotion = st.selectbox(
            "11. Which promotional offer would most encourage you to consider purchasing a product?",
            [
                "Select an option", "Immediate price discount", "Cashback after purchase",
                "Coupon or promo code", "Free delivery", "Buy-one-get-one offer",
                "None of these",
            ],
        )
        choose_promoted = st.radio(
            "12. If two products were otherwise equally suitable, how likely would you be to choose the one with a promotional offer?",
            [
                "Select an option", "1 — Very unlikely", "2 — Unlikely",
                "3 — Neither likely nor unlikely", "4 — Likely", "5 — Very likely",
            ],
        )

        consent = st.checkbox("I understand the study information and voluntarily agree to participate.")
        submitted = st.form_submit_button("Submit survey and continue")

    if submitted:
        answers = {
            "Age group": age_group,
            "Discovery platform": discovery_platform,
            "Ad frequency": ad_frequency,
            "Product category": product_category,
            "Clicked ad": clicked_ad,
            "Purchased after ad": purchased_after_ad,
            "Ad trust": ad_trust,
            "Purchase influence": purchase_influence,
            "Reason for ignoring ads": ignore_reason,
            "Discount importance": discount_importance,
            "Preferred promotion": preferred_promotion,
            "Likelihood of choosing promoted product": choose_promoted,
        }
        missing = [key for key, value in answers.items() if value == "Select an option"]
        if purchased_after_ad == "Yes" and purchase_ad_platform == "Select an option":
            missing.append("Platform where the advertisement was seen before purchasing")
        if not consent:
            st.error("Please confirm consent before participating.")
        elif missing:
            st.error("Please answer the following required questions: " + ", ".join(missing))
        else:
            counts = quota_status()
            current_count, target = counts[age_group]
            if current_count >= target:
                st.error(
                    f"The target for '{age_group}' has been reached. Please contact the research team "
                    "if you believe you should still participate."
                )
            else:
                condition = random.choice(list(CONDITIONS.keys()))
                st.session_state.assigned_condition = condition
                st.session_state.survey_data = {
                    "age_group": age_group,
                    "discovery_platform": discovery_platform,
                    "ad_frequency": ad_frequency,
                    "online_product_category": product_category,
                    "clicked_ad_3m": clicked_ad,
                    "purchased_after_ad_3m": purchased_after_ad,
                    "purchase_ad_platform": purchase_ad_platform if purchased_after_ad == "Yes" else "Not applicable",
                    "ad_trust": int(ad_trust[0]),
                    "purchase_influence": purchase_influence,
                    "ad_ignore_reason": ignore_reason,
                    "discount_importance": int(discount_importance[0]),
                    "preferred_promotion": preferred_promotion,
                    "choose_promoted_product_likelihood": int(choose_promoted[0]),
                }
                try:
                    survey_ws.append_row([
                        st.session_state.respondent_id,
                        datetime.now(timezone.utc).isoformat(),
                        age_group, discovery_platform, ad_frequency, product_category,
                        clicked_ad, purchased_after_ad,
                        purchase_ad_platform if purchased_after_ad == "Yes" else "Not applicable",
                        int(ad_trust[0]), purchase_influence, ignore_reason,
                        int(discount_importance[0]), preferred_promotion, int(choose_promoted[0]),
                        condition, "In progress",
                    ], value_input_option="USER_ENTERED")
                    st.session_state.page = "experiment"
                    st.rerun()
                except Exception as exc:
                    st.error("Your survey could not be saved. Please contact the research team.")
                    st.code(str(exc))

elif st.session_state.page == "experiment":
    condition = st.session_state.assigned_condition
    scenario = CONDITIONS[condition]
    st.header("Part 2 — Promotional Experiment")
    st.write(
        "Imagine you are considering purchasing the same pair of wireless headphones. "
        "Assume the product meets your needs. Read the assigned scenario and answer based on your own opinion."
    )
    st.subheader(scenario["title"])
    st.info(scenario["description"])
    st.caption("Only one scenario is shown to each participant. Please answer based on the scenario above.")

    with st.form("experiment_form"):
        purchase_intention = st.radio(
            "13. How likely are you to purchase this product?",
            [
                "Select an option", "1 — Very unlikely", "2 — Unlikely",
                "3 — Neither likely nor unlikely", "4 — Likely", "5 — Very likely",
            ],
        )
        if condition == "No Promotion":
            attractiveness_prompt = "14. How attractive do you find this purchase option?"
        else:
            attractiveness_prompt = "14. How attractive do you find this offer?"
        attractiveness = st.radio(
            attractiveness_prompt,
            [
                "Select an option", "1 — Not attractive at all",
                "2 — Slightly attractive", "3 — Moderately attractive",
                "4 — Very attractive", "5 — Extremely attractive",
            ],
        )
        value_confidence = st.radio(
            "15. How confident are you that this option provides good value for money?",
            [
                "Select an option", "1 — Not confident at all",
                "2 — Slightly confident", "3 — Moderately confident",
                "4 — Very confident", "5 — Extremely confident",
            ],
        )
        submitted = st.form_submit_button("Finish and submit")

    if submitted:
        if "Select an option" in (purchase_intention, attractiveness, value_confidence):
            st.error("Please answer all three questions.")
        else:
            try:
                experiment_ws.append_row([
                    st.session_state.respondent_id,
                    condition,
                    int(purchase_intention[0]),
                    int(attractiveness[0]),
                    int(value_confidence[0]),
                    datetime.now(timezone.utc).isoformat(),
                ], value_input_option="USER_ENTERED")

                survey_ids = survey_ws.col_values(1)
                row_number = max(
                    (i for i, value in enumerate(survey_ids, start=1)
                     if value == st.session_state.respondent_id),
                    default=None,
                )
                if row_number:
                    survey_ws.update(
                        range_name=f"Q{row_number}:Q{row_number}",
                        values=[["Completed"]],
                    )
                st.session_state.page = "complete"
                st.rerun()
            except Exception as exc:
                st.error("We could not save the experimental response. Please contact the research team.")
                st.code(str(exc))

elif st.session_state.page == "complete":
    st.success("Thank you! Your responses have been recorded.")
    st.write("Your consumer survey and one randomly assigned promotional scenario have been submitted.")
    st.write("You may now close this page.")
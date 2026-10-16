import random
import uuid
from datetime import datetime, timezone

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

st.set_page_config(
    page_title="Marketing Campaign Effectiveness Analysis",
    page_icon="📊",
    layout="centered",
)

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------
SPREADSHEET_ID = "1M4rsS-XfIA2IPjAWTzqaMylK6gm95tEKqGIqSvHc2G0"

AGE_QUOTAS = {
    "Under 18": 8,
    "18–22": 24,
    "23–30": 24,
    "31–45": 16,
    "Above 45": 8,
}

SURVEY_HEADERS = [
    "respondent_id",
    "timestamp_utc",
    "age_group",
    "discovery_platform",
    "ad_frequency",
    "online_product_category",
    "clicked_ad_3m",
    "purchased_after_ad_3m",
    "purchase_ad_platform",
    "ad_trust",
    "purchase_influence",
    "ad_ignore_reason",
    "discount_importance",
    "preferred_promotion",
    "choose_promoted_product_likelihood",
    "status",
]

EXPERIMENT_HEADERS = [
    "respondent_id",
    "scenario_order",
    "scenario_condition",
    "purchase_intention",
    "offer_attractiveness",
    "value_for_money_confidence",
    "timestamp_utc",
]

PLATFORMS = [
    "Instagram",
    "YouTube",
    "Google/Search",
    "E-commerce apps (Amazon, Flipkart, etc.)",
    "Facebook",
    "Other",
]

SCENARIOS = {
    "Discount": {
        "title": "Scenario A — ₹400 Immediate Discount",
        "description": (
            "Imagine you are considering the same pair of wireless headphones, "
            "priced at ₹2,000 before any offer. You can buy them now for ₹1,600 "
            "after an immediate ₹400 discount."
        ),
        "image": "assets/scenario_discount.png",
    },
    "Cashback": {
        "title": "Scenario B — ₹400 Cashback",
        "description": (
            "Imagine you are considering the same pair of wireless headphones, "
            "priced at ₹2,000. You pay ₹2,000 at checkout and receive ₹400 "
            "cashback to your original payment method within 7 days. No extra "
            "purchase or minimum-spend condition applies."
        ),
        "image": "assets/scenario_cashback.png",
    },
    "No Promotion": {
        "title": "Scenario C — No Promotion",
        "description": (
            "Imagine you are considering the same pair of wireless headphones, "
            "priced at ₹2,000. No discount, cashback, coupon, or other promotional "
            "offer is available."
        ),
        "image": "assets/scenario_no_promotion.png",
    },
}

LIKELIHOOD_OPTIONS = [
    "1 — Very unlikely",
    "2 — Unlikely",
    "3 — Neither likely nor unlikely",
    "4 — Likely",
    "5 — Very likely",
]
ATTRACTIVENESS_OPTIONS = [
    "1 — Not attractive at all",
    "2 — Slightly attractive",
    "3 — Moderately attractive",
    "4 — Very attractive",
    "5 — Extremely attractive",
]
VALUE_OPTIONS = [
    "1 — Not confident at all",
    "2 — Slightly confident",
    "3 — Moderately confident",
    "4 — Very confident",
    "5 — Extremely confident",
]


# ---------------------------------------------------------------------
# Google Sheets connection
# ---------------------------------------------------------------------
@st.cache_resource
def connect_to_spreadsheet():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]

    # Prefer the local JSON file when running on your computer.
    # Streamlit Cloud can use the gcp_service_account secret instead.
    from pathlib import Path

    local_credentials_file = Path(__file__).resolve().parent / "service_account.json"

    if local_credentials_file.is_file():
        credentials = Credentials.from_service_account_file(
            str(local_credentials_file),
            scopes=scopes,
        )
    else:
        try:
            service_account_secret = st.secrets.get("gcp_service_account")
        except Exception:
            service_account_secret = None

        if not service_account_secret:
            raise FileNotFoundError(
                "Google credentials were not found. For local use, place "
                "'service_account.json' beside app.py. For Streamlit Cloud, "
                "configure [gcp_service_account] in the app's Secrets settings."
            )

        credentials = Credentials.from_service_account_info(
            dict(service_account_secret),
            scopes=scopes,
        )

    client = gspread.authorize(credentials)
    spreadsheet = client.open_by_key(SPREADSHEET_ID)

    def get_or_create_worksheet(title, headers, rows=1000):
        try:
            worksheet = spreadsheet.worksheet(title)
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(
                title=title,
                rows=rows,
                cols=len(headers),
            )

        first_row = worksheet.row_values(1)
        if not first_row:
            worksheet.append_row(headers)
        elif first_row != headers:
            raise ValueError(
                f"Worksheet '{title}' has different headers. "
                "Back up its contents and use a new worksheet or correct "
                "the header row before collecting data."
            )
        return worksheet

    survey_ws = get_or_create_worksheet("Survey Responses", SURVEY_HEADERS)
    experiment_ws = get_or_create_worksheet(
        "Experiment Responses", EXPERIMENT_HEADERS, rows=2000
    )
    return spreadsheet, survey_ws, experiment_ws


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------
def age_counts(survey_ws):
    counts = {age: 0 for age in AGE_QUOTAS}
    for row in survey_ws.get_all_records():
        if row.get("status") == "Completed" and row.get("age_group") in counts:
            counts[row["age_group"]] += 1
    return counts


def score_from_option(option):
    return int(option[0])


def ensure_session_state():
    defaults = {
        "page": "survey",
        "respondent_id": str(uuid.uuid4()),
        "survey_saved": False,
        "survey_row": None,
        "survey_data": {},
        "scenario_order": [],
        "scenario_index": 0,
        "experiment_responses": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


# ---------------------------------------------------------------------
# App
# ---------------------------------------------------------------------
st.title("Marketing Campaign Effectiveness Analysis")
st.write(
    "This voluntary academic study examines online advertising behaviour, "
    "consumer trust, price sensitivity, and responses to promotional offers."
)
st.caption(
    "Do not enter your name, email address, phone number, or other direct "
    "identifying information. Participation is voluntary."
)

try:
    spreadsheet, survey_ws, experiment_ws = connect_to_spreadsheet()
except Exception as exc:
    st.error(
        "Could not connect to Google Sheets. Check the spreadsheet ID, "
        "credentials, sharing permissions, and worksheet headers."
    )
    st.code(str(exc))
    st.stop()

ensure_session_state()

if st.session_state.page == "survey":
    st.header("Part 1 — Consumer Behaviour Survey")
    st.write("12 core questions plus one conditional follow-up. Approx. 3–5 minutes.")

    counts = age_counts(survey_ws)
    with st.expander("Participant targets by age group"):
        for age, target in AGE_QUOTAS.items():
            st.write(f"**{age}:** {counts[age]}/{target}")
        st.write(f"**Total:** {sum(counts.values())}/80")

    st.info(
        "If you are under 18, participate only if the research team's consent "
        "and institutional requirements permit it."
    )

    with st.form("survey_form"):
        age_group = st.selectbox(
            "1. What is your age group?",
            list(AGE_QUOTAS.keys()),
            index=None,
            placeholder="Choose your age group",
        )
        discovery_platform = st.selectbox(
            "2. Which platform do you most often use to discover or learn about products?",
            PLATFORMS,
            index=None,
            placeholder="Choose a platform",
        )
        ad_frequency = st.radio(
            "3. How frequently do you encounter online advertisements?",
            ["Never", "Rarely", "Sometimes", "Often", "Very often"],
            index=None,
        )
        product_category = st.selectbox(
            "4. Which type of product do you most frequently purchase online?",
            [
                "Electronics",
                "Clothing and fashion",
                "Food and groceries",
                "Beauty and personal care",
                "Home and lifestyle",
                "Other / I rarely purchase online",
            ],
            index=None,
            placeholder="Choose a product category",
        )
        clicked_ad = st.radio(
            "5. In the past three months, have you clicked on an online advertisement to learn more about a product?",
            ["Yes", "No", "Not sure"],
            index=None,
        )
        purchased_after_ad = st.radio(
            "6. In the past three months, have you purchased a product after seeing an online advertisement?",
            ["Yes", "No", "Not sure"],
            index=None,
        )
        purchase_ad_platform = st.selectbox(
            "6a. If yes, which platform did you see the advertisement on before purchasing?",
            [*PLATFORMS, "Cannot remember"],
            index=None,
            placeholder="Choose a platform",
        )
        ad_trust = st.radio(
            "7. How much do you trust product information shown in online advertisements?",
            [
                "1 — Do not trust at all",
                "2 — Trust very little",
                "3 — Neither trust nor distrust",
                "4 — Trust somewhat",
                "5 — Trust completely",
            ],
            index=None,
        )
        purchase_influence = st.selectbox(
            "8. What most influences your decision to purchase a product after seeing an advertisement?",
            [
                "Price or discount",
                "Product quality and features",
                "Customer reviews and ratings",
                "Brand reputation",
                "Recommendation from friends or family",
                "The advertisement itself",
                "Other",
            ],
            index=None,
            placeholder="Choose an answer",
        )
        ignore_reason = st.selectbox(
            "9. What is the main reason you ignore an online advertisement?",
            [
                "The product is irrelevant to me",
                "I do not trust the advertisement",
                "The price or offer is unattractive",
                "I see too many advertisements",
                "I prefer researching products independently",
                "Other",
            ],
            index=None,
            placeholder="Choose an answer",
        )
        discount_importance = st.radio(
            "10. How important is a discount when deciding whether to purchase a product?",
            [
                "1 — Not at all important",
                "2 — Slightly important",
                "3 — Moderately important",
                "4 — Very important",
                "5 — Extremely important",
            ],
            index=None,
        )
        preferred_promotion = st.selectbox(
            "11. Which promotional offer would most encourage you to consider purchasing a product?",
            [
                "Immediate price discount",
                "Cashback after purchase",
                "Coupon or promo code",
                "Free delivery",
                "Buy-one-get-one offer",
                "None of these",
            ],
            index=None,
            placeholder="Choose a promotion",
        )
        choose_promoted = st.radio(
            "12. If two products were otherwise equally suitable, how likely would you be to choose the one with a promotional offer?",
            LIKELIHOOD_OPTIONS,
            index=None,
        )
        consent = st.checkbox(
            "I have read the study information and voluntarily agree to participate."
        )
        submitted = st.form_submit_button("Submit survey and continue")

    if submitted:
        required = {
            "Age group": age_group,
            "Discovery platform": discovery_platform,
            "Ad frequency": ad_frequency,
            "Product category": product_category,
            "Clicked advertisement": clicked_ad,
            "Purchased after advertisement": purchased_after_ad,
            "Advertising trust": ad_trust,
            "Purchase influence": purchase_influence,
            "Reason for ignoring advertisements": ignore_reason,
            "Discount importance": discount_importance,
            "Preferred promotion": preferred_promotion,
            "Likelihood of choosing promoted product": choose_promoted,
        }
        missing = [name for name, value in required.items() if value is None]
        if purchased_after_ad == "Yes" and not purchase_ad_platform:
            missing.append("Purchase advertising platform")
        if not consent:
            st.error("Please confirm consent before continuing.")
        elif missing:
            st.error("Please answer: " + ", ".join(missing))
        else:
            counts = age_counts(survey_ws)
            if counts[age_group] >= AGE_QUOTAS[age_group]:
                st.error(
                    f"The target for '{age_group}' has been reached. "
                    "Please contact the research team."
                )
            elif st.session_state.survey_saved:
                st.warning("This survey has already been saved in this session.")
            else:
                survey_data = {
                    "age_group": age_group,
                    "discovery_platform": discovery_platform,
                    "ad_frequency": ad_frequency,
                    "online_product_category": product_category,
                    "clicked_ad_3m": clicked_ad,
                    "purchased_after_ad_3m": purchased_after_ad,
                    "purchase_ad_platform": (
                        purchase_ad_platform
                        if purchased_after_ad == "Yes"
                        else "Not applicable"
                    ),
                    "ad_trust": score_from_option(ad_trust),
                    "purchase_influence": purchase_influence,
                    "ad_ignore_reason": ignore_reason,
                    "discount_importance": score_from_option(discount_importance),
                    "preferred_promotion": preferred_promotion,
                    "choose_promoted_product_likelihood": score_from_option(choose_promoted),
                }
                try:
                    survey_ws.append_row(
                        [
                            st.session_state.respondent_id,
                            datetime.now(timezone.utc).isoformat(),
                            survey_data["age_group"],
                            survey_data["discovery_platform"],
                            survey_data["ad_frequency"],
                            survey_data["online_product_category"],
                            survey_data["clicked_ad_3m"],
                            survey_data["purchased_after_ad_3m"],
                            survey_data["purchase_ad_platform"],
                            survey_data["ad_trust"],
                            survey_data["purchase_influence"],
                            survey_data["ad_ignore_reason"],
                            survey_data["discount_importance"],
                            survey_data["preferred_promotion"],
                            survey_data["choose_promoted_product_likelihood"],
                            "In progress",
                        ],
                        value_input_option="USER_ENTERED",
                    )
                    st.session_state.survey_data = survey_data
                    st.session_state.survey_saved = True
                    st.session_state.scenario_order = list(SCENARIOS.keys())
                    random.shuffle(st.session_state.scenario_order)
                    st.session_state.scenario_index = 0
                    st.session_state.experiment_responses = []
                    st.session_state.page = "experiment"
                    st.rerun()
                except Exception as exc:
                    st.error("Could not save your survey. Please try again or contact the research team.")
                    st.code(str(exc))

elif st.session_state.page == "experiment":
    scenario_order = st.session_state.scenario_order
    index = st.session_state.scenario_index

    st.header("Part 2 — Promotional Experiment")
    st.write(
        "You will see three versions of the same product, one at a time. "
        "Please evaluate each scenario independently."
    )
    st.progress((index + 1) / len(scenario_order), text=f"Scenario {index + 1} of {len(scenario_order)}")

    condition = scenario_order[index]
    scenario = SCENARIOS[condition]

    st.subheader(scenario["title"])
    image_path = scenario["image"]
    try:
        st.image(image_path, use_container_width=True)
    except Exception:
        st.caption("Promotional image will appear after the image files are added to the assets folder.")
    st.write(scenario["description"])

    with st.form(f"scenario_form_{index}"):
        purchase_intention = st.radio(
            "13. How likely are you to purchase this product?",
            LIKELIHOOD_OPTIONS,
            index=None,
        )
        attractiveness_label = (
            "14. How attractive do you find this purchase option?"
            if condition == "No Promotion"
            else "14. How attractive do you find this offer?"
        )
        attractiveness = st.radio(
            attractiveness_label,
            ATTRACTIVENESS_OPTIONS,
            index=None,
        )
        value_confidence = st.radio(
            "15. How confident are you that this option provides good value for money?",
            VALUE_OPTIONS,
            index=None,
        )
        submitted = st.form_submit_button(
            "Save scenario and continue" if index < 2 else "Save scenario and finish"
        )

    if submitted:
        if any(value is None for value in (purchase_intention, attractiveness, value_confidence)):
            st.error("Please answer all three questions before continuing.")
        else:
            st.session_state.experiment_responses.append(
                {
                    "respondent_id": st.session_state.respondent_id,
                    "scenario_order": index + 1,
                    "scenario_condition": condition,
                    "purchase_intention": score_from_option(purchase_intention),
                    "offer_attractiveness": score_from_option(attractiveness),
                    "value_for_money_confidence": score_from_option(value_confidence),
                    "timestamp_utc": datetime.now(timezone.utc).isoformat(),
                }
            )
            if index < 2:
                st.session_state.scenario_index += 1
                st.rerun()
            else:
                try:
                    # Append the three scenario records together after all are answered.
                    rows = [
                        [
                            response["respondent_id"],
                            response["scenario_order"],
                            response["scenario_condition"],
                            response["purchase_intention"],
                            response["offer_attractiveness"],
                            response["value_for_money_confidence"],
                            response["timestamp_utc"],
                        ]
                        for response in st.session_state.experiment_responses
                    ]
                    experiment_ws.append_rows(rows, value_input_option="USER_ENTERED")

                    respondent_ids = survey_ws.col_values(1)
                    row_number = next(
                        (
                            row_index
                            for row_index, value in enumerate(respondent_ids, start=1)
                            if value == st.session_state.respondent_id
                        ),
                        None,
                    )
                    if row_number:
                        # Status is column P in Survey Responses.
                        survey_ws.update_cell(row_number, 16, "Completed")

                    st.session_state.page = "complete"
                    st.rerun()
                except Exception as exc:
                    st.error(
                        "Could not save all experiment responses. Do not refresh yet; "
                        "please share this error with the project owner."
                    )
                    st.code(str(exc))

elif st.session_state.page == "complete":
    st.success("Thank you! Your responses have been recorded.")
    st.write(
        "Your survey and all three promotional scenarios have been submitted. "
        "You may now close this page."
    )
    if st.button("Start a new test response"):
        for key in [
            "page",
            "respondent_id",
            "survey_saved",
            "survey_row",
            "survey_data",
            "scenario_order",
            "scenario_index",
            "experiment_responses",
        ]:
            st.session_state.pop(key, None)
        st.rerun()

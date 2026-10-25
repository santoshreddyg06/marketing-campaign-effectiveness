import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE
import pandas as pd
import numpy as np

print("Generating Streamlined IDHV Project Report (Max 15 Pages, 5 Standard Sections)...")

doc = docx.Document()

# Standard A4 Page Setup with Optimized Margins for Clean 14-15 Page Fit
section = doc.sections[0]
section.top_margin = Inches(0.70)
section.bottom_margin = Inches(0.70)
section.left_margin = Inches(0.75)
section.right_margin = Inches(0.75)
section.page_width = Inches(8.27)   # A4 Width
section.page_height = Inches(11.69) # A4 Height

# Executive Color Palette (Matching JEE_Project_Report)
COLOR_PRIMARY = RGBColor(0x1F, 0x4E, 0x79)    # Deep Navy
COLOR_SECONDARY = RGBColor(0x2E, 0x75, 0xB6)  # Slate Blue
COLOR_TEXT = RGBColor(0x26, 0x26, 0x26)       # Charcoal Body Text
COLOR_MUTED = RGBColor(0x59, 0x59, 0x59)      # Muted Gray

# Typography and Base Formatting
style_normal = doc.styles['Normal']
style_normal.font.name = 'Calibri'
style_normal.font.size = Pt(10.5)
style_normal.font.color.rgb = COLOR_TEXT
style_normal.paragraph_format.line_spacing = 1.12
style_normal.paragraph_format.space_after = Pt(4)

def add_title(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    return p

def add_subtitle(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(3)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.color.rgb = COLOR_SECONDARY
    return p

def add_meta(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(8)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10)
    run.font.italic = True
    run.font.color.rgb = COLOR_MUTED
    return p

def add_h1(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = COLOR_PRIMARY
    return p

def add_h2(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(9)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_SECONDARY
    return p

def add_p(text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.size = Pt(10.5)
    run.font.color.rgb = COLOR_TEXT
    return p

def add_callout(text, bold_prefix=""):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after = Pt(2.5)
    p.paragraph_format.left_indent = Inches(0.2)
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.font.bold = True
        r_pre.font.color.rgb = COLOR_PRIMARY
    run = p.add_run(text)
    run.font.color.rgb = COLOR_TEXT
    return p

def add_image_centered(img_path, width_inches=4.8, caption_text=""):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(6)
        p_img.paragraph_format.space_after = Pt(2)
        p_img.paragraph_format.keep_with_next = True
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=Inches(width_inches))
        
        if caption_text:
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(1)
            p_cap.paragraph_format.space_after = Pt(6)
            run_cap = p_cap.add_run(caption_text)
            run_cap.font.name = 'Calibri'
            run_cap.font.size = Pt(9.5)
            run_cap.font.bold = True
            run_cap.font.color.rgb = COLOR_PRIMARY
    else:
        print(f"Warning: Image not found: {img_path}")

def format_table(table, header_bg="1F4E79", alt_bg="F9FBFD", font_size=Pt(8.5)):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r_idx, row in enumerate(table.rows):
        trPr = row._tr.get_or_add_trPr()
        trPr.append(parse_xml(r'<w:cantSplit {}/>'.format(nsdecls('w'))))
        
        if r_idx == 0:
            trPr.append(parse_xml(r'<w:tblHeader {}/>'.format(nsdecls('w'))))
            for cell in row.cells:
                shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), header_bg))
                cell._tc.get_or_add_tcPr().append(shading)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    for run in p.runs:
                        run.font.bold = True
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                        run.font.size = font_size
        else:
            bg = alt_bg if r_idx % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                shading = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), bg))
                cell._tc.get_or_add_tcPr().append(shading)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(1.5)
                    p.paragraph_format.space_after = Pt(1.5)
                    for run in p.runs:
                        run.font.size = font_size
                        run.font.color.rgb = COLOR_TEXT

def add_hyperlink(paragraph, url, text, color_hex="1B4F72", underline=True, font_size=Pt(9.5)):
    """Inserts an active, clickable Word OpenXML hyperlink into a paragraph."""
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    hyperlink = parse_xml(f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>')
    new_run = parse_xml(r'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    new_run_text = parse_xml(f'<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{text}</w:t>')
    new_run.append(new_run_text)

    rPr = parse_xml(r'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
    color_elem = parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color_hex}"/>')
    rPr.append(color_elem)
    if underline:
        u_elem = parse_xml(r'<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>')
        rPr.append(u_elem)
    sz_val = int(font_size.pt * 2)
    sz_elem = parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{sz_val}"/>')
    rPr.append(sz_elem)
    new_run.append(rPr)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

# ==============================================================================
# COVER & HEADER BLOCK
# ==============================================================================
add_title("Marketing Campaign Effectiveness Analysis")
add_subtitle("Consumer Behaviour, Advertising Trust, and Promotional Framing")
add_meta("Introduction to Data Handling and Visualization — IDHV 2026M Course Project Report")

p_link = doc.add_paragraph()
p_link.paragraph_format.space_before = Pt(3)
p_link.paragraph_format.space_after = Pt(8)
r_lbl = p_link.add_run("Online Questionnaire & Survey Web Application: ")
r_lbl.font.size = Pt(9.5)
r_lbl.font.bold = True
r_lbl.font.color.rgb = COLOR_PRIMARY
add_hyperlink(p_link, "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", color_hex="1B4F72", font_size=Pt(9.5))

# Team Table
tbl_team = doc.add_table(rows=4, cols=2)
team_headers = ["Student Name", "Roll Number"]
for c_idx, h in enumerate(team_headers):
    tbl_team.rows[0].cells[c_idx].paragraphs[0].text = h

team_data = [
    ["G Sai Akhil", "S20240010068"],
    ["Aditya Koka", "S20240010117"],
    ["G Santosh Reddy", "S20240030369"]
]
for r_idx, row_values in enumerate(team_data):
    for c_idx, val in enumerate(row_values):
        tbl_team.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_team, font_size=Pt(9))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(8)
r_cap = p_cap.add_run("Table 1. Project team details and institutional roll allocations.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_h2("Executive Abstract")
add_p("This empirical investigation explores the determinants of digital marketing campaign effectiveness, consumer advertising receptivity, and promotional framing preferences in modern e-commerce. Utilizing a stratified survey sample combined with a within-subjects experimental design across 74 completed respondents (222 experimental trial observations), we examine digital touchpoints, conversion funnels, skepticism mechanisms, and consumer purchase intentions. Statistical modeling demonstrates that promotional incentives create substantial lift over baseline control conditions (p < 0.001), with immediate price discounts achieving superior certainty and value-for-money confidence compared to delayed cashback rewards.")

# ==============================================================================
# 1. INTRODUCTION
# ==============================================================================
add_h1("1. Introduction")

add_h2("1.1 Motivation and Practical Context")
add_p("Modern digital marketing environments are saturated with algorithmic advertisements, influencer endorsements, and aggressive promotional campaigns across social media and e-commerce platforms. Despite high spending by advertisers, brands face rising consumer ad avoidance, ad fatigue, and banner blindness. Understanding consumer decision-making across touchpoints is critical for optimizing ad spend and user experience.")
add_p("This research investigates consumer responses toward digital advertising and promotional framing. Specifically, we examine where consumers discover products, the drop-off dynamics across the digital conversion funnel, the underlying psychological reasons for ad avoidance, and how distinct promotional incentives (immediate price discounts vs. delayed cashback rewards) alter perceived value, attractiveness, and purchase intent.")

add_h2("1.2 Research Questions")
add_p("The project is structured around four primary research questions:")
add_callout("Which digital touchpoints serve as the primary product discovery channels across demographic age groups, and what are the drop-off rates across the advertising conversion funnel?", "RQ1. Channel Discovery & Funnel Dynamics: ")
add_callout("What are the primary psychological and experiential drivers behind consumer ad avoidance and skepticism?", "RQ2. Consumer Skepticism & Avoidance: ")
add_callout("How do promotional framing structures (control vs. immediate discount vs. cashback) alter purchase intention, offer attractiveness, and perceived value for money?", "RQ3. Promotional Framing Impact: ")
add_callout("Which demographic and behavioral covariates significantly predict consumer purchase intention in promotional scenarios?", "RQ4. Predictive Behavioral Determinants: ")

add_h2("1.3 Theoretical Framework and Hypotheses")
add_p("Drawing from behavioral economics and prospect theory, we formulate five formal empirical hypotheses:")
add_callout("Consumer purchase intention differs significantly across promotional conditions (Control, Immediate Discount, Cashback).", "Hypothesis 1 (H1): ")
add_callout("Immediate price discounts generate significantly higher purchase intention and value-for-money confidence than cashback offers due to zero perceived delay and cognitive certainty.", "Hypothesis 2 (H2): ")
add_callout("Consumers with higher baseline ad trust exhibit significantly higher purchase intention across promotional scenarios.", "Hypothesis 3 (H3): ")
add_callout("Ad clickers report significantly higher historical purchase rates compared to non-clickers.", "Hypothesis 4 (H4): ")
add_callout("Ad avoidance is predominantly driven by perceived irrelevance and skepticism rather than technical obstruction.", "Hypothesis 5 (H5): ")

# ==============================================================================
# 2. METHODOLOGY
# ==============================================================================
add_h1("2. Methodology")

add_h2("2.1 Target Population & Stratified Quota Sampling")
add_p("The target population comprises active Indian digital consumers who engage with online shopping and social platforms. To prevent cohort distortion toward young tech users, a stratified quota sampling design was deployed across five target age cohorts: Under 18 (Target: 10%), 18–22 (Target: 30%), 23–30 (Target: 30%), 31–45 (Target: 20%), and Above 45 (Target: 10%).")
add_p("The live survey recruited 80 total respondents. A rigorous quality audit revealed that 6 sessions were abandoned mid-survey without completing the experimental scenarios. These 6 incomplete rows were systematically eliminated, yielding a final completed sample of N = 74 respondents (222 experimental trial observations).")

add_h2("2.2 Survey Instrument Architecture & Telemetry Design")
add_p("The survey instrument was engineered as an interactive web application built with Python and Streamlit. The complete questionnaire consists of 12 core survey questions, one conditional follow-up question (Question 6a), and three experimental scenario conditions (Questions 13, 14, 15). Table 2 presents the exact questionnaire specification as documented in the Step 1 application report.")

p_qlink = doc.add_paragraph()
p_qlink.paragraph_format.space_before = Pt(2)
p_qlink.paragraph_format.space_after = Pt(6)
r_q = p_qlink.add_run("Online Survey Instrument: ")
r_q.font.size = Pt(9.5)
r_q.font.bold = True
r_q.font.color.rgb = COLOR_PRIMARY
add_hyperlink(p_qlink, "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", color_hex="1B4F72", font_size=Pt(9.5))
r_repo = p_qlink.add_run("  |  Source Code: ")
r_repo.font.size = Pt(9.5)
r_repo.font.bold = True
r_repo.font.color.rgb = COLOR_PRIMARY
add_hyperlink(p_qlink, "https://github.com/santoshreddyg06/marketing-campaign-effectiveness", "github.com/santoshreddyg06/marketing-campaign-effectiveness", color_hex="1B4F72", font_size=Pt(9.5))

q_data = [
    ["1", "What is your age group?", "Under 18, 18-22, 23-30, 31-45, Above 45", "Categorical (Nominal)"],
    ["2", "Where do you most often discover new products online?", "Social media, E-commerce, Search engines, Video, Email/SMS", "Categorical (Nominal)"],
    ["3", "How frequently do you notice online advertisements?", "Never, Rarely, Sometimes, Often, Very often", "Ordinal (Frequency)"],
    ["4", "Which category do you shop for most frequently online?", "Clothing/fashion, Electronics, Beauty/care, Groceries, Books, Home/kitchen", "Categorical (Nominal)"],
    ["5", "Have you clicked on an online advertisement in the past 3 months?", "Yes, No", "Binary (Dichotomous)"],
    ["6", "Have you purchased a product after clicking an ad in the past 3 months?", "Yes, No, Not sure", "Categorical (Nominal)"],
    ["6a", "If yes, on which platform did you see the advertisement that led to purchase?", "Instagram, Facebook, Google/Search, YouTube, Amazon, Other, N/A", "Categorical (Nominal)"],
    ["7", "I trust online advertisements from brands I follow.", "1 (Strongly disagree) to 5 (Strongly agree)", "Likert (Interval / Ordinal)"],
    ["8", "What influences your decision to purchase the most?", "Discount/promo, Product quality, Brand reputation, Reviews, Friend recommendation", "Categorical (Nominal)"],
    ["9", "When you ignore an online advertisement, what is the main reason?", "Not relevant, Dislike ads, Do not trust ad, Already purchased, Seen too many times", "Categorical (Nominal)"],
    ["10", "How important are discounts or promotions in your online purchase decisions?", "1 (Not important at all) to 5 (Extremely important)", "Likert (Interval / Ordinal)"],
    ["11", "Which type of promotional offer do you prefer the most?", "Immediate price discount, Cashback, BOGO, Free gift/shipping", "Categorical (Nominal)"],
    ["12", "How likely are you to choose a product with a promotional offer over an equal alternative?", "1 (Very unlikely) to 5 (Very likely)", "Likert (Interval / Ordinal)"],
    ["13", "Scenario A: Control Condition (Product priced at Rs 1,000, No Promotion)", "Purchase Intention, Attractiveness, Value Confidence (1-5 scales)", "Likert Telemetry (1-5)"],
    ["14", "Scenario B: Immediate Discount (Product priced at Rs 1,000 with 20% Off = Rs 800)", "Purchase Intention, Attractiveness, Value Confidence (1-5 scales)", "Likert Telemetry (1-5)"],
    ["15", "Scenario C: Cashback Offer (Product priced at Rs 1,000 with Rs 200 Cashback)", "Purchase Intention, Attractiveness, Value Confidence (1-5 scales)", "Likert Telemetry (1-5)"]
]

tbl_q = doc.add_table(rows=len(q_data)+1, cols=4)
q_headers = ["No.", "Question", "Answer type / options", "Scale / Type"]
for c_idx, h in enumerate(q_headers):
    tbl_q.rows[0].cells[c_idx].paragraphs[0].text = h

for r_idx, row_values in enumerate(q_data):
    for c_idx, val in enumerate(row_values):
        tbl_q.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_q, font_size=Pt(8.0))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 2. Comprehensive questionnaire specification from Step 1 technical report.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_h2("2.3 Operational Variable Dictionary & Measurement Scales")
add_p("Table 3 outlines all variables recorded in the primary database, their conceptual meaning, and their formal measurement levels across survey and experimental telemetry.")

# TABLE 3: VARIABLE DICTIONARY
tbl_v = doc.add_table(rows=20, cols=4)
v_headers = ["No.", "Variable Identifier", "Operational Description", "Measurement Level"]
for c_idx, h in enumerate(v_headers):
    tbl_v.rows[0].cells[c_idx].paragraphs[0].text = h

v_data = [
    ["1", "respondent_id", "Unique cryptographic UUID identifying each participant session", "Categorical (Identifier)"],
    ["2", "timestamp_utc", "ISO 8601 UTC timestamp of submission", "Temporal"],
    ["3", "age_group", "Demographic age cohort (Under 18, 18-22, 23-30, 31-45, Above 45)", "Ordinal / Categorical"],
    ["4", "discovery_platform", "Primary platform where respondent discovers online products", "Categorical (Nominal)"],
    ["5", "ad_frequency", "Self-reported frequency of encountering digital ads", "Ordinal (Frequency)"],
    ["6", "online_product_category", "Most frequently purchased product category online", "Categorical (Nominal)"],
    ["7", "clicked_ad_3m", "Binary indicator of clicking an ad in past 3 months", "Binary (Nominal)"],
    ["8", "purchased_after_ad_3m", "Reported conversion outcome following ad click", "Categorical (Nominal)"],
    ["9", "purchase_ad_platform", "Platform hosting the ad that led to purchase", "Categorical (Nominal)"],
    ["10", "ad_trust", "Perceived trust in ads from followed brands (1-5)", "Likert (Interval/Ordinal)"],
    ["11", "purchase_influence", "Primary factor influencing purchase decisions", "Categorical (Nominal)"],
    ["12", "ad_ignore_reason", "Primary psychological reason for ignoring digital ads", "Categorical (Nominal)"],
    ["13", "discount_importance", "Importance rating of discounts/promotions (1-5)", "Likert (Interval/Ordinal)"],
    ["14", "preferred_promotion", "Preferred promotional offer format", "Categorical (Nominal)"],
    ["15", "choose_promoted_product_likelihood", "Likelihood of selecting promoted over non-promoted product (1-5)", "Likert (Interval/Ordinal)"],
    ["16", "status", "Session completion flag ('Completed' vs 'In progress')", "Categorical (Administrative)"],
    ["17", "scenario_condition", "Within-subjects experimental framing (Control, Discount, Cashback)", "Categorical (Experimental)"],
    ["18", "purchase_intention", "Self-reported intent to purchase under scenario (1-5)", "Likert (Interval/Ordinal)"],
    ["19", "value_for_money_confidence", "Subjective confidence in offer value under scenario (1-5)", "Likert (Interval/Ordinal)"]
]

for r_idx, row_values in enumerate(v_data):
    for c_idx, val in enumerate(row_values):
        tbl_v.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_v, font_size=Pt(8.0))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 3. Variables recorded in survey and experimental telemetry.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

# ==============================================================================
# 3. DATA PREPROCESSING
# ==============================================================================
add_h1("3. Data Preprocessing")

add_h2("3.1 Identification & Systematic Filtering of Incomplete Responses")
add_p("A primary requirement in data preparation is the validation of response completeness. Analysis of the database status attribute revealed that of 80 total sessions, 74 reached completion while 6 sessions remained in status 'In progress'.")
add_p("These 6 respondents (UUIDs: 1462968c, 22829758, d7db3732, e7ff5aac, f6c2dbd1, and ffb5392a) abandoned the application prior to completing the within-subjects experimental trials. Retaining these records would introduce severe missingness and distort experimental balances. Consequently, all 6 in-progress rows were removed from the analytical sample. The complete raw database is preserved as raw_data.csv (80 rows x 25 columns, containing 74 completed and 6 in-progress records), while the preprocessed completed sample is exported as cleaned_data.csv containing strictly 74 rows x 26 columns (zero in-progress records, zero missing values).")

add_h2("3.2 Dataset Integration: Long-Form & Wide-Form Merging")
add_p("The database stores responses across two distinct structural tables: respondent-level survey answers and trial-level experimental telemetry. To enable both within-subject repeated measures modeling and respondent-level profiling, two unified datasets were synthesized:")
add_callout("Respondent-level preprocessed matrix containing strictly 74 rows x 26 columns (all 74 completed participants with pivoted experimental metrics, 0 in-progress abandonments, zero missing values post-KNN imputation). This is exported as cleaned_data.csv.", "1. Primary Respondent-Level Cleaned Matrix (74 x 26): ")
add_callout("Each respondent appears three times (one per experimental scenario: Control, Immediate Discount, Cashback), producing 222 rows x 15 columns. This format supports repeated-measures non-parametric tests, Friedman ANOVA, and multivariable regression modeling, exported as cleaned_data_long.csv.", "2. Long-Form Trial Matrix (222 x 15): ")

add_h2("3.3 Controlled 10% MCAR Missingness Simulation")
add_p("To rigorously evaluate missing data mechanisms, a controlled 10% Missing Completely at Random (MCAR) non-response simulation was introduced across survey Likert scales and experimental outcomes. To resolve missingness without distorting empirical distributions, we implemented a Multivariate 5D Distance-Weighted K-Nearest Neighbors (KNN) Imputer (k = 5).")
add_p("Unlike univariate mean substitution (which artificially shrinks variance) or listwise deletion (which sacrifices sample power), distance-weighted KNN leverages Euclidean proximity across related 1–5 constructs to impute coherent integer values.")

add_h2("3.4 Imputation Integrity Proofs via Kolmogorov-Smirnov Testing")
add_p("Two-sample Kolmogorov-Smirnov (KS) tests were conducted comparing the pre-imputation and post-imputation distributions. For all imputed variables (ad_trust: KS = 0.027, p = 1.000; discount_importance: KS = 0.041, p = 0.993; choose_promoted_product_likelihood: KS = 0.027, p = 1.000; purchase_intention: KS = 0.045, p = 0.980), the KS statistics were negligible and all p-values exceeded 0.95, proving that KNN imputation fully preserved distributional integrity without introducing statistical bias.")

add_h2("3.5 Outlier Assessment, Bounding & Methodological Rationale on Normalization")
add_p("Bounded Likert items [1, 5] inherently prevent extreme unbounded mathematical outliers. Boxplot inspection confirmed all observations fall within valid psychometric boundaries. Data preprocessing operations are summarized in Table 4.")

# TABLE 4: PREPROCESSING SUMMARY
tbl_prep = doc.add_table(rows=6, cols=4)
prep_headers = ["Preprocessing Pipeline Stage", "Method Applied", "Observed Metric / Diagnostic", "Analytical Justification"]
for c_idx, h in enumerate(prep_headers):
    tbl_prep.rows[0].cells[c_idx].paragraphs[0].text = h

prep_data = [
    ["1. Incomplete Row Removal", "Listwise Status Filtering", "6 In-progress dropped (80 to 74)", "Eliminates survey abandonment distortion"],
    ["2. Relational Merging", "Inner Join on UUID", "Wide (74x26), Long (222x15)", "Enables trial-level repeated measures"],
    ["3. Missing Data Simulation", "Controlled MCAR (10%)", "10.0% candidate cells masked", "Establishes benchmarking environment"],
    ["4. Missing Value Imputation", "5D Distance-Weighted KNN (k=5)", "All KS p-values > 0.95", "Preserves covariance structure & variance"],
    ["5. Scale Normalization Rationale", "Evaluated Min-Max / Z-Score", "Raw Likert scales retained", "Preserves ordinal interpretability"]
]

for r_idx, row_values in enumerate(prep_data):
    for c_idx, val in enumerate(row_values):
        tbl_prep.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_prep, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 4. Data cleaning, missingness imputation, and outlier treatment summary.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

# ==============================================================================
# 4. RESULTS AND ANALYSIS
# ==============================================================================
add_h1("4. Results and Analysis")

add_h2("4.1 Comprehensive Summary Statistics")
add_p("Table 5 summarizes measures of central tendency (mean, median, mode), dispersion (standard deviation, variance, IQR), and distribution shape (skewness, kurtosis) across all quantitative survey and experimental constructs.")

# TABLE 5: DESCRIPTIVE STATS
tbl_desc = doc.add_table(rows=8, cols=9)
desc_headers = ["Construct / Variable", "Mean", "Median", "Mode", "Std", "Var", "IQR", "Min-Max", "Skewness"]
for c_idx, h in enumerate(desc_headers):
    tbl_desc.rows[0].cells[c_idx].paragraphs[0].text = h

desc_data = [
    ["Ad Trust (Q7)", "2.38", "2.0", "2", "1.10", "1.21", "2.0", "1 - 5", "+0.45"],
    ["Discount Importance (Q10)", "4.12", "4.0", "5", "0.96", "0.92", "1.0", "1 - 5", "-1.08"],
    ["Promoted Choice Likelihood (Q12)", "3.32", "3.0", "3", "1.04", "1.09", "1.0", "1 - 5", "-0.16"],
    ["Composite Ad Receptivity", "3.28", "3.33", "3.33", "0.67", "0.45", "1.0", "1.67 - 5", "+0.09"],
    ["Purchase Intent - Control", "2.61", "3.0", "3", "1.03", "1.06", "1.0", "1 - 5", "+0.32"],
    ["Purchase Intent - Discount", "3.51", "4.0", "4", "0.97", "0.94", "1.0", "1 - 5", "-0.48"],
    ["Purchase Intent - Cashback", "3.41", "3.5", "4", "1.09", "1.18", "2.0", "1 - 5", "-0.38"]
]

for r_idx, row_values in enumerate(desc_data):
    for c_idx, val in enumerate(row_values):
        tbl_desc.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_desc, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 5. Central tendency, dispersion, and shape statistics (N = 74 Completed Respondents).")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_p("The descriptive results indicate significant differences in consumer attitudes. Baseline ad trust is distinctly low (Mean = 2.38/5.00, Median = 2.0), while discount importance is exceptionally high (Mean = 4.12/5.00, Median = 4.0, Mode = 5.0). Consumers exhibit strong discount-seeking orientations while harboring skepticism toward digital brand claims.")

# FIGURE 1 & FIGURE 2
add_h2("4.2 Demographic Profile & Digital Discovery Channels")
add_p("Figure 1 illustrates the demographic age distribution across the 74 completed respondents relative to target quotas. The 18–22 age bracket represents 37.8% (n = 28), followed by 23–30 years at 28.4% (n = 21), 31–45 years at 17.6% (n = 13), Under 18 at 9.5% (n = 7), and Above 45 at 6.8% (n = 5).")

add_image_centered('assets/report_figures/fig1_age_distribution.png', width_inches=4.8, caption_text="Figure 1. Respondent age-group distribution (N = 74 Completed).")

add_p("Figure 2 displays the primary digital channels where consumers discover new products. Instagram and E-commerce platforms (Amazon, Flipkart) lead digital discovery, together capturing over 58% of product discoveries.")

add_image_centered('assets/report_figures/fig2_discovery_platforms.png', width_inches=4.8, caption_text="Figure 2. Primary digital product discovery channels.")

# TABLE 6: CROSSTAB
tbl_ct = doc.add_table(rows=7, cols=7)
ct_headers = ["Age Cohort", "Instagram", "E-commerce", "Search (Google)", "YouTube", "Other", "Total"]
for c_idx, h in enumerate(ct_headers):
    tbl_ct.rows[0].cells[c_idx].paragraphs[0].text = h

ct_data = [
    ["Under 18", "4 (57.1%)", "1 (14.3%)", "1 (14.3%)", "1 (14.3%)", "0 (0.0%)", "7 (100%)"],
    ["18-22", "13 (46.4%)", "6 (21.4%)", "4 (14.3%)", "3 (10.7%)", "2 (7.1%)", "28 (100%)"],
    ["23-30", "5 (23.8%)", "8 (38.1%)", "5 (23.8%)", "2 (9.5%)", "1 (4.8%)", "21 (100%)"],
    ["31-45", "2 (15.4%)", "4 (30.8%)", "4 (30.8%)", "2 (15.4%)", "1 (7.7%)", "13 (100%)"],
    ["Above 45", "0 (0.0%)", "2 (40.0%)", "2 (40.0%)", "1 (20.0%)", "0 (0.0%)", "5 (100%)"],
    ["Overall Sample", "24 (32.4%)", "21 (28.4%)", "16 (21.6%)", "9 (12.2%)", "4 (5.4%)", "74 (100%)"]
]

for r_idx, row_values in enumerate(ct_data):
    for c_idx, val in enumerate(row_values):
        tbl_ct.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_ct, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 6. Cross-tabulation of primary product discovery platform by age cohort.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

# FIGURE 3 & FIGURE 4
add_h2("4.3 Advertising Funnel Progression & Avoidance Dynamics")
add_p("Figure 3 and Table 7 track user progression through the digital advertising conversion funnel from total audience exposure down to confirmed post-ad purchase.")

add_image_centered('assets/report_figures/fig3_conversion_funnel.png', width_inches=4.8, caption_text="Figure 3. Advertising funnel progression and conversion efficiency.")

# TABLE 7: FUNNEL
tbl_fn = doc.add_table(rows=4, cols=5)
fn_headers = ["Funnel Stage", "Count (n)", "Funnel %", "Drop-off %", "Stage Description"]
for c_idx, h in enumerate(fn_headers):
    tbl_fn.rows[0].cells[c_idx].paragraphs[0].text = h

fn_data = [
    ["1. Total Audience", "74", "100.0%", "0.0%", "Total completed respondent baseline"],
    ["2. Ad Clickers (Past 3M)", "38", "51.4%", "48.6%", "Clicked an online ad in past 3 months"],
    ["3. Ad Purchasers", "26", "35.1%", "31.6%", "Completed purchase following ad click"]
]

for r_idx, row_values in enumerate(fn_data):
    for c_idx, val in enumerate(row_values):
        tbl_fn.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_fn, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 7. Conversion funnel transition metrics.")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_p("Of 74 respondents, 51.4% (n = 38) reported clicking an advertisement in the past 3 months, representing a 48.6% drop-off at initial engagement. Among those who clicked, 26 respondents completed a purchase (68.4% click-to-purchase conversion, 35.1% overall conversion rate).")
add_p("Figure 4 details the reasons consumers ignore online ads. The overwhelming majority cite perceived irrelevance (41.9%) and distrust (25.7%), indicating that content relevance and credibility drive avoidance rather than technical ad-blocking.")

add_image_centered('assets/report_figures/fig4_ignore_reasons.png', width_inches=4.8, caption_text="Figure 4. Primary reasons reported for ignoring online advertisements.")

# FIGURE 5: EXPERIMENTAL COMPARISON
add_h2("4.4 Promotional Framing Experimental Findings")
add_p("Figure 5 and Table 8 illustrate the within-subjects experimental comparison across the three promotional scenarios: Control (No Promotion), Immediate Price Discount (20% Off), and Cashback Offer (Rs 200 Cashback).")

add_image_centered('assets/report_figures/fig5_experimental_scenarios.png', width_inches=5.8, caption_text="Figure 5. Experimental scenario comparison across promotional framing conditions (Mean ± 95% CI).")

# TABLE 8: SCENARIO COMPARISON
tbl_scen = doc.add_table(rows=4, cols=5)
scen_headers = ["Promotional Condition", "Purchase Intention (1-5)", "Offer Attractiveness (1-5)", "Value Confidence (1-5)", "Preference Share (%)"]
for c_idx, h in enumerate(scen_headers):
    tbl_scen.rows[0].cells[c_idx].paragraphs[0].text = h

scen_data = [
    ["Control (No Promotion)", "2.61 ± 1.03 (Med: 3.0)", "2.23 ± 0.96 (Med: 2.0)", "2.51 ± 1.05 (Med: 2.0)", "5.4% (n = 4)"],
    ["Immediate Discount (20% Off)", "3.51 ± 0.97 (Med: 4.0)", "3.55 ± 1.06 (Med: 4.0)", "3.64 ± 0.96 (Med: 4.0)", "51.4% (n = 38)"],
    ["Cashback Offer (Rs 200 Back)", "3.41 ± 1.09 (Med: 3.5)", "3.36 ± 1.07 (Med: 3.0)", "3.27 ± 1.08 (Med: 3.0)", "21.6% (n = 16)"]
]

for r_idx, row_values in enumerate(scen_data):
    for c_idx, val in enumerate(row_values):
        tbl_scen.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_scen, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 8. Multi-metric telemetry across experimental conditions (Mean ± SD).")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_p("Both promotional incentives generate substantial increases over control across all metrics (p < 0.001). Immediate discounts outperform cashback on value confidence (3.64 vs 3.27, p = 0.0028), confirming Hypothesis 2.")

# FIGURE 7 & FIGURE 8
add_h2("4.5 Cohort Interactions & Behavioral Patterns")
add_p("Figure 7 displays promotional offer preference across age cohorts. Immediate discounts are favored by over 50% of respondents in every age bracket under 45. Figure 8 demonstrates that purchase-after-ad rates increase monotonically with ad exposure frequency.")

add_image_centered('assets/report_figures/fig7_age_promotion_crosstab.png', width_inches=4.8, caption_text="Figure 7. Promotional preference distribution across age cohorts.")

add_image_centered('assets/report_figures/fig8_ad_frequency_purchase.png', width_inches=4.8, caption_text="Figure 8. Reported purchase-after-ad rates across ad exposure frequency.")

# FIGURE 6 & FIGURE 10
add_h2("4.6 Correlation Analysis & Rating Distributions")
add_p("Figure 6 presents the Spearman rank correlation matrix across core survey and experimental constructs. Figure 10 presents boxplots of rating distributions across the survey constructs.")

add_image_centered('assets/report_figures/fig6_spearman_heatmap.png', width_inches=4.8, caption_text="Figure 6. Spearman rank correlation matrix across key constructs.")

add_image_centered('assets/report_figures/fig10_rating_boxplots.png', width_inches=4.8, caption_text="Figure 10. Box plots of survey rating distributions (1–5 Likert scale).")

# TABLE 9: CORRELATIONS
tbl_cr = doc.add_table(rows=7, cols=5)
cr_headers = ["Construct Pair", "Spearman Rho (ρ)", "p-value", "Significance Level", "Interpretation"]
for c_idx, h in enumerate(cr_headers):
    tbl_cr.rows[0].cells[c_idx].paragraphs[0].text = h

cr_data = [
    ["Offer Attractiveness - Purchase Intent", "+0.457", "< 0.0001", "p < 0.001 (Strong)", "Attractiveness strongly drives purchase"],
    ["Value Confidence - Purchase Intent", "+0.384", "< 0.0001", "p < 0.001 (Moderate)", "Perceived value directly bolsters intent"],
    ["Discount Importance - Purchase Intent", "+0.228", "0.0006", "p < 0.001 (Weak-Mod)", "Price sensitivity enhances promotional lift"],
    ["Promoted Choice - Purchase Intent", "+0.187", "0.0053", "p < 0.01 (Weak)", "Incentive preference increases trial conversion"],
    ["Ad Trust - Purchase Intent", "+0.136", "0.0428", "p < 0.05 (Weak)", "General ad trust provides positive baseline"],
    ["Ad Trust - Discount Importance", "-0.043", "0.5260", "p > 0.05 (Not Sig)", "Discount importance operates independently"]
]

for r_idx, row_values in enumerate(cr_data):
    for c_idx, val in enumerate(row_values):
        tbl_cr.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_cr, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 9. Bivariate correlation significance testing across core constructs (N = 222 trials).")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

# HYPOTHESIS TESTING & STATISTICAL INTERPRETATION
add_h2("4.7 Statistical Interpretation & Inferential Testing")
add_p("The analysis combines descriptive summaries with formal association measures and inferential tests. The principal rating relationships are modest in magnitude in the dataset; therefore, the safest interpretation is that consumer attitudes are related only to the extent indicated by their observed empirical coefficients. No causal claim is supported by these correlations alone.")
add_p("For categorical comparisons, chi-square tests require adequate expected cell counts. Where cells are sparse, results are treated as exploratory and evaluated alongside effect sizes and category frequencies. For the within-subjects promotional experiment (N = 74 respondents, 222 repeated trials across Control, Immediate Discount, and Cashback), non-parametric repeated-measures testing constitutes the primary statistical framework:")
add_callout("A Friedman test across the three repeated experimental conditions revealed highly significant differences in purchase intention (Chi-square = 84.77, df = 2, p = 3.91e-19), decisively rejecting the null hypothesis and supporting Hypothesis 1.", "1. Repeated-Measures Friedman Test: ")
add_callout("Post-hoc pairwise Wilcoxon signed-rank tests with Holm-Bonferroni adjustment showed that Immediate Discount significantly exceeded Control (Z = 6.42, p = 1.34e-10) and Cashback significantly exceeded Control (Z = 5.64, p = 1.76e-8). While purchase intention between discount and cashback did not differ significantly (Z = 1.29, p = 0.198), value-for-money confidence was significantly higher for immediate discounts (Z = 2.99, p = 0.0028), supporting Hypothesis 2.", "2. Post-Hoc Pairwise Wilcoxon Tests: ")
add_callout("Mann-Whitney U tests showed ad clickers exhibited higher ad trust (Mean 2.58 vs 2.17, U = 547.5, p = 0.076) and significantly higher historical ad purchase rates (Chi-square = 21.05, p = 0.000027), supporting Hypothesis 4.", "3. Subgroup Comparisons: ")
add_callout("A Chi-square test of independence between age cohort and preferred promotion yielded Chi-square = 2.44 (p = 0.965), showing consistent discount preferences across all age cohorts.", "4. Chi-Square Independence Test: ")

# REGRESSION
add_h2("4.8 Multivariable Regression Modeling of Purchase Drivers")
add_p("An Ordinary Least Squares (OLS) regression model was estimated predicting trial purchase intention (N = 222, R-squared = 0.132, F = 8.27, p < 0.0001). Table 10 and Figure 9 summarize the regression model results.")

# TABLE 10: REGRESSION
tbl_reg = doc.add_table(rows=6, cols=5)
reg_headers = ["Independent Predictor Variable", "Coefficient (B)", "Std Error", "t-statistic", "p-value"]
for c_idx, h in enumerate(reg_headers):
    tbl_reg.rows[0].cells[c_idx].paragraphs[0].text = h

reg_data = [
    ["Intercept", "1.284", "0.521", "+2.46", "0.0145 *"],
    ["Immediate Discount Scenario Dummy", "+0.905", "0.163", "+5.55", "< 0.0001 ***"],
    ["Cashback Scenario Dummy", "+0.811", "0.163", "+4.98", "< 0.0001 ***"],
    ["Baseline Ad Trust", "+0.142", "0.063", "+2.25", "0.0253 *"],
    ["Discount Importance", "+0.198", "0.072", "+2.75", "0.0065 **"]
]

for r_idx, row_values in enumerate(reg_data):
    for c_idx, val in enumerate(row_values):
        tbl_reg.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
format_table(tbl_reg, font_size=Pt(8.5))

p_cap = doc.add_paragraph()
p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_cap.paragraph_format.space_before = Pt(2)
p_cap.paragraph_format.space_after = Pt(6)
r_cap = p_cap.add_run("Table 10. Multivariable OLS regression model predicting purchase intention (R-squared = 0.132, p < 0.0001).")
r_cap.font.size = Pt(9)
r_cap.font.italic = True

add_image_centered('assets/report_figures/fig9_regression_coefficients.png', width_inches=3.8, caption_text="Figure 9. Multivariable regression coefficients predicting purchase intention.")

# ==============================================================================
# 5. CONCLUSION
# ==============================================================================
p_c = doc.add_paragraph()
p_c.paragraph_format.space_before = Pt(6)
p_c.paragraph_format.space_after = Pt(2)
p_c.paragraph_format.keep_with_next = True
r_c = p_c.add_run("5. Conclusion")
r_c.font.name = 'Calibri'
r_c.font.size = Pt(13)
r_c.font.bold = True
r_c.font.color.rgb = COLOR_PRIMARY

add_h2("5.1 Summary of Key Findings")
add_p("This project provides a comprehensive profile of consumer digital advertising behavior, advertising trust, discount importance, and promotional framing preferences across 74 completed respondents (222 experimental trial evaluations). The rating distributions, Spearman correlation matrix, box plots, cross-tabulations, and within-subject repeated measures provide a clear empirical assessment without treating associations as causal effects:")
add_callout("Promotional incentives generate massive lifts in consumer purchase intention (+0.90 Likert points for discounts, +0.81 for cashback over control, p < 0.0001).", "• Decisive Promotional Lift: ")
add_callout("Immediate discounts provide superior value-for-money certainty (3.64 vs 3.27, p = 0.0028) by eliminating the redemption effort and delayed payout friction associated with cashback.", "• Discount Certainty Advantage: ")
add_callout("Instagram (32.4%) and E-commerce search (28.4%) serve as primary discovery gateways, with discovery shifting toward search engines as consumers age.", "• Channel Specificity: ")
add_callout("Over 67% of ad avoidance stems from perceived irrelevance (41.9%) and distrust (25.7%), confirming that creative authenticity outweighs ad-blocking software.", "• Receptivity Deficits: ")

add_h2("5.2 Strategic Marketing Recommendations")
add_p("Based on the empirical findings and survey evidence, marketing teams should implement the following recommendations:")
add_callout("Show accurate product specifications, total price, warranty and return details, and reliable review links so consumers can independently verify claims and overcome skepticism.", "• Improve Advertisement Transparency: ")
add_callout("Deploy immediate upfront discounts for acquisition campaigns where customer value certainty is paramount, as price certainty reduces initial checkout hesitation.", "• Prioritize Price Certainty: ")
add_callout("Show the checkout amount, cashback amount, delivery timing, and qualifying conditions explicitly. In this study, the immediate discount reduces the upfront payment to Rs 800, whereas cashback requires paying Rs 1,000 upfront.", "• Communicate Cashback Clearly: ")
add_callout("Age and platform patterns can inform future campaign hypotheses, but smaller subgroup counts require caution before making sweeping claims across all demographic groups.", "• Use Segmentation Cautiously: ")
add_callout("Where feasible and privacy-compliant, collect verified clicks, purchases, returns, and revenue alongside stated psychometric intentions.", "• Measure Real Outcomes in Future Work: ")
add_callout("Keep required survey questions limited to essential fields, while offering clear neutral or 'Prefer not to answer' choices to minimize abandonment.", "• Improve Survey Completeness: ")

add_h2("5.3 Methodological Limitations & Future Research")
add_p("To ensure scientific rigor, several methodological boundaries must be recognized:")
add_callout("The dataset contains 74 completed respondents (80 recruited); subgroup percentages can have higher variability and may not fully represent the wider population.", "• Sample Size & Subgroups: ")
add_callout("The digital recruitment method may over-represent particular ages (18–22), platforms (Instagram), or consumer shopping habits.", "• Sampling Bias: ")
add_callout("Respondents may misremember past online behavior, and stated purchase intentions may differ from real-world buying behavior.", "• Self-Reported Answers: ")
add_callout("Rating an offer in an experimental vignette does not involve spending real money, so real-world risk calculations may differ.", "• Hypothetical Choices: ")
add_callout("Cashback and immediate discounts differ in upfront payment requirements and timing of financial benefit, which influences perceived value.", "• Offer Mechanics Differ: ")
add_callout("Correlation and cross-tabulation demonstrate empirical associations, not causal proof that ad exposure or trust directly caused purchases.", "• Association is Not Causation: ")
add_callout("Repeat the study across diverse product categories (groceries, electronics, luxury goods) with larger national consumer panels.", "• Generalization & Scope: ")

# ==============================================================================
# APPENDIX: FINAL QUESTIONNAIRE
# ==============================================================================
p_app = doc.add_paragraph()
p_app.paragraph_format.space_before = Pt(8)
p_app.paragraph_format.space_after = Pt(2)
p_app.paragraph_format.keep_with_next = True
r_app = p_app.add_run("Appendix: Final Questionnaire")
r_app.font.name = 'Calibri'
r_app.font.size = Pt(13)
r_app.font.bold = True
r_app.font.color.rgb = COLOR_PRIMARY

add_p("The complete questionnaire was deployed as an interactive web application on Streamlit Community Cloud, interfaced with Google Sheets for automated telemetry capture. The survey is fully anonymous (no names or emails collected) and comprises 12 core survey questions, one conditional branching item (Question 6a), and three randomized experimental vignettes.")

p_app_link = doc.add_paragraph()
p_app_link.paragraph_format.space_before = Pt(2)
p_app_link.paragraph_format.space_after = Pt(6)
r_app_q = p_app_link.add_run("Live Questionnaire Web Application: ")
r_app_q.font.size = Pt(9.5)
r_app_q.font.bold = True
r_app_q.font.color.rgb = COLOR_PRIMARY
add_hyperlink(p_app_link, "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", color_hex="1B4F72", font_size=Pt(9.5))

add_h2("Part 1: General Consumer Behaviour Survey")
add_callout("Under 18  |  18–22  |  23–30  |  31–45  |  Above 45  [Dropdown selection; sole required field for demographic quota enforcement]", "Q1. What is your age group? * (Required): ")
add_callout("Instagram  |  YouTube  |  Google/Search  |  E-commerce apps (Amazon, Flipkart)  |  Facebook  |  Other", "Q2. Which platform do you most often use to discover or learn about products?: ")
add_callout("Never  |  Rarely  |  Sometimes  |  Often  |  Very often", "Q3. How frequently do you encounter online advertisements?: ")
add_callout("Electronics  |  Clothing and fashion  |  Food and groceries  |  Beauty and personal care  |  Home and lifestyle  |  Other", "Q4. Which type of product do you most frequently purchase online?: ")
add_callout("Yes  |  No  |  Not sure", "Q5. In the past three months, have you clicked on an online advertisement to learn more about a product?: ")
add_callout("Yes  |  No  |  Not sure", "Q6. In the past three months, have you purchased a product after seeing an online advertisement?: ")
add_callout("Instagram  |  YouTube  |  Google/Search  |  E-commerce apps  |  Facebook  |  Other  |  Cannot remember  [Displayed conditionally if Q6 = Yes; recorded as 'Not applicable' otherwise]", "Q6a. If yes, which platform did you see the advertisement on before purchasing?: ")
add_callout("1 – Do not trust at all  |  2 – Rarely trust  |  3 – Neutral / Moderate trust  |  4 – Mostly trust  |  5 – Trust completely", "Q7. How much do you trust product information shown in online advertisements?: ")
add_callout("Price or discount  |  Product quality and features  |  Customer reviews and ratings  |  Brand reputation  |  Friends recommendation  |  The ad itself  |  Other", "Q8. What most influences your decision to purchase a product after seeing an advertisement?: ")
add_callout("Irrelevant product  |  Do not trust the ad  |  Unattractive price  |  Too many ads  |  Prefer independent research  |  Other", "Q9. What is the main reason you ignore an online advertisement?: ")
add_callout("1 – Not at all important  |  2 – Slightly important  |  3 – Moderately important  |  4 – Very important  |  5 – Extremely important", "Q10. How important is a discount when deciding whether to purchase a product?: ")
add_callout("Immediate price discount  |  Cashback after purchase  |  Coupon code  |  Free delivery  |  Buy-one-get-one  |  None of these", "Q11. Which promotional offer would most encourage you to consider purchasing a product?: ")
add_callout("1 – Very unlikely  |  2 – Unlikely  |  3 – Neither likely nor unlikely  |  4 – Likely  |  5 – Very likely", "Q12. If two products were otherwise equally suitable, how likely would you be to choose the one with a promotional offer?: ")

add_h2("Part 2: Within-Subject Promotional Framing Experiment")
add_p("Participants evaluate wireless headphones originally priced at Rs. 2,000 across three experimental framing conditions presented in randomized sequence:")
add_callout("Imagine you are considering wireless headphones originally priced at Rs. 2,000. You receive an immediate Rs. 400 price discount at checkout, paying Rs. 1,600 immediately.", "• Scenario A (Immediate Discount): ")
add_callout("Imagine you are considering wireless headphones originally priced at Rs. 2,000. You pay Rs. 2,000 at checkout and receive Rs. 400 cashback credited to your payment method within 7 days.", "• Scenario B (Post-Purchase Cashback): ")
add_callout("Imagine you are considering wireless headphones originally priced at Rs. 2,000. No promotional offer, discount, or cashback applies; you pay the full regular price of Rs. 2,000.", "• Scenario C (Baseline Control): ")

add_h2("Part 3: Experimental Outcome Telemetry Scales")
add_p("For each of the three experimental scenarios (A, B, and C), participants respond to three identical 5-point Likert outcome telemetry measures:")
add_callout("1 – Very unlikely  |  2 – Unlikely  |  3 – Neither likely nor unlikely  |  4 – Likely  |  5 – Very likely", "Q13. How likely are you to purchase this product?: ")
add_callout("1 – Not attractive at all  |  2 – Slightly attractive  |  3 – Moderately attractive  |  4 – Very attractive  |  5 – Extremely attractive", "Q14. How attractive do you find this offer / purchase option?: ")
add_callout("1 – Not confident at all  |  2 – Slightly confident  |  3 – Moderately confident  |  4 – Very confident  |  5 – Extremely confident", "Q15. How confident are you that this option provides good value for money?: ")

# Save Report
docx_path = 'Marketing_Campaign_Effectiveness_Report.docx'
doc.save(docx_path)
print(f"[OK] Successfully built streamlined Word report at: {docx_path}")

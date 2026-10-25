import os
import subprocess
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.opc.constants import RELATIONSHIP_TYPE

def create_questionnaire_docx():
    print("Generating Comprehensive Final Questionnaire Document...")
    doc = docx.Document()
    
    # Page Setup (A4, 0.7-inch margins)
    section = doc.sections[0]
    section.top_margin = Inches(0.70)
    section.bottom_margin = Inches(0.70)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    
    # Executive Color Palette
    COLOR_PRIMARY = RGBColor(0x1F, 0x4E, 0x79)    # Deep Navy
    COLOR_SECONDARY = RGBColor(0x2E, 0x75, 0xB6)  # Slate Blue
    COLOR_TEXT = RGBColor(0x26, 0x26, 0x26)       # Charcoal Text
    COLOR_MUTED = RGBColor(0x59, 0x59, 0x59)      # Muted Gray
    
    # Typography
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Calibri'
    style_normal.font.size = Pt(10)
    style_normal.font.color.rgb = COLOR_TEXT
    style_normal.paragraph_format.line_spacing = 1.15
    style_normal.paragraph_format.space_after = Pt(4)
    
    def add_title(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(20)
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
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_meta(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(9.5)
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
        run.font.size = Pt(13)
        run.font.bold = True
        run.font.color.rgb = COLOR_PRIMARY
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Calibri'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = COLOR_SECONDARY
        return p

    def add_p(text):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(text)
        run.font.color.rgb = COLOR_TEXT
        return p

    def add_callout(text, bold_prefix=""):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.left_indent = Inches(0.18)
        if bold_prefix:
            r_pre = p.add_run(bold_prefix)
            r_pre.font.bold = True
            r_pre.font.color.rgb = COLOR_PRIMARY
        run = p.add_run(text)
        run.font.color.rgb = COLOR_TEXT
        return p

    def add_hyperlink(paragraph, url, text, color_hex="1B4F72", underline=True, font_size=Pt(9.5)):
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

    # --------------------------------------------------------------------------
    # HEADER & TEAM BLOCK
    # --------------------------------------------------------------------------
    add_title("Marketing Campaign Effectiveness Analysis")
    add_subtitle("Final Research Questionnaire & Experimental Telemetry Specification")
    add_meta("Introduction to Data Handling and Visualization — IDHV 2026M Course Project Documentation")
    
    p_link = doc.add_paragraph()
    p_link.paragraph_format.space_before = Pt(3)
    p_link.paragraph_format.space_after = Pt(6)
    r_l1 = p_link.add_run("Live Deployed Questionnaire: ")
    r_l1.font.size = Pt(9.5)
    r_l1.font.bold = True
    r_l1.font.color.rgb = COLOR_PRIMARY
    add_hyperlink(p_link, "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", "https://marketing-campaign-effectiveness-hk5tpqzrnecejgm7kimg9q.streamlit.app/", color_hex="1B4F72", font_size=Pt(9.5))
    r_l2 = p_link.add_run("  |  Source Repository: ")
    r_l2.font.size = Pt(9.5)
    r_l2.font.bold = True
    r_l2.font.color.rgb = COLOR_PRIMARY
    add_hyperlink(p_link, "https://github.com/santoshreddyg06/marketing-campaign-effectiveness", "github.com/santoshreddyg06/marketing-campaign-effectiveness", color_hex="1B4F72", font_size=Pt(9.5))

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
    format_table(tbl_team, font_size=Pt(8.5))

    # --------------------------------------------------------------------------
    # OVERVIEW & ETHICAL PROTOCOL
    # --------------------------------------------------------------------------
    add_h1("1. Survey Overview & Administration Protocol")
    add_p("This document specifies the complete final questionnaire administered for the empirical investigation on digital advertising effectiveness and promotional framing. Data collection was conducted through a bespoke web application engineered in Python using Streamlit, deployed to Streamlit Community Cloud, and synchronized in real time with Google Sheets via a Google Cloud API Service Account.")
    add_p("Ethical & Privacy Assurance: The survey is strictly anonymous. No personal identifiers (names, email addresses, phone numbers, IP addresses) were collected. Participation was voluntary. To prevent response distortion and demographic skew, stratified quota sampling was enforced across five age brackets (Under 18: 10%, 18–22: 30%, 23–30: 30%, 31–45: 20%, Above 45: 10%; Target N = 80).")
    add_p("The questionnaire is structured into two sequential phases:")
    add_callout("12 consumer behavior questions capturing product discovery channels, ad exposure frequency, historical ad engagement (clicks and purchases), advertising skepticism mechanisms, discount importance, and promotional framing preferences.", "Phase 1 (Consumer Behavior Survey): ")
    add_callout("A within-subjects experimental trial evaluating consumer responses to Rs. 2,000 wireless headphones across three distinct promotional framing conditions (Immediate ₹400 Discount, Delayed ₹400 Cashback, and Baseline Control), evaluated on three 5-point Likert telemetry outcome scales.", "Phase 2 (Promotional Framing Experiment): ")

    # --------------------------------------------------------------------------
    # PART 1: GENERAL CONSUMER BEHAVIOR SURVEY
    # --------------------------------------------------------------------------
    add_h1("2. Part 1: General Consumer Behaviour Questionnaire")
    add_p("Introductory Participant Prompt: 'Thank you for participating! This survey is completely anonymous: no names or personal details are collected. Please answer the following questions regarding your online shopping habits and advertising perceptions.'")

    q_items = [
        {
            "num": "Question 1",
            "text": "What is your age group? * (Required)",
            "type": "Categorical (Ordered Quota Filter)",
            "options": ["Under 18", "18–22", "23–30", "31–45", "Above 45"],
            "note": "Configured as the sole mandatory field in the application to enforce demographic quota constraints."
        },
        {
            "num": "Question 2",
            "text": "Which platform do you most often use to discover or learn about products?",
            "type": "Categorical (Nominal Channel Attribution)",
            "options": ["Instagram", "YouTube", "Google/Search", "E-commerce apps (Amazon, Flipkart, etc.)", "Facebook", "Other"],
            "note": "Identifies the respondent's primary digital discovery gateway."
        },
        {
            "num": "Question 3",
            "text": "How frequently do you encounter online advertisements?",
            "type": "Ordinal (5-Point Exposure Frequency)",
            "options": ["Never", "Rarely", "Sometimes", "Often", "Very often"],
            "note": "Measures self-reported volume of exposure to digital advertising."
        },
        {
            "num": "Question 4",
            "text": "Which type of product do you most frequently purchase online?",
            "type": "Categorical (Nominal Product Category)",
            "options": ["Electronics", "Clothing and fashion", "Food and groceries", "Beauty and personal care", "Home and lifestyle", "Other"],
            "note": "Captures the respondent's dominant online purchasing category."
        },
        {
            "num": "Question 5",
            "text": "In the past three months, have you clicked on an online advertisement to learn more about a product?",
            "type": "Binary / Categorical (Upper-Funnel Engagement)",
            "options": ["Yes", "No", "Not sure"],
            "note": "Stages audience transition from ad exposure to active consideration."
        },
        {
            "num": "Question 6",
            "text": "In the past three months, have you purchased a product after seeing an online advertisement?",
            "type": "Binary / Categorical (Lower-Funnel Conversion)",
            "options": ["Yes", "No", "Not sure"],
            "note": "Stages audience transition to actual commercial acquisition."
        },
        {
            "num": "Question 6a",
            "text": "If yes, which platform did you see the advertisement on before purchasing?",
            "type": "Categorical (Conditional Channel Attribution)",
            "options": ["Instagram", "YouTube", "Google/Search", "E-commerce apps (Amazon, Flipkart, etc.)", "Facebook", "Other", "Cannot remember"],
            "note": "Conditional logic: Displayed only if respondent answered 'Yes' to Question 6. Recorded as 'Not applicable' for non-purchasers."
        },
        {
            "num": "Question 7",
            "text": "How much do you trust product information shown in online advertisements?",
            "type": "Continuous Rating (1–5 Likert Scale)",
            "options": ["1 – Do not trust at all", "2 – Rarely trust", "3 – Neutral / Moderate trust", "4 – Mostly trust", "5 – Trust completely"],
            "note": "Quantifies baseline institutional advertising trust (ad_trust)."
        },
        {
            "num": "Question 8",
            "text": "What most influences your decision to purchase a product after seeing an advertisement?",
            "type": "Categorical (Nominal Influence Driver)",
            "options": ["Price or discount", "Product quality and features", "Customer reviews and ratings", "Brand reputation", "Recommendation from friends or family", "The advertisement itself", "Other"],
            "note": "Identifies the psychological factor driving conversion."
        },
        {
            "num": "Question 9",
            "text": "What is the main reason you ignore an online advertisement?",
            "type": "Categorical (Nominal Skepticism / Avoidance Driver)",
            "options": ["The product is irrelevant to me", "I do not trust the advertisement", "The price or offer is unattractive", "I see too many advertisements (ad fatigue)", "I prefer researching products independently", "Other"],
            "note": "Identifies the core driver behind ad resistance."
        },
        {
            "num": "Question 10",
            "text": "How important is a discount when deciding whether to purchase a product?",
            "type": "Continuous Rating (1–5 Likert Scale)",
            "options": ["1 – Not at all important", "2 – Slightly important", "3 – Moderately important", "4 – Very important", "5 – Extremely important"],
            "note": "Measures subjective price sensitivity (discount_importance)."
        },
        {
            "num": "Question 11",
            "text": "Which promotional offer would most encourage you to consider purchasing a product?",
            "type": "Categorical (Nominal Promotional Preference)",
            "options": ["Immediate price discount", "Cashback after purchase", "Coupon or promo code", "Free delivery", "Buy-one-get-one (BOGO) offer", "None of these"],
            "note": "Evaluates baseline stated preference across deal structures."
        },
        {
            "num": "Question 12",
            "text": "If two products were otherwise equally suitable, how likely would you be to choose the one with a promotional offer?",
            "type": "Continuous Rating (1–5 Likert Scale)",
            "options": ["1 – Very unlikely", "2 – Unlikely", "3 – Neither likely nor unlikely", "4 – Likely", "5 – Very likely"],
            "note": "Measures competitive promotion preference (choose_promoted_product_likelihood)."
        }
    ]

    for q in q_items:
        p_q = doc.add_paragraph()
        p_q.paragraph_format.space_before = Pt(5)
        p_q.paragraph_format.space_after = Pt(1)
        p_q.paragraph_format.keep_with_next = True
        r_num = p_q.add_run(f"{q['num']}: {q['text']}")
        r_num.font.bold = True
        r_num.font.size = Pt(10)
        r_num.font.color.rgb = COLOR_PRIMARY
        
        p_opt = doc.add_paragraph()
        p_opt.paragraph_format.space_before = Pt(1)
        p_opt.paragraph_format.space_after = Pt(1)
        p_opt.paragraph_format.left_indent = Inches(0.2)
        r_type = p_opt.add_run(f"Type: {q['type']}\nResponse Options: ")
        r_type.font.size = Pt(9)
        r_type.font.italic = True
        r_type.font.color.rgb = COLOR_SECONDARY
        
        opt_str = "   •   ".join(q['options'])
        r_opts = p_opt.add_run(opt_str)
        r_opts.font.size = Pt(9)
        r_opts.font.color.rgb = COLOR_TEXT
        
        p_nt = doc.add_paragraph()
        p_nt.paragraph_format.space_before = Pt(0)
        p_nt.paragraph_format.space_after = Pt(4)
        p_nt.paragraph_format.left_indent = Inches(0.2)
        r_nt = p_nt.add_run(f"Logic / Purpose: {q['note']}")
        r_nt.font.size = Pt(8.5)
        r_nt.font.color.rgb = COLOR_MUTED

    # --------------------------------------------------------------------------
    # PART 2: PROMOTIONAL FRAMING EXPERIMENT
    # --------------------------------------------------------------------------
    add_h1("3. Part 2: Within-Subject Promotional Framing Experiment")
    add_p("Experimental Context: 'Imagine you are shopping online and considering a pair of wireless Bluetooth headphones originally priced at ₹2,000. Assume the product is relevant to your needs and its price is affordable. You evaluate the following three purchase scenarios.'")
    add_p("Presentation Order Randomization: To prevent order and fatigue bias, the display sequence of the three scenarios was dynamically randomized for each participant session using the Fisher-Yates shuffle algorithm (random.shuffle()).")

    # Table of Scenarios
    tbl_scen = doc.add_table(rows=4, cols=4)
    scen_headers = ["Condition", "Offer Title", "Commercial Stimulus Description", "Effective Price"]
    for c_idx, h in enumerate(scen_headers):
        tbl_scen.rows[0].cells[c_idx].paragraphs[0].text = h
        
    scen_data = [
        [
            "Scenario A\n(Discount)",
            "₹400 Immediate Discount",
            "You receive an immediate ₹400 discount applied at checkout. You pay ₹1,600 immediately. No waiting period or post-purchase redemption required.",
            "₹1,600\n(Pay Now)"
        ],
        [
            "Scenario B\n(Cashback)",
            "₹400 Post-Purchase Cashback",
            "You pay the full ₹2,000 at checkout and receive ₹400 cashback credited to your original payment method within 7 days after purchase. No minimum spend applies.",
            "₹1,600\n(₹2,000 Upfront)"
        ],
        [
            "Scenario C\n(Control)",
            "No Promotion (Control)",
            "You see the headphones advertised at regular price. No promotional discount, cashback, coupon, or special offer is available. You pay ₹2,000 at checkout.",
            "₹2,000\n(Full Price)"
        ]
    ]
    for r_idx, row_values in enumerate(scen_data):
        for c_idx, val in enumerate(row_values):
            tbl_scen.rows[r_idx+1].cells[c_idx].paragraphs[0].text = val
    format_table(tbl_scen, font_size=Pt(8.5))

    p_scen_cap = doc.add_paragraph()
    p_scen_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_scen_cap.paragraph_format.space_before = Pt(2)
    p_scen_cap.paragraph_format.space_after = Pt(6)
    r_scen_cap = p_scen_cap.add_run("Table A1. Experimental treatment conditions and commercial stimulus framing.")
    r_scen_cap.font.size = Pt(9)
    r_scen_cap.font.italic = True

    # Visual Stimuli Comparison Table
    add_h2("Visual Stimuli Presented in Web Application")
    add_p("The following visual stimuli were rendered in the Streamlit application interface for each scenario:")

    tbl_imgs = doc.add_table(rows=2, cols=3)
    tbl_imgs.alignment = WD_TABLE_ALIGNMENT.CENTER
    headers_img = ["Scenario A: Immediate Discount", "Scenario B: Delayed Cashback", "Scenario C: Baseline Control"]
    for c_idx, h in enumerate(headers_img):
        tbl_imgs.rows[0].cells[c_idx].paragraphs[0].text = h
    format_table(tbl_imgs, header_bg="2E75B6", font_size=Pt(9))

    img_files = ['assets/scenario_discount.png', 'assets/scenario_cashback.png', 'assets/scenario_no_promotion.png']
    for c_idx, img_p in enumerate(img_files):
        cell = tbl_imgs.rows[1].cells[c_idx]
        p_c = cell.paragraphs[0]
        p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if os.path.exists(img_p):
            r_c = p_c.add_run()
            r_c.add_picture(img_p, width=Inches(1.85))

    p_img_cap = doc.add_paragraph()
    p_img_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_img_cap.paragraph_format.space_before = Pt(2)
    p_img_cap.paragraph_format.space_after = Pt(6)
    r_img_cap = p_img_cap.add_run("Figure A1. Visual stimulus cards rendered in the Streamlit questionnaire application.")
    r_img_cap.font.size = Pt(9)
    r_img_cap.font.italic = True

    # --------------------------------------------------------------------------
    # PART 3: EXPERIMENTAL OUTCOME TELEMETRY SCALES
    # --------------------------------------------------------------------------
    add_h1("4. Part 3: Experimental Outcome Telemetry Measures")
    add_p("For each of the three randomized scenarios (Discount, Cashback, Control), respondents evaluated the purchase option using three continuous 5-point Likert psychometric measures:")

    exp_qs = [
        {
            "num": "Question 13",
            "var": "purchase_intention",
            "prompt": "How likely are you to purchase this product?",
            "scale": "1 – Very unlikely  |  2 – Unlikely  |  3 – Neither likely nor unlikely  |  4 – Likely  |  5 – Very likely",
            "desc": "Primary dependent variable measuring commercial conversion probability."
        },
        {
            "num": "Question 14",
            "var": "offer_attractiveness",
            "prompt": "How attractive do you find this offer / purchase option?",
            "scale": "1 – Not attractive at all  |  2 – Slightly attractive  |  3 – Moderately attractive  |  4 – Very attractive  |  5 – Extremely attractive",
            "desc": "Psychological valuation measuring immediate perceptual deal appeal."
        },
        {
            "num": "Question 15",
            "var": "value_for_money_confidence",
            "prompt": "How confident are you that this option provides good value for money?",
            "scale": "1 – Not confident at all  |  2 – Slightly confident  |  3 – Moderately confident  |  4 – Very confident  |  5 – Extremely confident",
            "desc": "Cognitive assessment measuring perceived economic certainty and transaction security."
        }
    ]

    for eq in exp_qs:
        p_eq = doc.add_paragraph()
        p_eq.paragraph_format.space_before = Pt(4)
        p_eq.paragraph_format.space_after = Pt(1)
        p_eq.paragraph_format.keep_with_next = True
        r_eq = p_eq.add_run(f"{eq['num']} [{eq['var']}]: \"{eq['prompt']}\"")
        r_eq.font.bold = True
        r_eq.font.size = Pt(10)
        r_eq.font.color.rgb = COLOR_PRIMARY
        
        p_eq_opt = doc.add_paragraph()
        p_eq_opt.paragraph_format.space_before = Pt(1)
        p_eq_opt.paragraph_format.space_after = Pt(1)
        p_eq_opt.paragraph_format.left_indent = Inches(0.2)
        r_sc = p_eq_opt.add_run(f"Scale: {eq['scale']}")
        r_sc.font.size = Pt(9)
        r_sc.font.color.rgb = COLOR_TEXT
        
        p_eq_nt = doc.add_paragraph()
        p_eq_nt.paragraph_format.space_before = Pt(0)
        p_eq_nt.paragraph_format.space_after = Pt(4)
        p_eq_nt.paragraph_format.left_indent = Inches(0.2)
        r_desc = p_eq_nt.add_run(f"Measurement Rationale: {eq['desc']}")
        r_desc.font.size = Pt(8.5)
        r_desc.font.italic = True
        r_desc.font.color.rgb = COLOR_MUTED

    # --------------------------------------------------------------------------
    # PART 4: TELEMETRY ARCHITECTURE & STATUS LIFECYCLE
    # --------------------------------------------------------------------------
    add_h1("5. Telemetry Capture & Status Lifecycle Management")
    add_p("The web application incorporates automated state tracking to record session completeness and eliminate manual data entry:")
    add_callout("Upon initiating the survey, a unique UUID is assigned (e.g., e87556d4-8853-467e). After submitting Part 1, the respondent's answers are appended to the 'Survey Responses' Google Sheet and assigned status = 'In progress'.", "1. Initial Survey Submission (Phase 1): ")
    add_callout("Upon completing all three experimental scenario trials, the telemetry ratings are committed as 3 rows to the 'Experiment Responses' sheet in a single atomic batch (append_rows()), and the survey record is updated to status = 'Completed'.", "2. Atomic Experiment Commit (Phase 2): ")
    add_callout("Participants who abandon the application after Part 1 retain status = 'In progress'. In data preprocessing, these 6 incomplete records were audited and dropped to preserve experimental balance, establishing the final N = 74 completed sample.", "3. Quality Audit & Incomplete Dropping: ")

    # Save DOCX
    docx_path = 'Questionnaire_Specification.docx'
    doc.save(docx_path)
    print(f"[OK] Successfully built {docx_path}")
    return docx_path

def convert_to_pdf(docx_path, pdf_path):
    import win32com.client as win32
    # Ensure Word is not locked
    subprocess.run(['powershell', '-Command', 'Get-Process -Name WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force'])
    
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    try:
        doc = word.Documents.Open(os.path.abspath(docx_path))
        pages = doc.ComputeStatistics(2)
        doc.SaveAs(os.path.abspath(pdf_path), FileFormat=17)
        doc.Close(False)
        print(f"[OK] Exported PDF: {pdf_path} (Total Pages: {pages})")
    finally:
        word.Quit()

if __name__ == '__main__':
    docx_file = create_questionnaire_docx()
    
    # Export primary Questionnaire.pdf
    convert_to_pdf(docx_file, 'Questionnaire.pdf')
    
    # Also export Step1_Application_Development_and_Deployment_Workflow_Report.pdf copy
    convert_to_pdf(docx_file, 'Final_Questionnaire.pdf')
    convert_to_pdf(docx_file, 'Step1_Application_Development_and_Deployment_Workflow_Report.pdf')

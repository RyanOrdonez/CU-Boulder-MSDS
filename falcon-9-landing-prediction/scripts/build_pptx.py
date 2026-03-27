"""
Build the SpaceY Capstone PowerPoint presentation.
Run from the project root:  python scripts/build_pptx.py
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ── Paths ────────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG = os.path.join(ROOT, "images")
OUT = os.path.join(ROOT, "slides")
os.makedirs(OUT, exist_ok=True)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── Color palette ────────────────────────────────────────────────────────────
DARK_BLUE = RGBColor(0x1B, 0x2A, 0x4A)
ACCENT_BLUE = RGBColor(0x2E, 0x86, 0xC1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xF2, 0xF2, 0xF2)
DARK_GRAY = RGBColor(0x33, 0x33, 0x33)


# ── Helpers ──────────────────────────────────────────────────────────────────
def add_bg(slide, color=DARK_BLUE):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_text_box(slide, left, top, width, height, text, font_size=18,
                 bold=False, color=WHITE, alignment=PP_ALIGN.LEFT, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(Inches(left), Inches(top),
                                     Inches(width), Inches(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = font_name
    p.alignment = alignment
    return tf


def add_para(tf, text, font_size=16, bold=False, color=WHITE, alignment=PP_ALIGN.LEFT):
    p = tf.add_paragraph()
    p.text = text
    p.font.size = Pt(font_size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Calibri"
    p.alignment = alignment
    p.space_before = Pt(4)
    return p


def add_image_slide(title_text, img_filename, subtitle="", notes=""):
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    add_bg(slide, WHITE)
    # Title bar
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Inches(0), Inches(0),
                                    Inches(13.333), Inches(1.0))
    shape.fill.solid()
    shape.fill.fore_color.rgb = DARK_BLUE
    shape.line.fill.background()
    add_text_box(slide, 0.5, 0.15, 12, 0.7, title_text,
                 font_size=28, bold=True, color=WHITE)
    if subtitle:
        add_text_box(slide, 0.5, 0.6, 12, 0.4, subtitle,
                     font_size=14, color=RGBColor(0xBB, 0xBB, 0xBB))

    img_path = os.path.join(IMG, img_filename)
    if os.path.exists(img_path):
        slide.shapes.add_picture(img_path, Inches(1.5), Inches(1.3),
                                 Inches(10.3), Inches(5.8))
    else:
        add_text_box(slide, 2, 3, 9, 1, f"[Image not found: {img_filename}]",
                     font_size=20, color=DARK_GRAY)
    return slide


# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — Cover
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, DARK_BLUE)
add_text_box(slide, 1, 1.5, 11, 1.5,
             "SpaceY Launch Cost Prediction Analysis",
             font_size=44, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)
add_text_box(slide, 1, 3.2, 11, 0.8,
             "Predicting Falcon 9 First Stage Landing Success",
             font_size=24, color=ACCENT_BLUE, alignment=PP_ALIGN.CENTER)
add_text_box(slide, 1, 4.5, 11, 0.6,
             "IBM Data Science Professional Certificate — Capstone Project",
             font_size=18, color=RGBColor(0xBB, 0xBB, 0xBB), alignment=PP_ALIGN.CENTER)
add_text_box(slide, 1, 5.5, 11, 0.5,
             "Ryan Ordonez  |  2025",
             font_size=16, color=RGBColor(0x99, 0x99, 0x99), alignment=PP_ALIGN.CENTER)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — Executive Summary
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Executive Summary",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 11.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
bullets = [
    "SpaceX advertises Falcon 9 launches at $62M vs. $165M+ from competitors — savings driven by first stage reuse.",
    "SpaceY (our scenario company) aims to predict landing success to estimate competitive launch pricing.",
    "We collected data via SpaceX REST API and Wikipedia web scraping (90 Falcon 9 launches).",
    "Exploratory analysis reveals: success rate improved from 0% (2010) to >80% (2020); KSC LC-39A is the most successful site.",
    "SQL analysis confirmed payload mass, orbit type, and booster version as key differentiators.",
    "Folium maps show all launch sites are coastal and close to transportation infrastructure.",
    "Four ML models (Logistic Regression, SVM, Decision Tree, KNN) all achieved 83.3% test accuracy.",
    "Decision Tree had the highest cross-validation score (88.75%); recommended as the primary model.",
    "Business impact: Accurate landing prediction enables SpaceY to bid competitively at ~$62M per launch."
]
for b in bullets:
    add_para(tf, f"• {b}", font_size=15, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — Table of Contents
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Table of Contents",
             font_size=32, bold=True, color=WHITE)

toc_items = [
    "1.  Introduction & Problem Statement",
    "2.  Data Collection — SpaceX REST API",
    "3.  Data Collection — Web Scraping",
    "4.  Data Wrangling",
    "5.  EDA with Visualization",
    "6.  EDA with SQL",
    "7.  Interactive Visual Analytics — Folium Maps",
    "8.  Interactive Visual Analytics — Plotly Dash",
    "9.  Predictive Modeling",
    "10. Model Evaluation & Results",
    "11. Conclusion & Business Implications",
]
tf = add_text_box(slide, 2, 1.3, 9, 5.5, "", font_size=18, color=DARK_GRAY)
tf.paragraphs[0].text = ""
for item in toc_items:
    add_para(tf, item, font_size=20, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — Introduction
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Introduction & Problem Statement",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 11.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Background", font_size=22, bold=True, color=DARK_BLUE)
add_para(tf, "SpaceX has revolutionized the rocket launch industry by successfully landing and reusing Falcon 9 first stages, reducing launch costs from ~$165M to ~$62M.", font_size=16, color=DARK_GRAY)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Problem Statement", font_size=22, bold=True, color=DARK_BLUE)
add_para(tf, "SpaceY, a competing company, wants to predict whether the Falcon 9 first stage will land successfully. This prediction determines whether the launch cost will be ~$62M (reusable) or ~$165M (expendable).", font_size=16, color=DARK_GRAY)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Methodology", font_size=22, bold=True, color=DARK_BLUE)
add_para(tf, "1. Collect data via SpaceX API & Wikipedia web scraping", font_size=16, color=DARK_GRAY)
add_para(tf, "2. Wrangle and clean datasets; engineer target variable (Class)", font_size=16, color=DARK_GRAY)
add_para(tf, "3. Perform EDA with visualizations and SQL", font_size=16, color=DARK_GRAY)
add_para(tf, "4. Build interactive maps (Folium) and dashboards (Plotly Dash)", font_size=16, color=DARK_GRAY)
add_para(tf, "5. Train and compare ML models: Logistic Regression, SVM, Decision Tree, KNN", font_size=16, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — API Data Collection
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Data Collection — SpaceX REST API",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "API Endpoint: api.spacexdata.com/v4", font_size=14, bold=True, color=ACCENT_BLUE)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Process:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "1. GET /v4/launches/past — all past launches", font_size=14, color=DARK_GRAY)
add_para(tf, "2. Filter for Falcon 9 (rocket ID)", font_size=14, color=DARK_GRAY)
add_para(tf, "3. Resolve IDs via helper endpoints:", font_size=14, color=DARK_GRAY)
add_para(tf, "   • /v4/rockets/{id} → BoosterVersion", font_size=13, color=DARK_GRAY)
add_para(tf, "   • /v4/launchpads/{id} → LaunchSite, Lat, Lon", font_size=13, color=DARK_GRAY)
add_para(tf, "   • /v4/payloads/{id} → PayloadMass, Orbit", font_size=13, color=DARK_GRAY)
add_para(tf, "   • /v4/cores/{id} → Serial, Block, Landing", font_size=13, color=DARK_GRAY)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Output:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "90 Falcon 9 launches × 17 columns", font_size=14, color=DARK_GRAY)
add_para(tf, "Saved as dataset_part_1.csv", font_size=14, color=DARK_GRAY)

# Right side — key columns table
tf2 = add_text_box(slide, 7, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf2.paragraphs[0].text = ""
add_para(tf2, "Key Columns:", font_size=18, bold=True, color=DARK_BLUE)
cols = ["FlightNumber, Date, BoosterVersion",
        "PayloadMass, Orbit, LaunchSite",
        "Outcome (landing_success + landing_type)",
        "Flights, GridFins, Reused, Legs",
        "LandingPad, Block, ReusedCount, Serial",
        "Longitude, Latitude"]
for c in cols:
    add_para(tf2, f"  • {c}", font_size=13, color=DARK_GRAY)
add_para(tf2, "", font_size=10, color=DARK_GRAY)
add_para(tf2, "Outcome column format:", font_size=16, bold=True, color=DARK_BLUE)
add_para(tf2, '  "True ASDS"  → successful drone ship', font_size=13, color=DARK_GRAY)
add_para(tf2, '  "True RTLS"  → successful ground pad', font_size=13, color=DARK_GRAY)
add_para(tf2, '  "False Ocean" → failed ocean landing', font_size=13, color=DARK_GRAY)
add_para(tf2, '  "None None"  → no landing attempt', font_size=13, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — Web Scraping
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Data Collection — Web Scraping",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 11.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Source: Wikipedia — List of Falcon 9 and Falcon Heavy Launches", font_size=16, bold=True, color=ACCENT_BLUE)
add_para(tf, "(Snapshot: 9th June 2021 revision)", font_size=13, color=DARK_GRAY)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Tools: requests + BeautifulSoup (html.parser)", font_size=16, bold=True, color=DARK_BLUE)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Process:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "1. HTTP GET request to static Wikipedia URL", font_size=15, color=DARK_GRAY)
add_para(tf, "2. Parse HTML with BeautifulSoup", font_size=15, color=DARK_GRAY)
add_para(tf, "3. Find all <table> elements (launch record tables)", font_size=15, color=DARK_GRAY)
add_para(tf, "4. Extract column names from <th> headers", font_size=15, color=DARK_GRAY)
add_para(tf, "5. Iterate through <tr> rows, extracting cell data", font_size=15, color=DARK_GRAY)
add_para(tf, "6. Build Pandas DataFrame from dictionary of lists", font_size=15, color=DARK_GRAY)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Columns: Flight No., Date, Time, Version Booster, Launch site,", font_size=14, color=DARK_GRAY)
add_para(tf, "Payload, Payload mass, Orbit, Customer, Launch outcome, Booster landing", font_size=14, color=DARK_GRAY)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "Output: spacex_web_scraped.csv", font_size=14, bold=True, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — Data Wrangling
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Data Wrangling",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Steps:", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf, "1. Load dataset_part_1.csv (90 rows × 17 cols)", font_size=15, color=DARK_GRAY)
add_para(tf, "2. Identify missing values (LandingPad: 28.9%)", font_size=15, color=DARK_GRAY)
add_para(tf, "3. Impute PayloadMass nulls with column mean", font_size=15, color=DARK_GRAY)
add_para(tf, "4. Count launches per site and per orbit", font_size=15, color=DARK_GRAY)
add_para(tf, "5. Analyze landing outcome categories", font_size=15, color=DARK_GRAY)
add_para(tf, "6. Engineer target variable: Class", font_size=15, color=DARK_GRAY)

tf2 = add_text_box(slide, 7, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf2.paragraphs[0].text = ""
add_para(tf2, "Target Variable — Class:", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf2, "", font_size=8, color=DARK_GRAY)
add_para(tf2, "Class = 1  (Successful landing)", font_size=16, bold=True, color=RGBColor(0x27, 0xAE, 0x60))
add_para(tf2, "  • True ASDS, True RTLS, True Ocean", font_size=14, color=DARK_GRAY)
add_para(tf2, "", font_size=8, color=DARK_GRAY)
add_para(tf2, "Class = 0  (Failed / no attempt)", font_size=16, bold=True, color=RGBColor(0xE7, 0x4C, 0x3C))
add_para(tf2, "  • False ASDS, False Ocean, False RTLS", font_size=14, color=DARK_GRAY)
add_para(tf2, "  • None None, None ASDS", font_size=14, color=DARK_GRAY)
add_para(tf2, "", font_size=10, color=DARK_GRAY)
add_para(tf2, "Distribution:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf2, "  Success (1): 60 launches (66.7%)", font_size=15, color=DARK_GRAY)
add_para(tf2, "  Failure (0): 30 launches (33.3%)", font_size=15, color=DARK_GRAY)
add_para(tf2, "", font_size=8, color=DARK_GRAY)
add_para(tf2, "Output: dataset_part_2.csv", font_size=14, bold=True, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDES 8-13 — EDA Visuals
# ══════════════════════════════════════════════════════════════════════════════
eda_slides = [
    ("EDA — Flight Number vs Payload Mass", "eda_flightnumber_vs_payloadmass.png",
     "Higher flight numbers show more successful landings. Payload mass alone does not determine outcome."),
    ("EDA — Flight Number vs Launch Site", "eda_flightnumber_vs_launchsite.png",
     "KSC LC-39A has the most launches and highest concentration of successes at higher flight numbers."),
    ("EDA — Payload Mass vs Launch Site", "eda_payloadmass_vs_launchsite.png",
     "KSC LC-39A handles the heaviest payloads. VAFB SLC-4E has no launches above 10,000 kg."),
    ("EDA — Success Rate by Orbit Type", "eda_success_rate_by_orbit.png",
     "ES-L1, GEO, HEO, and SSO orbits show 100% success. GTO has mixed results."),
    ("EDA — Yearly Success Rate Trend", "eda_yearly_success_trend.png",
     "Success rate improved from 0% in 2010 to >80% by 2019-2020, showing SpaceX's learning curve."),
    ("EDA — Success Rate by Launch Site", "eda_success_rate_by_site.png",
     "KSC LC-39A leads with the highest success rate among all launch sites."),
]

for title, img, subtitle in eda_slides:
    add_image_slide(title, img, subtitle)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 14 — SQL Results
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "EDA with SQL — Key Results",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.2, 5.8, 5.8, "", font_size=14, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "10 SQL Queries Executed:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "", font_size=6, color=DARK_GRAY)
add_para(tf, "1. Unique launch sites: CCAFS LC-40,", font_size=13, color=DARK_GRAY)
add_para(tf, "   CCAFS SLC-40, KSC LC-39A, VAFB SLC-4E", font_size=13, color=DARK_GRAY)
add_para(tf, "2. CCA site records (5 sample rows)", font_size=13, color=DARK_GRAY)
add_para(tf, "3. NASA (CRS) total payload: 45,596 kg", font_size=13, color=DARK_GRAY)
add_para(tf, "4. F9 v1.1 avg payload: 2,928 kg", font_size=13, color=DARK_GRAY)
add_para(tf, "5. First ground pad landing: 2015-12-22", font_size=13, color=DARK_GRAY)

tf2 = add_text_box(slide, 7, 1.2, 5.8, 5.8, "", font_size=14, color=DARK_GRAY)
tf2.paragraphs[0].text = ""
add_para(tf2, "", font_size=18, color=DARK_GRAY)
add_para(tf2, "", font_size=6, color=DARK_GRAY)
add_para(tf2, "6. Drone ship success (4000-6000 kg):", font_size=13, color=DARK_GRAY)
add_para(tf2, "   F9 FT B1021.2, F9 FT B1031.2", font_size=13, color=DARK_GRAY)
add_para(tf2, "7. Mission outcomes: 98 Success, 1 Failure,", font_size=13, color=DARK_GRAY)
add_para(tf2, "   1 Success (payload status unclear)", font_size=13, color=DARK_GRAY)
add_para(tf2, "8. Max payload boosters: F9 B5 series", font_size=13, color=DARK_GRAY)
add_para(tf2, "9. 2015 drone ship failures: Jan & Apr", font_size=13, color=DARK_GRAY)
add_para(tf2, "10. Landing outcome rankings (2010-2017):", font_size=13, color=DARK_GRAY)
add_para(tf2, "    No Attempt: 21, Failure (drone ship): 9,", font_size=13, color=DARK_GRAY)
add_para(tf2, "    Success (drone ship): 5, others...", font_size=13, color=DARK_GRAY)
add_para(tf2, "", font_size=10, color=DARK_GRAY)
add_para(tf2, "Database: SQLite  |  Table: SPACEXTABLE", font_size=14, bold=True, color=ACCENT_BLUE)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 15 — Folium Maps
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Interactive Visual Analytics — Folium Maps",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Launch Site Map", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf, "• All 4 launch sites marked with circles", font_size=14, color=DARK_GRAY)
add_para(tf, "• Success (green) / Failure (red) markers", font_size=14, color=DARK_GRAY)
add_para(tf, "• MarkerCluster for interactive exploration", font_size=14, color=DARK_GRAY)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "KSC LC-39A Proximity Analysis", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf, "• Coastline: ~1.2 km", font_size=14, color=DARK_GRAY)
add_para(tf, "• Highway: ~7.1 km", font_size=14, color=DARK_GRAY)
add_para(tf, "• Railway: ~5.7 km", font_size=14, color=DARK_GRAY)
add_para(tf, "• City (Melbourne, FL): ~52 km", font_size=14, color=DARK_GRAY)

tf2 = add_text_box(slide, 7, 1.3, 5.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf2.paragraphs[0].text = ""
add_para(tf2, "Key Findings:", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf2, "", font_size=8, color=DARK_GRAY)
add_para(tf2, "• All sites are within 1-2 km of the coast", font_size=15, color=DARK_GRAY)
add_para(tf2, "  (safety: failed rockets fall into ocean)", font_size=13, color=RGBColor(0x77, 0x77, 0x77))
add_para(tf2, "", font_size=6, color=DARK_GRAY)
add_para(tf2, "• All sites are near equatorial latitudes", font_size=15, color=DARK_GRAY)
add_para(tf2, "  (physics: Earth's rotation provides velocity)", font_size=13, color=RGBColor(0x77, 0x77, 0x77))
add_para(tf2, "", font_size=6, color=DARK_GRAY)
add_para(tf2, "• Transport infrastructure nearby for logistics", font_size=15, color=DARK_GRAY)
add_para(tf2, "", font_size=6, color=DARK_GRAY)
add_para(tf2, "• Safe distance (50+ km) from populated areas", font_size=15, color=DARK_GRAY)
add_para(tf2, "", font_size=10, color=DARK_GRAY)
add_para(tf2, "Interactive maps: output/launch_site_map.html", font_size=13, bold=True, color=ACCENT_BLUE)
add_para(tf2, "                  output/ksc_proximity_map.html", font_size=13, bold=True, color=ACCENT_BLUE)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 16 — Plotly Dash
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Interactive Visual Analytics — Plotly Dash",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 11.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Dashboard Features:", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf, "", font_size=8, color=DARK_GRAY)
add_para(tf, "1. Launch Site Dropdown — Filter by individual site or view all sites", font_size=16, color=DARK_GRAY)
add_para(tf, "2. Pie Chart — Success vs Failure distribution for selected site", font_size=16, color=DARK_GRAY)
add_para(tf, "3. Payload Range Slider — Filter launches by payload mass (0–10,000 kg)", font_size=16, color=DARK_GRAY)
add_para(tf, "4. Scatter Plot — Payload Mass vs Outcome, colored by Booster Version", font_size=16, color=DARK_GRAY)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Implementation:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "• Framework: Plotly Dash", font_size=15, color=DARK_GRAY)
add_para(tf, "• Callbacks: Two reactive callbacks (dropdown + slider → charts)", font_size=15, color=DARK_GRAY)
add_para(tf, "• Data source: spacex_launch_dash.csv from IBM cloud", font_size=15, color=DARK_GRAY)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Run: python dash/spacex_dash_app.py → http://localhost:8050", font_size=15, bold=True, color=ACCENT_BLUE)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDES 17-20 — Confusion Matrices
# ══════════════════════════════════════════════════════════════════════════════
cm_slides = [
    ("Logistic Regression — Confusion Matrix", "cm_logistic_regression.png",
     "Test Accuracy: 83.33%  |  Best CV Score: 84.64%  |  Best C=0.01, penalty=l2"),
    ("SVM — Confusion Matrix", "cm_svm.png",
     "Test Accuracy: 83.33%  |  Best CV Score: 84.82%  |  Best kernel=sigmoid, C=1.0"),
    ("Decision Tree — Confusion Matrix", "cm_decision_tree.png",
     "Test Accuracy: 83.33%  |  Best CV Score: 88.75%  |  Best criterion=entropy, max_depth=14"),
    ("KNN — Confusion Matrix", "cm_knn.png",
     "Test Accuracy: 83.33%  |  Best CV Score: 84.82%  |  Best n_neighbors=10, p=1"),
]

for title, img, subtitle in cm_slides:
    add_image_slide(title, img, subtitle)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 21 — Model Comparison
# ══════════════════════════════════════════════════════════════════════════════
add_image_slide("Model Comparison — CV Score vs Test Accuracy",
                "model_comparison.png",
                "All models achieve 83.33% test accuracy. Decision Tree has the highest CV score (88.75%).")

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 22 — Conclusion
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Conclusion & Business Implications",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.2, 5.8, 5.8, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "Key Findings:", font_size=22, bold=True, color=DARK_BLUE)
add_para(tf, "", font_size=6, color=DARK_GRAY)
add_para(tf, "• Landing success rate improved from 0% to >80%", font_size=15, color=DARK_GRAY)
add_para(tf, "  over 10 years of Falcon 9 launches", font_size=13, color=RGBColor(0x77, 0x77, 0x77))
add_para(tf, "• KSC LC-39A is the most reliable launch site", font_size=15, color=DARK_GRAY)
add_para(tf, "• Orbit type strongly predicts landing success", font_size=15, color=DARK_GRAY)
add_para(tf, "• Higher flight numbers → higher success rates", font_size=15, color=DARK_GRAY)
add_para(tf, "  (SpaceX learns from each launch)", font_size=13, color=RGBColor(0x77, 0x77, 0x77))
add_para(tf, "• All 4 ML models achieve ~83% test accuracy", font_size=15, color=DARK_GRAY)
add_para(tf, "• Decision Tree recommended (highest CV: 88.75%)", font_size=15, color=DARK_GRAY)

tf2 = add_text_box(slide, 7, 1.2, 5.8, 5.8, "", font_size=16, color=DARK_GRAY)
tf2.paragraphs[0].text = ""
add_para(tf2, "Business Implications:", font_size=22, bold=True, color=DARK_BLUE)
add_para(tf2, "", font_size=6, color=DARK_GRAY)
add_para(tf2, "• Successful landing → ~$62M launch cost", font_size=15, color=RGBColor(0x27, 0xAE, 0x60))
add_para(tf2, "• Failed landing → ~$165M launch cost", font_size=15, color=RGBColor(0xE7, 0x4C, 0x3C))
add_para(tf2, "", font_size=8, color=DARK_GRAY)
add_para(tf2, "SpaceY can use these predictions to:", font_size=16, bold=True, color=DARK_BLUE)
add_para(tf2, "1. Estimate competitor launch costs accurately", font_size=15, color=DARK_GRAY)
add_para(tf2, "2. Identify which launch conditions favor", font_size=15, color=DARK_GRAY)
add_para(tf2, "   successful first-stage recovery", font_size=15, color=DARK_GRAY)
add_para(tf2, "3. Develop competitive pricing strategies", font_size=15, color=DARK_GRAY)
add_para(tf2, "4. Prioritize reusable rocket development", font_size=15, color=DARK_GRAY)
add_para(tf2, "", font_size=10, color=DARK_GRAY)
add_para(tf2, "Prediction accuracy of 83%+ makes this model", font_size=15, bold=True, color=DARK_BLUE)
add_para(tf2, "actionable for competitive intelligence.", font_size=15, bold=True, color=DARK_BLUE)

# ══════════════════════════════════════════════════════════════════════════════
# SLIDE 23 — Appendix
# ══════════════════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, WHITE)
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
                                Inches(13.333), Inches(1.0))
shape.fill.solid()
shape.fill.fore_color.rgb = DARK_BLUE
shape.line.fill.background()
add_text_box(slide, 0.5, 0.15, 12, 0.7, "Appendix",
             font_size=32, bold=True, color=WHITE)

tf = add_text_box(slide, 0.8, 1.3, 11.5, 5.5, "", font_size=16, color=DARK_GRAY)
tf.paragraphs[0].text = ""
add_para(tf, "GitHub Repository", font_size=20, bold=True, color=DARK_BLUE)
add_para(tf, "https://github.com/<your-username>/spacey-capstone", font_size=15, color=ACCENT_BLUE)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Repository Structure:", font_size=18, bold=True, color=DARK_BLUE)
add_para(tf, "  spacey-capstone/", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── data/          — CSVs, SQLite database", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── notebooks/     — 7 Jupyter notebooks (01–07)", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── scripts/       — Helper scripts (build_pptx.py)", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── sql/           — SQL query scripts", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── images/        — Exported charts and confusion matrices", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── dash/          — Plotly Dash application", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── output/        — Folium HTML maps", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── slides/        — PowerPoint and PDF", font_size=14, color=DARK_GRAY)
add_para(tf, "  ├── requirements.txt", font_size=14, color=DARK_GRAY)
add_para(tf, "  └── README.md", font_size=14, color=DARK_GRAY)
add_para(tf, "", font_size=10, color=DARK_GRAY)
add_para(tf, "Tech Stack: Python, Pandas, NumPy, Matplotlib, Seaborn, Scikit-learn,", font_size=14, color=DARK_GRAY)
add_para(tf, "BeautifulSoup, Folium, Plotly Dash, SQLite, Jupyter", font_size=14, color=DARK_GRAY)

# ══════════════════════════════════════════════════════════════════════════════
# Save
# ══════════════════════════════════════════════════════════════════════════════
output_path = os.path.join(OUT, "SpaceY_Launch_Cost_Prediction_Analysis.pptx")
prs.save(output_path)
print(f"Saved {len(prs.slides)} slides to {output_path}")

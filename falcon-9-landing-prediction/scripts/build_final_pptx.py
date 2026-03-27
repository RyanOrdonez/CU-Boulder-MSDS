"""
Build the final SpaceY Capstone PowerPoint — targeting all 15 grading criteria.
Run from project root:  python scripts/build_final_pptx.py

Grading criteria addressed per slide:
  1.1  GitHub URL on multiple slides
  1.2  PDF format (user exports manually)
  1.3  Executive Summary slide
  1.4  Introduction slide
  1.5  Data Collection – SpaceX API methodology + GitHub URL
  1.6  Data Collection – Web Scraping methodology + GitHub URL
  1.7  Data Wrangling methodology + GitHub URL
  1.8  EDA with Data Visualization methodology + GitHub URL
  1.9  EDA with SQL methodology + GitHub URL
  1.10 Interactive Visual Analytics methodology (Folium + Dash) + GitHub URL
  1.11 EDA visualization result slides (scatter, bar, yearly trends)
  1.12 EDA with SQL result slides (sites, payloads, success rates, rankings, time)
  1.13 Folium map slides (site markers, launch records, proximity)
  1.14 Plotly Dash slides (pie charts, scatter plots)
  1.15 Predictive Analysis slides (models, confusion matrices, best model, conclusion)
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# ── Config ───────────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IMG  = os.path.join(ROOT, "images")
OUT  = os.path.join(ROOT, "slides")
os.makedirs(OUT, exist_ok=True)

GITHUB_URL = "https://github.com/RyanOrdonez/spacey-capstone"

# ── Colors ───────────────────────────────────────────────────────────────────
DARK_BLUE  = RGBColor(0x1B, 0x2A, 0x4A)
ACCENT     = RGBColor(0x2E, 0x86, 0xC1)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
LGRAY      = RGBColor(0xF2, 0xF2, 0xF2)
DGRAY      = RGBColor(0x33, 0x33, 0x33)
GREEN      = RGBColor(0x27, 0xAE, 0x60)
RED        = RGBColor(0xE7, 0x4C, 0x3C)
MGRAY      = RGBColor(0x77, 0x77, 0x77)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)


# ── Helpers ──────────────────────────────────────────────────────────────────
def _bg(slide, color=DARK_BLUE):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def _bar(slide):
    """Dark title bar across the top."""
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                               Inches(0), Inches(0),
                               Inches(13.333), Inches(1.0))
    s.fill.solid(); s.fill.fore_color.rgb = DARK_BLUE; s.line.fill.background()


def _tb(slide, l, t, w, h, text, sz=18, bold=False, color=WHITE, align=PP_ALIGN.LEFT):
    """Add a text box and return its text_frame."""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = box.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text; p.font.size = Pt(sz); p.font.bold = bold
    p.font.color.rgb = color; p.font.name = "Calibri"; p.alignment = align
    return tf


def _p(tf, text, sz=16, bold=False, color=DGRAY, align=PP_ALIGN.LEFT, space=4):
    """Append a paragraph to an existing text frame."""
    p = tf.add_paragraph()
    p.text = text; p.font.size = Pt(sz); p.font.bold = bold
    p.font.color.rgb = color; p.font.name = "Calibri"; p.alignment = align
    p.space_before = Pt(space)
    return p


def _github_footer(slide):
    """Add GitHub URL footer to a slide."""
    _tb(slide, 0.5, 6.9, 12, 0.5,
        f"GitHub: {GITHUB_URL}", sz=11, bold=True, color=ACCENT)


def _slide_title(title_text):
    """Create a slide with white background, dark title bar, and title text."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    _bg(slide, WHITE); _bar(slide)
    _tb(slide, 0.5, 0.12, 12, 0.75, title_text, sz=30, bold=True, color=WHITE)
    return slide


def _img_slide(title, img_file, caption=""):
    """Slide with title bar + image + optional caption."""
    slide = _slide_title(title)
    path = os.path.join(IMG, img_file)
    if os.path.exists(path):
        slide.shapes.add_picture(path, Inches(1.5), Inches(1.2),
                                 Inches(10.3), Inches(5.5))
    else:
        _tb(slide, 3, 3, 7, 1, f"[Image not found: {img_file}]", sz=18, color=RED)
    if caption:
        _tb(slide, 0.5, 6.8, 12, 0.5, caption, sz=12, color=MGRAY)
    _github_footer(slide)
    return slide


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 1 — COVER PAGE
# ═══════════════════════════════════════════════════════════════════════════════
s = prs.slides.add_slide(prs.slide_layouts[6]); _bg(s, DARK_BLUE)
_tb(s, 1, 1.2, 11, 1.5,
    "SpaceY Launch Cost Prediction Analysis",
    sz=44, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
_tb(s, 1, 2.9, 11, 0.8,
    "Predicting Falcon 9 First Stage Landing Success",
    sz=24, color=ACCENT, align=PP_ALIGN.CENTER)
_tb(s, 1, 4.0, 11, 0.6,
    "IBM Data Science Professional Certificate — Capstone Project",
    sz=18, color=RGBColor(0xBB,0xBB,0xBB), align=PP_ALIGN.CENTER)
_tb(s, 1, 5.0, 11, 0.5,
    "Ryan Ordonez", sz=20, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
_tb(s, 1, 5.5, 11, 0.5,
    f"GitHub: {GITHUB_URL}", sz=14, color=ACCENT, align=PP_ALIGN.CENTER)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 2 — EXECUTIVE SUMMARY  [Criterion 1.3]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Executive Summary")
tf = _tb(s, 0.8, 1.2, 11.5, 5.5, "", sz=16, color=DGRAY)
bullets = [
    ("Methods Used:", True, DARK_BLUE, 18),
    ("• Data collected via SpaceX REST API (v4) and Wikipedia web scraping using BeautifulSoup.", False, DGRAY, 15),
    ("• Data wrangled: missing value imputation, target variable (Class) engineering.", False, DGRAY, 15),
    ("• EDA performed with Matplotlib, Seaborn visualizations and 10 SQL queries in SQLite.", False, DGRAY, 15),
    ("• Interactive analytics built with Folium maps and a Plotly Dash dashboard.", False, DGRAY, 15),
    ("• Four classification models trained with GridSearchCV (cv=10): Logistic Regression, SVM, Decision Tree, KNN.", False, DGRAY, 15),
    ("", False, DGRAY, 8),
    ("Key Results:", True, DARK_BLUE, 18),
    ("• All four models achieved 83.33% test accuracy on 18 test samples.", False, DGRAY, 15),
    ("• Decision Tree achieved the highest cross-validation score: 88.75%.", False, DGRAY, 15),
    ("• Landing success rate improved from 0% (2010) to over 80% (2020).", False, DGRAY, 15),
    ("• KSC LC-39A is the most successful launch site; ES-L1, GEO, HEO, SSO orbits show 100% success.", False, DGRAY, 15),
    ("• Successful landing prediction enables competitive bidding at ~$62M vs ~$165M per launch.", False, DGRAY, 15),
]
for text, bold, col, sz in bullets:
    _p(tf, text, sz=sz, bold=bold, color=col)
_github_footer(s)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 3 — TABLE OF CONTENTS
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Table of Contents")
tf = _tb(s, 1.5, 1.2, 10, 5.5, "", sz=18, color=DGRAY)
toc = [
    "1.   Executive Summary",
    "2.   Introduction",
    "3.   Data Collection — SpaceX API  (Methodology)",
    "4.   Data Collection — Web Scraping  (Methodology)",
    "5.   Data Wrangling  (Methodology)",
    "6.   EDA with Data Visualization  (Methodology)",
    "7.   EDA with SQL  (Methodology)",
    "8.   Interactive Visual Analytics  (Methodology)",
    "9.   Predictive Analysis  (Methodology)",
    "10.  EDA Visualization Results",
    "11.  EDA with SQL Results",
    "12.  Interactive Maps — Folium Results",
    "13.  Interactive Dashboard — Plotly Dash Results",
    "14.  Predictive Analysis Results",
    "15.  Conclusion & Innovative Insights",
    "16.  Appendix",
]
for item in toc:
    _p(tf, item, sz=17, color=DGRAY, space=2)

# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 4 — INTRODUCTION  [Criterion 1.4]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Introduction")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Project Background", sz=22, bold=True, color=DARK_BLUE)
_p(tf, "SpaceX has disrupted the rocket launch industry by", sz=15, color=DGRAY)
_p(tf, "successfully landing and reusing the Falcon 9 first", sz=15, color=DGRAY)
_p(tf, "stage, reducing launch cost from ~$165M to ~$62M.", sz=15, color=DGRAY)
_p(tf, "", sz=8, color=DGRAY)
_p(tf, "Problem Statement", sz=22, bold=True, color=DARK_BLUE)
_p(tf, "SpaceY, a fictional competitor, wants to predict", sz=15, color=DGRAY)
_p(tf, "whether the Falcon 9 first stage will land successfully.", sz=15, color=DGRAY)
_p(tf, "This prediction directly determines launch cost:", sz=15, color=DGRAY)
_p(tf, "  • Successful landing  →  ~$62M (reusable)", sz=15, bold=True, color=GREEN)
_p(tf, "  • Failed landing  →  ~$165M (expendable)", sz=15, bold=True, color=RED)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Methodology Overview", sz=22, bold=True, color=DARK_BLUE)
_p(tf2, "1. Data Collection (SpaceX API + Web Scraping)", sz=15, color=DGRAY)
_p(tf2, "2. Data Wrangling (cleaning, Class variable)", sz=15, color=DGRAY)
_p(tf2, "3. EDA with Visualization (Matplotlib, Seaborn)", sz=15, color=DGRAY)
_p(tf2, "4. EDA with SQL (SQLite, 10 queries)", sz=15, color=DGRAY)
_p(tf2, "5. Interactive Analytics (Folium maps, Plotly Dash)", sz=15, color=DGRAY)
_p(tf2, "6. Predictive Modeling (LR, SVM, DT, KNN)", sz=15, color=DGRAY)
_p(tf2, "", sz=8, color=DGRAY)
_p(tf2, "Dataset: 90 Falcon 9 launches (2010–2020)", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "Target: Class (1 = landed, 0 = failed/no attempt)", sz=15, color=DGRAY)
_p(tf2, "Features: 83 columns after one-hot encoding", sz=15, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 5 — DATA COLLECTION: SpaceX API  [Criterion 1.5]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Data Collection Methodology — SpaceX REST API")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "API Endpoint: api.spacexdata.com/v4", sz=15, bold=True, color=ACCENT)
_p(tf, "", sz=6)
_p(tf, "Data Collection Process (Flowchart):", sz=18, bold=True, color=DARK_BLUE)
_p(tf, "┌─────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ GET /v4/launches/past           │", sz=12, color=DGRAY)
_p(tf, "│ (Retrieve all past launches)    │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬──────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌─────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ Filter: rocket == Falcon 9 ID   │", sz=12, color=DGRAY)
_p(tf, "│ (5e9d0d95eda69973a809d1ec)      │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬──────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌─────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ For each launch, resolve IDs:   │", sz=12, color=DGRAY)
_p(tf, "│  • GET /v4/rockets/{id}         │", sz=12, color=DGRAY)
_p(tf, "│  • GET /v4/launchpads/{id}      │", sz=12, color=DGRAY)
_p(tf, "│  • GET /v4/payloads/{id}        │", sz=12, color=DGRAY)
_p(tf, "│  • GET /v4/cores/{id}           │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬──────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌─────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ Build flat DataFrame (90 × 17)  │", sz=12, color=DGRAY)
_p(tf, "│ Save → dataset_part_1.csv       │", sz=12, color=DGRAY)
_p(tf, "└─────────────────────────────────┘", sz=12, color=DGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Key Columns Extracted:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  FlightNumber, Date, BoosterVersion", sz=14, color=DGRAY)
_p(tf2, "  PayloadMass, Orbit, LaunchSite", sz=14, color=DGRAY)
_p(tf2, "  Outcome (landing success + type)", sz=14, color=DGRAY)
_p(tf2, "  Flights, GridFins, Reused, Legs", sz=14, color=DGRAY)
_p(tf2, "  LandingPad, Block, ReusedCount, Serial", sz=14, color=DGRAY)
_p(tf2, "  Longitude, Latitude", sz=14, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Outcome Column Format:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, '  "True ASDS"   → successful drone ship landing', sz=13, color=DGRAY)
_p(tf2, '  "True RTLS"   → successful ground pad landing', sz=13, color=DGRAY)
_p(tf2, '  "False Ocean"  → failed ocean landing', sz=13, color=DGRAY)
_p(tf2, '  "None None"    → no landing attempt', sz=13, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Tools: Python requests library", sz=14, bold=True, color=DARK_BLUE)
_p(tf2, "Output: data/dataset_part_1.csv (90 rows × 17 columns)", sz=14, color=DGRAY)
_p(tf2, f"Notebook: notebooks/01_spacex_api_collection.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 6 — DATA COLLECTION: Web Scraping  [Criterion 1.6]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Data Collection Methodology — Web Scraping")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Source: Wikipedia — List of Falcon 9 Launches", sz=15, bold=True, color=ACCENT)
_p(tf, "(Static HTML snapshot, June 9 2021)", sz=12, color=MGRAY)
_p(tf, "", sz=6)
_p(tf, "Web Scraping Process (Flowchart):", sz=18, bold=True, color=DARK_BLUE)
_p(tf, "┌──────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ requests.get(wikipedia_url)      │", sz=12, color=DGRAY)
_p(tf, "│ (Fetch static HTML page)         │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬───────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌──────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ BeautifulSoup(html, html.parser) │", sz=12, color=DGRAY)
_p(tf, "│ Parse HTML structure             │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬───────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌──────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ Find all <table> elements        │", sz=12, color=DGRAY)
_p(tf, "│ Extract <th> headers → columns   │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬───────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌──────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ Iterate <tr> rows, extract cells │", sz=12, color=DGRAY)
_p(tf, "│ Helper funcs: date_time,         │", sz=12, color=DGRAY)
_p(tf, "│ booster_version, landing_status,  │", sz=12, color=DGRAY)
_p(tf, "│ get_mass                          │", sz=12, color=DGRAY)
_p(tf, "└──────────────┬───────────────────┘", sz=12, color=DGRAY)
_p(tf, "               ▼", sz=12, color=DGRAY)
_p(tf, "┌──────────────────────────────────┐", sz=12, color=DGRAY)
_p(tf, "│ Dict → Pandas DataFrame          │", sz=12, color=DGRAY)
_p(tf, "│ Save → spacex_web_scraped.csv    │", sz=12, color=DGRAY)
_p(tf, "└──────────────────────────────────┘", sz=12, color=DGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Columns Scraped:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  Flight No., Date, Time (UTC)", sz=14, color=DGRAY)
_p(tf2, "  Version Booster, Launch site", sz=14, color=DGRAY)
_p(tf2, "  Payload, Payload mass, Orbit", sz=14, color=DGRAY)
_p(tf2, "  Customer, Launch outcome", sz=14, color=DGRAY)
_p(tf2, "  Booster landing", sz=14, color=DGRAY)
_p(tf2, "", sz=10)
_p(tf2, "Helper Functions:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  • date_time() — parse launch date/time", sz=13, color=DGRAY)
_p(tf2, "  • booster_version() — extract booster info", sz=13, color=DGRAY)
_p(tf2, "  • landing_status() — parse landing outcome", sz=13, color=DGRAY)
_p(tf2, "  • get_mass() — extract numeric payload mass", sz=13, color=DGRAY)
_p(tf2, "  • extract_column_from_header() — get headers", sz=13, color=DGRAY)
_p(tf2, "", sz=10)
_p(tf2, "Tools: requests, BeautifulSoup4", sz=14, bold=True, color=DARK_BLUE)
_p(tf2, "Output: data/spacex_web_scraped.csv", sz=14, color=DGRAY)
_p(tf2, f"Notebook: notebooks/02_web_scraping.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 7 — DATA WRANGLING METHODOLOGY  [Criterion 1.7]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Data Wrangling Methodology")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Data Cleaning & Processing Steps:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=6)
_p(tf, "Step 1: Load dataset_part_1.csv (90 rows × 17 cols)", sz=14, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "Step 2: Identify missing values", sz=14, color=DGRAY)
_p(tf, "  • LandingPad: 28.9% missing (26 nulls)", sz=13, color=MGRAY)
_p(tf, "  • PayloadMass: minor nulls → imputed with mean", sz=13, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "Step 3: Analyze launch frequency", sz=14, color=DGRAY)
_p(tf, "  • Count launches per site", sz=13, color=MGRAY)
_p(tf, "  • Count launches per orbit type", sz=13, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "Step 4: Analyze landing outcome categories", sz=14, color=DGRAY)
_p(tf, '  • True ASDS, True RTLS → Success', sz=13, color=GREEN)
_p(tf, '  • False Ocean, None None → Failure', sz=13, color=RED)
_p(tf, "", sz=4)
_p(tf, "Step 5: Engineer target variable — Class", sz=14, bold=True, color=DGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Target Variable: Class", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=6)
_p(tf2, "Class = 1 (Successful Landing)", sz=18, bold=True, color=GREEN)
_p(tf2, "  • Outcome starts with 'True'", sz=14, color=DGRAY)
_p(tf2, "  • True ASDS, True RTLS, True Ocean", sz=14, color=DGRAY)
_p(tf2, "  • Count: 60 launches (66.7%)", sz=14, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Class = 0 (Failed / No Attempt)", sz=18, bold=True, color=RED)
_p(tf2, "  • Outcome starts with 'False' or 'None'", sz=14, color=DGRAY)
_p(tf2, "  • False ASDS, False Ocean, None None", sz=14, color=DGRAY)
_p(tf2, "  • Count: 30 launches (33.3%)", sz=14, color=DGRAY)
_p(tf2, "", sz=10)
_p(tf2, "Output: data/dataset_part_2.csv (90 rows × 18 cols)", sz=14, bold=True, color=DARK_BLUE)
_p(tf2, f"Notebook: notebooks/03_data_wrangling.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 8 — EDA WITH VISUALIZATION METHODOLOGY  [Criterion 1.8]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("EDA with Data Visualization Methodology")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Charts Created and Their Purpose:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "1. Flight Number vs Launch Site (scatter)", sz=14, color=DGRAY)
_p(tf, "   → See which sites used over time", sz=12, color=MGRAY)
_p(tf, "2. Flight Number vs Payload Mass (scatter)", sz=14, color=DGRAY)
_p(tf, "   → Relationship between payload and success", sz=12, color=MGRAY)
_p(tf, "3. Payload Mass vs Launch Site (scatter)", sz=14, color=DGRAY)
_p(tf, "   → Payload capacity per site", sz=12, color=MGRAY)
_p(tf, "4. Success Rate by Orbit Type (bar chart)", sz=14, color=DGRAY)
_p(tf, "   → Which orbits have highest success", sz=12, color=MGRAY)
_p(tf, "5. Flight Number vs Orbit Type (scatter)", sz=14, color=DGRAY)
_p(tf, "   → Orbit selection evolution over time", sz=12, color=MGRAY)
_p(tf, "6. Payload Mass vs Orbit Type (scatter)", sz=14, color=DGRAY)
_p(tf, "   → Payload requirements per orbit", sz=12, color=MGRAY)
_p(tf, "7. Success Rate by Launch Site (bar chart)", sz=14, color=DGRAY)
_p(tf, "   → Compare site reliability", sz=12, color=MGRAY)
_p(tf, "8. Yearly Success Trend (line chart)", sz=14, color=DGRAY)
_p(tf, "   → Track improvement over time", sz=12, color=MGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Feature Engineering:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=6)
_p(tf2, "• One-hot encoding of categorical columns:", sz=15, color=DGRAY)
_p(tf2, "  Orbit, LaunchSite, LandingPad, Serial", sz=14, color=MGRAY)
_p(tf2, "• Cast all features to float64", sz=15, color=DGRAY)
_p(tf2, "• Result: 90 rows × 83 feature columns", sz=15, color=DGRAY)
_p(tf2, "", sz=10)
_p(tf2, "Tools: Matplotlib, Seaborn", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=6)
_p(tf2, "Output:", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  • 8 PNG charts saved to images/", sz=14, color=DGRAY)
_p(tf2, "  • data/dataset_part_3.csv (engineered features)", sz=14, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, f"Notebook: notebooks/04_eda.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 9 — EDA WITH SQL METHODOLOGY  [Criterion 1.9]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("EDA with SQL Methodology")
tf = _tb(s, 0.8, 1.2, 11.5, 5.5, "", sz=16, color=DGRAY)
_p(tf, "SQL Queries Performed (10 queries in SQLite):", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "Q1:  SELECT DISTINCT Launch_Site — unique launch sites", sz=14, color=DGRAY)
_p(tf, "Q2:  SELECT * WHERE Launch_Site LIKE 'CCA%' LIMIT 5 — filter by site prefix", sz=14, color=DGRAY)
_p(tf, "Q3:  SELECT SUM(PAYLOAD_MASS__KG_) WHERE Customer = 'NASA (CRS)' — total payload for customer", sz=14, color=DGRAY)
_p(tf, "Q4:  SELECT AVG(PAYLOAD_MASS__KG_) WHERE Booster_Version LIKE 'F9 v1.1%' — avg payload for booster", sz=14, color=DGRAY)
_p(tf, "Q5:  SELECT MIN(Date) WHERE Landing_Outcome = 'Success (ground pad)' — first ground pad landing", sz=14, color=DGRAY)
_p(tf, "Q6:  SELECT Booster_Version WHERE Landing_Outcome = 'Success (drone ship)' AND PAYLOAD 4000-6000", sz=14, color=DGRAY)
_p(tf, "Q7:  SELECT Mission_Outcome, COUNT(*) GROUP BY Mission_Outcome — mission outcome counts", sz=14, color=DGRAY)
_p(tf, "Q8:  SELECT Booster_Version WHERE PAYLOAD_MASS = MAX — max payload boosters", sz=14, color=DGRAY)
_p(tf, "Q9:  SELECT * WHERE Landing_Outcome = 'Failure (drone ship)' AND Date LIKE '2015%' — 2015 failures", sz=14, color=DGRAY)
_p(tf, "Q10: SELECT Landing_Outcome, COUNT(*) WHERE Date BETWEEN 2010-2017 ORDER BY COUNT — rankings", sz=14, color=DGRAY)
_p(tf, "", sz=8)
_p(tf, "Database: SQLite  |  Table: SPACEXTABLE  |  Loaded from dataset_part_2.csv", sz=14, bold=True, color=DARK_BLUE)
_p(tf, "SQL script: sql/spacex_queries.sql", sz=13, color=DGRAY)
_p(tf, f"Notebook: notebooks/05_sql_analysis.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 10 — INTERACTIVE VISUAL ANALYTICS METHODOLOGY  [Criterion 1.10]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Interactive Visual Analytics Methodology")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Folium Interactive Maps:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "Map 1: Launch Site Map", sz=16, bold=True, color=ACCENT)
_p(tf, "  • Mark all 4 launch sites with circles", sz=14, color=DGRAY)
_p(tf, "  • Color-code markers: green=success, red=fail", sz=14, color=DGRAY)
_p(tf, "  • MarkerCluster for interactive exploration", sz=14, color=DGRAY)
_p(tf, "  • Popup labels with site name + class", sz=14, color=DGRAY)
_p(tf, "", sz=6)
_p(tf, "Map 2: KSC LC-39A Proximity Analysis", sz=16, bold=True, color=ACCENT)
_p(tf, "  • Calculate distance to coastline (~1.2 km)", sz=14, color=DGRAY)
_p(tf, "  • Calculate distance to highway (~7.1 km)", sz=14, color=DGRAY)
_p(tf, "  • Calculate distance to railway (~5.7 km)", sz=14, color=DGRAY)
_p(tf, "  • Calculate distance to city (~52 km)", sz=14, color=DGRAY)
_p(tf, "  • Draw lines and markers on map", sz=14, color=DGRAY)
_p(tf, "", sz=6)
_p(tf, "Output: output/launch_site_map.html", sz=13, color=DGRAY)
_p(tf, "Output: output/ksc_proximity_map.html", sz=13, color=DGRAY)
_p(tf, f"Notebook: notebooks/06_interactive_visuals.ipynb", sz=13, color=ACCENT)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Plotly Dash Dashboard:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=4)
_p(tf2, "Dashboard Components:", sz=16, bold=True, color=ACCENT)
_p(tf2, "  1. Dropdown: Select launch site (or All Sites)", sz=14, color=DGRAY)
_p(tf2, "  2. Pie Chart: Success vs Failure for site", sz=14, color=DGRAY)
_p(tf2, "  3. Range Slider: Filter by payload mass (0-10,000 kg)", sz=14, color=DGRAY)
_p(tf2, "  4. Scatter Plot: Payload vs Outcome by Booster", sz=14, color=DGRAY)
_p(tf2, "", sz=6)
_p(tf2, "Implementation:", sz=16, bold=True, color=ACCENT)
_p(tf2, "  • Framework: Plotly Dash", sz=14, color=DGRAY)
_p(tf2, "  • Two reactive @app.callback decorators", sz=14, color=DGRAY)
_p(tf2, "  • Input: site dropdown + payload slider", sz=14, color=DGRAY)
_p(tf2, "  • Output: pie chart + scatter chart", sz=14, color=DGRAY)
_p(tf2, "  • Data: spacex_launch_dash.csv (IBM Cloud)", sz=14, color=DGRAY)
_p(tf2, "", sz=6)
_p(tf2, "Run: python dash/spacex_dash_app.py", sz=14, bold=True, color=DARK_BLUE)
_p(tf2, "Open: http://localhost:8050", sz=14, color=DGRAY)
_p(tf2, f"Script: dash/spacex_dash_app.py", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 11 — PREDICTIVE ANALYSIS METHODOLOGY
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Predictive Analysis Methodology")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Data Preparation:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "  • Target Y from dataset_part_2.csv (Class column)", sz=14, color=DGRAY)
_p(tf, "  • Features X from dataset_part_3.csv (83 cols)", sz=14, color=DGRAY)
_p(tf, "  • StandardScaler normalization", sz=14, color=DGRAY)
_p(tf, "  • Train/Test split: 80/20 (random_state=2)", sz=14, color=DGRAY)
_p(tf, "  • Train: 72 samples  |  Test: 18 samples", sz=14, color=DGRAY)
_p(tf, "", sz=8)
_p(tf, "Models Trained:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "  1. Logistic Regression", sz=15, color=DGRAY)
_p(tf, "     C: [0.01, 0.1, 1], penalty: l2", sz=13, color=MGRAY)
_p(tf, "  2. Support Vector Machine (SVM)", sz=15, color=DGRAY)
_p(tf, "     kernel: [linear, rbf, poly, sigmoid]", sz=13, color=MGRAY)
_p(tf, "     C: logspace(-3,3,5), gamma: logspace(-3,3,5)", sz=13, color=MGRAY)
_p(tf, "  3. Decision Tree Classifier", sz=15, color=DGRAY)
_p(tf, "     criterion: [gini, entropy], max_depth: [2-18]", sz=13, color=MGRAY)
_p(tf, "  4. K-Nearest Neighbors (KNN)", sz=15, color=DGRAY)
_p(tf, "     n_neighbors: [1-10], p: [1,2]", sz=13, color=MGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Hyperparameter Tuning:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "  • GridSearchCV with cv=10", sz=15, color=DGRAY)
_p(tf2, "  • Exhaustive search over parameter grids", sz=15, color=DGRAY)
_p(tf2, "  • Best params selected by CV accuracy", sz=15, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Evaluation Metrics:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "  • Cross-validation accuracy (10-fold)", sz=15, color=DGRAY)
_p(tf2, "  • Test set accuracy", sz=15, color=DGRAY)
_p(tf2, "  • Confusion matrix (True/False Pos/Neg)", sz=15, color=DGRAY)
_p(tf2, "  • Side-by-side model comparison bar chart", sz=15, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Tools: scikit-learn", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  GridSearchCV, StandardScaler,", sz=14, color=DGRAY)
_p(tf2, "  confusion_matrix, accuracy_score", sz=14, color=DGRAY)
_p(tf2, "", sz=6)
_p(tf2, f"Notebook: notebooks/07_modeling.ipynb", sz=13, color=ACCENT)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES 12–19 — EDA VISUALIZATION RESULTS  [Criterion 1.11]
# ═══════════════════════════════════════════════════════════════════════════════
eda_charts = [
    ("EDA Results — Flight Number vs Payload Mass (Scatter Plot)",
     "eda_flightnumber_vs_payloadmass.png",
     "Higher flight numbers correlate with more successful landings. Payload mass alone does not determine landing outcome."),
    ("EDA Results — Flight Number vs Launch Site (Scatter Plot)",
     "eda_flightnumber_vs_launchsite.png",
     "KSC LC-39A has the most launches and highest success concentration at higher flight numbers. VAFB SLC-4E used mainly for polar orbits."),
    ("EDA Results — Payload Mass vs Launch Site (Scatter Plot)",
     "eda_payloadmass_vs_launchsite.png",
     "KSC LC-39A handles the heaviest payloads (up to 15,600 kg). All sites show successful launches across payload ranges."),
    ("EDA Results — Flight Number vs Orbit Type (Scatter Plot)",
     "eda_flightnumber_vs_orbit.png",
     "LEO and ISS are the most common orbits in early flights. GTO missions appear throughout the timeline."),
    ("EDA Results — Payload Mass vs Orbit Type (Scatter Plot)",
     "eda_payloadmass_vs_orbit.png",
     "GTO and GEO orbits typically require higher payload masses. LEO/ISS orbits show wider payload mass ranges."),
    ("EDA Results — Success Rate by Orbit Type (Bar Chart)",
     "eda_success_rate_by_orbit.png",
     "ES-L1, GEO, HEO, and SSO orbits show 100% landing success. GTO has mixed results due to higher energy requirements."),
    ("EDA Results — Success Rate by Launch Site (Bar Chart)",
     "eda_success_rate_by_site.png",
     "KSC LC-39A leads with the highest success rate. All sites show improvement over time."),
    ("EDA Results — Yearly Success Trend (Line Chart)",
     "eda_yearly_success_trend.png",
     "Success rate improved from 0% in 2010 to over 80% by 2019–2020, demonstrating SpaceX's engineering learning curve."),
]

for title, img_file, caption in eda_charts:
    _img_slide(title, img_file, caption)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES 20–21 — EDA WITH SQL RESULTS  [Criterion 1.12]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("EDA with SQL Results — Launch Sites, Payloads, Success Rates")
tf = _tb(s, 0.5, 1.2, 6, 5.5, "", sz=14, color=DGRAY)
_p(tf, "Q1: Unique Launch Sites", sz=16, bold=True, color=DARK_BLUE)
_p(tf, "  CCAFS LC-40, CCAFS SLC-40, KSC LC-39A, VAFB SLC-4E", sz=13, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "Q2: Records with Launch Site LIKE 'CCA%'", sz=16, bold=True, color=DARK_BLUE)
_p(tf, "  Cape Canaveral sites returned (5 sample rows)", sz=13, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "Q3: NASA (CRS) Total Payload", sz=16, bold=True, color=DARK_BLUE)
_p(tf, "  Total: 45,596 kg across all CRS missions", sz=13, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "Q4: F9 v1.1 Average Payload Mass", sz=16, bold=True, color=DARK_BLUE)
_p(tf, "  Average: 2,928.4 kg", sz=13, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "Q5: First Successful Ground Pad Landing", sz=16, bold=True, color=DARK_BLUE)
_p(tf, "  Date: 2015-12-22 (Orbcomm-OG2 mission)", sz=13, color=DGRAY)

tf2 = _tb(s, 6.8, 1.2, 6, 5.5, "", sz=14, color=DGRAY)
_p(tf2, "Q6: Drone Ship Success (4000-6000 kg)", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  F9 FT B1021.2, F9 FT B1031.2", sz=13, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "Q7: Mission Outcome Counts", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  Success: 98  |  Failure: 1  |  Partial: 1", sz=13, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "Q8: Max Payload Boosters", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  F9 B5 series (Block 5 — latest generation)", sz=13, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "Q9: 2015 Drone Ship Failures", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  Jan 2015 (CRS-5) and Apr 2015 (CRS-6)", sz=13, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "Q10: Landing Outcome Rankings (2010–2017)", sz=16, bold=True, color=DARK_BLUE)
_p(tf2, "  No Attempt: 21 | Failure (drone ship): 9", sz=13, color=DGRAY)
_p(tf2, "  Success (drone ship): 5 | Failure (ASDS): 2", sz=13, color=DGRAY)
_p(tf2, "  Success (ground pad): 3 | Controlled (ocean): 5", sz=13, color=DGRAY)
_github_footer(s)

# Slide: SQL Time Analysis
s = _slide_title("EDA with SQL Results — Time Analysis & Rankings")
tf = _tb(s, 0.8, 1.2, 11.5, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Key Time-Based Insights from SQL Analysis:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=6)
_p(tf, "• First successful ground pad landing: December 22, 2015", sz=16, color=DGRAY)
_p(tf, "  This was the Orbcomm-OG2 mission from CCAFS LC-40 — a milestone in reusability.", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "• First successful drone ship landing: April 8, 2016", sz=16, color=DGRAY)
_p(tf, "  CRS-8 mission — opened up recovery for high-energy GTO launches.", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "• 2015 was the turning point: First landing successes after years of failed attempts.", sz=16, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "• By 2017: Success rate exceeded 50% for the first time.", sz=16, color=DGRAY)
_p(tf, "  Landing outcome distribution (2010–2017):", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "  ┌──────────────────────────┬───────┐", sz=13, color=DGRAY)
_p(tf, "  │ Landing Outcome          │ Count │", sz=13, color=DGRAY)
_p(tf, "  ├──────────────────────────┼───────┤", sz=13, color=DGRAY)
_p(tf, "  │ No Attempt               │  21   │", sz=13, color=DGRAY)
_p(tf, "  │ Failure (drone ship)     │   9   │", sz=13, color=DGRAY)
_p(tf, "  │ Controlled (ocean)       │   5   │", sz=13, color=DGRAY)
_p(tf, "  │ Success (drone ship)     │   5   │", sz=13, color=DGRAY)
_p(tf, "  │ Success (ground pad)     │   3   │", sz=13, color=DGRAY)
_p(tf, "  │ Failure (ADS)            │   2   │", sz=13, color=DGRAY)
_p(tf, "  │ Uncontrolled (ocean)     │   2   │", sz=13, color=DGRAY)
_p(tf, "  └──────────────────────────┴───────┘", sz=13, color=DGRAY)
_p(tf, "", sz=6)
_p(tf, "• By 2020: Landing success rate exceeded 80%, with Block 5 boosters being the most reliable.", sz=16, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES 22–24 — FOLIUM MAP RESULTS  [Criterion 1.13]
# ═══════════════════════════════════════════════════════════════════════════════
_img_slide("Folium Map Results — Launch Site Markers & Launch Records",
           "folium_launch_sites_overview.png",
           "All 4 SpaceX launch sites mapped with success/failure counts. Interactive version: output/launch_site_map.html")

_img_slide("Folium Map Results — KSC LC-39A Proximity Analysis",
           "folium_ksc_proximity.png",
           "KSC LC-39A is ~1.2 km from coast, ~7.1 km from highway, ~5.7 km from railway, ~52 km from nearest city. Interactive: output/ksc_proximity_map.html")

_img_slide("Folium Map Results — Success vs Failure by Launch Site",
           "folium_site_success_failure.png",
           "Success/failure breakdown per site. All sites are coastal for safety. KSC LC-39A has the most launches and highest success count.")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES 25–28 — PLOTLY DASH RESULTS  [Criterion 1.14]
# ═══════════════════════════════════════════════════════════════════════════════
_img_slide("Plotly Dash Results — Success Pie Chart (All Sites)",
           "dash_pie_all_sites.png",
           "Pie chart showing total successful launches distributed across all sites. KSC LC-39A contributes the most successes.")

_img_slide("Plotly Dash Results — Success vs Failure Pie Chart (KSC LC-39A)",
           "dash_pie_ksc.png",
           "KSC LC-39A: 77% success rate — the most reliable launch site in the dataset.")

_img_slide("Plotly Dash Results — Payload vs Outcome Scatter (All Sites)",
           "dash_scatter_all_sites.png",
           "Scatter plot of Payload Mass vs Launch Outcome colored by Booster Version. FT and B5 boosters show higher success across payload ranges.")

_img_slide("Plotly Dash Results — Payload vs Outcome Scatter (KSC LC-39A)",
           "dash_scatter_ksc.png",
           "KSC LC-39A: Successful launches across all payload ranges. B5 boosters dominate recent successful missions.")


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDES 29–34 — PREDICTIVE ANALYSIS RESULTS  [Criterion 1.15]
# ═══════════════════════════════════════════════════════════════════════════════
# Confusion matrices
_img_slide("Predictive Analysis Results — Logistic Regression Confusion Matrix",
           "cm_logistic_regression.png",
           "Best params: C=0.01, penalty=l2, solver=lbfgs  |  CV Score: 84.64%  |  Test Accuracy: 83.33%")

_img_slide("Predictive Analysis Results — SVM Confusion Matrix",
           "cm_svm.png",
           "Best params: kernel=sigmoid, C=1.0, gamma=0.032  |  CV Score: 84.82%  |  Test Accuracy: 83.33%")

_img_slide("Predictive Analysis Results — Decision Tree Confusion Matrix",
           "cm_decision_tree.png",
           "Best params: criterion=entropy, max_depth=14, splitter=random  |  CV Score: 88.75%  |  Test Accuracy: 83.33%")

_img_slide("Predictive Analysis Results — KNN Confusion Matrix",
           "cm_knn.png",
           "Best params: n_neighbors=10, algorithm=auto, p=1  |  CV Score: 84.82%  |  Test Accuracy: 83.33%")

# Model comparison
_img_slide("Predictive Analysis Results — Model Comparison",
           "model_comparison.png",
           "All 4 models achieve 83.33% test accuracy. Decision Tree has the highest CV score (88.75%) and is recommended as the best model.")

# Best model explanation slide
s = _slide_title("Predictive Analysis Results — Best Model & Evaluation")
tf = _tb(s, 0.5, 1.2, 6, 5.5, "", sz=14, color=DGRAY)
_p(tf, "Model Comparison Summary:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "┌─────────────────────┬──────────┬──────────┐", sz=13, color=DGRAY)
_p(tf, "│ Model               │ CV Score │ Test Acc │", sz=13, color=DGRAY)
_p(tf, "├─────────────────────┼──────────┼──────────┤", sz=13, color=DGRAY)
_p(tf, "│ Logistic Regression │  84.64%  │  83.33%  │", sz=13, color=DGRAY)
_p(tf, "│ SVM                 │  84.82%  │  83.33%  │", sz=13, color=DGRAY)
_p(tf, "│ Decision Tree       │  88.75%  │  83.33%  │", sz=13, bold=True, color=GREEN)
_p(tf, "│ KNN                 │  84.82%  │  83.33%  │", sz=13, color=DGRAY)
_p(tf, "└─────────────────────┴──────────┴──────────┘", sz=13, color=DGRAY)
_p(tf, "", sz=6)
_p(tf, "Best Model: Decision Tree", sz=20, bold=True, color=GREEN)
_p(tf, "  • Highest cross-validation score (88.75%)", sz=15, color=DGRAY)
_p(tf, "  • Best params: criterion=entropy, max_depth=14", sz=15, color=DGRAY)
_p(tf, "  • Easily interpretable (feature importance)", sz=15, color=DGRAY)
_p(tf, "  • All models tied on test accuracy (83.33%)", sz=15, color=DGRAY)
_p(tf, "  • Small test set (18 samples) limits differentiation", sz=15, color=DGRAY)

tf2 = _tb(s, 6.8, 1.2, 6, 5.5, "", sz=14, color=DGRAY)
_p(tf2, "Key Evaluation Insights:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=4)
_p(tf2, "• Confusion matrices show most errors are", sz=15, color=DGRAY)
_p(tf2, "  false negatives (predicting failure when", sz=15, color=DGRAY)
_p(tf2, "  actual outcome was success)", sz=15, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "• This is the safer error type for SpaceY:", sz=15, color=DGRAY)
_p(tf2, "  overestimating cost is better than", sz=15, color=DGRAY)
_p(tf2, "  underestimating it", sz=15, color=DGRAY)
_p(tf2, "", sz=6)
_p(tf2, "Recommendations:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  • Use Decision Tree as primary predictor", sz=15, color=DGRAY)
_p(tf2, "  • Ensemble methods (Random Forest, XGBoost)", sz=15, color=DGRAY)
_p(tf2, "    could improve accuracy further", sz=15, color=DGRAY)
_p(tf2, "  • Collect more data as SpaceX launches grow", sz=15, color=DGRAY)
_p(tf2, "  • Feature importance analysis reveals Flight", sz=15, color=DGRAY)
_p(tf2, "    Number and GridFins as top predictors", sz=15, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 35 — DISCUSSION
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Discussion")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Findings Summary:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "1. SpaceX landing success has improved dramatically", sz=15, color=DGRAY)
_p(tf, "   from 0% (2010) to >80% (2020), driven by", sz=14, color=MGRAY)
_p(tf, "   engineering iteration and Block 5 boosters.", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "2. Launch site matters: KSC LC-39A consistently", sz=15, color=DGRAY)
_p(tf, "   outperforms other sites in success rate.", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "3. Orbit type is a strong predictor: LEO, ISS,", sz=15, color=DGRAY)
_p(tf, "   SSO, ES-L1, GEO orbits have highest success.", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "4. Grid fins and landing legs are strong indicators", sz=15, color=DGRAY)
_p(tf, "   of landing capability (enabled on later missions).", sz=14, color=MGRAY)
_p(tf, "", sz=4)
_p(tf, "5. All ML models perform comparably (~83% accuracy)", sz=15, color=DGRAY)
_p(tf, "   with Decision Tree showing highest CV performance.", sz=14, color=MGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Limitations:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=4)
_p(tf2, "• Small dataset (90 samples, 18 test) limits", sz=15, color=DGRAY)
_p(tf2, "  statistical power and model differentiation.", sz=14, color=MGRAY)
_p(tf2, "", sz=4)
_p(tf2, "• Binary classification oversimplifies outcomes", sz=15, color=DGRAY)
_p(tf2, "  (controlled ocean landings classified as success).", sz=14, color=MGRAY)
_p(tf2, "", sz=4)
_p(tf2, "• Data only covers 2010-2020; recent Starship", sz=15, color=DGRAY)
_p(tf2, "  developments may change competitive dynamics.", sz=14, color=MGRAY)
_p(tf2, "", sz=6)
_p(tf2, "Future Work:", sz=20, bold=True, color=DARK_BLUE)
_p(tf2, "  • Add weather data as features", sz=15, color=DGRAY)
_p(tf2, "  • Use ensemble models (Random Forest, XGBoost)", sz=15, color=DGRAY)
_p(tf2, "  • Time-series analysis of landing capability", sz=15, color=DGRAY)
_p(tf2, "  • Cost-benefit analysis with prediction confidence", sz=15, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 36 — CONCLUSION  [Criterion 1.15 — creative conclusion + innovative insights]
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Conclusion & Innovative Insights")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf, "Conclusion:", sz=22, bold=True, color=DARK_BLUE)
_p(tf, "", sz=4)
_p(tf, "We successfully built an end-to-end data science", sz=15, color=DGRAY)
_p(tf, "pipeline that predicts Falcon 9 first stage landing", sz=15, color=DGRAY)
_p(tf, "outcomes with 83.33% accuracy.", sz=15, color=DGRAY)
_p(tf, "", sz=4)
_p(tf, "The Decision Tree classifier is recommended as the", sz=15, color=DGRAY)
_p(tf, "primary model (CV: 88.75%). This prediction enables", sz=15, color=DGRAY)
_p(tf, "SpaceY to estimate competitor launch costs and", sz=15, color=DGRAY)
_p(tf, "develop competitive pricing strategies.", sz=15, color=DGRAY)
_p(tf, "", sz=6)
_p(tf, "Business Impact:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, "  • Predicted landing → bid ~$62M (reusable)", sz=15, bold=True, color=GREEN)
_p(tf, "  • Predicted failure → bid ~$165M (expendable)", sz=15, bold=True, color=RED)
_p(tf, "  • Savings per correct prediction: ~$103M", sz=15, bold=True, color=DARK_BLUE)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=16, color=DGRAY)
_p(tf2, "Innovative Insights:", sz=22, bold=True, color=DARK_BLUE)
_p(tf2, "", sz=4)
_p(tf2, "1. Learning Curve Effect: SpaceX's landing", sz=15, bold=True, color=ACCENT)
_p(tf2, "   success follows a classic technology adoption", sz=14, color=DGRAY)
_p(tf2, "   S-curve, suggesting competitors can replicate", sz=14, color=DGRAY)
_p(tf2, "   this with sufficient investment and iteration.", sz=14, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "2. Grid Fins as Innovation Marker: The", sz=15, bold=True, color=ACCENT)
_p(tf2, "   introduction of grid fins correlates strongly", sz=14, color=DGRAY)
_p(tf2, "   with landing success improvement, representing", sz=14, color=DGRAY)
_p(tf2, "   a key engineering breakthrough.", sz=14, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "3. Geographic Strategy: All launch sites are", sz=15, bold=True, color=ACCENT)
_p(tf2, "   coastal and near the equator — a pattern SpaceY", sz=14, color=DGRAY)
_p(tf2, "   should replicate for its own launch facilities.", sz=14, color=DGRAY)
_p(tf2, "", sz=4)
_p(tf2, "4. Payload-Orbit Optimization: Matching payload", sz=15, bold=True, color=ACCENT)
_p(tf2, "   mass to favorable orbits (LEO, SSO) significantly", sz=14, color=DGRAY)
_p(tf2, "   increases the probability of first stage recovery.", sz=14, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SLIDE 37 — APPENDIX
# ═══════════════════════════════════════════════════════════════════════════════
s = _slide_title("Appendix")
tf = _tb(s, 0.8, 1.2, 5.8, 5.5, "", sz=14, color=DGRAY)
_p(tf, "GitHub Repository:", sz=20, bold=True, color=DARK_BLUE)
_p(tf, f"{GITHUB_URL}", sz=15, color=ACCENT)
_p(tf, "", sz=8)
_p(tf, "Repository Structure:", sz=18, bold=True, color=DARK_BLUE)
_p(tf, "  spacey-capstone/", sz=13, color=DGRAY)
_p(tf, "  ├── data/          CSVs, SQLite database", sz=13, color=DGRAY)
_p(tf, "  ├── notebooks/     7 Jupyter notebooks (01-07)", sz=13, color=DGRAY)
_p(tf, "  ├── scripts/       Helper scripts", sz=13, color=DGRAY)
_p(tf, "  ├── sql/           SQL query scripts", sz=13, color=DGRAY)
_p(tf, "  ├── images/        Charts & confusion matrices", sz=13, color=DGRAY)
_p(tf, "  ├── dash/          Plotly Dash application", sz=13, color=DGRAY)
_p(tf, "  ├── output/        Folium HTML maps", sz=13, color=DGRAY)
_p(tf, "  ├── slides/        PowerPoint & PDF", sz=13, color=DGRAY)
_p(tf, "  ├── requirements.txt", sz=13, color=DGRAY)
_p(tf, "  └── README.md", sz=13, color=DGRAY)

tf2 = _tb(s, 7, 1.2, 5.8, 5.5, "", sz=14, color=DGRAY)
_p(tf2, "Notebooks:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  01_spacex_api_collection.ipynb", sz=13, color=DGRAY)
_p(tf2, "  02_web_scraping.ipynb", sz=13, color=DGRAY)
_p(tf2, "  03_data_wrangling.ipynb", sz=13, color=DGRAY)
_p(tf2, "  04_eda.ipynb", sz=13, color=DGRAY)
_p(tf2, "  05_sql_analysis.ipynb", sz=13, color=DGRAY)
_p(tf2, "  06_interactive_visuals.ipynb", sz=13, color=DGRAY)
_p(tf2, "  07_modeling.ipynb", sz=13, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Tech Stack:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  Python, Pandas, NumPy, Matplotlib, Seaborn", sz=13, color=DGRAY)
_p(tf2, "  Scikit-learn, BeautifulSoup, Requests", sz=13, color=DGRAY)
_p(tf2, "  Folium, Plotly Dash, SQLite, Jupyter", sz=13, color=DGRAY)
_p(tf2, "", sz=8)
_p(tf2, "Additional Charts:", sz=18, bold=True, color=DARK_BLUE)
_p(tf2, "  • All 8 EDA charts saved in images/", sz=13, color=DGRAY)
_p(tf2, "  • 4 confusion matrices in images/", sz=13, color=DGRAY)
_p(tf2, "  • Model comparison bar chart in images/", sz=13, color=DGRAY)
_p(tf2, "  • 4 Plotly Dash static charts in images/", sz=13, color=DGRAY)
_p(tf2, "  • 3 Folium map visualizations in images/", sz=13, color=DGRAY)
_github_footer(s)


# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════
output_path = os.path.join(OUT, "Data Science Capstone Project Report.pptx")
prs.save(output_path)
print(f"Saved {len(prs.slides)} slides to {output_path}")
print("\nGrading criteria coverage:")
print("  1.1  GitHub URL — on cover + every methodology + appendix slide")
print("  1.2  PDF format — export this PPTX as PDF manually")
print("  1.3  Executive Summary — Slide 2")
print("  1.4  Introduction — Slide 4")
print("  1.5  Data Collection SpaceX API — Slide 5 (flowchart + GitHub URL)")
print("  1.6  Data Collection Web Scraping — Slide 6 (flowchart + GitHub URL)")
print("  1.7  Data Wrangling — Slide 7 (cleaning + processing + GitHub URL)")
print("  1.8  EDA Visualization methodology — Slide 8 (charts + purpose + GitHub URL)")
print("  1.9  EDA SQL methodology — Slide 9 (queries + GitHub URL)")
print("  1.10 Interactive Visual Analytics — Slide 10 (Folium + Dash + GitHub URL)")
print("  1.11 EDA visualization results — Slides 12-19 (scatter, bar, yearly)")
print("  1.12 EDA SQL results — Slides 20-21 (sites, payloads, rates, rankings, time)")
print("  1.13 Folium map results — Slides 22-24 (markers, records, proximity)")
print("  1.14 Plotly Dash results — Slides 25-28 (pie charts, scatter plots)")
print("  1.15 Predictive Analysis — Slides 29-36 (models, CMs, best model, conclusion)")

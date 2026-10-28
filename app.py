"""Study of Nutrient Contents in Breakfast Items — single-file Flask application.

Run locally:
    python app.py

The application creates instance/survey.db automatically. Public survey results
are aggregate-only. The /admin dashboard is intentionally password-free and is
not linked from any public page.

Charts use Chart.js from a pinned CDN. An internet connection is required for
the chart library; readable fallback tables are shown when it is unavailable.
"""
from __future__ import annotations

import base64
import binascii
import csv
import hashlib
import hmac
import io
import os
import re
import secrets
import sqlite3
import time
from datetime import date, datetime, time as dt_time, timedelta, timezone
from pathlib import Path
from urllib.parse import urlencode

from flask import (
    Flask,
    Response,
    abort,
    jsonify,
    redirect,
    render_template_string,
    request,
    send_file,
    session,
    url_for,
)
from flask_limiter import Limiter
from flask_limiter.util import get_remote_address
from markupsafe import Markup
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from reportlab.lib import colors as rl_colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)
from xml.sax.saxutils import escape as xml_escape


BASE_DIR = Path(__file__).resolve().parent
INSTANCE_DIR = BASE_DIR / "instance"
INSTANCE_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = Path(os.environ.get("SURVEY_DB_PATH", INSTANCE_DIR / "survey.db"))
CHART_CDN = "https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"


def load_or_create_secret_key() -> str:
    configured = os.environ.get("SECRET_KEY")
    if configured:
        return configured
    key_path = INSTANCE_DIR / ".secret_key"
    if key_path.exists():
        value = key_path.read_text(encoding="utf-8").strip()
        if value:
            return value
    value = secrets.token_hex(32)
    key_path.write_text(value, encoding="utf-8")
    return value


app = Flask(__name__)
app.config.update(
    SECRET_KEY=load_or_create_secret_key(),
    MAX_CONTENT_LENGTH=64 * 1024,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("SESSION_COOKIE_SECURE") == "1",
    RATELIMIT_ENABLED=os.environ.get("RATELIMIT_ENABLED", "1") != "0",
    RATELIMIT_HEADERS_ENABLED=True,
)
app.jinja_env.autoescape = True

limiter = Limiter(
    key_func=get_remote_address,
    app=app,
    default_limits=[],
    storage_uri=os.environ.get("RATELIMIT_STORAGE_URI", "memory://"),
)

PROJECT = {
    "brand": "NutriLab",
    "title": "Study of Nutrient Contents in Breakfast Items",
    "subtitle": "A Qualitative Chemical & Bio-Nutritional Analysis",
    "eyebrow": "CBSE Chemistry Investigatory Project",
    "class_label": "Class XII – Science",
    "year": "2026–2027",
    "submitted_by": ["R Adhithya", "Darshini K", "Kaivalya P"],
    "source_pdf": "docs/CHEMISTRY-IP-ROUGH-DRAFT.pdf",
    "source_pdf_label": "CHEMISTRY IP ROUGH DRAFT.pdf",
    "footer_lede": (
        "A qualitative chemical and bio-nutritional analysis of four common breakfast "
        "items — built from the submitted Class XII investigatory project."
    ),
    "footer_meta": "CBSE Chemistry Investigatory Project · Class XII – Science · 2026–2027",
}


def slugify(value: str) -> str:
    value = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return value or "item"


def radio_options(values: list[str]) -> list[dict]:
    return [{"value": value, "label": value} for value in values]


CLASS_OPTIONS = ["Class IX", "Class X", "Class XI", "Class XII", "Other"]
FREQUENCY_OPTIONS = [
    "Every day",
    "4–6 days a week",
    "1–3 days a week",
    "Rarely",
    "Never",
]
PREFERRED_OPTIONS = ["Idli", "Multi Grain Porridge", "Muesli", "Bread with Cheese", "Other"]
REASON_OPTIONS = [
    "Taste",
    "Nutritional value",
    "Convenience",
    "Cost",
    "Family preference",
    "Availability",
    "Other",
]
EATING_TIME_OPTIONS = [
    "Less than 5 minutes",
    "5–10 minutes",
    "10–20 minutes",
    "More than 20 minutes",
    "I usually skip breakfast",
]
NUTRIENT_OPTIONS = ["Carbohydrates", "Proteins", "Fats", "Calcium", "Iron", "All are important"]
FAMILIARITY_OPTIONS = [
    "Very familiar",
    "Somewhat familiar",
    "Have heard about them",
    "Not familiar",
]
YES_NO_NOT_SURE_OPTIONS = ["Yes", "No", "Not sure"]
GREASE_OPTIONS = ["Idli", "Multi Grain Porridge", "Muesli", "Bread with Cheese", "Not sure"]
IODINE_KNOWLEDGE_OPTIONS = ["Starch", "Protein", "Fat", "Calcium", "Not sure"]
LEARNING_OPTIONS = [
    "Interactive experiments",
    "Infographics and graphs",
    "Short videos",
    "Classroom demonstrations",
    "Interesting facts about everyday foods",
    "Other",
]
INFLUENCE_OPTIONS = ["Definitely", "Probably", "Not sure", "Probably not", "Definitely not"]
INTEREST_OPTIONS = [
    "Testing for starch and sugars",
    "Identifying proteins",
    "Testing for fats",
    "Detecting calcium and iron",
    "Comparing all four breakfast items",
    "Seeing the chemical reactions",
]

PREFERRED_IMAGE_OPTIONS = [
    {
        "value": "Idli",
        "label": "Idli",
        "image": "images/sample-idli.jpg",
        "image_webp": "images/sample-idli.webp",
        "alt": "Idli served with sambar and chutney.",
    },
    {
        "value": "Multi Grain Porridge",
        "label": "Multi Grain Porridge",
        "image": "images/sample-porridge.jpg",
        "image_webp": "images/sample-porridge.webp",
        "alt": "A warm bowl of whole-grain porridge with fruit.",
    },
    {
        "value": "Muesli",
        "label": "Muesli",
        "image": "images/sample-muesli.jpg",
        "image_webp": "images/sample-muesli.webp",
        "alt": "A bowl of muesli with berries, seeds and fruit.",
    },
    {
        "value": "Bread with Cheese",
        "label": "Bread with Cheese",
        "image": "images/sample-bread-cheese.jpg",
        "image_webp": "images/sample-bread-cheese.webp",
        "alt": "Sliced bread served with wedges of cheese.",
    },
    {
        "value": "Other",
        "label": "Other",
        "image": None,
        "image_webp": None,
        "alt": "",
    },
]

SURVEY_SECTIONS = [
    {
        "id": "about",
        "title": "About you",
        "description": (
            "These details help the project team understand who responded. Only the project "
            "administrator can see them; they never appear in public charts."
        ),
        "questions": [
            {
                "number": 1,
                "field": "name",
                "type": "text",
                "label": "What is your name?",
                "help": "Used only by the project administrator. It is never shown in public results.",
                "required": True,
                "maxlength": 80,
                "autocomplete": "name",
                "placeholder": "Your full name",
                "full_width": True,
            },
            {
                "number": 2,
                "field": "class_level",
                "type": "radio",
                "label": "What is your class?",
                "help": "Choose the class you are currently studying in.",
                "required": True,
                "options": radio_options(CLASS_OPTIONS),
            },
            {
                "number": 3,
                "field": "email",
                "type": "email",
                "label": "What is your email address?",
                "help": (
                    "We collect your email only so the project team can contact you about this "
                    "survey if needed. It is never shown publicly."
                ),
                "required": True,
                "maxlength": 254,
                "autocomplete": "email",
                "placeholder": "you@example.com",
                "full_width": True,
            },
        ],
    },
    {
        "id": "habits",
        "title": "Breakfast habits",
        "description": "Tell us what a normal school morning looks like for you.",
        "questions": [
            {
                "number": 4,
                "field": "breakfast_frequency",
                "type": "radio",
                "label": "How often do you eat breakfast?",
                "help": "Think about a typical school week.",
                "required": True,
                "options": radio_options(FREQUENCY_OPTIONS),
            },
            {
                "number": 5,
                "field": "preferred_breakfast",
                "type": "image-radio",
                "label": "Which breakfast would you choose on a normal school morning?",
                "help": "Choose the option closest to your usual choice.",
                "required": True,
                "options": PREFERRED_IMAGE_OPTIONS,
                "full_width": True,
            },
            {
                "number": 6,
                "field": "main_reason",
                "type": "radio",
                "label": "What matters most when choosing your breakfast?",
                "help": "Pick the single factor that matters most to you.",
                "required": True,
                "options": radio_options(REASON_OPTIONS),
            },
            {
                "number": 7,
                "field": "eating_time",
                "type": "radio",
                "label": "How much time do you usually spend eating breakfast?",
                "help": "Include the time you spend sitting down to eat.",
                "required": True,
                "options": radio_options(EATING_TIME_OPTIONS),
            },
        ],
    },
    {
        "id": "awareness",
        "title": "Nutritional awareness",
        "description": (
            "There are no trick questions here. Choose what you believe, even if you are not sure."
        ),
        "questions": [
            {
                "number": 8,
                "field": "important_nutrient",
                "type": "radio",
                "label": "Which nutrient do you think is most important for a balanced breakfast?",
                "help": "Choose the answer that best matches your understanding.",
                "required": True,
                "options": radio_options(NUTRIENT_OPTIONS),
            },
            {
                "number": 9,
                "field": "test_familiarity",
                "type": "radio",
                "label": "How familiar are you with chemical tests used to identify nutrients in food?",
                "help": "This helps us understand how familiar students are with food chemistry.",
                "required": True,
                "options": radio_options(FAMILIARITY_OPTIONS),
            },
            {
                "number": 10,
                "field": "iodine_awareness",
                "type": "radio",
                "label": "Have you heard of the iodine test for starch?",
                "help": "Answer based on what you already know.",
                "required": True,
                "options": radio_options(YES_NO_NOT_SURE_OPTIONS),
            },
            {
                "number": 11,
                "field": "perceived_grease_spot",
                "type": "radio",
                "label": (
                    "Which breakfast item do you think produces the strongest positive result in "
                    "the grease-spot test according to our project's observations?"
                ),
                "help": (
                    "Project observation: Bread with Cheese produced a strong persistent grease "
                    "spot. This question records what you think — it is separate from the "
                    "documented project result."
                ),
                "required": True,
                "options": radio_options(GREASE_OPTIONS),
                "project_note": (
                    "The project's qualitative observation was a strong persistent grease spot "
                    "for Bread with Cheese. Your answer below is your own prediction."
                ),
            },
            {
                "number": 12,
                "field": "iodine_knowledge",
                "type": "radio",
                "label": "What do you think a blue-black colour in the iodine test indicates?",
                "help": "Think about the chemistry you have seen in this project.",
                "required": True,
                "options": radio_options(IODINE_KNOWLEDGE_OPTIONS),
            },
            {
                "number": 13,
                "field": "learning_method",
                "type": "radio",
                "label": "What would most encourage you to learn more about food nutrients?",
                "help": "Choose the learning style that would help you most.",
                "required": True,
                "options": radio_options(LEARNING_OPTIONS),
            },
        ],
    },
    {
        "id": "opinion",
        "title": "Opinion & impact",
        "description": "Your view helps the project team understand how the project lands with students.",
        "questions": [
            {
                "number": 14,
                "field": "choice_influence",
                "type": "radio",
                "label": "Would learning about nutrients in breakfast foods influence your food choices?",
                "help": "Be honest — there is no expected answer.",
                "required": True,
                "options": radio_options(INFLUENCE_OPTIONS),
            },
            {
                "number": 15,
                "field": "interesting_topic",
                "type": "radio",
                "label": "Which part of this Chemistry IP project interests you most?",
                "help": "Pick the part you would most like to explore further.",
                "required": True,
                "options": radio_options(INTEREST_OPTIONS),
            },
        ],
    },
    {
        "id": "voice",
        "title": "Your voice",
        "description": "Optional. Share anything that could help improve this project.",
        "questions": [
            {
                "number": 16,
                "field": "comments",
                "type": "textarea",
                "label": "Any suggestions or comments?",
                "help": (
                    "Optional. Please do not include phone numbers, addresses or other sensitive "
                    "personal information."
                ),
                "required": False,
                "maxlength": 500,
                "placeholder": "Share a suggestion, a question or an observation…",
                "full_width": True,
            },
        ],
    },
]

SURVEY_QUESTIONS = [question for section in SURVEY_SECTIONS for question in section["questions"]]
SURVEY_FIELDS = [question["field"] for question in SURVEY_QUESTIONS]
SURVEY_QUESTION_BY_FIELD = {question["field"]: question for question in SURVEY_QUESTIONS}

HEADER_NAV = [
    {"label": "Overview", "slug": "overview"},
    {"label": "Aim & Objectives", "slug": "aim"},
    {"label": "Introduction", "slug": "introduction"},
    {"label": "Food Samples", "slug": "samples"},
    {"label": "Materials", "slug": "materials"},
    {"label": "Chemical Tests", "slug": "tests"},
    {"label": "Results", "slug": "results"},
    {"label": "Conclusion", "slug": "conclusion"},
    {"label": "Survey", "endpoint": "survey"},
    {"label": "Survey Results", "endpoint": "survey_results"},
]

SIDEBAR_GROUPS = [
    {
        "label": "Project",
        "items": [
            {"label": "Overview", "slug": "overview"},
            {"label": "Aim & Objectives", "slug": "aim"},
            {"label": "Introduction", "slug": "introduction"},
        ],
    },
    {
        "label": "Investigation",
        "items": [
            {"label": "Food Samples", "slug": "samples"},
            {"label": "Materials Required", "slug": "materials"},
            {"label": "Chemicals & Reagents", "slug": "chemicals"},
            {"label": "Preparation of Samples", "slug": "preparation"},
        ],
    },
    {
        "label": "Findings",
        "items": [
            {"label": "Chemical Tests", "slug": "tests"},
            {"label": "Comparative Results", "slug": "results"},
            {"label": "Comprehensive Conclusion", "slug": "conclusion"},
            {"label": "Precautions & Safety", "slug": "safety"},
        ],
    },
    {
        "label": "Participate",
        "items": [
            {"label": "Survey", "endpoint": "survey"},
            {"label": "Survey Results", "endpoint": "survey_results"},
        ],
    },
    {
        "label": "About",
        "items": [
            {"label": "Project Information", "slug": "project_info"},
        ],
    },
]


APPARATUS = [
    "Clean test tubes",
    "Test tube stands",
    "Test tube holders",
    "Boiling tubes",
    "Silica crucible",
    "Clay pipe triangle",
    "Tripod stand",
    "Bunsen burner setup",
    "Regulated water bath with temperature monitoring",
    "Mortar and pestle",
    "Glass funnels",
    "Whatman No. 42 filter papers",
    "Glass stirring rods",
    "Volumetric droppers",
    "100 mL measuring cylinder",
    "Analytical balance",
    "Calibrated pH paper or pH meter",
]
APPARATUS_ITEMS = [
    {"id": slugify(name), "name": name, "kind": "apparatus"} for name in APPARATUS
]

CHEMICALS = [
    {
        "name": "Distilled or deionized water",
        "formula": "",
        "hazard": "",
        "hazard_class": "",
        "note": "Used for extraction and for avoiding cross-contamination.",
    },
    {
        "name": "Iodine solution",
        "formula": "0.1 N in aqueous KI",
        "hazard": "Irritant",
        "hazard_class": "irritant",
        "note": "Stains skin and clothing; use fresh reagent.",
    },
    {
        "name": "Benedict's qualitative reagent",
        "formula": "CuSO₄ · sodium citrate · Na₂CO₃",
        "hazard": "Irritant",
        "hazard_class": "irritant",
        "note": "Alkaline copper(II) reagent for reducing sugars.",
    },
    {
        "name": "Concentrated nitric acid",
        "formula": "HNO₃",
        "hazard": "Corrosive",
        "hazard_class": "corrosive",
        "note": "Corrosive and oxidising; use a fume hood.",
    },
    {
        "name": "Concentrated ammonium hydroxide",
        "formula": "NH₄OH",
        "hazard": "Corrosive",
        "hazard_class": "corrosive",
        "note": "Volatile base; add slowly and avoid inhalation.",
    },
    {
        "name": "Sodium hydroxide",
        "formula": "10% NaOH",
        "hazard": "Corrosive",
        "hazard_class": "corrosive",
        "note": "Strong base; avoid contact with skin and eyes.",
    },
    {
        "name": "Lead acetate solution",
        "formula": "10% Pb(CH₃COO)₂",
        "hazard": "Toxic",
        "hazard_class": "toxic",
        "note": "Toxic lead-containing reagent; collect waste separately.",
    },
    {
        "name": "Ammonium oxalate solution",
        "formula": "(NH₄)₂C₂O₄",
        "hazard": "Harmful",
        "hazard_class": "harmful",
        "note": "Harmful oxalate solution; handle with care.",
    },
    {
        "name": "Potassium thiocyanate",
        "formula": "KSCN",
        "hazard": "Harmful",
        "hazard_class": "harmful",
        "note": "Harmful irritant reagent for the iron test.",
    },
    {
        "name": "Dilute hydrochloric acid",
        "formula": "2 M HCl",
        "hazard": "Corrosive",
        "hazard_class": "corrosive",
        "note": "Corrosive acid; dilute carefully.",
    },
    {
        "name": "Dilute nitric acid",
        "formula": "2 M HNO₃",
        "hazard": "Corrosive",
        "hazard_class": "corrosive",
        "note": "Corrosive acid; use appropriate PPE.",
    },
]
for item in CHEMICALS:
    item["id"] = slugify(item["name"])
    item["kind"] = "chemical"

SAMPLES = [
    {
        "slug": "idli",
        "index": "01",
        "name": "Idli",
        "tag": "Fermented · Traditional",
        "description": (
            "A traditional South Indian breakfast made from a fermented batter of rice and "
            "black gram."
        ),
        "image": "images/sample-idli.jpg",
        "image_webp": "images/sample-idli.webp",
        "alt": "Idli served with sambar and chutney on a dark table.",
        "findings": [
            {"label": "Strong starch", "level": "strong"},
            {"label": "Fat negative", "level": "negative"},
            {"label": "Moderate protein", "level": "moderate"},
            {"label": "pH 5.8–6.2", "level": "ph"},
        ],
        "detail": (
            "The fermented rice-and-gram batter produced a clear extract with a strong iodine "
            "response. Fermentation lowered the pH slightly, while the grease-spot test showed no "
            "visible lipid stain. The sample is a useful example of a low-fat, starch-rich "
            "traditional breakfast."
        ),
        "nutrition": [
            "Rice contributes starch and some reducing sugars; black gram adds protein and dietary fibre.",
            "Fermentation produces lactic acid, which is consistent with the mildly acidic pH range.",
            "The low fat signal does not mean the food is nutritionally empty; it simply describes the qualitative lipid response.",
        ],
        "observations": [
            "Iodine: strong blue-black colour.",
            "Benedict's: slight positive response.",
            "Grease spot: negative.",
            "Heat coagulation: moderate positive.",
        ],
    },
    {
        "slug": "porridge",
        "index": "02",
        "name": "Multi Grain Porridge",
        "tag": "Whole grain · Fibre-rich",
        "description": (
            "A warm, whole-grain breakfast preparation built around complex carbohydrates and "
            "dietary fibre."
        ),
        "image": "images/sample-porridge.jpg",
        "image_webp": "images/sample-porridge.webp",
        "alt": "A warm whole-grain porridge served in bowls with milk and fruit.",
        "findings": [
            {"label": "Strong starch", "level": "strong"},
            {"label": "Moderate sugars", "level": "moderate"},
            {"label": "Slight fat trace", "level": "slight"},
            {"label": "Strong iron", "level": "strong"},
        ],
        "detail": (
            "The porridge extract gave a strong starch response with a moderate reducing-sugar "
            "signal. The grease-spot test showed only a slight trace, while the ash test produced "
            "a strong iron response. This makes the sample a strong example of a nutrient-dense, "
            "low-fat whole-grain breakfast."
        ),
        "nutrition": [
            "Whole grains supply complex carbohydrates, fibre and a range of minerals.",
            "The moderate reducing-sugar response reflects some free sugars released during preparation.",
            "The strong iron signal is a useful contrast to the lower iron response of Idli.",
        ],
        "observations": [
            "Iodine: strong blue-black colour.",
            "Benedict's: moderate positive response.",
            "Grease spot: slight trace spot.",
            "Iron test: strong blood-red complex.",
        ],
    },
    {
        "slug": "muesli",
        "index": "03",
        "name": "Muesli",
        "tag": "Processed cereal · Grain-based",
        "description": (
            "A grain-based breakfast cereal, often combined with milk, fruit, nuts and seeds."
        ),
        "image": "images/sample-muesli.jpg",
        "image_webp": "images/sample-muesli.webp",
        "alt": "A bowl of muesli with fruit, seeds and berries.",
        "findings": [
            {"label": "Moderate starch", "level": "moderate"},
            {"label": "Strong sugars", "level": "strong"},
            {"label": "Moderate fat", "level": "moderate"},
            {"label": "Strong iron", "level": "strong"},
        ],
        "detail": (
            "Muesli produced the strongest reducing-sugar response of the four samples, alongside "
            "a moderate starch signal and a moderate grease spot. The ash test gave a strong iron "
            "response. The result is consistent with a grain-and-fruit cereal that can be "
            "nutrient-dense but may also carry more simple sugars."
        ),
        "nutrition": [
            "Oats and other grains contribute starch, fibre and minerals.",
            "Dried fruit and added sweeteners can increase the reducing-sugar response.",
            "Nuts and seeds contribute lipids, which is consistent with the moderate grease spot.",
        ],
        "observations": [
            "Iodine: moderate positive response.",
            "Benedict's: strong positive response.",
            "Grease spot: moderate spot.",
            "Iron test: strong blood-red complex.",
        ],
    },
    {
        "slug": "bread",
        "index": "04",
        "name": "Bread with Cheese",
        "tag": "Dairy-rich · Bakery",
        "description": (
            "Bread combined with dairy cheese — a pairing that brings carbohydrates, protein, "
            "lipids and calcium together."
        ),
        "image": "images/sample-bread-cheese.jpg",
        "image_webp": "images/sample-bread-cheese.webp",
        "alt": "Sliced bread served with wedges of cheese on a wooden board.",
        "findings": [
            {"label": "Strong starch", "level": "strong"},
            {"label": "Strong fat", "level": "strong"},
            {"label": "Strong protein", "level": "strong"},
            {"label": "Strong calcium", "level": "strong"},
        ],
        "detail": (
            "Bread with Cheese gave the strongest grease spot and the strongest protein and "
            "calcium responses. The bread supplied a strong starch signal, while the cheese "
            "contributed lipid, sulfur-containing protein and calcium. This is the most "
            "lipid- and protein-dense sample in the comparison."
        ),
        "nutrition": [
            "Bread is a source of refined starch and some protein.",
            "Cheese contributes fat, protein, calcium and sulfur-containing amino acids.",
            "The strong fat response is a qualitative observation, not a measured fat content.",
        ],
        "observations": [
            "Iodine: strong blue-black colour.",
            "Benedict's: slight positive response.",
            "Grease spot: strong persistent spot.",
            "Calcium test: strong white precipitate.",
        ],
    },
]

TESTS = [
    {
        "slug": "iodine",
        "index": "01",
        "name": "Iodine Test for Starch",
        "target": "Detects starch / polysaccharides",
        "icon": "starch",
        "materials": (
            "2 mL clear aqueous extract; 2–3 drops freshly prepared iodine solution (0.1 N "
            "iodine in aqueous potassium iodide); clean test tube."
        ),
        "procedure": (
            "Take 2 mL of clear aqueous extract in a clean test tube. Add 2–3 drops of freshly "
            "prepared iodine solution. Shake gently to mix."
        ),
        "observation": (
            "Immediate development of a deep blue-black colour confirms the presence of starch."
        ),
        "basis": (
            "Polyiodide ions (I₃⁻ and I₅⁻) slide into the inner core of helical amylose chains in "
            "starch, forming a dark blue charge-transfer complex."
        ),
        "safety": (
            "Iodine solution can stain skin and clothing and may irritate. Handle it carefully "
            "and avoid contact with eyes."
        ),
        "danger": False,
    },
    {
        "slug": "benedicts",
        "index": "02",
        "name": "Benedict's Test for Reducing Sugars",
        "target": "Detects reducing sugars",
        "icon": "sugar",
        "materials": (
            "2 mL extract; 2 mL Benedict's Qualitative Reagent; boiling tube; boiling water bath."
        ),
        "procedure": (
            "Mix 2 mL of extract with 2 mL of Benedict's Qualitative Reagent in a boiling tube. "
            "Place in a boiling water bath for 3 to 5 minutes."
        ),
        "observation": (
            "A colour transition from blue to green, yellow, orange, or brick-red precipitate "
            "indicates reducing sugars."
        ),
        "basis": (
            "Alkaline cupric ions (Cu²⁺) are reduced by free aldehyde/ketone groups of sugars to "
            "insoluble cuprous oxide (Cu₂O), forming a brick-red precipitate."
        ),
        "safety": (
            "Use a test-tube holder and keep the boiling tube pointed away from yourself and "
            "others. The water bath and tube are hot."
        ),
        "danger": False,
    },
    {
        "slug": "grease",
        "index": "03",
        "name": "Grease Spot Test for Fats & Oils",
        "target": "Detects lipids",
        "icon": "fat",
        "materials": (
            "A small portion of solid sample or a drop of lipid extract; filter paper; low flame."
        ),
        "procedure": (
            "Press a small portion of solid sample or a drop of lipid extract onto filter paper. "
            "Warm gently over a low flame to evaporate moisture."
        ),
        "observation": (
            "Formation of a persistent, non-drying translucent spot confirms the presence of "
            "lipids."
        ),
        "basis": (
            "Non-polar triglycerides penetrate paper fibres and permanently alter light "
            "scattering."
        ),
        "safety": (
            "Keep the flame low and away from clothing and hair. Use tongs and never leave the "
            "flame unattended."
        ),
        "danger": False,
    },
    {
        "slug": "coagulation",
        "index": "04",
        "name": "Heat Coagulation Test for Proteins",
        "target": "Detects denaturable proteins",
        "icon": "protein",
        "materials": "Clear extract; test tube; flame.",
        "procedure": (
            "Fill a test tube two-thirds full with clear extract. Heat the upper portion over a "
            "flame until boiling."
        ),
        "observation": (
            "Turbidity or coagulation in the upper layer confirms the presence of denaturable "
            "proteins."
        ),
        "basis": (
            "Heat breaks intra-molecular hydrogen bonds, unfolding protein chains and causing "
            "aggregation."
        ),
        "safety": (
            "Heat gently and point the mouth of the tube away from people to prevent bumping or "
            "spitting."
        ),
        "danger": False,
    },
    {
        "slug": "ph",
        "index": "05",
        "name": "pH Determination",
        "target": "Measures acidity / alkalinity",
        "icon": "ph",
        "materials": (
            "5 mL extract; universal pH paper or calibrated digital pH probe; standard colour "
            "chart."
        ),
        "procedure": (
            "Dip a strip of universal pH paper or a calibrated digital probe into 5 mL of "
            "extract. Compare the colour against the standard chart."
        ),
        "observation": (
            "Identifies acidic or alkaline properties. For example, fermented Idli registers pH "
            "5.8–6.2 due to lactic fermentation."
        ),
        "basis": (
            "pH paper indicators change colour in response to hydrogen-ion concentration; a "
            "calibrated digital probe measures the corresponding electrode potential. The project "
            "treats pH as a physical parameter."
        ),
        "safety": (
            "Handle glass probes carefully. Clean the probe between samples to avoid "
            "cross-contamination."
        ),
        "danger": False,
    },
    {
        "slug": "xanthoproteic",
        "index": "06",
        "name": "Xanthoproteic Test for Aromatic Amino Acids",
        "target": "Detects tyrosine & tryptophan",
        "icon": "aromatic",
        "materials": "2 mL extract; 2 mL concentrated HNO₃; excess concentrated NH₄OH.",
        "procedure": (
            "Add 2 mL conc. HNO₃ to 2 mL extract. Heat gently, cool, then carefully add excess "
            "conc. NH₄OH along the sides."
        ),
        "observation": (
            "White precipitate turns yellow on heating and orange with NH₄OH, confirming "
            "aromatic amino acids (tyrosine, tryptophan)."
        ),
        "basis": (
            "Nitration of aromatic benzene rings yields yellow nitrated derivatives, converting "
            "to deep orange ammonium salts in basic solution."
        ),
        "safety": (
            "Concentrated HNO₃ and NH₄OH are corrosive. Use safety goggles, gloves and a fume "
            "hood; add NH₄OH slowly and carefully."
        ),
        "danger": True,
    },
    {
        "slug": "lead",
        "index": "07",
        "name": "Lead Acetate Test for Sulfur-Containing Proteins",
        "target": "Detects cysteine & methionine",
        "icon": "sulfur",
        "materials": (
            "2 mL extract; 2 mL 10% NaOH; 4–5 drops 10% lead acetate solution; boiling tube."
        ),
        "procedure": (
            "Mix 2 mL extract with 2 mL 10% NaOH in a boiling tube and boil for 2 minutes. Add "
            "4–5 drops of 10% lead acetate solution."
        ),
        "observation": (
            "Formation of a black precipitate confirms sulfur amino acids (cysteine, methionine)."
        ),
        "basis": (
            "Boiling with alkali hydrolyses organic sulfur into inorganic sulfide ions (S²⁻), "
            "forming insoluble black lead sulfide (PbS)."
        ),
        "safety": (
            "Lead acetate is toxic and NaOH is corrosive. Avoid contact and inhalation, use a "
            "fume hood, and dispose of lead-containing waste properly."
        ),
        "danger": True,
    },
    {
        "slug": "ash",
        "index": "08",
        "name": "Mineral Ash Analysis for Calcium & Iron",
        "target": "Detects Ca²⁺ and Fe³⁺",
        "icon": "mineral",
        "materials": (
            "5 g sample; silica crucible; 2 M HCl / HNO₃; ammonium oxalate; KSCN; "
            "filter paper."
        ),
        "procedure": (
            "Incinerate 5 g of sample in a crucible to white ash. Dissolve the ash in 5 mL of "
            "2 M HCl/HNO₃ and filter. For calcium, make the filtrate basic with NH₄OH and add "
            "ammonium oxalate. For iron, add conc. HNO₃, boil, cool and add KSCN."
        ),
        "observation": (
            "Calcium gives a white precipitate of calcium oxalate (CaC₂O₄). Iron gives a "
            "blood-red complex [Fe(SCN)]²⁺."
        ),
        "basis": (
            "Calcium ions react with oxalate ions to form insoluble white calcium oxalate. Ferric "
            "ions react with thiocyanate ions to form a deeply coloured blood-red complex."
        ),
        "safety": (
            "Ash incineration requires a high-temperature silica crucible in a well-ventilated "
            "fume hood or laboratory setup. Handle hot crucibles with tongs and use appropriate "
            "PPE."
        ),
        "danger": True,
    },
]

RESULT_ROWS = [
    {
        "category": "carbohydrates",
        "test": "Starch (Iodine)",
        "values": [
            {"label": "Strong Positive", "level": "strong"},
            {"label": "Strong Positive", "level": "strong"},
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Strong Positive", "level": "strong"},
        ],
    },
    {
        "category": "carbohydrates",
        "test": "Reducing Sugars",
        "values": [
            {"label": "Slight Positive", "level": "slight"},
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Strong Positive", "level": "strong"},
            {"label": "Slight Positive", "level": "slight"},
        ],
    },
    {
        "category": "lipids",
        "test": "Fats (Grease Spot)",
        "values": [
            {"label": "Negative", "level": "negative"},
            {"label": "Slight Trace Spot", "level": "slight"},
            {"label": "Moderate Spot", "level": "moderate"},
            {"label": "Strong Persistent Spot", "level": "strong"},
        ],
    },
    {
        "category": "proteins",
        "test": "Proteins (Coagulation)",
        "values": [
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Strong Positive", "level": "strong"},
        ],
    },
    {
        "category": "proteins",
        "test": "Sulfur Proteins",
        "values": [
            {"label": "Slight Positive", "level": "slight"},
            {"label": "Moderate Positive", "level": "moderate"},
            {"label": "Slight Positive", "level": "slight"},
            {"label": "Strong Positive", "level": "strong"},
        ],
    },
    {
        "category": "minerals",
        "test": "Calcium (Ca²⁺)",
        "values": [
            {"label": "Positive", "level": "positive"},
            {"label": "Positive", "level": "positive"},
            {"label": "Positive", "level": "positive"},
            {"label": "Strong Positive", "level": "strong"},
        ],
    },
    {
        "category": "minerals",
        "test": "Iron (Fe³⁺)",
        "values": [
            {"label": "Positive", "level": "positive"},
            {"label": "Strong Positive", "level": "strong"},
            {"label": "Strong Positive", "level": "strong"},
            {"label": "Trace Positive", "level": "slight"},
        ],
    },
    {
        "category": "physical",
        "test": "pH Level",
        "values": [
            {"label": "5.8–6.2", "level": "ph"},
            {"label": "6.5–6.8", "level": "ph"},
            {"label": "6.2–6.5", "level": "ph"},
            {"label": "6.0–6.4", "level": "ph"},
        ],
    },
]

RESULT_FILTERS = [
    {"value": "all", "label": "All tests"},
    {"value": "carbohydrates", "label": "Carbohydrates"},
    {"value": "proteins", "label": "Proteins"},
    {"value": "lipids", "label": "Lipids"},
    {"value": "minerals", "label": "Minerals"},
    {"value": "physical", "label": "Physical"},
]

CONCLUSIONS = [
    {
        "index": "01",
        "name": "Idli",
        "tag": "Fermented",
        "text": (
            "High starch content with low fat presence. Fermentation lowers the pH slightly "
            "(5.8–6.2) due to lactic acid formation, enhancing bioavailability of essential "
            "minerals like calcium and iron without a heavy lipid load."
        ),
        "findings": [
            {"label": "Starch", "level": "strong"},
            {"label": "Fat", "level": "negative"},
            {"label": "pH 5.8–6.2", "level": "ph"},
        ],
    },
    {
        "index": "02",
        "name": "Multi Grain Porridge",
        "tag": "Whole grain",
        "text": (
            "An optimal balance of complex starch, dietary fibres, moderate reducing sugars and "
            "rich iron (Fe³⁺) content. It represents a nutrient-dense, low-fat breakfast choice "
            "suitable for sustained energy release."
        ),
        "findings": [
            {"label": "Starch", "level": "strong"},
            {"label": "Sugars", "level": "moderate"},
            {"label": "Iron", "level": "strong"},
        ],
    },
    {
        "index": "03",
        "name": "Muesli",
        "tag": "Processed cereal",
        "text": (
            "High reducing sugar content alongside moderate complex carbohydrates and iron. "
            "While nutrient-dense, processed variants often contain higher simple sugars."
        ),
        "findings": [
            {"label": "Reducing sugars", "level": "strong"},
            {"label": "Starch", "level": "moderate"},
            {"label": "Iron", "level": "strong"},
        ],
    },
    {
        "index": "04",
        "name": "Bread with Cheese",
        "tag": "Dairy-rich",
        "text": (
            "High total lipid content, dense proteins rich in sulfur amino acids, and high "
            "calcium levels (Ca²⁺) derived from dairy cheese, balanced by refined carbohydrates "
            "from bread."
        ),
        "findings": [
            {"label": "Fat", "level": "strong"},
            {"label": "Protein", "level": "strong"},
            {"label": "Calcium", "level": "strong"},
        ],
    },
]

SAFETY_PRECAUTIONS = [
    {
        "icon": "flask",
        "title": "Corrosive reagents",
        "text": (
            "Handle concentrated acids (HNO₃, HCl) and strong bases (NaOH) with extreme care "
            "using test tube holders and safety goggles."
        ),
    },
    {
        "icon": "heat",
        "title": "Heating & boiling tubes",
        "text": (
            "Heat boiling tubes gently while keeping them pointed away from self and peers to "
            "prevent sudden liquid bumping."
        ),
    },
    {
        "icon": "glass",
        "title": "Clean glassware & water",
        "text": (
            "Use clean, deionized distilled water and thoroughly washed glassware to prevent "
            "cross-contamination during sensitive precipitation and colorimetric tests."
        ),
    },
    {
        "icon": "crucible",
        "title": "Ash incineration",
        "text": (
            "Perform ash incineration in a high-temperature silica crucible inside a "
            "well-ventilated fume hood or laboratory setup."
        ),
    },
]

PPE_ITEMS = [
    {"icon": "goggles", "label": "Safety goggles"},
    {"icon": "coat", "label": "Lab coat"},
    {"icon": "gloves", "label": "Gloves"},
    {"icon": "hood", "label": "Fume hood"},
]

IMAGE_CREDITS = [
    {
        "label": "Hero",
        "title": "Healthy breakfast",
        "author": "Arek Adeoye",
        "licence": "CC0",
        "url": "https://commons.wikimedia.org/wiki/File:Healthy_breakfast_(Unsplash).jpg",
    },
    {
        "label": "Idli",
        "title": "Idli Sambar-Noida-UP-SP004",
        "author": "Sutapa Pal",
        "licence": "CC BY-SA 4.0",
        "url": "https://commons.wikimedia.org/wiki/File:Idli_Sambar-Noida-UP-SP004.jpg",
    },
    {
        "label": "Porridge",
        "title": "Jennifer Pallian 2016-10-07",
        "author": "Jennifer Pallian",
        "licence": "CC0",
        "url": "https://commons.wikimedia.org/wiki/File:Jennifer_Pallian_2016-10-07_(Unsplash_DfbrRpHTLy0).jpg",
    },
    {
        "label": "Muesli",
        "title": "Hamburg, Germany",
        "author": "Jannis Brandt",
        "licence": "CC0",
        "url": "https://commons.wikimedia.org/wiki/File:Hamburg,_Germany_(Unsplash_mmsQUgMLqUo).jpg",
    },
    {
        "label": "Bread with Cheese",
        "title": "Bread and Cheese at Faubourg Wines New Orleans",
        "author": "Infrogmation of New Orleans",
        "licence": "CC BY-SA 4.0",
        "url": "https://commons.wikimedia.org/wiki/File:Bread_and_Cheese_at_Faubourg_Wines_New_Orleans.jpg",
    },
    {
        "label": "Laboratory",
        "title": "Equipment and glassware",
        "author": "Belikov Maxim",
        "licence": "CC BY 4.0",
        "url": "https://commons.wikimedia.org/wiki/File:Equipment_and_glassware.jpg",
    },
]

TOPICS = {
    "overview": {
        "slug": "overview",
        "nav_label": "Overview",
        "eyebrow": "CBSE Chemistry Investigatory Project",
        "title": "Study of Nutrient Contents in Breakfast Items",
        "lede": (
            "A qualitative chemical and bio-nutritional analysis of four familiar breakfasts, "
            "supported by a live student survey."
        ),
        "blocks": [
            {
                "type": "hero",
                "eyebrow": "CBSE Chemistry Investigatory Project · Class XII · 2026–2027",
                "title_prefix": "Study of Nutrient Contents in",
                "title_accent": "Breakfast Items",
                "subtitle": "A Qualitative Chemical & Bio-Nutritional Analysis",
                "lede": (
                    "What is really in your morning meal? This project uses observable chemical "
                    "tests to identify starch, reducing sugars, proteins, lipids, calcium, iron "
                    "and pH across four familiar breakfasts."
                ),
                "image": "images/hero-breakfast.jpg",
                "image_webp": "images/hero-breakfast.webp",
                "image_alt": (
                    "A bright breakfast setting with a bowl of cereal, blueberries, banana and a "
                    "cup of coffee."
                ),
                "metrics": [
                    {"value": "4", "label": "Breakfast samples"},
                    {"value": "8", "label": "Chemical tests"},
                    {"value": "3", "label": "Macronutrient groups"},
                ],
                "ticker": [
                    "Idli",
                    "Multi Grain Porridge",
                    "Muesli",
                    "Bread with Cheese",
                ],
                "ticker_note": "Macronutrients · Minerals · pH",
            },
            {
                "type": "prose",
                "eyebrow": "01 · The brief",
                "title": "Breakfast is a chemical story.",
                "lead": (
                    "Breakfast breaks the overnight metabolic fast and restores systemic glucose "
                    "levels required for physical performance and cognitive function."
                ),
                "paragraphs": [
                    (
                        "A balanced breakfast supplies essential macronutrients — complex "
                        "carbohydrates for sustained glycemic release, proteins for tissue growth, "
                        "and essential lipids for membrane integrity — together with calcium and "
                        "iron."
                    ),
                    (
                        "This project compares four very different morning meals: a fermented "
                        "traditional food (Idli), a whole-grain preparation (Multi Grain "
                        "Porridge), a processed cereal (Muesli) and a dairy-rich bakery pairing "
                        "(Bread with Cheese)."
                    ),
                ],
                "note": (
                    "The aim is to identify what is present and compare the strength of each "
                    "response — not to measure exact nutrient quantities."
                ),
            },
            {
                "type": "aim",
                "eyebrow": "Aim · Project objective",
                "title": "What this project set out to do.",
                "text": (
                    "To perform qualitative chemical analysis to detect key macronutrients "
                    "(carbohydrates, proteins, fats), inorganic minerals (calcium, iron), and "
                    "physical parameters (pH, ash content) across four common breakfast items: "
                    "Idli, Multi Grain Porridge, Muesli, and Bread with Cheese."
                ),
                "objectives": [
                    (
                        "Extract aqueous fractions from food samples and perform qualitative "
                        "functional group identification tests."
                    ),
                    (
                        "Evaluate protein quality, heat denaturation characteristics, and the "
                        "presence of sulfur-containing amino acids."
                    ),
                    (
                        "Determine inorganic ash residue through incineration and test for "
                        "essential micronutrient cations (Ca²⁺, Fe³⁺)."
                    ),
                    (
                        "Compare the biochemical nutritional density of traditional fermented "
                        "foods versus processed modern breakfast choices."
                    ),
                ],
            },
            {
                "type": "visual",
                "eyebrow": "From kitchen to bench",
                "title": "A breakfast can be read through chemistry.",
                "text": (
                    "The project moves from four familiar breakfasts to a repeatable bench "
                    "method: extract, test, observe and compare. Each result is a qualitative "
                    "signal, interpreted alongside the method that produced it."
                ),
                "image": "images/lab-extraction.jpg",
                "image_alt": "Laboratory glassware arranged on a bench",
                "caption": "A clear extract is the starting point for most of the qualitative tests.",
                "note": (
                    "The photographs show the foods and the bench context. They are not "
                    "experimental evidence and do not replace the recorded observations."
                ),
            },
            {
                "type": "samples",
                "eyebrow": "02 · The samples",
                "title": "Four breakfasts. Four chemical fingerprints.",
                "lede": (
                    "Each sample was prepared the same way, then tested for the same set of "
                    "nutrients and properties. Open a sample to see its recorded observations, "
                    "preparation notes and the tests behind them."
                ),
                "items": SAMPLES,
            },
            {
                "type": "survey_insights",
                "eyebrow": "Live student survey",
                "title": "What students are saying about breakfast.",
                "lede": (
                    "A compact, anonymous view of the survey responses. Public charts show "
                    "aggregate counts only — never names, emails or individual comments."
                ),
            },
            {
                "type": "topics",
                "eyebrow": "Explore the project",
                "title": "Follow the investigation.",
                "lede": (
                    "Move from the aim and method to the bench work, observations and the final "
                    "comparison."
                ),
                "items": [
                    {
                        "slug": "aim",
                        "label": "Aim & Objectives",
                        "icon": "target",
                        "description": "The question, scope and four project objectives.",
                    },
                    {
                        "slug": "samples",
                        "label": "Food Samples",
                        "icon": "leaf",
                        "description": "Meet the four breakfasts and their qualitative profiles.",
                    },
                    {
                        "slug": "tests",
                        "label": "Chemical Tests",
                        "icon": "flask",
                        "description": "Eight tests, their reagents, observations and chemistry.",
                    },
                    {
                        "slug": "results",
                        "label": "Comparative Results",
                        "icon": "chart",
                        "description": "The consolidated observation table and legend.",
                    },
                    {
                        "slug": "conclusion",
                        "label": "Conclusion",
                        "icon": "spark",
                        "description": "What the four breakfasts revealed as a group.",
                    },
                    {
                        "slug": "project_info",
                        "label": "Project Information",
                        "icon": "info",
                        "description": "Submission details, source document and credits.",
                    },
                ],
            },
        ],
    },
    "aim": {
        "slug": "aim",
        "nav_label": "Aim & Objectives",
        "eyebrow": "Project objective",
        "title": "Aim & Objectives",
        "lede": "A focused qualitative question, explored through four breakfasts and eight bench tests.",
        "blocks": [
            {
                "type": "aim",
                "eyebrow": "Aim",
                "title": "The question at the centre of the project.",
                "text": (
                    "To perform qualitative chemical analysis to detect key macronutrients "
                    "(carbohydrates, proteins, fats), inorganic minerals (calcium, iron), and "
                    "physical parameters (pH, ash content) across four common breakfast items: "
                    "Idli, Multi Grain Porridge, Muesli, and Bread with Cheese."
                ),
                "objectives": [
                    (
                        "Extract aqueous fractions from food samples and perform qualitative "
                        "functional group identification tests."
                    ),
                    (
                        "Evaluate protein quality, heat denaturation characteristics, and the "
                        "presence of sulfur-containing amino acids."
                    ),
                    (
                        "Determine inorganic ash residue through incineration and test for "
                        "essential micronutrient cations (Ca²⁺, Fe³⁺)."
                    ),
                    (
                        "Compare the biochemical nutritional density of traditional fermented "
                        "foods versus processed modern breakfast choices."
                    ),
                ],
            },
            {
                "type": "visual",
                "eyebrow": "The question",
                "title": "Identify first. Compare second.",
                "text": (
                    "The investigation asks what is present in each breakfast and how "
                    "pronounced the visible response is. It does not claim to measure exact "
                    "grams, milligrams or nutrient quantities."
                ),
                "image": "images/hero-breakfast.jpg",
                "image_alt": "A breakfast table with several morning foods",
                "caption": "Four breakfasts, one shared qualitative method.",
                "reverse": True,
            },
            {
                "type": "callout",
                "tone": "info",
                "title": "Qualitative, not quantitative.",
                "text": (
                    "The tests describe whether a nutrient is present and how pronounced the "
                    "visible response is. They do not measure grams, milligrams or exact nutrient "
                    "quantities."
                ),
            },
        ],
    },
    "introduction": {
        "slug": "introduction",
        "nav_label": "Introduction",
        "eyebrow": "The nutritional frame",
        "title": "Introduction",
        "lede": "Why breakfast, and why chemistry can tell us something useful about it.",
        "blocks": [
            {
                "type": "prose",
                "eyebrow": "01 · The brief",
                "title": "Breakfast is a chemical story.",
                "lead": (
                    "Breakfast breaks the overnight metabolic fast and restores systemic glucose "
                    "levels required for optimal physical performance and cognitive functions."
                ),
                "paragraphs": [
                    (
                        "A balanced breakfast supplies essential macronutrients — complex "
                        "carbohydrates for sustained glycemic release, functional proteins for "
                        "tissue growth, and essential lipids for membrane integrity — together "
                        "with vital micronutrients like calcium and iron."
                    ),
                    (
                        "Modern dietary habits feature a diverse array of morning meals ranging "
                        "from fermented traditional foods (such as Idli) to nutrient-dense, "
                        "whole-grain mixtures (Multi Grain Porridge), processed breakfast cereals "
                        "(Muesli), and dairy-enriched bakery foods (Bread with Cheese)."
                    ),
                    (
                        "Qualitative chemical analysis enables direct identification of these "
                        "nutritional biochemicals through specific functional group reactions, "
                        "complexation, and precipitation tests. By systematic chemical testing, "
                        "this project evaluates the qualitative nutritional composition of common "
                        "breakfast items."
                    ),
                ],
                "note": (
                    "The aim is to identify what is present and compare the strength of each "
                    "response — not to measure exact nutrient quantities."
                ),
            },
            {
                "type": "gallery",
                "eyebrow": "The four breakfasts",
                "title": "One table, four different starting points.",
                "lede": (
                    "The samples span fermented, whole-grain, processed cereal and dairy-rich "
                    "breakfasts, giving the investigation a deliberately broad comparison."
                ),
                "items": [
                    {
                        "image": "images/sample-idli.jpg",
                        "alt": "Idli served with sambar",
                        "name": "Idli",
                        "tag": "Fermented · traditional",
                    },
                    {
                        "image": "images/sample-porridge.jpg",
                        "alt": "A bowl of multi grain porridge",
                        "name": "Multi Grain Porridge",
                        "tag": "Whole grain · fibre",
                    },
                    {
                        "image": "images/sample-muesli.jpg",
                        "alt": "Muesli with fruit and nuts",
                        "name": "Muesli",
                        "tag": "Processed cereal",
                    },
                    {
                        "image": "images/sample-bread-cheese.jpg",
                        "alt": "Bread served with cheese",
                        "name": "Bread with Cheese",
                        "tag": "Dairy-rich · bakery",
                    },
                ],
            },
            {
                "type": "nutrients",
                "eyebrow": "The nutritional frame",
                "title": "What the tests look for",
                "lede": (
                    "Four nutrient families and two mineral targets are explored through visible "
                    "chemical evidence."
                ),
                "items": [
                    {
                        "name": "Carbohydrates",
                        "description": "Starch & reducing sugars",
                        "tone": "amber",
                        "icon": "starch",
                    },
                    {
                        "name": "Proteins",
                        "description": "Coagulation & sulfur amino acids",
                        "tone": "blue",
                        "icon": "protein",
                    },
                    {
                        "name": "Lipids",
                        "description": "Fats & oils",
                        "tone": "rose",
                        "icon": "fat",
                    },
                    {
                        "name": "Micronutrients",
                        "description": "Calcium & iron",
                        "tone": "teal",
                        "icon": "mineral",
                    },
                ],
            },
            {
                "type": "callout",
                "tone": "note",
                "title": "One important distinction.",
                "text": (
                    "A stronger colour or precipitate means a more pronounced qualitative response. "
                    "It does not automatically mean a larger nutrient quantity in the food."
                ),
            },
        ],
    },
    "samples": {
        "slug": "samples",
        "nav_label": "Food Samples",
        "eyebrow": "The samples",
        "title": "Four breakfasts. Four chemical fingerprints.",
        "lede": "Each sample was prepared the same way, then tested for the same set of nutrients and properties.",
        "blocks": [
            {
                "type": "samples",
                "eyebrow": "Sample dossiers",
                "title": "Meet the four breakfasts.",
                "lede": (
                    "Open a sample to see its recorded observations, the tests behind them and a "
                    "short nutrition context."
                ),
                "items": SAMPLES,
            },
            {
                "type": "callout",
                "tone": "note",
                "title": "Read the cards as qualitative evidence.",
                "text": (
                    "The findings shown on these cards are the project's qualitative "
                    "observations. General nutritional context is kept separate inside each "
                    "sample dossier."
                ),
            },
        ],
    },
    "materials": {
        "slug": "materials",
        "nav_label": "Materials Required",
        "eyebrow": "The bench",
        "title": "Materials Required",
        "lede": "A compact set of glassware, equipment and measurement tools for the qualitative bench.",
        "blocks": [
            {
                "type": "materials",
                "eyebrow": "Apparatus & glassware",
                "title": "Everything the bench needs.",
                "lede": (
                    "Search the list, mark what is available, and read the hazard labels before "
                    "beginning any supervised work."
                ),
                "kind": "apparatus",
                "items": APPARATUS_ITEMS,
                "hazard_note": (
                    "Experiments involving concentrated acids, lead acetate, strong bases, "
                    "heating or incineration must be performed under qualified laboratory "
                    "supervision with appropriate protective equipment and waste disposal. Never "
                    "attempt these procedures at home."
                ),
            },
            {
                "type": "visual",
                "eyebrow": "The bench",
                "title": "Good results begin with organised glassware.",
                "text": (
                    "Clean, labelled glassware and a clear workspace make the qualitative "
                    "tests easier to observe, compare and repeat. The bench is treated as part "
                    "of the method, not as background."
                ),
                "image": "images/lab-extraction.jpg",
                "image_alt": "Laboratory glassware on a bench",
                "caption": "Glassware, measuring tools and a clear working area.",
            },
        ],
    },
    "chemicals": {
        "slug": "chemicals",
        "nav_label": "Chemicals & Reagents",
        "eyebrow": "The reagents",
        "title": "Chemicals & Reagents",
        "lede": "Eleven qualitative reagents, their roles, and the hazards that need respect.",
        "blocks": [
            {
                "type": "materials",
                "eyebrow": "Reagent cabinet",
                "title": "The reagents behind the reactions.",
                "lede": (
                    "Each reagent is listed with its formula or working concentration, hazard "
                    "class and a short bench note."
                ),
                "kind": "chemicals",
                "items": CHEMICALS,
                "hazard_note": (
                    "Concentrated acids and bases are corrosive; lead acetate is toxic; iodine, "
                    "Benedict's reagent, ammonium oxalate and KSCN are irritant or harmful. Use "
                    "goggles, gloves and a fume hood, and follow your institution's waste-disposal "
                    "rules."
                ),
            },
        ],
    },
    "preparation": {
        "slug": "preparation",
        "nav_label": "Preparation of Samples",
        "eyebrow": "The method",
        "title": "Preparation of Food Samples",
        "lede": "A single, repeatable extraction procedure keeps the comparison fair.",
        "blocks": [
            {
                "type": "prep",
                "eyebrow": "Aqueous extraction",
                "title": "From food sample to clear extract.",
                "lede": (
                    "The solid sample is used directly only for the grease-spot test. Everything "
                    "else starts with the same aqueous extract."
                ),
                "image": "images/lab-extraction.jpg",
                "image_webp": "images/lab-extraction.webp",
                "image_alt": "Laboratory glassware and test tubes holding coloured solutions.",
                "caption": "Clear filtrate, ready for qualitative testing.",
                "callout_title": "For lipids, skip the extract.",
                "callout_text": (
                    "A solid, unextracted portion of each sample is pressed directly onto filter "
                    "paper for the grease-spot test."
                ),
                "steps": [
                    {"title": "Weigh the sample", "text": "Take 10 g of each finely ground food sample."},
                    {"title": "Add water", "text": "Mix each sample with 50 mL of distilled water in a beaker."},
                    {"title": "Boil", "text": "Boil the mixture for 5 minutes."},
                    {"title": "Cool", "text": "Allow the mixture to cool to room temperature."},
                    {"title": "Filter", "text": "Filter through Whatman No. 42 filter paper."},
                    {"title": "Collect the extract", "text": "Collect the clear aqueous extract for chemical testing."},
                ],
            },
        ],
    },
    "tests": {
        "slug": "tests",
        "nav_label": "Chemical Tests",
        "eyebrow": "The chemistry",
        "title": "Chemical Tests",
        "lede": "Eight tests, one qualitative picture.",
        "blocks": [
            {
                "type": "tests",
                "eyebrow": "Bench procedures",
                "title": "Eight tests, one qualitative picture.",
                "lede": (
                    "Each test targets a specific nutrient or property. Expand a card for the "
                    "materials, procedure, expected positive observation, chemical basis and "
                    "safety note."
                ),
                "items": TESTS,
            },
        ],
    },
    "results": {
        "slug": "results",
        "nav_label": "Comparative Results",
        "eyebrow": "The evidence",
        "title": "Comparative Observation Table",
        "lede": "The consolidated observations from the project. Labels describe the relative strength of each qualitative response.",
        "blocks": [
            {
                "type": "gallery",
                "eyebrow": "The samples at a glance",
                "title": "Four foods behind the comparison.",
                "lede": (
                    "Keep the samples in view while reading the consolidated table. The "
                    "gallery is contextual; the table carries the recorded qualitative labels."
                ),
                "items": [
                    {
                        "image": "images/sample-idli.jpg",
                        "alt": "Idli served with sambar",
                        "name": "Idli",
                        "tag": "High starch · low fat",
                    },
                    {
                        "image": "images/sample-porridge.jpg",
                        "alt": "A bowl of multi grain porridge",
                        "name": "Multi Grain Porridge",
                        "tag": "Complex starch · iron",
                    },
                    {
                        "image": "images/sample-muesli.jpg",
                        "alt": "Muesli with fruit and nuts",
                        "name": "Muesli",
                        "tag": "Reducing sugars · iron",
                    },
                    {
                        "image": "images/sample-bread-cheese.jpg",
                        "alt": "Bread served with cheese",
                        "name": "Bread with Cheese",
                        "tag": "Fat · protein · calcium",
                    },
                ],
            },
            {
                "type": "results",
                "eyebrow": "The evidence",
                "title": "The comparative picture.",
                "lede": (
                    "The consolidated observations from the project. Labels describe the relative "
                    "strength of each qualitative response — they are not nutrient quantities."
                ),
                "columns": ["Test", "Idli", "Multi Grain Porridge", "Muesli", "Bread with Cheese"],
                "rows": RESULT_ROWS,
                "filters": RESULT_FILTERS,
                "legend": [
                    {"label": "Strong positive", "level": "strong"},
                    {"label": "Positive", "level": "positive"},
                    {"label": "Moderate", "level": "moderate"},
                    {"label": "Slight / trace", "level": "slight"},
                    {"label": "Negative", "level": "negative"},
                    {"label": "pH range", "level": "ph"},
                ],
                "note": (
                    "The table preserves the exact qualitative labels and pH ranges from the "
                    "project's observation table. A stronger label means a more pronounced test "
                    "response, not a larger nutrient quantity."
                ),
            },
        ],
    },
    "conclusion": {
        "slug": "conclusion",
        "nav_label": "Comprehensive Conclusion",
        "eyebrow": "The takeaway",
        "title": "Comprehensive Conclusion",
        "lede": "What the breakfasts revealed, sample by sample and as a group.",
        "blocks": [
            {
                "type": "conclusions",
                "eyebrow": "The takeaway",
                "title": "What the breakfasts revealed.",
                "lede": (
                    "The project's conclusion, sample by sample — followed by the overall "
                    "comparison between traditional, whole-grain, processed and dairy-rich "
                    "breakfasts."
                ),
                "items": CONCLUSIONS,
                "overall": {
                    "title": "Traditional, whole grain, processed, dairy-rich.",
                    "text": (
                        "Traditional fermented foods (Idli) and whole-grain preparations (Multi "
                        "Grain Porridge) provide rich complex carbohydrates and micronutrients with "
                        "minimal fats, whereas dairy-rich breakfasts (Bread with Cheese) deliver "
                        "higher fat and protein density."
                    ),
                    "items": [
                        {
                            "label": "Traditional & fermented",
                            "strong": "Complex carbohydrates",
                            "text": "Low fat · mildly acidic pH",
                        },
                        {
                            "label": "Whole grain",
                            "strong": "Complex starch & fibre",
                            "text": "Moderate sugars · strong iron",
                        },
                        {
                            "label": "Processed cereal",
                            "strong": "Reducing sugars",
                            "text": "Moderate starch · strong iron",
                        },
                        {
                            "label": "Dairy-rich",
                            "strong": "Fat, protein & calcium",
                            "text": "Refined carbohydrates from bread",
                        },
                    ],
                },
                "note": (
                    "These qualitative tests indicate the presence and relative intensity of "
                    "nutrients. They do not provide exact nutrient quantities or definitive "
                    "nutritional measurements."
                ),
            },
            {
                "type": "visual",
                "eyebrow": "The comparison",
                "title": "The pattern matters more than a single label.",
                "text": (
                    "Read together, the four breakfasts show different balances of starch, "
                    "sugars, protein, fat and minerals. The comparison is qualitative, but the "
                    "contrast between traditional, whole-grain, processed and dairy-rich foods "
                    "is clear."
                ),
                "image": "images/hero-breakfast.jpg",
                "image_alt": "A breakfast table with several morning foods",
                "caption": "A shared table, four distinct chemical profiles.",
                "reverse": True,
            },
        ],
    },
    "safety": {
        "slug": "safety",
        "nav_label": "Precautions & Safety",
        "eyebrow": "Safety",
        "title": "Precautions & Safety Measures",
        "lede": "Chemistry is only useful when it is safe.",
        "blocks": [
            {
                "type": "safety",
                "eyebrow": "Read before the bench",
                "title": "Chemistry is only useful when it is safe.",
                "lede": (
                    "The project's precautions, presented as they apply to the bench. These "
                    "procedures are for a properly equipped laboratory under qualified supervision."
                ),
                "banner_title": "Laboratory supervision is mandatory.",
                "banner_text": (
                    "Concentrated acids, lead-containing reagents, open flames and high-temperature "
                    "incineration require trained supervision and appropriate waste disposal. Never "
                    "attempt these procedures at home."
                ),
                "precautions": SAFETY_PRECAUTIONS,
                "ppe": PPE_ITEMS,
                "ppe_note": (
                    "Follow your institution's waste-disposal rules, especially for lead-containing "
                    "reagents."
                ),
            },
            {
                "type": "visual",
                "eyebrow": "The bench",
                "title": "Safety is part of the method.",
                "text": (
                    "Corrosive reagents, lead-containing solutions, open flames and "
                    "high-temperature incineration all require trained supervision, personal "
                    "protective equipment and proper waste disposal."
                ),
                "image": "images/lab-extraction.jpg",
                "image_alt": "Laboratory glassware arranged on a bench",
                "caption": "A controlled laboratory bench, not a home kitchen.",
            },
        ],
    },
    "project_info": {
        "slug": "project_info",
        "nav_label": "Project Information",
        "eyebrow": "Project information",
        "title": "Project Information",
        "lede": "Submission details, the source document and image credits.",
        "blocks": [
            {
                "type": "project_info",
                "eyebrow": "Submission details",
                "title": "About this project.",
                "details": [
                    {"label": "Project title", "value": PROJECT["title"]},
                    {"label": "Submitted by", "value": "R Adhithya, Darshini K, Kaivalya P"},
                    {"label": "Class", "value": PROJECT["class_label"]},
                    {"label": "Academic year", "value": PROJECT["year"]},
                    {"label": "Project type", "value": "CBSE Chemistry Investigatory Project"},
                ],
                "source_title": "Source document",
                "source_text": (
                    "The experimental content on this website is derived from the submitted "
                    "project document. No external references were listed in the source "
                    "document. Reagent compositions, procedures and observations are presented "
                    "as recorded in the project."
                ),
                "source_pdf": PROJECT["source_pdf"],
                "source_pdf_label": PROJECT["source_pdf_label"],
                "credits_title": "Image credits & licences",
                "credits": IMAGE_CREDITS,
            },
        ],
    },
}

TOPIC_ORDER = [
    "overview",
    "aim",
    "introduction",
    "samples",
    "materials",
    "chemicals",
    "preparation",
    "tests",
    "results",
    "conclusion",
    "safety",
    "project_info",
]

# ---------------------------------------------------------------------------
# Single-file refinements and shared project data
# ---------------------------------------------------------------------------
PROJECT["source_pdf"] = ""
PROJECT["source_pdf_label"] = "CHEMISTRY IP ROUGH DRAFT.pdf"

IMAGE_MAP = {
    "images/hero-breakfast.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Healthy_breakfast_%28Unsplash%29.jpg?width=1400"
    ),
    "images/hero-breakfast.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Healthy_breakfast_%28Unsplash%29.jpg?width=1400"
    ),
    "images/sample-idli.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Idli_Sambar-Noida-UP-SP004.jpg?width=900"
    ),
    "images/sample-idli.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Idli_Sambar-Noida-UP-SP004.jpg?width=900"
    ),
    "images/sample-porridge.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Jennifer_Pallian_2016-10-07_%28Unsplash_DfbrRpHTLy0%29.jpg?width=900"
    ),
    "images/sample-porridge.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Jennifer_Pallian_2016-10-07_%28Unsplash_DfbrRpHTLy0%29.jpg?width=900"
    ),
    "images/sample-muesli.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Hamburg%2C_Germany_%28Unsplash_mmsQUgMLqUo%29.jpg?width=900"
    ),
    "images/sample-muesli.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Hamburg%2C_Germany_%28Unsplash_mmsQUgMLqUo%29.jpg?width=900"
    ),
    "images/sample-bread-cheese.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Bread_and_Cheese_at_Faubourg_Wines_New_Orleans.jpg?width=900"
    ),
    "images/sample-bread-cheese.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Bread_and_Cheese_at_Faubourg_Wines_New_Orleans.jpg?width=900"
    ),
    "images/lab-extraction.jpg": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Equipment_and_glassware.jpg?width=1200"
    ),
    "images/lab-extraction.webp": (
        "https://commons.wikimedia.org/wiki/Special:FilePath/"
        "Equipment_and_glassware.jpg?width=1200"
    ),
}


def replace_local_images(value) -> None:
    if isinstance(value, dict):
        for key, child in list(value.items()):
            if key in {"image", "image_webp"} and isinstance(child, str):
                value[key] = IMAGE_MAP.get(child, child)
            else:
                replace_local_images(child)
    elif isinstance(value, list):
        for child in value:
            replace_local_images(child)


replace_local_images(PREFERRED_IMAGE_OPTIONS)
replace_local_images(SAMPLES)
replace_local_images(TOPICS)


HEADER_NAV = [
    {"label": "Overview", "slug": "overview"},
    {"label": "Aim & Objectives", "slug": "aim"},
    {"label": "Introduction", "slug": "introduction"},
    {"label": "Food Samples", "slug": "samples"},
    {"label": "Materials Required", "slug": "materials"},
    {"label": "Chemicals & Reagents", "slug": "chemicals"},
    {"label": "Sample Preparation", "slug": "preparation"},
    {"label": "Chemical Tests", "slug": "tests"},
    {"label": "Observation Table", "slug": "results"},
    {"label": "Conclusion", "slug": "conclusion"},
    {"label": "Safety & Precautions", "slug": "safety"},
    {"label": "Survey", "endpoint": "survey"},
    {"label": "Survey Results", "endpoint": "survey_results"},
]

SIDEBAR_GROUPS = [
    {
        "label": "Project",
        "items": [
            {"label": "Overview", "slug": "overview"},
            {"label": "Aim & Objectives", "slug": "aim"},
            {"label": "Introduction", "slug": "introduction"},
        ],
    },
    {
        "label": "Investigation",
        "items": [
            {"label": "Food Samples", "slug": "samples"},
            {"label": "Materials Required", "slug": "materials"},
            {"label": "Chemicals & Reagents", "slug": "chemicals"},
            {"label": "Sample Preparation", "slug": "preparation"},
        ],
    },
    {
        "label": "Findings",
        "items": [
            {"label": "Chemical Tests", "slug": "tests"},
            {"label": "Observation Table", "slug": "results"},
            {"label": "Conclusion", "slug": "conclusion"},
            {"label": "Safety & Precautions", "slug": "safety"},
        ],
    },
    {
        "label": "Participate",
        "items": [
            {"label": "Survey", "endpoint": "survey"},
            {"label": "Survey Results", "endpoint": "survey_results"},
        ],
    },
]

TOPICS["preparation"]["nav_label"] = "Sample Preparation"
TOPICS["preparation"]["title"] = "Sample Preparation"
TOPICS["results"]["nav_label"] = "Observation Table"
TOPICS["results"]["title"] = "Observation Table"
TOPICS["conclusion"]["nav_label"] = "Conclusion"
TOPICS["conclusion"]["title"] = "Conclusion"
TOPICS["safety"]["nav_label"] = "Safety & Precautions"
TOPICS["safety"]["title"] = "Safety & Precautions"
TOPICS["overview"]["blocks"][4]["lede"] = (
    "A compact, live view of the survey responses. Public charts show aggregate counts "
    "only — never names, emails or individual comments."
)

LEARNING_OPTIONS[:] = [
    "Interactive experiments",
    "Infographics and graphs",
    "Short videos",
    "Classroom demonstrations",
    "Interesting food facts",
    "Other",
]
INTEREST_OPTIONS[:] = [
    "Starch and sugar tests",
    "Protein tests",
    "Fat tests",
    "Calcium and iron tests",
    "Comparing all four foods",
    "Chemical reactions",
]

SURVEY_QUESTION_BY_FIELD["preferred_breakfast"]["label"] = "Which breakfast do you prefer?"
SURVEY_QUESTION_BY_FIELD["preferred_breakfast"]["help"] = (
    "Choose the breakfast you would most enjoy or usually choose."
)
SURVEY_QUESTION_BY_FIELD["eating_time"]["label"] = (
    "How long do you spend eating breakfast?"
)
SURVEY_QUESTION_BY_FIELD["eating_time"]["help"] = (
    "Think about a typical morning and include the time you spend sitting down to eat."
)
SURVEY_QUESTION_BY_FIELD["perceived_grease_spot"]["label"] = (
    "Which food produced the strongest grease-spot result in the supplied project?"
)
SURVEY_QUESTION_BY_FIELD["perceived_grease_spot"]["help"] = (
    "Answer from memory or choose Not sure. The supplied project's observation is "
    "kept separate from your answer."
)
SURVEY_QUESTION_BY_FIELD["perceived_grease_spot"]["project_note"] = (
    "The supplied project recorded a strong persistent grease spot for Bread with Cheese. "
    "This question records your own answer."
)
SURVEY_QUESTION_BY_FIELD["learning_method"]["label"] = (
    "What would encourage you to learn more about nutrients?"
)
SURVEY_QUESTION_BY_FIELD["learning_method"]["options"] = radio_options(LEARNING_OPTIONS)
SURVEY_QUESTION_BY_FIELD["interesting_topic"]["options"] = radio_options(INTEREST_OPTIONS)

SURVEY_FIELD_LABELS = {
    question["field"]: question["label"] for question in SURVEY_QUESTIONS
}


# ---------------------------------------------------------------------------
# SQLite storage
# ---------------------------------------------------------------------------
DB_COLUMNS = {
    "submitted_at": "TEXT NOT NULL DEFAULT ''",
    "name": "TEXT NOT NULL DEFAULT ''",
    "email": "TEXT NOT NULL DEFAULT ''",
    "class_level": "TEXT NOT NULL DEFAULT ''",
    "breakfast_frequency": "TEXT NOT NULL DEFAULT ''",
    "preferred_breakfast": "TEXT NOT NULL DEFAULT ''",
    "main_reason": "TEXT NOT NULL DEFAULT ''",
    "eating_time": "TEXT NOT NULL DEFAULT ''",
    "important_nutrient": "TEXT NOT NULL DEFAULT ''",
    "test_familiarity": "TEXT NOT NULL DEFAULT ''",
    "iodine_awareness": "TEXT NOT NULL DEFAULT ''",
    "perceived_grease_spot": "TEXT NOT NULL DEFAULT ''",
    "iodine_knowledge": "TEXT NOT NULL DEFAULT ''",
    "learning_method": "TEXT NOT NULL DEFAULT ''",
    "choice_influence": "TEXT NOT NULL DEFAULT ''",
    "interesting_topic": "TEXT NOT NULL DEFAULT ''",
    "comments": "TEXT",
    "consent_given": "INTEGER NOT NULL DEFAULT 1",
}
RESPONSE_COLUMNS = list(DB_COLUMNS)


def get_db() -> sqlite3.Connection:
    connection = sqlite3.connect(str(DB_PATH), timeout=15)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    connection = get_db()
    try:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS survey_responses (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                submitted_at TEXT NOT NULL,
                name TEXT NOT NULL,
                email TEXT NOT NULL,
                class_level TEXT NOT NULL,
                breakfast_frequency TEXT NOT NULL,
                preferred_breakfast TEXT NOT NULL,
                main_reason TEXT NOT NULL,
                eating_time TEXT NOT NULL,
                important_nutrient TEXT NOT NULL,
                test_familiarity TEXT NOT NULL,
                iodine_awareness TEXT NOT NULL,
                perceived_grease_spot TEXT NOT NULL,
                iodine_knowledge TEXT NOT NULL,
                learning_method TEXT NOT NULL,
                choice_influence TEXT NOT NULL,
                interesting_topic TEXT NOT NULL,
                comments TEXT,
                consent_given INTEGER NOT NULL DEFAULT 1
            )
            """
        )
        existing = {
            row["name"] for row in connection.execute("PRAGMA table_info(survey_responses)")
        }
        for column, definition in DB_COLUMNS.items():
            if column not in existing:
                connection.execute(
                    f"ALTER TABLE survey_responses ADD COLUMN {column} {definition}"
                )
        connection.execute(
            "CREATE INDEX IF NOT EXISTS ix_survey_responses_submitted_at "
            "ON survey_responses (submitted_at)"
        )
        connection.commit()
    finally:
        connection.close()


init_db()


IST = timezone(timedelta(hours=5, minutes=30), "IST")


def utcnow() -> datetime:
    return datetime.now(timezone.utc).replace(tzinfo=None)


def iso_utc(value: datetime | None) -> str | None:
    if value is None:
        return None
    return value.replace(tzinfo=timezone.utc).isoformat().replace("+00:00", "Z")


def parse_db_datetime(value) -> datetime | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        parsed = value
    else:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    if parsed.tzinfo is not None:
        parsed = parsed.astimezone(timezone.utc).replace(tzinfo=None)
    return parsed


def format_local(value, include_zone: bool = True) -> str:
    parsed = parse_db_datetime(value)
    if parsed is None:
        return "No responses yet"
    local = parsed.replace(tzinfo=timezone.utc).astimezone(IST)
    pattern = "%d %b %Y, %I:%M %p %Z" if include_zone else "%Y-%m-%d %H:%M:%S"
    return local.strftime(pattern)


def format_local_date(value) -> str:
    parsed = parse_db_datetime(value)
    if parsed is None:
        return ""
    return parsed.replace(tzinfo=timezone.utc).astimezone(IST).strftime("%Y-%m-%d %H:%M:%S")


def local_date_start_utc(value: date) -> str:
    local = datetime.combine(value, dt_time.min, tzinfo=IST)
    return local.astimezone(timezone.utc).replace(tzinfo=None).strftime("%Y-%m-%d %H:%M:%S")


def local_date_end_utc(value: date) -> str:
    local = datetime.combine(value, dt_time.max, tzinfo=IST)
    return local.astimezone(timezone.utc).replace(tzinfo=None).strftime("%Y-%m-%d %H:%M:%S")


def parse_int(value: str | None, default: int) -> int:
    try:
        return int(value) if value is not None else default
    except (TypeError, ValueError):
        return default


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    try:
        return date.fromisoformat(value)
    except ValueError:
        return None


def opaque_ref(record_id: int) -> str:
    return hmac.new(
        app.secret_key.encode("utf-8"),
        str(record_id).encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()[:16]


# ---------------------------------------------------------------------------
# Validation and submission protection
# ---------------------------------------------------------------------------
def clean_name(value: str) -> str:
    value = re.sub(r"[\x00-\x1f\x7f]", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def clean_comments(value: str) -> str:
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    value = "".join(
        character for character in value if character in "\n\t" or ord(character) >= 32
    )
    return value.strip()


def validate_email(value: str) -> str | None:
    if not value or len(value) > 254:
        return None
    if not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", value):
        return None
    local, _, domain = value.partition("@")
    if not local or not domain or domain.startswith(".") or domain.endswith("."):
        return None
    return f"{local.lower()}@{domain.lower()}"


def validate_survey(form) -> tuple[dict[str, str], dict[str, str]]:
    cleaned: dict[str, str] = {}
    errors: dict[str, str] = {}

    for question in SURVEY_QUESTIONS:
        field = question["field"]
        raw = form.get(field, "")
        value = raw.strip() if isinstance(raw, str) else ""

        if question["type"] in {"radio", "image-radio"}:
            allowed = {option["value"] for option in question["options"]}
            if not value:
                errors[field] = "Please choose an option."
            elif value not in allowed:
                errors[field] = "Please choose one of the listed options."
            else:
                cleaned[field] = value
        elif field == "name":
            value = clean_name(value)
            if not value:
                errors[field] = "Please enter your name."
            elif len(value) > 80:
                errors[field] = "Your name must be 80 characters or fewer."
            else:
                cleaned[field] = value
        elif field == "email":
            normalized = validate_email(value)
            if not value:
                errors[field] = "Please enter your email address."
            elif normalized is None:
                errors[field] = "Please enter a valid email address."
            else:
                cleaned[field] = normalized
        elif field == "comments":
            value = clean_comments(value)
            if len(value) > 500:
                errors[field] = "Comments must be 500 characters or fewer."
            else:
                cleaned[field] = value
        elif question.get("required") and not value:
            errors[field] = "This answer is required."
        else:
            cleaned[field] = value

    if form.get("consent") != "on":
        errors["consent"] = (
            "Please confirm that you consent to your name and email being stored for this project."
        )

    return cleaned, errors


def issue_form_token() -> str:
    token = secrets.token_urlsafe(32)
    session["survey_form_token"] = token
    return token


def submission_signature(cleaned: dict[str, str]) -> str:
    material = "|".join(cleaned.get(field, "") for field in SURVEY_FIELDS)
    return hmac.new(
        app.secret_key.encode("utf-8"),
        material.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()


def check_session_rate_limit() -> tuple[bool, str | None]:
    now = time.time()
    times = [
        float(item)
        for item in session.get("survey_submit_times", [])
        if isinstance(item, (int, float)) and now - float(item) < 3600
    ]
    if times and now - max(times) < 15:
        return False, "Please wait a few seconds before submitting again."
    if len(times) >= 12:
        return False, "Too many submissions from this browser. Please try again later."
    return True, None


def record_session_submission(signature: str) -> None:
    now = time.time()
    times = [
        float(item)
        for item in session.get("survey_submit_times", [])
        if isinstance(item, (int, float)) and now - float(item) < 3600
    ]
    times.append(now)
    session["survey_submit_times"] = times
    session["survey_last_submission"] = {"signature": signature, "at": now}


def is_duplicate_in_session(signature: str) -> bool:
    previous = session.get("survey_last_submission")
    if not isinstance(previous, dict):
        return False
    try:
        previous_at = float(previous.get("at", 0))
    except (TypeError, ValueError):
        return False
    return (
        previous.get("signature") == signature
        and time.time() - previous_at < 120
    )

# ---------------------------------------------------------------------------
# Aggregate analytics and admin queries
# ---------------------------------------------------------------------------
PALETTES = {
    "blue": ["#2f6fed", "#6ea8fe", "#0b3d91", "#9bc5ff", "#4b7bec", "#86b7ff"],
    "teal": ["#0f9f9a", "#5ed3c8", "#0b6e6a", "#9ae6df", "#2ab7ad"],
    "orange": ["#e89018", "#f6bd60", "#b86b00", "#ffd89b", "#f2a541"],
    "red": ["#d94f4f", "#f28b82", "#9e2a2b", "#ffb3ae", "#e76f51"],
    "purple": ["#6f5bd3", "#a99af0", "#4b3b9e", "#c9c0ff", "#8b7ae8"],
}

CHART_DEFS = [
    {
        "key": "class_distribution",
        "field": "class_level",
        "title": "Class Distribution",
        "type": "bar",
        "options": CLASS_OPTIONS,
        "palette": PALETTES["blue"],
    },
    {
        "key": "breakfast_frequency",
        "field": "breakfast_frequency",
        "title": "Breakfast Frequency",
        "type": "doughnut",
        "options": FREQUENCY_OPTIONS,
        "palette": PALETTES["teal"],
    },
    {
        "key": "preferred_breakfast",
        "field": "preferred_breakfast",
        "title": "Preferred Breakfast",
        "type": "bar",
        "options": PREFERRED_OPTIONS,
        "palette": PALETTES["orange"],
        "horizontal": True,
    },
    {
        "key": "main_reason",
        "field": "main_reason",
        "title": "Reasons for Choosing Breakfast",
        "type": "doughnut",
        "options": REASON_OPTIONS,
        "palette": PALETTES["purple"],
    },
    {
        "key": "eating_time",
        "field": "eating_time",
        "title": "Time Spent Eating Breakfast",
        "type": "bar",
        "options": EATING_TIME_OPTIONS,
        "palette": PALETTES["blue"],
        "horizontal": True,
    },
    {
        "key": "test_familiarity",
        "field": "test_familiarity",
        "title": "Nutritional Test Familiarity",
        "type": "bar",
        "options": FAMILIARITY_OPTIONS,
        "palette": PALETTES["teal"],
        "horizontal": True,
    },
    {
        "key": "iodine_awareness",
        "field": "iodine_awareness",
        "title": "Iodine Test Awareness",
        "type": "doughnut",
        "options": YES_NO_NOT_SURE_OPTIONS,
        "palette": PALETTES["orange"],
    },
    {
        "key": "perceived_grease_spot",
        "field": "perceived_grease_spot",
        "title": "Grease-Spot Test Knowledge",
        "type": "bar",
        "options": GREASE_OPTIONS,
        "palette": PALETTES["red"],
        "horizontal": True,
    },
    {
        "key": "iodine_knowledge",
        "field": "iodine_knowledge",
        "title": "Iodine Test Knowledge",
        "type": "doughnut",
        "options": IODINE_KNOWLEDGE_OPTIONS,
        "palette": PALETTES["purple"],
    },
    {
        "key": "learning_method",
        "field": "learning_method",
        "title": "Preferred Learning Method",
        "type": "bar",
        "options": LEARNING_OPTIONS,
        "palette": PALETTES["blue"],
        "horizontal": True,
    },
    {
        "key": "choice_influence",
        "field": "choice_influence",
        "title": "Influence on Food Choices",
        "type": "bar",
        "options": INFLUENCE_OPTIONS,
        "palette": PALETTES["teal"],
        "horizontal": True,
    },
    {
        "key": "interesting_topic",
        "field": "interesting_topic",
        "title": "Most Interesting Chemistry Topic",
        "type": "bar",
        "options": INTEREST_OPTIONS,
        "palette": PALETTES["orange"],
        "horizontal": True,
    },
]


def query_all(sql: str, params: tuple = ()) -> list[sqlite3.Row]:
    connection = get_db()
    try:
        return connection.execute(sql, params).fetchall()
    finally:
        connection.close()


def fetch_counts(field: str) -> dict[str, int]:
    rows = query_all(
        f"SELECT {field} AS label, COUNT(*) AS total "
        f"FROM survey_responses GROUP BY {field}"
    )
    return {str(row["label"]): int(row["total"]) for row in rows if row["label"]}


def top_labels(pairs: list[tuple[str, int]]) -> str:
    populated = [(label, count) for label, count in pairs if count > 0]
    if not populated:
        return "No responses yet"
    best = max(count for _, count in populated)
    labels = [label for label, count in populated if count == best]
    if len(labels) == 1:
        return labels[0]
    if len(labels) == 2:
        return f"{labels[0]} & {labels[1]}"
    return ", ".join(labels[:-1]) + f" & {labels[-1]}"


def build_summary() -> dict:
    total_row = query_all("SELECT COUNT(*) AS total FROM survey_responses")[0]
    total = int(total_row["total"] or 0)
    last_row = query_all("SELECT MAX(submitted_at) AS last_at FROM survey_responses")[0]
    last_at = parse_db_datetime(last_row["last_at"])
    generated_at = datetime.now(timezone.utc)

    charts: dict[str, dict] = {}
    for definition in CHART_DEFS:
        mapping = fetch_counts(definition["field"])
        labels = list(definition["options"])
        for label in mapping:
            if label not in labels:
                labels.append(label)
        counts = [int(mapping.get(label, 0)) for label in labels]
        chart_total = sum(counts)
        percentages = [
            round((count / chart_total) * 100, 1) if chart_total else 0.0 for count in counts
        ]
        charts[definition["key"]] = {
            "key": definition["key"],
            "title": definition["title"],
            "type": definition["type"],
            "horizontal": bool(definition.get("horizontal")),
            "labels": labels,
            "counts": counts,
            "percentages": percentages,
            "total": chart_total,
            "colors": definition["palette"],
        }

    class_pairs = list(
        zip(
            charts["class_distribution"]["labels"],
            charts["class_distribution"]["counts"],
        )
    )
    breakfast_pairs = list(
        zip(
            charts["preferred_breakfast"]["labels"],
            charts["preferred_breakfast"]["counts"],
        )
    )
    frequency_pairs = list(
        zip(
            charts["breakfast_frequency"]["labels"],
            charts["breakfast_frequency"]["counts"],
        )
    )

    return {
        "total": total,
        "unique_classes": sum(1 for _, count in class_pairs if count > 0),
        "most_preferred_breakfast": top_labels(breakfast_pairs),
        "most_common_frequency": top_labels(frequency_pairs),
        "last_response_at": format_local(last_at),
        "last_response_at_utc": iso_utc(last_at),
        "generated_at": generated_at.isoformat().replace("+00:00", "Z"),
        "generated_at_local": format_local(generated_at.replace(tzinfo=None)),
        "privacy": (
            "Public results contain aggregate counts only. Names, emails, comments and "
            "individual submissions are never included."
        ),
        "charts": charts,
    }


def parse_filters(args) -> dict:
    class_level = args.get("class_level", "").strip()
    preferred = args.get("preferred_breakfast", "").strip()
    frequency = args.get("breakfast_frequency", "").strip()
    sort = args.get("sort", "newest").strip()
    per_page = parse_int(args.get("per_page"), 10)
    return {
        "q": args.get("q", "").strip()[:160],
        "class_level": class_level if class_level in CLASS_OPTIONS else "",
        "preferred_breakfast": preferred if preferred in PREFERRED_OPTIONS else "",
        "breakfast_frequency": frequency if frequency in FREQUENCY_OPTIONS else "",
        "date_from": parse_date(args.get("date_from")),
        "date_to": parse_date(args.get("date_to")),
        "sort": "oldest" if sort == "oldest" else "newest",
        "page": max(1, parse_int(args.get("page"), 1)),
        "per_page": per_page if per_page in {10, 25, 50, 100} else 10,
    }


def build_where(filters: dict) -> tuple[str, list]:
    clauses: list[str] = []
    params: list = []
    if filters["q"]:
        search_fields = [
            field for field in RESPONSE_COLUMNS if field != "submitted_at"
        ]
        clauses.append(
            "("
            + " OR ".join(
                f"COALESCE({field}, '') LIKE ? COLLATE NOCASE"
                for field in search_fields
            )
            + ")"
        )
        params.extend([f"%{filters['q']}%"] * len(search_fields))
    if filters["class_level"]:
        clauses.append("class_level = ?")
        params.append(filters["class_level"])
    if filters["preferred_breakfast"]:
        clauses.append("preferred_breakfast = ?")
        params.append(filters["preferred_breakfast"])
    if filters["breakfast_frequency"]:
        clauses.append("breakfast_frequency = ?")
        params.append(filters["breakfast_frequency"])
    if filters["date_from"]:
        clauses.append("submitted_at >= ?")
        params.append(local_date_start_utc(filters["date_from"]))
    if filters["date_to"]:
        clauses.append("submitted_at <= ?")
        params.append(local_date_end_utc(filters["date_to"]))
    if not clauses:
        return "", params
    return " WHERE " + " AND ".join(clauses), params


def ordered_rows(filters: dict) -> list[sqlite3.Row]:
    where_sql, params = build_where(filters)
    direction = "ASC" if filters["sort"] == "oldest" else "DESC"
    return query_all(
        "SELECT * FROM survey_responses"
        + where_sql
        + f" ORDER BY submitted_at {direction}, id {direction}",
        tuple(params),
    )


def fetch_admin_page(filters: dict) -> tuple[list[sqlite3.Row], int, int, int]:
    where_sql, params = build_where(filters)
    count_row = query_all(
        "SELECT COUNT(*) AS total FROM survey_responses" + where_sql,
        tuple(params),
    )[0]
    total = int(count_row["total"] or 0)
    pages = max(1, (total + filters["per_page"] - 1) // filters["per_page"])
    page = min(max(1, filters["page"]), pages)
    direction = "ASC" if filters["sort"] == "oldest" else "DESC"
    offset = (page - 1) * filters["per_page"]
    rows = query_all(
        "SELECT * FROM survey_responses"
        + where_sql
        + f" ORDER BY submitted_at {direction}, id {direction} LIMIT ? OFFSET ?",
        tuple(params) + (filters["per_page"], offset),
    )
    return rows, total, page, pages


def serialize_response(row: sqlite3.Row) -> dict:
    payload = {
        "ref": opaque_ref(int(row["id"])),
        "submitted_at": format_local(row["submitted_at"]),
        "submitted_at_utc": iso_utc(parse_db_datetime(row["submitted_at"])),
    }
    for field in RESPONSE_COLUMNS:
        if field == "submitted_at":
            continue
        if field == "consent_given":
            payload[field] = bool(row[field])
        else:
            payload[field] = row[field] or ""
    payload["comment_available"] = bool(str(payload.get("comments", "")).strip())
    return payload


def filters_payload(filters: dict) -> dict:
    return {
        "q": filters["q"],
        "class_level": filters["class_level"],
        "preferred_breakfast": filters["preferred_breakfast"],
        "breakfast_frequency": filters["breakfast_frequency"],
        "date_from": filters["date_from"].isoformat() if filters["date_from"] else "",
        "date_to": filters["date_to"].isoformat() if filters["date_to"] else "",
        "sort": filters["sort"],
        "page": filters["page"],
        "per_page": filters["per_page"],
    }


def spreadsheet_safe(value) -> str:
    text_value = "" if value is None else str(value)
    if text_value and text_value[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + text_value
    return text_value


EXPORT_HEADERS = [
    "Submission time (IST)",
    "Name",
    "Class",
    "Email",
    "Breakfast frequency",
    "Preferred breakfast",
    "Main reason",
    "Eating time",
    "Important nutrient",
    "Test familiarity",
    "Iodine test awareness",
    "Grease-spot answer",
    "Iodine test knowledge",
    "Preferred learning method",
    "Food-choice influence",
    "Most interesting topic",
    "Comments",
    "Consent given",
]

EXPORT_FIELDS = [
    "submitted_at",
    "name",
    "class_level",
    "email",
    "breakfast_frequency",
    "preferred_breakfast",
    "main_reason",
    "eating_time",
    "important_nutrient",
    "test_familiarity",
    "iodine_awareness",
    "perceived_grease_spot",
    "iodine_knowledge",
    "learning_method",
    "choice_influence",
    "interesting_topic",
    "comments",
    "consent_given",
]


def export_row(row: sqlite3.Row) -> list[str]:
    values: list[str] = []
    for field in EXPORT_FIELDS:
        if field == "submitted_at":
            values.append(format_local_date(row[field]))
        elif field == "consent_given":
            values.append("Yes" if row[field] else "No")
        else:
            values.append(str(row[field] or ""))
    return values


def build_csv(filters: dict) -> bytes:
    output = io.StringIO()
    writer = csv.writer(output, quoting=csv.QUOTE_MINIMAL, lineterminator="\n")
    writer.writerow(EXPORT_HEADERS)
    for row in ordered_rows(filters):
        writer.writerow([spreadsheet_safe(value) for value in export_row(row)])
    return output.getvalue().encode("utf-8-sig")


EXPORT_COLUMN_WEIGHTS = [
    1.4, 1.3, 0.7, 1.7, 1.3, 1.4, 1.3, 1.1, 1.1, 1.1, 1.1, 1.1, 1.1, 1.3, 1.1, 1.4, 2.1, 0.6,
]


def filters_summary(filters: dict) -> str:
    parts: list[str] = []
    if filters.get("q"):
        parts.append(f'search "{filters["q"]}"')
    if filters.get("class_level"):
        parts.append(f"class {filters['class_level']}")
    if filters.get("preferred_breakfast"):
        parts.append(f"preferred breakfast {filters['preferred_breakfast']}")
    if filters.get("breakfast_frequency"):
        parts.append(f"frequency {filters['breakfast_frequency']}")
    if filters.get("date_from"):
        parts.append(f"from {filters['date_from'].isoformat()}")
    if filters.get("date_to"):
        parts.append(f"to {filters['date_to'].isoformat()}")
    if filters.get("sort") == "oldest":
        parts.append("oldest first")
    return "Filters: " + "; ".join(parts) if parts else "Filters: none"


def shade_docx_cell(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:val"), "clear")
    shading.set(qn("w:color"), "auto")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def build_docx(filters: dict) -> io.BytesIO:
    rows = ordered_rows(filters)
    document = Document()
    section = document.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width = Cm(29.7)
    section.page_height = Cm(21.0)
    for margin in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
        setattr(section, margin, Cm(1.27))

    title = document.add_paragraph()
    title_run = title.add_run("Chemistry IP — Survey Responses")
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)

    meta = document.add_paragraph()
    meta_run = meta.add_run(
        f"Generated {format_local(datetime.now(timezone.utc))} · {len(rows)} response(s)"
    )
    meta_run.font.size = Pt(8)
    meta_run.font.color.rgb = RGBColor(0x41, 0x50, 0x6B)

    filter_line = document.add_paragraph()
    filter_run = filter_line.add_run(filters_summary(filters))
    filter_run.font.size = Pt(8)
    filter_run.font.color.rgb = RGBColor(0x41, 0x50, 0x6B)

    table = document.add_table(rows=1, cols=len(EXPORT_HEADERS))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    usable_width = section.page_width - section.left_margin - section.right_margin
    weight_total = sum(EXPORT_COLUMN_WEIGHTS)
    column_widths = [
        int(usable_width * weight / weight_total) for weight in EXPORT_COLUMN_WEIGHTS
    ]
    for index, width in enumerate(column_widths):
        table.columns[index].width = width

    header_cells = table.rows[0].cells
    header_row = table.rows[0]
    tr_pr = header_row._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    tr_pr.append(repeat)
    for index, header in enumerate(EXPORT_HEADERS):
        cell = header_cells[index]
        cell.width = column_widths[index]
        shade_docx_cell(cell, "0B2545")
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run(header)
        run.font.size = Pt(7)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    if not rows:
        empty_row = table.add_row()
        empty_cell = empty_row.cells[0]
        empty_cell.merge(empty_row.cells[-1])
        empty_run = empty_cell.paragraphs[0].add_run(
            "No survey responses match the current filters."
        )
        empty_run.font.size = Pt(8)
    else:
        for row_index, row in enumerate(rows):
            cells = table.add_row().cells
            row_cant_split = OxmlElement("w:cantSplit")
            table.rows[-1]._tr.get_or_add_trPr().append(row_cant_split)
            for cell_index, value in enumerate(export_row(row)):
                cell = cells[cell_index]
                cell.width = column_widths[cell_index]
                if row_index % 2 == 1:
                    shade_docx_cell(cell, "F5F9FF")
                run = cell.paragraphs[0].add_run(str(value))
                run.font.size = Pt(6.5)

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output


def build_pdf(filters: dict) -> io.BytesIO:
    rows = ordered_rows(filters)
    output = io.BytesIO()
    page_size = landscape(A4)
    document = SimpleDocTemplate(
        output,
        pagesize=page_size,
        leftMargin=12 * mm,
        rightMargin=12 * mm,
        topMargin=12 * mm,
        bottomMargin=12 * mm,
        title="Chemistry IP Survey Responses",
        author="Chemistry IP",
    )
    title_style = ParagraphStyle(
        "ExportTitle",
        fontName="Helvetica-Bold",
        fontSize=15,
        leading=18,
        textColor=rl_colors.HexColor("#0B2545"),
        spaceAfter=3,
    )
    meta_style = ParagraphStyle(
        "ExportMeta",
        fontName="Helvetica",
        fontSize=8,
        leading=11,
        textColor=rl_colors.HexColor("#41506B"),
    )
    header_style = ParagraphStyle(
        "ExportHeader",
        fontName="Helvetica-Bold",
        fontSize=6.5,
        leading=8,
        textColor=rl_colors.white,
    )
    cell_style = ParagraphStyle(
        "ExportCell",
        fontName="Helvetica",
        fontSize=6,
        leading=7.5,
        textColor=rl_colors.HexColor("#111827"),
    )

    def pdf_cell(value: str, style: ParagraphStyle) -> Paragraph:
        text_value = xml_escape(str(value or "")).replace("\r\n", "\n").replace("\n", "<br/>")
        return Paragraph(text_value or "&nbsp;", style)

    story = [
        Paragraph("Chemistry IP — Survey Responses", title_style),
        Paragraph(
            f"Generated {xml_escape(format_local(datetime.now(timezone.utc)))}"
            f" · {len(rows)} response(s)",
            meta_style,
        ),
        Paragraph(xml_escape(filters_summary(filters)), meta_style),
        Spacer(1, 6),
    ]

    data: list[list] = [
        [pdf_cell(header, header_style) for header in EXPORT_HEADERS]
    ]
    if rows:
        for row in rows:
            data.append([pdf_cell(value, cell_style) for value in export_row(row)])
    else:
        data.append(
            [pdf_cell("No survey responses match the current filters.", cell_style)]
            + [Paragraph("", cell_style) for _ in EXPORT_HEADERS[1:]]
        )

    usable_width = page_size[0] - 24 * mm
    weight_total = sum(EXPORT_COLUMN_WEIGHTS)
    column_widths = [
        usable_width * weight / weight_total for weight in EXPORT_COLUMN_WEIGHTS
    ]
    table = Table(data, colWidths=column_widths, repeatRows=1)
    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0B2545")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.25, rl_colors.HexColor("#C9D6E8")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F5F9FF")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 2),
        ("RIGHTPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 1.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
    ]
    if not rows:
        table_style.append(("SPAN", (0, 1), (-1, 1)))
    table.setStyle(TableStyle(table_style))
    story.append(table)
    document.build(story)
    output.seek(0)
    return output


def decode_chart_images(payload) -> dict[str, bytes]:
    images: dict[str, bytes] = {}
    if not isinstance(payload, dict):
        return images
    submitted = payload.get("charts")
    if not isinstance(submitted, list):
        return images
    valid_keys = {definition["key"] for definition in CHART_DEFS}
    for item in submitted[: len(valid_keys) + 5]:
        if not isinstance(item, dict):
            continue
        key = item.get("key")
        data = item.get("image")
        if key not in valid_keys or not isinstance(data, str):
            continue
        if not data.startswith("data:image/png;base64,"):
            continue
        try:
            raw = base64.b64decode(data.split(",", 1)[1], validate=True)
        except (ValueError, binascii.Error):
            continue
        if 0 < len(raw) <= 4_000_000:
            images[key] = raw
    return images


def chart_kind_label(chart: dict) -> str:
    if chart.get("type") == "doughnut":
        return "Doughnut chart"
    return "Horizontal bar chart" if chart.get("horizontal") else "Bar chart"


def chart_report_summary_rows(summary: dict) -> list[tuple[str, str]]:
    return [
        ("Total responses", str(summary["total"])),
        ("Classes represented", str(summary["unique_classes"])),
        ("Most preferred breakfast", summary["most_preferred_breakfast"]),
        ("Most common frequency", summary["most_common_frequency"]),
        ("Last response (IST)", summary["last_response_at"]),
    ]


def build_chart_report_pdf(images: dict[str, bytes]) -> io.BytesIO:
    summary = build_summary()
    charts = summary["charts"]
    generated = format_local(datetime.now(timezone.utc))
    output = io.BytesIO()
    document = SimpleDocTemplate(
        output,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=16 * mm,
        bottomMargin=18 * mm,
        title="Chemistry IP — Aggregate Survey Charts",
        author="Chemistry IP",
    )
    title_style = ParagraphStyle(
        "ChartTitle",
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=rl_colors.HexColor("#0B2545"),
        spaceAfter=4,
    )
    heading_style = ParagraphStyle(
        "ChartHeading",
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=17,
        textColor=rl_colors.HexColor("#0B2545"),
        spaceBefore=2,
        spaceAfter=3,
    )
    meta_style = ParagraphStyle(
        "ChartMeta",
        fontName="Helvetica",
        fontSize=8.5,
        leading=11,
        textColor=rl_colors.HexColor("#41506B"),
    )
    cell_style = ParagraphStyle(
        "ChartCell",
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=rl_colors.HexColor("#17263B"),
    )
    header_cell_style = ParagraphStyle(
        "ChartHeaderCell",
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=10,
        textColor=rl_colors.white,
    )

    story: list = [
        Paragraph("Aggregate Survey Charts", title_style),
        Paragraph(f"Chemistry IP · Generated {xml_escape(generated)}", meta_style),
        Spacer(1, 8),
        Paragraph(
            "A structured report of every aggregate chart on the admin dashboard. "
            "The report contains aggregate counts only — no names, emails, comments or "
            "individual records are included.",
            meta_style,
        ),
        Spacer(1, 10),
        Paragraph("Survey overview", heading_style),
    ]

    summary_rows = [
        [
            Paragraph("Metric", header_cell_style),
            Paragraph("Value", header_cell_style),
        ]
    ]
    for label, value in chart_report_summary_rows(summary):
        summary_rows.append(
            [Paragraph(xml_escape(label), cell_style), Paragraph(xml_escape(str(value)), cell_style)]
        )
    summary_table = Table(summary_rows, colWidths=[58 * mm, document.width - 58 * mm])
    summary_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0B2545")),
                ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#C9D6E8")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [rl_colors.white, rl_colors.HexColor("#F5F9FF")]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ]
        )
    )
    story.append(summary_table)

    for index, definition in enumerate(CHART_DEFS, start=1):
        chart = charts.get(definition["key"])
        if not chart:
            continue
        story.append(PageBreak())
        story.append(Paragraph(f"{index}. {xml_escape(definition['title'])}", heading_style))
        story.append(
            Paragraph(
                xml_escape(f"{chart_kind_label(chart)} · {chart['total']} response(s)"),
                meta_style,
            )
        )
        story.append(Spacer(1, 6))
        raw = images.get(definition["key"])
        if raw:
            try:
                reader = ImageReader(io.BytesIO(raw))
                pixel_width, pixel_height = reader.getSize()
                max_width = document.width
                max_height = 120 * mm
                scale = min(max_width / pixel_width, max_height / pixel_height)
                story.append(
                    Image(
                        io.BytesIO(raw),
                        width=pixel_width * scale,
                        height=pixel_height * scale,
                    )
                )
                story.append(Spacer(1, 8))
            except Exception:
                app.logger.exception("Could not embed chart image in PDF report")

        table_rows = [
            [
                Paragraph("Option", header_cell_style),
                Paragraph("Responses", header_cell_style),
                Paragraph("Share", header_cell_style),
            ]
        ]
        for label, count, percentage in zip(
            chart["labels"], chart["counts"], chart["percentages"]
        ):
            table_rows.append(
                [
                    Paragraph(xml_escape(str(label)), cell_style),
                    Paragraph(str(count), cell_style),
                    Paragraph(f"{percentage}%", cell_style),
                ]
            )
        table_rows.append(
            [
                Paragraph("<b>Total</b>", cell_style),
                Paragraph(f"<b>{chart['total']}</b>", cell_style),
                Paragraph("<b>100%</b>" if chart["total"] else "<b>0%</b>", cell_style),
            ]
        )
        data_table = Table(
            table_rows,
            colWidths=[document.width - 40 * mm, 20 * mm, 20 * mm],
            repeatRows=1,
        )
        data_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), rl_colors.HexColor("#0B2545")),
                    ("GRID", (0, 0), (-1, -1), 0.4, rl_colors.HexColor("#C9D6E8")),
                    ("ROWBACKGROUNDS", (0, 1), (-1, -2), [rl_colors.white, rl_colors.HexColor("#F5F9FF")]),
                    ("BACKGROUND", (0, -1), (-1, -1), rl_colors.HexColor("#EAF1FF")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ]
            )
        )
        story.append(data_table)

    def draw_footer(canvas, doc) -> None:
        canvas.saveState()
        canvas.setStrokeColor(rl_colors.HexColor("#D8E7F5"))
        canvas.setLineWidth(0.5)
        canvas.line(18 * mm, 14 * mm, A4[0] - 18 * mm, 14 * mm)
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(rl_colors.HexColor("#5F7087"))
        canvas.drawString(18 * mm, 10 * mm, "Chemistry IP — Aggregate Survey Charts")
        canvas.drawRightString(
            A4[0] - 18 * mm, 10 * mm, f"Page {canvas.getPageNumber()}"
        )
        canvas.restoreState()

    document.build(story, onFirstPage=draw_footer, onLaterPages=draw_footer)
    output.seek(0)
    return output


def add_docx_page_number_footer(document) -> None:
    footer = document.sections[0].footer
    paragraph = footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run("Chemistry IP — Aggregate Survey Charts · Page ")
    run.font.size = Pt(8)
    field_run = paragraph.add_run()
    field_run.font.size = Pt(8)
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    field_run._r.append(begin)
    field_run._r.append(instruction)
    field_run._r.append(end)


def build_chart_report_docx(images: dict[str, bytes]) -> io.BytesIO:
    summary = build_summary()
    charts = summary["charts"]
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    for margin in ("left_margin", "right_margin"):
        setattr(section, margin, Cm(1.9))
    section.top_margin = Cm(1.7)
    section.bottom_margin = Cm(1.7)
    add_docx_page_number_footer(document)

    title = document.add_paragraph()
    title_run = title.add_run("Aggregate Survey Charts")
    title_run.font.size = Pt(20)
    title_run.font.bold = True
    title_run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)

    meta = document.add_paragraph()
    meta_run = meta.add_run(
        f"Chemistry IP · Generated {format_local(datetime.now(timezone.utc))}"
    )
    meta_run.font.size = Pt(9)
    meta_run.font.color.rgb = RGBColor(0x41, 0x50, 0x6B)

    note = document.add_paragraph()
    note_run = note.add_run(
        "A structured report of every aggregate chart on the admin dashboard. "
        "The report contains aggregate counts only — no names, emails, comments or "
        "individual records are included."
    )
    note_run.font.size = Pt(9)
    note_run.font.color.rgb = RGBColor(0x41, 0x50, 0x6B)

    overview_heading = document.add_paragraph()
    overview_run = overview_heading.add_run("Survey overview")
    overview_run.font.size = Pt(14)
    overview_run.font.bold = True
    overview_run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)

    summary_table = document.add_table(rows=1, cols=2)
    summary_table.style = "Table Grid"
    summary_table.autofit = False
    summary_widths = (Cm(6.0), Cm(11.2))
    for index, header in enumerate(("Metric", "Value")):
        cell = summary_table.rows[0].cells[index]
        cell.width = summary_widths[index]
        shade_docx_cell(cell, "0B2545")
        run = cell.paragraphs[0].add_run(header)
        run.font.size = Pt(8.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for row_index, (label, value) in enumerate(chart_report_summary_rows(summary)):
        cells = summary_table.add_row().cells
        for index, text_value in enumerate((label, value)):
            cells[index].width = summary_widths[index]
            if row_index % 2 == 1:
                shade_docx_cell(cells[index], "F5F9FF")
            run = cells[index].paragraphs[0].add_run(str(text_value))
            run.font.size = Pt(8.5)

    for index, definition in enumerate(CHART_DEFS, start=1):
        chart = charts.get(definition["key"])
        if not chart:
            continue
        document.add_page_break()
        heading = document.add_paragraph()
        heading_run = heading.add_run(f"{index}. {definition['title']}")
        heading_run.font.size = Pt(14)
        heading_run.font.bold = True
        heading_run.font.color.rgb = RGBColor(0x0B, 0x25, 0x45)

        subtitle = document.add_paragraph()
        subtitle_run = subtitle.add_run(
            f"{chart_kind_label(chart)} · {chart['total']} response(s)"
        )
        subtitle_run.font.size = Pt(8.5)
        subtitle_run.font.color.rgb = RGBColor(0x41, 0x50, 0x6B)

        raw = images.get(definition["key"])
        if raw:
            try:
                document.add_picture(io.BytesIO(raw), width=Cm(16.4))
                document.add_paragraph()
            except Exception:
                app.logger.exception("Could not embed chart image in Word report")

        table = document.add_table(rows=1, cols=3)
        table.style = "Table Grid"
        table.autofit = False
        column_widths = (Cm(10.4), Cm(3.0), Cm(3.0))
        for column_index, header in enumerate(("Option", "Responses", "Share")):
            cell = table.rows[0].cells[column_index]
            cell.width = column_widths[column_index]
            shade_docx_cell(cell, "0B2545")
            run = cell.paragraphs[0].add_run(header)
            run.font.size = Pt(8.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        body_rows = list(zip(chart["labels"], chart["counts"], chart["percentages"]))
        for row_index, (label, count, percentage) in enumerate(body_rows):
            cells = table.add_row().cells
            for column_index, text_value in enumerate((label, count, f"{percentage}%")):
                cells[column_index].width = column_widths[column_index]
                if row_index % 2 == 1:
                    shade_docx_cell(cells[column_index], "F5F9FF")
                run = cells[column_index].paragraphs[0].add_run(str(text_value))
                run.font.size = Pt(8.5)

        total_cells = table.add_row().cells
        total_values = ("Total", chart["total"], "100%" if chart["total"] else "0%")
        for column_index, text_value in enumerate(total_values):
            total_cells[column_index].width = column_widths[column_index]
            shade_docx_cell(total_cells[column_index], "EAF1FF")
            run = total_cells[column_index].paragraphs[0].add_run(str(text_value))
            run.font.size = Pt(8.5)
            run.font.bold = True

    output = io.BytesIO()
    document.save(output)
    output.seek(0)
    return output

# ---------------------------------------------------------------------------
# Embedded styles and behaviour
# ---------------------------------------------------------------------------
CSS = r""":root {
  --navy: #0b2545;
  --navy-2: #13315c;
  --navy-3: #1d4475;
  --ink: #17263b;
  --muted: #5f7087;
  --soft: #8290a4;
  --line: #d8e7f5;
  --line-strong: #bdd5ea;
  --paper: #ffffff;
  --bg: #f4f8fd;
  --bg-2: #eaf3fb;
  --blue: #2f6fed;
  --blue-soft: #eaf1ff;
  --teal: #0f9f9a;
  --teal-soft: #e6f8f6;
  --orange: #e89018;
  --orange-soft: #fff4df;
  --red: #d94f4f;
  --red-soft: #fff0ef;
  --purple: #6f5bd3;
  --purple-soft: #f1efff;
  --green: #2d8a65;
  --green-soft: #e9f8f0;
  --shadow: 0 18px 44px rgba(11, 37, 69, 0.07);
  --shadow-sm: 0 8px 24px rgba(11, 37, 69, 0.06);
  --radius: 22px;
  --radius-sm: 14px;
  --serif: Georgia, "Times New Roman", "Book Antiqua", serif;
  --sans: "Segoe UI", system-ui, -apple-system, BlinkMacSystemFont, Arial, sans-serif;
  --mono: "Cascadia Mono", Consolas, "Courier New", monospace;
  --sidebar: 278px;
}

* { box-sizing: border-box; }
html { scroll-behavior: smooth; }
body {
  margin: 0;
  color: var(--ink);
  background:
    radial-gradient(circle at 86% 8%, rgba(47, 111, 237, 0.075), transparent 30rem),
    radial-gradient(circle at 8% 82%, rgba(15, 159, 154, 0.055), transparent 28rem),
    var(--bg);
  font-family: var(--sans);
  font-size: 16px;
  line-height: 1.65;
  -webkit-font-smoothing: antialiased;
  text-rendering: optimizeLegibility;
}
body.nav-open { overflow: hidden; }
a { color: var(--blue); text-decoration-thickness: 1px; text-underline-offset: 3px; }
a:hover { color: var(--navy); }
img { max-width: 100%; display: block; }
button, input, select, textarea { font: inherit; }
button { cursor: pointer; }
h1, h2, h3, h4, p { margin-top: 0; }
h1, h2, h3, h4 {
  font-family: var(--serif);
  font-weight: 500;
  line-height: 1.12;
  letter-spacing: -0.025em;
  color: var(--navy);
}
h1 { font-size: clamp(2.35rem, 5vw, 4.65rem); }
h2 { font-size: clamp(1.7rem, 3vw, 2.65rem); }
h3 { font-size: clamp(1.18rem, 2vw, 1.5rem); }
h4 { font-size: 1.03rem; }
::selection { background: #cfe2ff; color: var(--navy); }
:focus-visible { outline: 3px solid rgba(47, 111, 237, 0.42); outline-offset: 3px; }

.skip-link {
  position: fixed; top: 10px; left: 10px; z-index: 100;
  transform: translateY(-160%);
  background: var(--navy); color: #fff; padding: 10px 16px; border-radius: 8px;
  font-weight: 700; text-decoration: none;
}
.skip-link:focus { transform: none; color: #fff; }
.icon { width: 1.15em; height: 1.15em; flex: 0 0 auto; vertical-align: -0.15em; fill: none; stroke: currentColor; stroke-width: 1.8; stroke-linecap: round; stroke-linejoin: round; }
.icon-sprite { position: absolute; width: 0; height: 0; overflow: hidden; }

/* Sidebar and top bar */
.app-shell { min-height: 100vh; }
.sidebar {
  position: fixed; inset: 0 auto 0 0; width: var(--sidebar); z-index: 40;
  display: flex; flex-direction: column;
  padding: 26px 20px 20px;
  color: #e9f2fb;
  background:
    radial-gradient(circle at 20% 0%, rgba(86, 151, 235, 0.22), transparent 19rem),
    linear-gradient(168deg, #0b2545 0%, #0d2b50 58%, #0a203c 100%);
  border-right: 1px solid rgba(255, 255, 255, 0.08);
  overflow-y: auto;
}
.brand {
  display: flex; align-items: center; gap: 12px; color: #fff; text-decoration: none;
  padding: 2px 8px 18px;
}
.brand:hover { color: #fff; }
.brand-mark {
  width: 42px; height: 42px; display: grid; place-items: center;
  border-radius: 13px; color: #d9f8f3;
  background: linear-gradient(145deg, rgba(49, 157, 214, 0.3), rgba(15, 159, 154, 0.18));
  border: 1px solid rgba(180, 235, 255, 0.26);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, 0.2);
}
.brand-mark .icon { width: 22px; height: 22px; }
.brand strong { display: block; font-family: var(--serif); font-size: 1.34rem; font-weight: 600; letter-spacing: -0.02em; }
.brand small { display: block; color: #a9c6e0; font-size: 0.7rem; letter-spacing: 0.08em; text-transform: uppercase; margin-top: 1px; }
.sidebar-project {
  margin: 0 8px 22px; padding: 12px 13px;
  color: #c5d9eb; font-size: 0.78rem; line-height: 1.45;
  border-left: 2px solid rgba(94, 211, 200, 0.75);
  background: rgba(255, 255, 255, 0.045);
  border-radius: 0 9px 9px 0;
}
.side-nav { display: grid; gap: 16px; }
.nav-group > p {
  margin: 0 0 5px 10px; color: #7699b8; font-size: 0.64rem;
  letter-spacing: 0.16em; text-transform: uppercase; font-weight: 700;
}
.nav-link {
  display: block; color: #d7e6f4; text-decoration: none; font-size: 0.88rem;
  padding: 7px 11px; border-radius: 9px; border: 1px solid transparent;
  transition: background 160ms ease, color 160ms ease, border-color 160ms ease;
}
.nav-link:hover { background: rgba(255, 255, 255, 0.07); color: #fff; }
.nav-link.active {
  background: rgba(255, 255, 255, 0.12); color: #fff; font-weight: 700;
  border-color: rgba(160, 220, 255, 0.18);
  box-shadow: inset 3px 0 0 #5ed3c8;
}
.sidebar-foot { margin-top: auto; padding: 24px 9px 4px; border-top: 1px solid rgba(255, 255, 255, 0.1); }
.sidebar-foot a { color: #d8ecfa; font-size: 0.8rem; }
.sidebar-foot p { color: #7699b8; font-size: 0.68rem; margin: 8px 0 0; line-height: 1.45; }
.scrim { display: none; }
.main-wrap { margin-left: var(--sidebar); min-height: 100vh; display: flex; flex-direction: column; }
.topbar {
  position: sticky; top: 0; z-index: 30;
  min-height: 70px; display: flex; align-items: center; gap: 18px;
  padding: 10px clamp(18px, 4vw, 48px);
  background: rgba(250, 253, 255, 0.88);
  border-bottom: 1px solid rgba(189, 213, 234, 0.7);
  backdrop-filter: blur(16px);
}
.menu-btn {
  display: none; align-items: center; gap: 7px; border: 1px solid var(--line);
  background: #fff; color: var(--navy); border-radius: 9px; padding: 8px 11px;
  font-size: 0.8rem; font-weight: 700;
}
.topbar-context { min-width: 0; display: flex; flex-direction: column; line-height: 1.25; }
.topbar-context span { color: var(--soft); font-size: 0.66rem; letter-spacing: 0.12em; text-transform: uppercase; font-weight: 700; }
.topbar-context strong { color: var(--navy); font-family: var(--serif); font-size: 1.04rem; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.topbar .btn { margin-left: auto; }
main { flex: 1; }
.page { width: min(1240px, 100%); margin: 0 auto; padding: 34px clamp(18px, 4vw, 48px) 72px; }
.page.narrow { width: min(980px, 100%); }

/* Buttons and chips */
.btn {
  display: inline-flex; align-items: center; justify-content: center; gap: 9px;
  min-height: 46px; padding: 11px 18px; border-radius: 12px;
  border: 1px solid transparent; background: var(--navy); color: #fff;
  font-size: 0.86rem; font-weight: 700; letter-spacing: 0.01em; text-decoration: none;
  box-shadow: 0 8px 18px rgba(11, 37, 69, 0.14);
  transition: transform 160ms ease, box-shadow 160ms ease, background 160ms ease;
}
.btn:hover { background: var(--navy-2); color: #fff; transform: translateY(-1px); box-shadow: 0 11px 24px rgba(11, 37, 69, 0.19); }
.btn:active { transform: none; }
.btn .icon { width: 17px; height: 17px; }
.btn-small { min-height: 38px; padding: 8px 13px; font-size: 0.78rem; border-radius: 10px; box-shadow: none; }
.btn-light { background: #fff; color: var(--navy); border-color: var(--line-strong); box-shadow: none; }
.btn-light:hover { background: var(--bg); color: var(--navy); }
.btn-teal { background: var(--teal); box-shadow: 0 8px 18px rgba(15, 159, 154, 0.17); }
.btn-teal:hover { background: #0c8a86; }
.btn-ghost { background: transparent; color: var(--navy); border-color: var(--line); box-shadow: none; }
.btn-ghost:hover { background: #fff; color: var(--navy); }
.btn-danger { background: var(--red); }
.btn[disabled], .btn.is-disabled { opacity: 0.55; cursor: not-allowed; transform: none; }
.text-link { font-size: 0.82rem; font-weight: 700; text-decoration: none; display: inline-flex; align-items: center; gap: 6px; }
.text-link .icon { width: 15px; height: 15px; }
.pill {
  display: inline-flex; align-items: center; gap: 7px; width: fit-content;
  padding: 6px 10px; border: 1px solid var(--line); border-radius: 999px;
  background: #fff; color: var(--muted); font-size: 0.67rem; font-weight: 700;
  letter-spacing: 0.08em; text-transform: uppercase;
}
.pill .dot { width: 6px; height: 6px; border-radius: 50%; background: var(--teal); box-shadow: 0 0 0 3px var(--teal-soft); }
.kicker { display: flex; align-items: center; gap: 8px; color: var(--blue); font-size: 0.7rem; font-weight: 800; letter-spacing: 0.15em; text-transform: uppercase; margin-bottom: 10px; }
.kicker::before { content: ""; width: 22px; height: 2px; background: var(--teal); border-radius: 2px; }
.lede { color: var(--muted); font-size: 1.02rem; line-height: 1.7; max-width: 68ch; }
.muted { color: var(--muted); }
.small { font-size: 0.78rem; }
.mono { font-family: var(--mono); }

/* Hero and section blocks */
.hero {
  position: relative; overflow: hidden; color: #eaf4ff;
  display: grid; grid-template-columns: minmax(0, 1.16fr) minmax(280px, 0.84fr);
  min-height: 470px; margin-bottom: 28px;
  background:
    radial-gradient(circle at 87% 10%, rgba(53, 158, 220, 0.24), transparent 25rem),
    linear-gradient(142deg, #0b2545 0%, #102f57 68%, #0d294a 100%);
  border-radius: 30px; box-shadow: 0 26px 70px rgba(11, 37, 69, 0.18);
  border: 1px solid rgba(255, 255, 255, 0.1);
}
.hero-copy { padding: clamp(30px, 5vw, 58px); display: flex; flex-direction: column; justify-content: center; align-items: flex-start; position: relative; z-index: 2; }
.hero .kicker { color: #8cded8; }
.hero .kicker::before { background: #8cded8; }
.hero h1 { color: #fff; margin: 0 0 15px; max-width: 12ch; }
.hero h1 .accent { color: #9be7df; font-style: italic; }
.hero-subtitle { color: #b7cee3; font-size: clamp(0.95rem, 1.6vw, 1.16rem); margin-bottom: 15px; }
.hero-lede { color: #c8dbee; max-width: 55ch; margin-bottom: 25px; }
.hero-actions { display: flex; flex-wrap: wrap; gap: 10px; }
.hero .btn-light { background: #fff; border-color: #fff; color: var(--navy); }
.hero .btn-ghost { color: #e6f2fb; border-color: rgba(255, 255, 255, 0.24); }
.hero .btn-ghost:hover { background: rgba(255, 255, 255, 0.09); color: #fff; }
.hero-metrics { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 28px; }
.hero-metric { min-width: 104px; padding: 10px 13px; border: 1px solid rgba(255, 255, 255, 0.15); border-radius: 12px; background: rgba(255, 255, 255, 0.06); }
.hero-metric strong { display: block; font-family: var(--serif); color: #fff; font-size: 1.3rem; line-height: 1; }
.hero-metric span { display: block; color: #a9c5df; font-size: 0.67rem; margin-top: 5px; }
.hero-media { position: relative; min-height: 330px; overflow: hidden; background: #173b62; }
.hero-media::after { content: ""; position: absolute; inset: 0; background: linear-gradient(90deg, rgba(11, 37, 69, 0.48), transparent 42%), linear-gradient(0deg, rgba(11, 37, 69, 0.46), transparent 32%); pointer-events: none; }
.hero-media img { width: 100%; height: 100%; object-fit: cover; position: absolute; inset: 0; }
.hero-media.media-fallback { background: radial-gradient(circle at 70% 20%, #2d6d91, #15385e 56%, #0b2545); }
.hero-media.media-fallback::before { content: ""; position: absolute; width: 240px; height: 240px; border: 1px solid rgba(255,255,255,.18); border-radius: 50%; right: 12%; top: 17%; box-shadow: 0 0 0 26px rgba(255,255,255,.04), 0 0 0 54px rgba(255,255,255,.025); }
.hero-ticker {
  position: absolute; left: 18px; right: 18px; bottom: 18px; z-index: 3;
  display: flex; flex-wrap: wrap; gap: 6px; align-items: center;
  padding: 9px 11px; border-radius: 11px; background: rgba(7, 27, 51, 0.68);
  border: 1px solid rgba(255,255,255,.12); backdrop-filter: blur(8px);
  color: #cfe4f4; font-size: 0.67rem;
}
.hero-ticker span { padding: 2px 7px; border-radius: 999px; background: rgba(255,255,255,.08); }
.hero-ticker small { margin-left: auto; color: #9db9d1; }
.section { margin: 52px 0; }
.section-head { display: flex; justify-content: space-between; align-items: flex-end; gap: 20px; margin-bottom: 20px; }
.section-head > div { max-width: 720px; }
.section-head h2 { margin-bottom: 8px; }
.section-head .lede { margin-bottom: 0; }
.card {
  background: var(--paper); border: 1px solid var(--line); border-radius: var(--radius);
  box-shadow: var(--shadow); padding: clamp(20px, 3vw, 30px);
}
.card.flat { box-shadow: none; }
.card-tight { padding: 18px; }
.grid { display: grid; gap: 16px; }
.grid-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.grid-3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.grid-4 { grid-template-columns: repeat(4, minmax(0, 1fr)); }
.split { display: grid; grid-template-columns: minmax(0, 1.08fr) minmax(280px, 0.92fr); gap: 22px; align-items: stretch; }
.prose { max-width: 74ch; }
.prose p { color: #45566d; margin-bottom: 16px; }
.prose .lead { color: var(--ink); font-size: 1.1rem; line-height: 1.65; }
.note {
  margin-top: 20px; padding: 15px 17px; border-left: 3px solid var(--teal);
  border-radius: 0 12px 12px 0; background: var(--teal-soft); color: #265a57; font-size: 0.87rem;
}
.note.warning { border-color: var(--orange); background: var(--orange-soft); color: #735018; }
.note.info { border-color: var(--blue); background: var(--blue-soft); color: #244d93; }
.note.danger { border-color: var(--red); background: var(--red-soft); color: #7c2b2b; }
.callout { display: flex; gap: 14px; padding: 20px; border-radius: 16px; border: 1px solid var(--line); background: #fff; box-shadow: var(--shadow-sm); }
.callout .callout-icon { width: 40px; height: 40px; flex: 0 0 40px; display: grid; place-items: center; border-radius: 12px; background: var(--blue-soft); color: var(--blue); }
.callout.note { border-left: 1px solid var(--line); margin: 0; background: #fff; color: var(--ink); }
.callout.note .callout-icon { background: var(--teal-soft); color: var(--teal); }
.callout.warning .callout-icon { background: var(--orange-soft); color: var(--orange); }
.callout h3 { margin-bottom: 5px; }
.callout p { margin: 0; color: var(--muted); font-size: 0.9rem; }
.aim-card { display: grid; grid-template-columns: 1fr 1fr; gap: 28px; align-items: start; }
.aim-text { padding: 8px 4px; }
.aim-text h2 { margin-bottom: 14px; }
.aim-text p { color: #45566d; }
.objectives { list-style: none; padding: 0; margin: 0; display: grid; gap: 12px; counter-reset: aim; }
.objectives li { position: relative; padding: 15px 16px 15px 54px; border: 1px solid var(--line); border-radius: 14px; background: #fbfdff; color: #43546b; font-size: 0.9rem; counter-increment: aim; }
.objectives li::before { content: counter(aim, decimal-leading-zero); position: absolute; left: 16px; top: 14px; color: var(--teal); font-family: var(--mono); font-size: 0.72rem; font-weight: 700; }
.metric-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
.metric { padding: 18px; border: 1px solid var(--line); border-radius: 15px; background: #fff; }
.metric span { display: block; color: var(--soft); font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.1em; font-weight: 700; }
.metric strong { display: block; margin-top: 8px; font-family: var(--serif); font-size: clamp(1.3rem, 2.5vw, 2rem); color: var(--navy); line-height: 1.1; }
.metric small { display: block; color: var(--muted); font-size: 0.72rem; margin-top: 5px; }

/* Sample cards */
.sample-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 18px; }
.sample-card { overflow: hidden; padding: 0; }
.sample-media { height: 210px; position: relative; background: #dbe9f5; overflow: hidden; }
.sample-media img { width: 100%; height: 100%; object-fit: cover; transition: transform 500ms ease; }
.sample-card:hover .sample-media img { transform: scale(1.035); }
.sample-media.media-fallback { background: linear-gradient(145deg, #c9dff0, #edf6fb); }
.sample-media.media-fallback::after { content: ""; position: absolute; inset: 24px; border: 1px dashed #9bbbd3; border-radius: 50%; }
.sample-index { position: absolute; top: 14px; left: 14px; z-index: 2; padding: 5px 8px; border-radius: 7px; background: rgba(11,37,69,.75); color: #fff; font-family: var(--mono); font-size: 0.68rem; }
.sample-body { padding: 21px 22px 22px; }
.sample-body h3 { margin-bottom: 4px; }
.sample-tag { color: var(--teal); font-size: 0.68rem; font-weight: 800; letter-spacing: 0.08em; text-transform: uppercase; }
.sample-body p { color: var(--muted); font-size: 0.88rem; margin: 10px 0 14px; }
.finding-row { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 15px; }
.level-chip { display: inline-flex; align-items: center; gap: 5px; padding: 4px 8px; border-radius: 999px; font-size: 0.65rem; font-weight: 700; border: 1px solid transparent; white-space: nowrap; }
.level-chip::before { content: ""; width: 5px; height: 5px; border-radius: 50%; background: currentColor; }
.level-strong { background: #e4f6ee; color: #1c7454; border-color: #c6ead9; }
.level-positive { background: #e9f2ff; color: #275fb8; border-color: #cfe1fb; }
.level-moderate { background: #fff4dd; color: #936013; border-color: #f3dfb6; }
.level-slight { background: #f5f1ff; color: #6750b8; border-color: #e2d9fa; }
.level-negative { background: #f2f5f8; color: #68798c; border-color: #dde7ef; }
.level-ph { background: #e7f8f7; color: #16726e; border-color: #c9ebe8; }
details { border-top: 1px solid var(--line); padding-top: 13px; }
details summary { cursor: pointer; color: var(--navy); font-size: 0.78rem; font-weight: 800; list-style: none; display: flex; align-items: center; gap: 6px; }
details summary::-webkit-details-marker { display: none; }
details summary::after { content: "+"; margin-left: auto; color: var(--teal); font-size: 1.1rem; line-height: 1; }
details[open] summary::after { content: "−"; }
.detail-copy { padding-top: 12px; color: var(--muted); font-size: 0.82rem; }
.detail-copy p { margin-bottom: 10px; }
.detail-copy ul { padding-left: 18px; margin: 8px 0 0; }
.detail-copy li { margin-bottom: 5px; }

/* Nutrient cards */
.nutrient-card { padding: 18px; border: 1px solid var(--line); border-radius: 15px; background: #fff; }
.nutrient-card .nutrient-icon { width: 38px; height: 38px; display: grid; place-items: center; border-radius: 11px; margin-bottom: 13px; }
.nutrient-card h3 { font-size: 1.08rem; margin-bottom: 3px; }
.nutrient-card p { margin: 0; color: var(--muted); font-size: 0.78rem; }
.tone-amber .nutrient-icon { background: var(--orange-soft); color: var(--orange); }
.tone-blue .nutrient-icon { background: var(--blue-soft); color: var(--blue); }
.tone-rose .nutrient-icon { background: var(--red-soft); color: var(--red); }
.tone-teal .nutrient-icon { background: var(--teal-soft); color: var(--teal); }

/* Materials and tests */
.list-toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; justify-content: space-between; margin-bottom: 16px; }
.search-box { display: flex; align-items: center; gap: 8px; padding: 9px 12px; border: 1px solid var(--line-strong); border-radius: 10px; background: #fff; min-width: 220px; color: var(--soft); }
.search-box input { border: 0; outline: 0; background: transparent; width: 100%; color: var(--ink); font-size: 0.84rem; }
.search-box .icon { width: 16px; height: 16px; }
.item-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; }
.item-card { display: flex; align-items: flex-start; gap: 11px; padding: 14px 15px; border: 1px solid var(--line); border-radius: 12px; background: #fbfdff; }
.item-card .item-num { font-family: var(--mono); font-size: 0.65rem; color: var(--teal); padding-top: 2px; min-width: 22px; }
.item-card strong { display: block; color: var(--navy); font-size: 0.86rem; line-height: 1.35; }
.item-card small { display: block; color: var(--muted); font-size: 0.72rem; margin-top: 3px; }
.hazard { display: inline-flex; padding: 2px 7px; border-radius: 999px; font-size: 0.61rem; font-weight: 800; letter-spacing: 0.04em; text-transform: uppercase; margin-top: 5px; }
.hazard-corrosive { background: var(--red-soft); color: #a32f2f; }
.hazard-toxic { background: #f6e8ff; color: #71358c; }
.hazard-irritant { background: var(--orange-soft); color: #8a5a0e; }
.hazard-harmful { background: #fff1e8; color: #9b4b1e; }
.hazard-note { margin-top: 20px; }
.prep-card { display: grid; grid-template-columns: minmax(0, .88fr) minmax(0, 1.12fr); gap: 26px; }
.prep-media { min-height: 360px; border-radius: 18px; overflow: hidden; position: relative; background: #dfeef8; }
.prep-media img { width: 100%; height: 100%; object-fit: cover; position: absolute; inset: 0; }
.prep-media.media-fallback { background: linear-gradient(145deg, #d7e8f5, #eef7fc); }
.prep-media-caption { position: absolute; left: 15px; right: 15px; bottom: 15px; z-index: 2; padding: 9px 12px; border-radius: 9px; background: rgba(11,37,69,.74); color: #e7f2fb; font-size: 0.7rem; }
.steps { list-style: none; padding: 0; margin: 18px 0 0; display: grid; gap: 9px; counter-reset: step; }
.steps li { counter-increment: step; position: relative; padding: 12px 14px 12px 48px; border: 1px solid var(--line); border-radius: 11px; background: #fbfdff; }
.steps li::before { content: counter(step, decimal-leading-zero); position: absolute; left: 14px; top: 12px; font-family: var(--mono); font-size: 0.66rem; color: var(--teal); font-weight: 700; }
.steps strong { display: block; font-size: 0.83rem; color: var(--navy); }
.steps span { display: block; color: var(--muted); font-size: 0.76rem; margin-top: 2px; }
.test-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.test-card { padding: 20px; }
.test-head { display: flex; align-items: flex-start; gap: 12px; margin-bottom: 10px; }
.test-icon { width: 40px; height: 40px; flex: 0 0 40px; display: grid; place-items: center; border-radius: 12px; background: var(--blue-soft); color: var(--blue); }
.test-head h3 { margin: 0 0 3px; font-size: 1.12rem; }
.test-target { color: var(--muted); font-size: 0.72rem; }
.test-card details { border-top: 0; padding-top: 0; }
.test-card details summary { color: var(--teal); }
.test-detail { display: grid; gap: 11px; margin-top: 14px; }
.test-detail section { padding: 11px 13px; border-radius: 10px; background: #f8fbfe; border: 1px solid #e8f1f8; }
.test-detail h4 { font-family: var(--sans); font-size: 0.65rem; letter-spacing: 0.1em; text-transform: uppercase; color: var(--soft); margin-bottom: 4px; }
.test-detail p { margin: 0; color: #4a5c72; font-size: 0.78rem; }
.danger-pill { margin-left: auto; padding: 3px 7px; border-radius: 999px; background: var(--red-soft); color: #a63333; font-size: 0.6rem; font-weight: 800; text-transform: uppercase; white-space: nowrap; }

/* Results table */
.filter-row { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 15px; }
.filter-chip { border: 1px solid var(--line); background: #fff; color: var(--muted); border-radius: 999px; padding: 6px 11px; font-size: 0.72rem; font-weight: 700; }
.filter-chip:hover, .filter-chip.active { background: var(--navy); border-color: var(--navy); color: #fff; }
.table-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: 15px; background: #fff; }
table { width: 100%; border-collapse: collapse; }
.results-table { min-width: 760px; }
.results-table th, .results-table td { padding: 13px 14px; border-bottom: 1px solid #e8f1f8; text-align: left; vertical-align: top; font-size: 0.78rem; }
.results-table th { background: #f7fbfe; color: var(--navy); font-size: 0.68rem; text-transform: uppercase; letter-spacing: 0.07em; font-weight: 800; position: sticky; top: 0; z-index: 1; }
.results-table th:first-child { min-width: 170px; }
.results-table tbody tr:hover { background: #fbfdff; }
.results-table tbody tr:last-child td { border-bottom: 0; }
.legend { display: flex; flex-wrap: wrap; gap: 9px 16px; margin-top: 15px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; color: var(--muted); font-size: 0.7rem; }
.legend i { width: 9px; height: 9px; border-radius: 3px; display: inline-block; }
.legend .lg-strong { background: #2d8a65; }
.legend .lg-positive { background: #2f6fed; }
.legend .lg-moderate { background: #e89018; }
.legend .lg-slight { background: #6f5bd3; }
.legend .lg-negative { background: #9aa8b6; }
.legend .lg-ph { background: #0f9f9a; }

/* Conclusion */
.conclusion-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; }
.conclusion-card { padding: 22px; }
.conclusion-card h3 { margin: 4px 0 7px; }
.conclusion-card .sample-tag { font-size: 0.64rem; }
.conclusion-card p { color: var(--muted); font-size: 0.86rem; margin: 10px 0 13px; }
.overall { margin-top: 22px; background: var(--navy); color: #e8f2fb; border: 0; }
.overall h2, .overall h3 { color: #fff; }
.overall p { color: #bfd3e7; }
.overall-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 10px; margin-top: 17px; }
.overall-item { padding: 14px; border: 1px solid rgba(255,255,255,.14); border-radius: 12px; background: rgba(255,255,255,.055); }
.overall-item small { display: block; color: #8fb1cd; font-size: 0.62rem; text-transform: uppercase; letter-spacing: 0.1em; font-weight: 700; }
.overall-item strong { display: block; color: #fff; font-size: 0.84rem; margin-top: 5px; }
.overall-item span { display: block; color: #aec7dd; font-size: 0.7rem; margin-top: 4px; }

/* Safety */
.safety-banner { display: flex; gap: 16px; align-items: flex-start; padding: 20px; border-radius: 16px; border: 1px solid #f0d8b0; background: var(--orange-soft); color: #694a16; margin-bottom: 20px; }
.safety-banner .callout-icon { background: #ffe5b7; color: #9a650e; }
.safety-banner h3 { margin-bottom: 4px; color: #5e4315; }
.safety-banner p { margin: 0; font-size: 0.88rem; }
.safety-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 13px; }
.safety-card { padding: 18px; }
.safety-card .nutrient-icon { background: var(--red-soft); color: var(--red); }
.safety-card h3 { font-size: 1.04rem; margin-bottom: 6px; }
.safety-card p { color: var(--muted); font-size: 0.82rem; margin: 0; }
.ppe-row { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 18px; }
.ppe-item { display: flex; align-items: center; gap: 8px; padding: 9px 12px; border: 1px solid var(--line); border-radius: 999px; background: #fff; color: var(--navy); font-size: 0.75rem; font-weight: 700; }
.ppe-item .icon { color: var(--teal); }

/* Survey */
.survey-hero { display: grid; grid-template-columns: 1.15fr .85fr; gap: 18px; align-items: stretch; margin-bottom: 22px; }
.survey-hero .card { display: flex; flex-direction: column; justify-content: center; }
.survey-hero h1 { font-size: clamp(2rem, 4vw, 3.4rem); margin-bottom: 12px; }
.survey-meta { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 17px; }
.survey-side { background: var(--navy); color: #dceafa; border: 0; }
.survey-side h2 { color: #fff; font-size: 1.45rem; margin-bottom: 10px; }
.survey-side p { color: #b8cee3; font-size: 0.84rem; }
.survey-side ul { list-style: none; padding: 0; margin: 13px 0 0; display: grid; gap: 8px; }
.survey-side li { display: flex; gap: 8px; align-items: flex-start; color: #c7dbea; font-size: 0.78rem; }
.survey-side li .icon { color: #7de0d7; margin-top: 2px; }
.progress-shell { position: sticky; top: 70px; z-index: 20; margin-bottom: 18px; padding: 12px 15px; border: 1px solid var(--line); border-radius: 13px; background: rgba(255,255,255,.92); backdrop-filter: blur(12px); box-shadow: var(--shadow-sm); }
.progress-top { display: flex; justify-content: space-between; gap: 12px; font-size: 0.7rem; font-weight: 800; color: var(--muted); margin-bottom: 7px; }
.progress-track { height: 7px; border-radius: 999px; background: #e5eff8; overflow: hidden; }
.progress-fill { height: 100%; width: 0; border-radius: inherit; background: linear-gradient(90deg, var(--teal), var(--blue)); transition: width 260ms ease; }
.survey-form { display: grid; gap: 16px; }
.form-section { padding: clamp(20px, 3vw, 28px); }
.form-section-head { display: flex; gap: 13px; align-items: flex-start; margin-bottom: 20px; padding-bottom: 15px; border-bottom: 1px solid var(--line); }
.section-number { width: 34px; height: 34px; flex: 0 0 34px; display: grid; place-items: center; border-radius: 10px; background: var(--blue-soft); color: var(--blue); font-family: var(--mono); font-size: 0.7rem; font-weight: 700; }
.form-section-head h2 { font-size: 1.35rem; margin-bottom: 3px; }
.form-section-head p { margin: 0; color: var(--muted); font-size: 0.8rem; }
.question { padding: 0 0 21px; margin-bottom: 21px; border-bottom: 1px solid #edf4fa; }
.question:last-child { padding-bottom: 0; margin-bottom: 0; border-bottom: 0; }
.question-label { display: flex; gap: 8px; align-items: baseline; color: var(--navy); font-weight: 700; font-size: 0.94rem; line-height: 1.4; margin-bottom: 5px; }
.question-number { color: var(--teal); font-family: var(--mono); font-size: 0.68rem; min-width: 22px; }
.required-mark { color: var(--red); font-size: 0.78rem; }
.question-help { color: var(--muted); font-size: 0.76rem; margin: 0 0 12px 30px; }
.project-note { margin: 10px 0 0 30px; padding: 9px 11px; border-radius: 9px; background: var(--blue-soft); color: #31578c; font-size: 0.73rem; }
.field-control { margin-left: 30px; }
.text-input, .textarea, .select-input {
  width: 100%; border: 1px solid var(--line-strong); border-radius: 10px;
  background: #fff; color: var(--ink); padding: 11px 12px; font-size: 0.88rem;
  transition: border-color 160ms ease, box-shadow 160ms ease;
}
.text-input:focus, .textarea:focus, .select-input:focus { border-color: var(--blue); box-shadow: 0 0 0 3px rgba(47,111,237,.12); outline: 0; }
.textarea { min-height: 105px; resize: vertical; }
.option-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }
.option { position: relative; display: flex; align-items: center; gap: 9px; padding: 10px 12px; border: 1px solid var(--line); border-radius: 10px; background: #fbfdff; color: #42546a; font-size: 0.82rem; cursor: pointer; transition: border-color 160ms ease, background 160ms ease, box-shadow 160ms ease; }
.option:hover { border-color: #a9c8e5; background: #fff; }
.option input { accent-color: var(--teal); width: 16px; height: 16px; flex: 0 0 16px; margin: 0; }
.option:has(input:checked) { border-color: var(--teal); background: var(--teal-soft); color: #1e5b56; box-shadow: inset 0 0 0 1px rgba(15,159,154,.18); }
.image-options { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 8px; }
.image-option { position: relative; border: 1px solid var(--line); border-radius: 12px; overflow: hidden; background: #fff; cursor: pointer; transition: border-color 160ms ease, box-shadow 160ms ease, transform 160ms ease; }
.image-option:hover { border-color: #a9c8e5; transform: translateY(-1px); }
.image-option input { position: absolute; opacity: 0; pointer-events: none; }
.image-option .image-option-media { height: 84px; background: #eaf2f9; overflow: hidden; position: relative; }
.image-option .image-option-media img { width: 100%; height: 100%; object-fit: cover; }
.image-option .image-option-media.media-fallback { background: linear-gradient(145deg, #d8e8f4, #f1f8fc); }
.image-option .image-option-media.media-fallback::after { content: ""; position: absolute; inset: 13px; border: 1px dashed #aac6dc; border-radius: 50%; }
.image-option span.image-option-label { display: block; padding: 8px 7px; font-size: 0.68rem; line-height: 1.25; color: #405269; font-weight: 700; text-align: center; }
.image-option:has(input:checked) { border-color: var(--teal); box-shadow: 0 0 0 2px rgba(15,159,154,.15); }
.image-option:has(input:checked) span.image-option-label { background: var(--teal-soft); color: #1e5b56; }
.field-error { display: block; color: #a63333; font-size: 0.74rem; margin-top: 6px; font-weight: 700; }
.question.has-error .option, .question.has-error .text-input, .question.has-error .textarea { border-color: #e6a5a5; }
.form-error { padding: 12px 14px; border-radius: 10px; background: var(--red-soft); border: 1px solid #f1c9c7; color: #8d3030; font-size: 0.82rem; margin-bottom: 12px; }
.consent-box { padding: 17px 18px; border: 1px solid var(--line); border-radius: 13px; background: #fbfdff; }
.consent-label { display: flex; align-items: flex-start; gap: 10px; color: #43546b; font-size: 0.82rem; cursor: pointer; }
.consent-label input { accent-color: var(--teal); width: 17px; height: 17px; margin-top: 2px; flex: 0 0 17px; }
.form-actions { display: flex; flex-wrap: wrap; align-items: center; gap: 12px; padding: 4px 0 10px; }
.form-actions .btn { min-width: 170px; }
.form-actions small { color: var(--muted); font-size: 0.73rem; }
.success-card { max-width: 760px; margin: 50px auto; text-align: center; padding: clamp(30px, 6vw, 58px); }
.success-mark { width: 62px; height: 62px; display: grid; place-items: center; border-radius: 50%; background: var(--green-soft); color: var(--green); margin: 0 auto 18px; }
.success-mark .icon { width: 30px; height: 30px; }
.success-card h1 { font-size: clamp(2rem, 4vw, 3rem); margin-bottom: 12px; }
.success-card p { color: var(--muted); max-width: 55ch; margin: 0 auto 24px; }
.success-actions { display: flex; flex-wrap: wrap; gap: 10px; justify-content: center; }

/* Charts and live results */
.insights-band { margin: 54px 0; padding: clamp(22px, 4vw, 36px); border-radius: 26px; background: #eef6fc; border: 1px solid #d5e7f4; }
.insights-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 18px; margin-bottom: 18px; }
.insights-head h2 { margin-bottom: 6px; }
.insights-head p { margin: 0; color: var(--muted); font-size: 0.86rem; }
.summary-grid { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 11px; margin-bottom: 18px; }
.summary-card { padding: 15px; border: 1px solid #d5e7f4; border-radius: 13px; background: #fff; }
.summary-card span { display: block; color: var(--soft); font-size: 0.63rem; letter-spacing: 0.09em; text-transform: uppercase; font-weight: 800; }
.summary-card strong { display: block; color: var(--navy); font-family: var(--serif); font-size: 1.45rem; line-height: 1.1; margin-top: 6px; }
.summary-card small { display: block; color: var(--muted); font-size: 0.68rem; margin-top: 4px; }
.chart-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 14px; }
.chart-grid.three { grid-template-columns: repeat(3, minmax(0, 1fr)); }
.chart-card { padding: 17px; border: 1px solid var(--line); border-radius: 15px; background: #fff; min-width: 0; }
.chart-card.wide { grid-column: span 2; }
.chart-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 8px; margin-bottom: 10px; }
.chart-head h3 { font-size: 0.98rem; margin: 0; }
.chart-head span { color: var(--soft); font-family: var(--mono); font-size: 0.63rem; white-space: nowrap; }
.chart-canvas { position: relative; height: 250px; }
.chart-canvas.tall { height: 300px; }
.chart-fallback { display: none; }
.chart-fallback table { font-size: 0.7rem; }
.chart-fallback th, .chart-fallback td { padding: 5px 6px; border-bottom: 1px solid #edf4fa; text-align: left; }
.chart-fallback th { color: var(--soft); font-weight: 700; }
body.charts-unavailable .chart-canvas { display: none; }
body.charts-unavailable .chart-fallback { display: block; }
.empty-state { display: none; padding: 16px; border: 1px dashed var(--line-strong); border-radius: 12px; color: var(--muted); font-size: 0.82rem; background: #fbfdff; margin-bottom: 14px; }
.empty-state.visible { display: block; }
.chart-note { display: flex; flex-wrap: wrap; gap: 8px 16px; align-items: center; margin-top: 13px; color: var(--muted); font-size: 0.68rem; }
.chart-note .live-dot { width: 6px; height: 6px; border-radius: 50%; background: var(--teal); box-shadow: 0 0 0 3px var(--teal-soft); display: inline-block; }
.results-page-head { display: flex; justify-content: space-between; gap: 18px; align-items: flex-end; margin-bottom: 18px; }
.results-page-head h1 { font-size: clamp(2.2rem, 4vw, 3.7rem); margin-bottom: 7px; }
.results-actions { display: flex; flex-wrap: wrap; gap: 8px; }
.loading-line { display: inline-flex; align-items: center; gap: 6px; color: var(--muted); font-size: 0.72rem; }
.spinner { width: 12px; height: 12px; border: 2px solid #d3e4f2; border-top-color: var(--teal); border-radius: 50%; animation: spin 700ms linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.noscript-note { padding: 12px 14px; border-radius: 10px; background: var(--orange-soft); color: #755118; font-size: 0.8rem; }

/* Admin */
.admin-body .topbar { background: rgba(255, 250, 246, .92); }
.admin-warning { display: flex; gap: 14px; align-items: flex-start; padding: 17px 18px; border-radius: 14px; background: var(--orange-soft); border: 1px solid #f0d8b0; color: #6d4d16; margin-bottom: 18px; }
.admin-warning .icon { width: 22px; height: 22px; flex: 0 0 22px; margin-top: 1px; color: #a96c08; }
.admin-warning strong { display: block; color: #5d4213; margin-bottom: 2px; }
.admin-warning p { margin: 0; font-size: 0.82rem; }
.admin-toolbar { display: flex; flex-wrap: wrap; gap: 9px; align-items: center; margin-bottom: 16px; }
.admin-filters { display: grid; grid-template-columns: 1.5fr repeat(3, 1fr); gap: 9px; align-items: end; }
.admin-filters label { display: grid; gap: 4px; color: var(--muted); font-size: 0.67rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.06em; }
.admin-filters .text-input, .admin-filters .select-input { padding: 9px 10px; font-size: 0.8rem; }
.admin-filters .search-field { grid-column: span 1; }
.admin-filters .date-field { grid-column: span 1; }
.admin-filters-actions { display: flex; gap: 7px; align-items: center; }
.admin-table { min-width: 1750px; font-size: 0.73rem; }
.admin-table th { background: #f5f9fd; color: var(--navy); font-size: 0.61rem; text-transform: uppercase; letter-spacing: 0.05em; padding: 10px 9px; white-space: nowrap; position: sticky; top: 0; z-index: 1; }
.admin-table td { padding: 9px; border-bottom: 1px solid #edf4fa; vertical-align: top; color: #46576d; }
.admin-table tbody tr:hover { background: #fbfdff; }
.admin-table .name-button { border: 0; background: transparent; color: var(--blue); padding: 0; text-align: left; font-weight: 800; font-size: 0.73rem; text-decoration: underline; text-underline-offset: 2px; }
.admin-table .email-cell { max-width: 190px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.admin-table .comment-yes { color: var(--teal); font-weight: 800; }
.admin-table .comment-no { color: var(--soft); }
.pagination-bar { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; padding: 12px 4px 0; color: var(--muted); font-size: 0.76rem; }
.pagination-actions { display: flex; align-items: center; gap: 7px; }
.pagination-actions .btn { min-height: 34px; padding: 6px 10px; font-size: 0.72rem; }
.admin-dialog { width: min(680px, calc(100vw - 30px)); max-height: min(82vh, 760px); border: 1px solid var(--line); border-radius: 18px; padding: 0; background: #fff; color: var(--ink); box-shadow: 0 30px 80px rgba(11,37,69,.24); }
.admin-dialog::backdrop { background: rgba(7, 24, 44, 0.56); backdrop-filter: blur(3px); }
.dialog-head { display: flex; align-items: flex-start; gap: 14px; justify-content: space-between; padding: 20px 22px 14px; border-bottom: 1px solid var(--line); position: sticky; top: 0; background: #fff; z-index: 2; }
.dialog-head h2 { margin: 0 0 4px; font-size: 1.4rem; }
.dialog-head p { margin: 0; color: var(--muted); font-size: 0.76rem; }
.dialog-close { border: 1px solid var(--line); background: #fff; color: var(--muted); border-radius: 9px; width: 34px; height: 34px; display: grid; place-items: center; }
.dialog-body { padding: 18px 22px 24px; display: grid; gap: 9px; }
.detail-row { display: grid; grid-template-columns: 190px 1fr; gap: 12px; padding: 8px 0; border-bottom: 1px solid #edf4fa; font-size: 0.8rem; }
.detail-row dt { color: var(--soft); font-weight: 700; }
.detail-row dd { margin: 0; color: var(--ink); overflow-wrap: anywhere; }
.detail-row.comment dd { white-space: pre-wrap; background: #f8fbfe; border-radius: 8px; padding: 8px 10px; }
.admin-summary-grid { grid-template-columns: repeat(5, minmax(0, 1fr)); }
.admin-chart-grid { margin-top: 15px; }
.admin-section-title { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 10px; margin: 24px 0 12px; }
.admin-section-title h2 { font-size: 1.4rem; margin: 0; }
.admin-section-title p { margin: 3px 0 0; color: var(--muted); font-size: 0.78rem; }
.admin-error { display: none; margin: 10px 0; padding: 11px 13px; border-radius: 9px; background: var(--red-soft); color: #8d3030; font-size: 0.8rem; }
.admin-error.visible { display: block; }
.admin-loading { display: none; align-items: center; gap: 8px; color: var(--muted); font-size: 0.76rem; }
.admin-loading.visible { display: inline-flex; }

/* Topic map, project details and credits */
.topic-card { display: flex; flex-direction: column; align-items: flex-start; text-decoration: none; transition: transform 180ms ease, box-shadow 180ms ease, border-color 180ms ease; }
.topic-card:hover { transform: translateY(-3px); border-color: #b9d5ea; box-shadow: 0 20px 46px rgba(11,37,69,.1); color: inherit; }
.topic-icon { width: 42px; height: 42px; display: grid; place-items: center; border-radius: 12px; background: var(--blue-soft); color: var(--blue); margin-bottom: 15px; }
.topic-card h3 { font-size: 1.17rem; margin-bottom: 7px; }
.topic-card p { color: var(--muted); font-size: 0.82rem; margin-bottom: 16px; }
.topic-card .text-link { margin-top: auto; }
.project-details { margin: 18px 0 0; display: grid; gap: 0; }
.project-details > div { display: grid; grid-template-columns: 130px 1fr; gap: 12px; padding: 11px 0; border-bottom: 1px solid #edf4fa; }
.project-details > div:last-child { border-bottom: 0; }
.project-details dt { color: var(--soft); font-size: 0.7rem; font-weight: 800; text-transform: uppercase; letter-spacing: .06em; padding-top: 2px; }
.project-details dd { margin: 0; color: var(--ink); font-size: 0.84rem; }
.project-details dd a { color: var(--navy); font-weight: 700; text-decoration: none; }
.project-details dd a:hover { text-decoration: underline; }
.credit-list { list-style: none; padding: 0; margin: 18px 0 0; display: grid; gap: 0; }
.credit-list li { display: grid; grid-template-columns: 86px 1fr; gap: 12px; padding: 11px 0; border-bottom: 1px solid #edf4fa; }
.credit-list li:last-child { border-bottom: 0; }
.credit-label { color: var(--teal); font-size: 0.64rem; font-weight: 800; text-transform: uppercase; letter-spacing: .08em; padding-top: 3px; }
.credit-list a { color: var(--navy); font-size: 0.8rem; font-weight: 700; text-decoration: none; }
.credit-list a:hover { text-decoration: underline; }
.credit-list small { display: block; color: var(--muted); font-size: 0.68rem; margin-top: 3px; }
@media (max-width: 700px) {
  .project-details > div, .credit-list li { grid-template-columns: 1fr; gap: 3px; }
}

/* Footer */
.site-footer { margin-top: auto; padding: 30px clamp(18px, 4vw, 48px); border-top: 1px solid var(--line); background: rgba(255,255,255,.55); }
.footer-inner { width: min(1240px, 100%); margin: 0 auto; display: grid; grid-template-columns: 1.3fr .7fr; gap: 24px; }
.footer-brand { display: flex; gap: 11px; align-items: flex-start; }
.footer-mark { width: 34px; height: 34px; flex: 0 0 34px; display: grid; place-items: center; border-radius: 10px; background: var(--navy); color: #9be7df; }
.footer-brand strong { display: block; font-family: var(--serif); color: var(--navy); font-size: 1.1rem; }
.footer-brand p { margin: 5px 0 0; color: var(--muted); font-size: 0.76rem; max-width: 58ch; }
.footer-links { display: flex; flex-wrap: wrap; gap: 9px 16px; align-content: start; justify-content: flex-end; }
.footer-links a { color: var(--muted); font-size: 0.74rem; font-weight: 700; text-decoration: none; }
.footer-links a:hover { color: var(--navy); }
.footer-meta { grid-column: 1 / -1; padding-top: 16px; border-top: 1px solid #e6f0f8; color: var(--soft); font-size: 0.67rem; display: flex; flex-wrap: wrap; justify-content: space-between; gap: 8px 20px; }
.footer-meta p { margin: 0; }

/* Reveal and responsive */
.reveal { opacity: 0; transform: translateY(12px); transition: opacity 500ms ease, transform 500ms ease; }
.reveal.is-visible { opacity: 1; transform: none; }
@media (prefers-reduced-motion: reduce) {
  html { scroll-behavior: auto; }
  *, *::before, *::after { animation-duration: .01ms !important; animation-iteration-count: 1 !important; transition-duration: .01ms !important; }
  .reveal { opacity: 1; transform: none; }
}
@media (max-width: 1120px) {
  :root { --sidebar: 258px; }
  .hero { grid-template-columns: 1fr; }
  .hero-media { min-height: 290px; }
  .hero-copy { padding-bottom: 38px; }
  .grid-4, .metric-grid, .overall-grid, .summary-grid, .admin-summary-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .image-options { grid-template-columns: repeat(3, minmax(0, 1fr)); }
  .admin-filters { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
@media (max-width: 900px) {
  .sidebar { transform: translateX(-104%); transition: transform 240ms ease; box-shadow: 20px 0 60px rgba(0,0,0,.2); }
  .sidebar.is-open { transform: none; }
  .main-wrap { margin-left: 0; }
  .menu-btn { display: inline-flex; }
  .scrim { display: block; position: fixed; inset: 0; z-index: 35; background: rgba(6, 22, 42, .5); opacity: 0; pointer-events: none; transition: opacity 220ms ease; }
  .scrim.is-open { opacity: 1; pointer-events: auto; }
  .topbar .btn { min-height: 36px; padding: 8px 11px; font-size: 0.74rem; }
  .page { padding-top: 24px; }
  .aim-card, .split, .prep-card, .survey-hero { grid-template-columns: 1fr; }
  .prep-media { min-height: 270px; }
  .chart-grid.three { grid-template-columns: 1fr; }
  .chart-grid.three .chart-card.wide { grid-column: auto; }
  .footer-inner { grid-template-columns: 1fr; }
  .footer-links { justify-content: flex-start; }
  .progress-shell { top: 64px; }
}
@media (max-width: 700px) {
  body { font-size: 15px; }
  .topbar { gap: 10px; padding: 9px 14px; min-height: 62px; }
  .topbar-context span { display: none; }
  .topbar-context strong { font-size: 0.92rem; }
  .topbar .btn span { display: none; }
  .page { padding: 18px 14px 54px; }
  .hero { border-radius: 21px; }
  .hero-copy { padding: 28px 22px; }
  .hero-media { min-height: 230px; }
  .hero h1 { font-size: clamp(2.05rem, 10vw, 3.2rem); }
  .hero-metrics { margin-top: 20px; }
  .hero-metric { min-width: 0; flex: 1 1 90px; }
  .section { margin: 38px 0; }
  .section-head, .insights-head, .results-page-head { align-items: flex-start; flex-direction: column; }
  .grid-2, .grid-3, .grid-4, .sample-grid, .item-grid, .test-grid, .conclusion-grid, .safety-grid, .chart-grid, .metric-grid, .summary-grid, .admin-summary-grid, .admin-filters, .option-grid { grid-template-columns: 1fr; }
  .image-options { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .chart-card.wide { grid-column: auto; }
  .chart-canvas { height: 230px; }
  .survey-hero h1 { font-size: 2.25rem; }
  .form-section { padding: 19px 15px; }
  .question-help, .project-note, .field-control { margin-left: 0; }
  .question-label { font-size: 0.9rem; }
  .image-option .image-option-media { height: 96px; }
  .progress-shell { top: 62px; }
  .admin-table { min-width: 1450px; }
  .detail-row { grid-template-columns: 1fr; gap: 2px; }
  .footer-meta { display: block; }
  .footer-meta p + p { margin-top: 6px; }
}
@media (max-width: 430px) {
  .hero-actions .btn, .form-actions .btn, .success-actions .btn { width: 100%; }
  .hero-ticker small { width: 100%; margin-left: 0; }
  .results-actions { width: 100%; }
  .results-actions .btn { flex: 1; }
}

/* --------------------------------------------------------------------------
   UI refresh - editorial laboratory theme
   -------------------------------------------------------------------------- */
:root {
  --navy: #08233f;
  --navy-2: #103b5d;
  --navy-3: #1b557d;
  --ink: #182a39;
  --muted: #5d7180;
  --soft: #8396a4;
  --line: #dde8ed;
  --line-strong: #c4d7df;
  --paper: #fffefa;
  --bg: #f4f8f5;
  --bg-2: #e8f2ee;
  --blue: #1f6feb;
  --blue-soft: #eaf2ff;
  --teal: #0c8f84;
  --teal-soft: #e4f7f2;
  --orange: #d97a17;
  --orange-soft: #fff3df;
  --red: #cf4b4b;
  --red-soft: #fff0ef;
  --purple: #6b5bd2;
  --purple-soft: #f0eeff;
  --green: #2b8563;
  --green-soft: #e7f7ef;
  --shadow: 0 26px 64px rgba(8, 35, 63, .09);
  --shadow-sm: 0 10px 28px rgba(8, 35, 63, .07);
  --radius: 26px;
  --radius-sm: 16px;
  --sidebar: 288px;
}

body {
  background:
    radial-gradient(circle at 88% 4%, rgba(31, 111, 235, .08), transparent 34rem),
    radial-gradient(circle at 8% 78%, rgba(12, 143, 132, .08), transparent 30rem),
    linear-gradient(180deg, #f7faf7 0%, var(--bg) 48%, #f1f7f4 100%);
}

.sidebar {
  padding: 28px 22px 22px;
  background:
    radial-gradient(circle at 12% 0%, rgba(94, 211, 200, .22), transparent 18rem),
    radial-gradient(circle at 96% 88%, rgba(31, 111, 235, .2), transparent 20rem),
    linear-gradient(168deg, #061b31 0%, #0a2948 54%, #071f38 100%);
  border-right: 1px solid rgba(157, 219, 230, .14);
  box-shadow: 18px 0 60px rgba(8, 35, 63, .08);
}
.brand { padding: 2px 8px 20px; }
.brand-mark {
  width: 46px;
  height: 46px;
  border-radius: 15px;
  color: #b9fff3;
  background: linear-gradient(145deg, rgba(94, 211, 200, .3), rgba(31, 111, 235, .2));
  border-color: rgba(185, 255, 243, .32);
  box-shadow: inset 0 1px 0 rgba(255, 255, 255, .24), 0 10px 24px rgba(0, 0, 0, .16);
}
.brand-mark .icon { width: 24px; height: 24px; }
.brand strong { font-size: 1.45rem; }
.brand small { color: #8fb3c8; }
.sidebar-project {
  margin: 0 8px 24px;
  padding: 13px 14px;
  color: #cfe1ed;
  border-left: 3px solid #5ed3c8;
  background: linear-gradient(90deg, rgba(94, 211, 200, .14), rgba(255, 255, 255, .035));
  border-radius: 0 14px 14px 0;
}
.nav-group > p { color: #7e9db2; }
.nav-link { border-radius: 12px; padding: 8px 12px; }
.nav-link:hover { background: rgba(255, 255, 255, .08); }
.nav-link.active {
  background: linear-gradient(90deg, rgba(94, 211, 200, .24), rgba(31, 111, 235, .14));
  border-color: rgba(94, 211, 200, .34);
  box-shadow: inset 3px 0 0 #5ed3c8, 0 8px 20px rgba(0, 0, 0, .1);
}
.sidebar-foot { border-top-color: rgba(255, 255, 255, .12); }
.sidebar-foot p { color: #7e9db2; }

.topbar {
  min-height: 74px;
  background: rgba(255, 254, 250, .9);
  border-bottom-color: rgba(196, 215, 223, .85);
  box-shadow: 0 8px 24px rgba(8, 35, 63, .035);
}
.topbar-context strong { font-size: 1.08rem; }
.page { width: min(1280px, 100%); padding-top: 38px; padding-bottom: 84px; }

.hero {
  min-height: 500px;
  border-radius: 34px;
  background:
    radial-gradient(circle at 86% 10%, rgba(255, 195, 113, .16), transparent 24rem),
    radial-gradient(circle at 12% 100%, rgba(94, 211, 200, .14), transparent 28rem),
    linear-gradient(142deg, #061b31 0%, #0b2d50 64%, #0b3455 100%);
  border: 1px solid rgba(255, 255, 255, .14);
  box-shadow: 0 34px 84px rgba(8, 35, 63, .22);
}
.hero-copy { padding: clamp(34px, 5.4vw, 64px); }
.hero h1 { max-width: 13ch; letter-spacing: -.04em; }
.hero h1 .accent { color: #f6c177; }
.hero .kicker { color: #8cded8; }
.hero .kicker::before { background: #f6c177; }
.hero-subtitle { color: #c3d8e6; }
.hero-lede { color: #c6d9e7; }
.hero-metric { background: rgba(255, 255, 255, .075); border-color: rgba(255, 255, 255, .18); }
.hero-media { border-radius: 0 34px 34px 0; }
.hero-media img { filter: saturate(1.04) contrast(1.02); }
.hero-media::after {
  background:
    linear-gradient(90deg, rgba(8, 35, 63, .58), transparent 46%),
    linear-gradient(0deg, rgba(8, 35, 63, .58), transparent 38%);
}
.hero-ticker { background: rgba(6, 27, 49, .76); border-color: rgba(255, 255, 255, .16); }

.kicker { color: var(--teal); }
.kicker::before { background: var(--orange); }
.section { margin: 60px 0; }
.section-head { margin-bottom: 24px; }
.card {
  background: var(--paper);
  border-color: var(--line);
  border-radius: var(--radius);
  box-shadow: var(--shadow-sm);
}
.sample-card, .test-card, .topic-card, .nutrient-card {
  transition: transform 200ms ease, box-shadow 200ms ease, border-color 200ms ease;
}
.sample-card:hover, .test-card:hover, .nutrient-card:hover {
  transform: translateY(-3px);
  border-color: var(--line-strong);
  box-shadow: var(--shadow);
}
.sample-media { height: 235px; }
.sample-media::after {
  content: "";
  position: absolute;
  inset: 0;
  background: linear-gradient(0deg, rgba(8, 35, 63, .34), transparent 42%);
  pointer-events: none;
}
.sample-index { z-index: 3; background: rgba(8, 35, 63, .8); }
.test-icon, .nutrient-icon, .topic-icon, .callout-icon { box-shadow: inset 0 0 0 1px rgba(8, 35, 63, .05); }
.test-card { border-top: 3px solid rgba(12, 143, 132, .32); }
.test-card:nth-child(3n + 2) { border-top-color: rgba(31, 111, 235, .3); }
.test-card:nth-child(3n) { border-top-color: rgba(217, 122, 23, .34); }

/* Image-led story sections */
.visual-band {
  display: grid;
  grid-template-columns: minmax(0, 1.02fr) minmax(0, .98fr);
  gap: clamp(22px, 4vw, 48px);
  align-items: center;
  padding: clamp(18px, 3vw, 28px);
  border-radius: calc(var(--radius) + 4px);
  border: 1px solid var(--line);
  background: linear-gradient(135deg, #fffefa 0%, #f2f9f5 100%);
  box-shadow: var(--shadow-sm);
  overflow: hidden;
}
.visual-band.reverse .visual-media { order: 2; }
.visual-media {
  position: relative;
  min-height: 380px;
  border-radius: 22px;
  overflow: hidden;
  background: #dcebf0;
}
.visual-media img {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: cover;
  transition: transform 700ms ease;
}
.visual-band:hover .visual-media img { transform: scale(1.025); }
.visual-media::after {
  content: "";
  position: absolute;
  inset: 0;
  background: linear-gradient(0deg, rgba(8, 35, 63, .58), transparent 46%);
  pointer-events: none;
}
.visual-media.media-fallback {
  background:
    radial-gradient(circle at 70% 24%, rgba(94, 211, 200, .5), transparent 12rem),
    linear-gradient(145deg, #bcdbe2, #eaf4f1);
}
.visual-caption {
  position: absolute;
  left: 15px;
  right: 15px;
  bottom: 15px;
  z-index: 2;
  padding: 10px 13px;
  border-radius: 11px;
  background: rgba(6, 27, 49, .72);
  color: #e7f2f8;
  font-size: .72rem;
  backdrop-filter: blur(8px);
}
.visual-copy { padding: clamp(8px, 2.4vw, 32px); }
.visual-copy h2 { margin-bottom: 12px; }
.visual-copy p { color: var(--muted); max-width: 56ch; }
.visual-copy .note { margin-top: 18px; }

.image-gallery {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 14px;
}
.gallery-card {
  position: relative;
  min-height: 230px;
  margin: 0;
  overflow: hidden;
  border-radius: 20px;
  border: 1px solid var(--line);
  background: #dcebf0;
  box-shadow: var(--shadow-sm);
}
.gallery-card img {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  object-fit: cover;
  transition: transform 600ms ease;
}
.gallery-card:hover img { transform: scale(1.045); }
.gallery-card::after {
  content: "";
  position: absolute;
  inset: 0;
  background: linear-gradient(0deg, rgba(8, 35, 63, .7), transparent 56%);
  pointer-events: none;
}
.gallery-card.media-fallback {
  background:
    radial-gradient(circle at 50% 35%, rgba(94, 211, 200, .38), transparent 8rem),
    linear-gradient(145deg, #c6dfe5, #eef7f3);
}
.gallery-caption {
  position: absolute;
  left: 14px;
  right: 14px;
  bottom: 13px;
  z-index: 2;
  color: #fff;
}
.gallery-caption strong {
  display: block;
  font-family: var(--serif);
  font-size: 1rem;
  font-weight: 500;
  line-height: 1.15;
}
.gallery-caption small {
  display: block;
  margin-top: 3px;
  color: #d4e6ef;
  font-size: .64rem;
  letter-spacing: .06em;
  text-transform: uppercase;
}

/* Survey image */
.survey-side-media {
  height: 150px;
  margin-bottom: 18px;
  overflow: hidden;
  border-radius: 16px;
  border: 1px solid rgba(255, 255, 255, .18);
  background: rgba(255, 255, 255, .06);
}
.survey-side-media img {
  width: 100%;
  height: 100%;
  object-fit: cover;
  opacity: .92;
}
.survey-side { overflow: hidden; }

/* Admin table and scrollbars */
.table-wrap {
  box-shadow: var(--shadow-sm);
  scrollbar-width: thin;
  scrollbar-color: var(--line-strong) transparent;
}
.table-wrap::-webkit-scrollbar { height: 10px; }
.table-wrap::-webkit-scrollbar-thumb { background: var(--line-strong); border-radius: 999px; }
.table-wrap::-webkit-scrollbar-track { background: transparent; }
.admin-table { min-width: 1680px; }
.admin-table th, .admin-table td { white-space: nowrap; }
.admin-table td:nth-child(7) { min-width: 180px; white-space: normal; }
.admin-table th:nth-child(2), .admin-table td:nth-child(2) {
  position: sticky;
  left: 0;
  z-index: 2;
  background: var(--paper);
  box-shadow: 1px 0 0 var(--line);
}
.admin-table th:nth-child(2) { background: #f3f8f6; z-index: 3; }
.admin-body .topbar { background: rgba(255, 251, 246, .94); }
.admin-warning { box-shadow: var(--shadow-sm); }

/* Keep content visible without JavaScript, and reduce motion for users who ask. */
.reveal { opacity: 1; transform: none; }

@media (max-width: 1120px) {
  .visual-media { min-height: 320px; }
  .image-gallery { grid-template-columns: repeat(2, minmax(0, 1fr)); }
}
@media (max-width: 900px) {
  .visual-band { grid-template-columns: 1fr; }
  .visual-band.reverse .visual-media { order: 0; }
  .visual-media { min-height: 290px; }
  .hero-media { border-radius: 0 0 28px 28px; }
}
@media (max-width: 700px) {
  .topbar .btn { display: none; }
  .topbar-context { flex: 1; }
  .hero { border-radius: 25px; }
  .hero-media { border-radius: 0 0 24px 24px; }
  .visual-media { min-height: 230px; border-radius: 17px; }
  .visual-copy { padding: 4px 2px 8px; }
  .gallery-card { min-height: 175px; }
  .survey-side-media { height: 135px; }
  .admin-table { min-width: 1450px; }
}
@media (max-width: 430px) {
  .image-gallery { grid-template-columns: 1fr; }
  .gallery-card { min-height: 205px; }
}
"""

BASE_JS = r"""(function () {
  "use strict";

  const $ = (selector, scope) => (scope || document).querySelector(selector);
  const $$ = (selector, scope) => Array.from((scope || document).querySelectorAll(selector));

  function setText(element, value) {
    if (element) element.textContent = value == null ? "" : String(value);
  }

  /* Navigation */
  const sidebar = $("#site-nav");
  const scrim = $("[data-nav-close]");
  const navToggle = $("[data-nav-toggle]");
  function closeNav() {
    if (!sidebar) return;
    sidebar.classList.remove("is-open");
    if (scrim) scrim.classList.remove("is-open");
    if (navToggle) navToggle.setAttribute("aria-expanded", "false");
    document.body.classList.remove("nav-open");
  }
  if (navToggle && sidebar) {
    navToggle.addEventListener("click", () => {
      const opening = !sidebar.classList.contains("is-open");
      sidebar.classList.toggle("is-open", opening);
      if (scrim) scrim.classList.toggle("is-open", opening);
      navToggle.setAttribute("aria-expanded", String(opening));
      document.body.classList.toggle("nav-open", opening);
    });
  }
  if (scrim) scrim.addEventListener("click", closeNav);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") closeNav();
  });
  window.addEventListener("resize", () => {
    if (window.innerWidth > 900) closeNav();
  });

  /* Reveal on scroll */
  const revealItems = $$(".reveal");
  if (revealItems.length) {
    if (!("IntersectionObserver" in window)) {
      revealItems.forEach((item) => item.classList.add("is-visible"));
    } else {
      const observer = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (entry.isIntersecting) {
              entry.target.classList.add("is-visible");
              observer.unobserve(entry.target);
            }
          });
        },
        { threshold: 0.08, rootMargin: "0px 0px -30px 0px" }
      );
      revealItems.forEach((item) => observer.observe(item));
    }
  }

  /* Searchable material lists */
  $$("[data-list-search]").forEach((input) => {
    const list = input.closest("[data-list]");
    if (!list) return;
    const items = $$("[data-search]", list);
    const empty = $("[data-list-empty]", list);
    input.addEventListener("input", () => {
      const query = input.value.trim().toLowerCase();
      let shown = 0;
      items.forEach((item) => {
        const haystack = (item.dataset.search || "").toLowerCase();
        const visible = !query || haystack.includes(query);
        item.hidden = !visible;
        if (visible) shown += 1;
      });
      if (empty) empty.hidden = shown !== 0;
    });
  });

  /* Observation-table filters */
  $$("[data-results-filter]").forEach((button) => {
    button.addEventListener("click", () => {
      const category = button.dataset.resultsFilter;
      $$("[data-results-filter]").forEach((item) => {
        item.classList.toggle("active", item === button);
        item.setAttribute("aria-pressed", item === button ? "true" : "false");
      });
      $$("[data-result-row]").forEach((row) => {
        row.hidden = category !== "all" && row.dataset.category !== category;
      });
    });
  });

  /* Survey progress and submit state */
  const surveyForm = $("[data-survey-form]");
  if (surveyForm) {
    const progressFill = $("[data-progress-fill]");
    const progressLabel = $("[data-progress-label]");
    const requiredQuestions = $$("[data-required-question]", surveyForm);
    const consent = $("[data-consent]", surveyForm);
    const submitButton = $("[data-submit]", surveyForm);
    const comment = $("#comments", surveyForm);
    const commentCount = $("[data-comment-count]", surveyForm);

    function isAnswered(question) {
      const radios = $$('input[type="radio"]', question);
      if (radios.length) return radios.some((radio) => radio.checked);
      const field = $("input, textarea", question);
      return Boolean(field && field.value.trim());
    }

    function updateProgress() {
      const answered = requiredQuestions.filter(isAnswered).length;
      const requiredTotal = requiredQuestions.length + (consent ? 1 : 0);
      const completed = answered + (consent && consent.checked ? 1 : 0);
      const percent = requiredTotal ? Math.round((completed / requiredTotal) * 100) : 0;
      if (progressFill) progressFill.style.width = percent + "%";
      if (progressLabel) {
        progressLabel.textContent = completed + " of " + requiredTotal + " required answers";
      }
    }

    surveyForm.addEventListener("change", updateProgress);
    surveyForm.addEventListener("input", updateProgress);
    if (comment && commentCount) {
      const updateCount = () => {
        commentCount.textContent = comment.value.length + " / 500";
      };
      comment.addEventListener("input", updateCount);
      updateCount();
    }
    surveyForm.addEventListener("submit", () => {
      if (submitButton) {
        submitButton.disabled = true;
        submitButton.classList.add("is-disabled");
        submitButton.textContent = "Saving response…";
      }
    });
    updateProgress();
  }

  /* Chart helpers */
  const chartInstances = new Map();

  function fallbackTable(container, data) {
    if (!container) return;
    container.replaceChildren();
    const table = document.createElement("table");
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    ["Option", "Responses", "Share"].forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    table.appendChild(thead);
    const tbody = document.createElement("tbody");
    data.labels.forEach((label, index) => {
      const tr = document.createElement("tr");
      const labelCell = document.createElement("td");
      labelCell.textContent = label;
      const countCell = document.createElement("td");
      countCell.textContent = String(data.counts[index]);
      const shareCell = document.createElement("td");
      shareCell.textContent = String(data.percentages[index]) + "%";
      tr.append(labelCell, countCell, shareCell);
      tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    container.appendChild(table);
  }

  function buildChartConfig(data) {
    const isDoughnut = data.type === "doughnut";
    return {
      type: data.type,
      data: {
        labels: data.labels,
        datasets: [
          {
            label: "Responses",
            data: data.counts,
            backgroundColor: data.colors,
            borderColor: isDoughnut ? "#ffffff" : "transparent",
            borderWidth: isDoughnut ? 2 : 0,
            borderRadius: isDoughnut ? 0 : 6,
            maxBarThickness: isDoughnut ? undefined : 24,
            hoverOffset: isDoughnut ? 5 : 0,
          },
        ],
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        indexAxis: !isDoughnut && data.horizontal ? "y" : "x",
        cutout: isDoughnut ? "62%" : undefined,
        animation: { duration: 450 },
        plugins: {
          legend: {
            display: isDoughnut,
            position: "bottom",
            labels: {
              boxWidth: 9,
              boxHeight: 9,
              usePointStyle: true,
              pointStyle: "circle",
              padding: 14,
              color: "#5f7087",
              font: { size: 10, family: "Segoe UI, system-ui, sans-serif" },
            },
          },
          tooltip: {
            callbacks: {
              label: function (context) {
                const value = Number(context.raw) || 0;
                const total = Number(data.total) || 0;
                const percent = total ? ((value / total) * 100).toFixed(1) : "0.0";
                return context.label + ": " + value + " (" + percent + "%)";
              },
            },
          },
        },
        scales: isDoughnut
          ? {}
          : {
              x: {
                beginAtZero: true,
                ticks: {
                  precision: 0,
                  color: "#8290a4",
                  font: { size: 10 },
                },
                grid: { color: "#edf4fa" },
                title: {
                  display: !data.horizontal,
                  text: "Responses",
                  color: "#8290a4",
                  font: { size: 10 },
                },
              },
              y: {
                beginAtZero: true,
                ticks: {
                  precision: 0,
                  color: "#5f7087",
                  font: { size: 10 },
                  autoSkip: false,
                },
                grid: { display: false },
                title: {
                  display: Boolean(data.horizontal),
                  text: "Responses",
                  color: "#8290a4",
                  font: { size: 10 },
                },
              },
            },
      },
    };
  }

  function renderChartCard(card, data) {
    const canvas = $("[data-chart-canvas]", card);
    const fallback = $("[data-chart-fallback]", card);
    if (!canvas) return;
    fallbackTable(fallback, data);
    if (typeof window.Chart === "undefined") {
      document.body.classList.add("charts-unavailable");
      return;
    }
    const context = canvas.getContext("2d");
    if (!context) return;
    const existing = chartInstances.get(canvas.id);
    if (existing) existing.destroy();
    const chart = new window.Chart(context, buildChartConfig(data));
    chartInstances.set(canvas.id, chart);
  }

  function renderCharts(root, summary) {
    if (!root) return;
    $$("[data-chart-card]", root).forEach((card) => {
      const key = card.dataset.chartCard;
      const data = summary.charts && summary.charts[key];
      if (data) renderChartCard(card, data);
    });
    if (typeof window.Chart === "undefined") {
      document.body.classList.add("charts-unavailable");
    }
  }

  function renderSummaryCards(root, summary, scope) {
    if (!root) return;
    const values = {
      total: String(summary.total),
      classes: String(summary.unique_classes),
      preferred: summary.most_preferred_breakfast,
      frequency: summary.most_common_frequency,
      last: summary.last_response_at,
    };
    $$("[data-summary]", root).forEach((element) => {
      const key = element.dataset.summary;
      if (scope && element.dataset.scope && element.dataset.scope !== scope) return;
      if (Object.prototype.hasOwnProperty.call(values, key)) {
        setText(element, values[key]);
      }
    });
    $$("[data-generated-at]", root).forEach((element) => {
      setText(element, summary.generated_at_local);
    });
    $$("[data-empty-state]", root).forEach((element) => {
      element.classList.toggle("visible", summary.total === 0);
    });
  }

  async function fetchSummary() {
    const response = await fetch("/api/survey/summary", {
      headers: { Accept: "application/json" },
      cache: "no-store",
    });
    if (!response.ok) throw new Error("Summary request failed");
    return response.json();
  }

  function showChartError(message) {
    $$("[data-chart-error]").forEach((element) => {
      element.textContent = message;
      element.hidden = false;
    });
  }

  window.NutriLab = { $, $$, setText, fetchSummary, renderSummaryCards, renderCharts, chartInstances, showChartError };

  /* If Chart.js failed to load, still show fallbacks after data arrives. */
  if (typeof window.Chart === "undefined") {
    document.body.classList.add("charts-unavailable");
  }
})();
"""

LIVE_JS = r"""(function () {
  "use strict";
  const { $, $$, fetchSummary, renderSummaryCards, renderCharts, showChartError } = window.NutriLab;

/* Homepage compact insights */
  const homeInsights = $("[data-home-insights]");
  if (homeInsights) {
    let loading = false;
    async function loadHome() {
      if (loading) return;
      loading = true;
      try {
        const summary = await fetchSummary();
        renderSummaryCards(homeInsights, summary, "home");
        renderCharts(homeInsights, summary);
      } catch (error) {
        showChartError("Live survey insights could not load. Please refresh in a moment.");
      } finally {
        loading = false;
      }
    }
    loadHome();
    const refresh = $("[data-refresh-home]", homeInsights);
    if (refresh) refresh.addEventListener("click", loadHome);
  }

  /* Results page */
  const resultsPage = $("[data-results-page]");
  if (resultsPage) {
    let loading = false;
    let timer = null;
    async function loadResults() {
      if (loading) return;
      loading = true;
      const loadingElement = $("[data-loading]", resultsPage);
      if (loadingElement) loadingElement.hidden = false;
      try {
        const summary = await fetchSummary();
        renderSummaryCards(resultsPage, summary, "results");
        renderCharts(resultsPage, summary);
        const note = $("[data-chart-note]", resultsPage);
        if (note) note.hidden = false;
      } catch (error) {
        showChartError("Live results could not load. Check the server and refresh this page.");
      } finally {
        loading = false;
        if (loadingElement) loadingElement.hidden = true;
      }
    }
    loadResults();
    const refresh = $("[data-refresh-results]", resultsPage);
    if (refresh) refresh.addEventListener("click", loadResults);
    timer = window.setInterval(() => {
      if (!document.hidden) loadResults();
    }, 60000);
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) loadResults();
    });
    window.addEventListener("beforeunload", () => window.clearInterval(timer));
  }
})();
"""

ADMIN_JS = r"""(function () {
  "use strict";
  const { $, $$, setText, fetchSummary, renderSummaryCards, renderCharts, chartInstances } = window.NutriLab;

/* Admin dashboard */
  const adminRoot = $("[data-admin]");
  if (adminRoot) {
    const state = {
      q: "",
      class_level: "",
      preferred_breakfast: "",
      breakfast_frequency: "",
      date_from: "",
      date_to: "",
      sort: "newest",
      page: 1,
      per_page: 10,
    };
    const filterForm = $("[data-admin-filters]", adminRoot);
    const tbody = $("[data-admin-table-body]", adminRoot);
    const errorBox = $("[data-admin-error]", adminRoot);
    const loadingBox = $("[data-admin-loading]", adminRoot);
    const paginationInfo = $("[data-pagination-info]", adminRoot);
    const prevButton = $("[data-page-prev]", adminRoot);
    const nextButton = $("[data-page-next]", adminRoot);
    const exportCsv = $("[data-export-csv]", adminRoot);
    const exportPdf = $("[data-export-pdf]", adminRoot);
    const exportDocx = $("[data-export-docx]", adminRoot);
    const exportChartsPdf = $("[data-export-charts-pdf]", adminRoot);
    const exportChartsDocx = $("[data-export-charts-docx]", adminRoot);
    const dialog = $("[data-detail-dialog]", adminRoot);
    const detailBody = $("[data-detail-body]", adminRoot);
    let currentItems = new Map();
    let currentPage = 1;
    let currentPages = 1;
    let isLoading = false;

    const params = new URLSearchParams(window.location.search);
    ["q", "class_level", "preferred_breakfast", "breakfast_frequency", "date_from", "date_to", "sort", "per_page"].forEach((key) => {
      if (params.has(key)) state[key] = params.get(key);
    });
    state.page = Math.max(1, parseInt(params.get("page") || "1", 10) || 1);
    if (!["10", "25", "50", "100"].includes(String(state.per_page))) state.per_page = 10;
    if (!["newest", "oldest"].includes(state.sort)) state.sort = "newest";

    function syncFormFromState() {
      if (!filterForm) return;
      ["q", "class_level", "preferred_breakfast", "breakfast_frequency", "date_from", "date_to", "sort", "per_page"].forEach((key) => {
        const field = filterForm.elements.namedItem(key);
        if (field) field.value = state[key] || "";
      });
    }

    function queryParams(includePage) {
      const query = new URLSearchParams();
      ["q", "class_level", "preferred_breakfast", "breakfast_frequency", "date_from", "date_to", "sort", "per_page"].forEach((key) => {
        if (state[key]) query.set(key, state[key]);
      });
      if (includePage) query.set("page", String(state.page));
      return query;
    }

    function updateUrl() {
      const query = queryParams(true);
      const suffix = query.toString() ? "?" + query.toString() : "";
      window.history.replaceState(null, "", window.location.pathname + suffix);
    }

    function updateExports() {
      const query = queryParams(false).toString();
      const suffix = query ? "?" + query : "";
      if (exportCsv) exportCsv.href = "/admin/export.csv" + suffix;
      if (exportPdf) exportPdf.href = "/admin/export.pdf" + suffix;
      if (exportDocx) exportDocx.href = "/admin/export.docx" + suffix;
    }

    function collectChartImages() {
      const charts = [];
      $$("[data-chart-card]", adminRoot).forEach((card) => {
        const key = card.dataset.chartCard;
        const canvas = $("[data-chart-canvas]", card);
        const instance = canvas ? chartInstances.get(canvas.id) : null;
        let image = "";
        if (instance && typeof instance.toBase64Image === "function") {
          try {
            image = instance.toBase64Image("image/png", 1);
          } catch (error) {
            image = "";
          }
        }
        charts.push({ key, image });
      });
      return charts;
    }

    async function downloadChartReport(format, button) {
      if (!button) return;
      button.disabled = true;
      try {
        const response = await fetch("/admin/export.charts." + format, {
          method: "POST",
          headers: { "Content-Type": "application/json", Accept: "*/*" },
          body: JSON.stringify({ charts: collectChartImages() }),
        });
        if (!response.ok) throw new Error("Chart report request failed");
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "chemistry_ip_aggregate_charts." + format;
        document.body.appendChild(link);
        link.click();
        link.remove();
        window.setTimeout(() => URL.revokeObjectURL(url), 2000);
      } catch (error) {
        if (errorBox) {
          errorBox.textContent =
            "The aggregate chart report could not be generated. Please try again.";
          errorBox.classList.add("visible");
        }
      } finally {
        button.disabled = false;
      }
    }

    if (exportChartsPdf) {
      exportChartsPdf.addEventListener("click", () => {
        downloadChartReport("pdf", exportChartsPdf);
      });
    }
    if (exportChartsDocx) {
      exportChartsDocx.addEventListener("click", () => {
        downloadChartReport("docx", exportChartsDocx);
      });
    }

    function renderTable(items, pagination) {
      if (!tbody) return;
      tbody.replaceChildren();
      currentItems = new Map();
      if (!items.length) {
        const row = document.createElement("tr");
        const cell = document.createElement("td");
        cell.colSpan = 17;
        cell.textContent = "No survey responses match the current filters.";
        cell.style.color = "#8290a4";
        cell.style.padding = "24px 14px";
        row.appendChild(cell);
        tbody.appendChild(row);
        return;
      }
      items.forEach((item) => {
        currentItems.set(item.ref, item);
        const row = document.createElement("tr");
        const values = [
          item.submitted_at,
          null,
          item.class_level,
          item.email,
          item.breakfast_frequency,
          item.preferred_breakfast,
          item.main_reason,
          item.eating_time,
          item.important_nutrient,
          item.test_familiarity,
          item.iodine_awareness,
          item.perceived_grease_spot,
          item.iodine_knowledge,
          item.learning_method,
          item.choice_influence,
          item.interesting_topic,
        ];
        values.forEach((value, index) => {
          const cell = document.createElement("td");
          if (index === 1) {
            const button = document.createElement("button");
            button.type = "button";
            button.className = "name-button";
            button.textContent = item.name || "Unnamed";
            button.addEventListener("click", () => openDetails(item));
            cell.appendChild(button);
          } else {
            cell.textContent = value || "";
            if (index === 3) cell.className = "email-cell";
          }
          row.appendChild(cell);
        });
        const commentCell = document.createElement("td");
        commentCell.textContent = item.comment_available ? "Yes" : "No";
        commentCell.className = item.comment_available ? "comment-yes" : "comment-no";
        row.appendChild(commentCell);
        tbody.appendChild(row);
      });
      currentPage = pagination.page;
      currentPages = Math.max(1, pagination.pages);
      if (paginationInfo) {
        paginationInfo.textContent =
          "Page " + pagination.page + " of " + currentPages + " · " + pagination.total + " response" + (pagination.total === 1 ? "" : "s");
      }
      if (prevButton) prevButton.disabled = !pagination.has_prev;
      if (nextButton) nextButton.disabled = !pagination.has_next;
    }

    function appendDetail(label, value) {
      if (!detailBody) return;
      const row = document.createElement("div");
      row.className = "detail-row";
      const dt = document.createElement("dt");
      dt.textContent = label;
      const dd = document.createElement("dd");
      dd.textContent = value || "—";
      row.append(dt, dd);
      detailBody.appendChild(row);
    }

    function openDetails(item) {
      if (!dialog || !detailBody) return;
      detailBody.replaceChildren();
      appendDetail("Submitted", item.submitted_at);
      appendDetail("Name", item.name);
      appendDetail("Class", item.class_level);
      appendDetail("Email", item.email);
      appendDetail("Breakfast frequency", item.breakfast_frequency);
      appendDetail("Preferred breakfast", item.preferred_breakfast);
      appendDetail("Main reason", item.main_reason);
      appendDetail("Eating time", item.eating_time);
      appendDetail("Important nutrient", item.important_nutrient);
      appendDetail("Test familiarity", item.test_familiarity);
      appendDetail("Iodine test awareness", item.iodine_awareness);
      appendDetail("Grease-spot answer", item.perceived_grease_spot);
      appendDetail("Iodine test knowledge", item.iodine_knowledge);
      appendDetail("Preferred learning method", item.learning_method);
      appendDetail("Food-choice influence", item.choice_influence);
      appendDetail("Most interesting topic", item.interesting_topic);
      const commentRow = document.createElement("div");
      commentRow.className = "detail-row comment";
      const dt = document.createElement("dt");
      dt.textContent = "Comments";
      const dd = document.createElement("dd");
      dd.textContent = item.comments || "No comment provided.";
      commentRow.append(dt, dd);
      detailBody.appendChild(commentRow);
      appendDetail("Consent recorded", item.consent_given ? "Yes" : "No");
      if (typeof dialog.showModal === "function") dialog.showModal();
      else dialog.setAttribute("open", "");
    }

    if (dialog) {
      $$("[data-dialog-close]", dialog).forEach((button) => {
        button.addEventListener("click", () => dialog.close());
      });
      dialog.addEventListener("click", (event) => {
        if (event.target === dialog) dialog.close();
      });
    }

    async function loadAdmin() {
      if (isLoading) return;
      isLoading = true;
      if (loadingBox) loadingBox.classList.add("visible");
      if (errorBox) {
        errorBox.classList.remove("visible");
        errorBox.textContent = "";
      }
      try {
        const query = queryParams(true);
        const [responsesResponse, summary] = await Promise.all([
          fetch("/api/admin/responses?" + query.toString(), {
            headers: { Accept: "application/json" },
            cache: "no-store",
          }),
          fetchSummary(),
        ]);
        if (!responsesResponse.ok) throw new Error("Responses request failed");
        const payload = await responsesResponse.json();
        renderSummaryCards(adminRoot, summary, "admin");
        const filteredCount = $("[data-admin-filtered-count]", adminRoot);
        if (filteredCount) setText(filteredCount, String(payload.pagination.total));
        renderCharts(adminRoot, summary);
        renderTable(payload.items, payload.pagination);
        updateExports();
        updateUrl();
      } catch (error) {
        if (errorBox) {
          errorBox.textContent = "The survey responses could not be loaded. Please refresh or try again.";
          errorBox.classList.add("visible");
        }
      } finally {
        isLoading = false;
        if (loadingBox) loadingBox.classList.remove("visible");
      }
    }

    function readFilters() {
      if (!filterForm) return;
      ["q", "class_level", "preferred_breakfast", "breakfast_frequency", "date_from", "date_to", "sort", "per_page"].forEach((key) => {
        const field = filterForm.elements.namedItem(key);
        state[key] = field ? field.value : "";
      });
      state.page = 1;
    }

    if (filterForm) {
      filterForm.addEventListener("submit", (event) => {
        event.preventDefault();
        readFilters();
        loadAdmin();
      });
      filterForm.addEventListener("reset", () => {
        window.setTimeout(() => {
          readFilters();
          loadAdmin();
        }, 0);
      });
    }
    if (prevButton) {
      prevButton.addEventListener("click", () => {
        if (state.page > 1) {
          state.page -= 1;
          loadAdmin();
        }
      });
    }
    if (nextButton) {
      nextButton.addEventListener("click", () => {
        if (state.page < currentPages) {
          state.page += 1;
          loadAdmin();
        }
      });
    }
    const refresh = $("[data-refresh-admin]", adminRoot);
    if (refresh) refresh.addEventListener("click", loadAdmin);
    syncFormFromState();
    loadAdmin();
  }
})();
"""

# ---------------------------------------------------------------------------
# Shared layout, icon system and page renderer
# ---------------------------------------------------------------------------
ICON_SYMBOLS = """
<symbol id="i-leaf" viewBox="0 0 24 24"><path d="M20 4C13 4 5 7 5 14c0 3 2 5 5 5 7 0 10-8 10-15Z"/><path d="M4 21c3-7 8-10 13-13"/></symbol>
<symbol id="i-menu" viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></symbol>
<symbol id="i-close" viewBox="0 0 24 24"><path d="m6 6 12 12M18 6 6 18"/></symbol>
<symbol id="i-arrow" viewBox="0 0 24 24"><path d="M5 12h14M13 6l6 6-6 6"/></symbol>
<symbol id="i-check" viewBox="0 0 24 24"><path d="m5 12 4 4L19 6"/></symbol>
<symbol id="i-alert" viewBox="0 0 24 24"><path d="M12 3 2.8 19h18.4L12 3Z"/><path d="M12 9v4M12 17h.01"/></symbol>
<symbol id="i-clock" viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><path d="M12 7v5l3 2"/></symbol>
<symbol id="i-users" viewBox="0 0 24 24"><circle cx="9" cy="8" r="3"/><path d="M3.5 19c.4-3.3 2.3-5 5.5-5s5.1 1.7 5.5 5"/><path d="M16 5.5a3 3 0 0 1 0 5.8M17 14c2.4.3 3.7 2 4 5"/></symbol>
<symbol id="i-mail" viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="2"/><path d="m3.5 6.5 8.5 6 8.5-6"/></symbol>
<symbol id="i-chart" viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></symbol>
<symbol id="i-target" viewBox="0 0 24 24"><circle cx="12" cy="12" r="8.5"/><circle cx="12" cy="12" r="4.5"/><circle cx="12" cy="12" r="1"/></symbol>
<symbol id="i-flask" viewBox="0 0 24 24"><path d="M9 3h6M10 3v6l-5.2 8.6A2.2 2.2 0 0 0 6.7 21h10.6a2.2 2.2 0 0 0 1.9-3.4L14 9V3"/><path d="M7.2 15h9.6"/></symbol>
<symbol id="i-spark" viewBox="0 0 24 24"><path d="m12 2 1.7 6.3L20 10l-6.3 1.7L12 18l-1.7-6.3L4 10l6.3-1.7L12 2Z"/><path d="m19 16 .7 2.3L22 19l-2.3.7L19 22l-.7-2.3L16 19l2.3-.7L19 16Z"/></symbol>
<symbol id="i-info" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 10v6M12 7h.01"/></symbol>
<symbol id="i-starch" viewBox="0 0 24 24"><path d="m12 3 7.8 4.5v9L12 21l-7.8-4.5v-9L12 3Z"/><path d="m4.5 7.8 7.5 4.4 7.5-4.4M12 12.2V21"/></symbol>
<symbol id="i-sugar" viewBox="0 0 24 24"><path d="m12 3 7.8 4.5v9L12 21l-7.8-4.5v-9L12 3Z"/><circle cx="9" cy="10" r="1"/><circle cx="15" cy="10" r="1"/><circle cx="9" cy="15" r="1"/><circle cx="15" cy="15" r="1"/></symbol>
<symbol id="i-fat" viewBox="0 0 24 24"><path d="M12 3c3.5 4.2 6 7.2 6 10.2A6 6 0 0 1 6 13.2C6 10.2 8.5 7.2 12 3Z"/><path d="M9 14c.4 1.5 1.5 2.4 3 2.5"/></symbol>
<symbol id="i-protein" viewBox="0 0 24 24"><circle cx="7" cy="7" r="2.5"/><circle cx="17" cy="17" r="2.5"/><path d="M9 9c2 1 4 1 6 2M8 14c1 2 2 3 4 4"/></symbol>
<symbol id="i-mineral" viewBox="0 0 24 24"><path d="m12 3 7 5-3 11H8L5 8l7-5Z"/><path d="m5 8 7 3 7-3M12 11v8"/></symbol>
<symbol id="i-ph" viewBox="0 0 24 24"><path d="M12 3c3.5 4.2 6 7.2 6 10.2A6 6 0 0 1 6 13.2C6 10.2 8.5 7.2 12 3Z"/><path d="M9 12h6"/></symbol>
<symbol id="i-aromatic" viewBox="0 0 24 24"><path d="m12 3 7.8 4.5v9L12 21l-7.8-4.5v-9L12 3Z"/><circle cx="12" cy="12" r="3"/></symbol>
<symbol id="i-sulfur" viewBox="0 0 24 24"><path d="M16 6.5c-1-1.2-2.3-1.8-4-1.8-2.2 0-4 1.1-4 3.2 0 4.5 8 2.2 8 6.6 0 2.1-1.8 3.3-4 3.3-1.8 0-3.3-.7-4.4-2"/></symbol>
<symbol id="i-goggles" viewBox="0 0 24 24"><path d="M3 9h18M5 9v4a4 4 0 0 0 8 0v-1h-2a3 3 0 0 1-6 0V9M19 9v4a4 4 0 0 1-8 0v-1h2a3 3 0 0 0 6 0V9"/></symbol>
<symbol id="i-coat" viewBox="0 0 24 24"><path d="m8 4-4 3 2 4 2-1v10h8V10l2 1 2-4-4-3-4 3-4-3Z"/><path d="M12 7v13"/></symbol>
<symbol id="i-gloves" viewBox="0 0 24 24"><path d="M8 11V5a1.5 1.5 0 0 1 3 0v5M11 10V4a1.5 1.5 0 0 1 3 0v6M14 10V5a1.5 1.5 0 0 1 3 0v8c0 4-2.5 7-6 7s-6-2.4-6-6v-3a1.5 1.5 0 0 1 3 0"/></symbol>
<symbol id="i-hood" viewBox="0 0 24 24"><path d="M4 21V4h16v17M8 21v-7h8v7M9 8h6"/></symbol>
<symbol id="i-heat" viewBox="0 0 24 24"><path d="M12 21c4 0 7-2.6 7-6.4 0-3.5-2.4-6.5-5.4-9.6.2 2.5-1 4-2.1 5.1C10 7.5 8 5.8 8 3c-2.3 2.4-3 5-3 7.7C5 18 8 21 12 21Z"/><path d="M12 21c-2 0-3.4-1.4-3.4-3.2 0-1.7 1.2-2.9 3.4-4.8 2.2 1.9 3.4 3.1 3.4 4.8 0 1.8-1.4 3.2-3.4 3.2Z"/></symbol>
<symbol id="i-glass" viewBox="0 0 24 24"><path d="M8 3h8M9 3v5l-4 9a2.4 2.4 0 0 0 2.2 3.4h9.6A2.4 2.4 0 0 0 19 17l-4-9V3"/><path d="M7 14h10"/></symbol>
<symbol id="i-crucible" viewBox="0 0 24 24"><path d="M5 5h14l-1.2 9a3 3 0 0 1-3 2.6H9.2a3 3 0 0 1-3-2.6L5 5Z"/><path d="M8 21h8M9 2h6"/></symbol>
<symbol id="i-search" viewBox="0 0 24 24"><circle cx="10.5" cy="10.5" r="6.5"/><path d="m15.5 15.5 5 5"/></symbol>
<symbol id="i-download" viewBox="0 0 24 24"><path d="M12 3v12M7 10l5 5 5-5M4 21h16"/></symbol>
<symbol id="i-refresh" viewBox="0 0 24 24"><path d="M20 7v5h-5M4 17v-5h5"/><path d="M18.5 9A7.5 7.5 0 0 0 5 7.5M5.5 15A7.5 7.5 0 0 0 19 16.5"/></symbol>
<symbol id="i-table" viewBox="0 0 24 24"><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M3 9h18M3 15h18M9 9v11M15 9v11"/></symbol>
<symbol id="i-eye" viewBox="0 0 24 24"><path d="M2.5 12s3.5-6 9.5-6 9.5 6 9.5 6-3.5 6-9.5 6-9.5-6-9.5-6Z"/><circle cx="12" cy="12" r="2.5"/></symbol>
<symbol id="i-database" viewBox="0 0 24 24"><ellipse cx="12" cy="5" rx="8" ry="3"/><path d="M4 5v7c0 1.7 3.6 3 8 3s8-1.3 8-3V5M4 12v7c0 1.7 3.6 3 8 3s8-1.3 8-3v-7"/></symbol>
<symbol id="i-book" viewBox="0 0 24 24"><path d="M4 5.5A2.5 2.5 0 0 1 6.5 3H20v16H6.5A2.5 2.5 0 0 0 4 21.5v-16Z"/><path d="M4 18.5A2.5 2.5 0 0 1 6.5 16H20"/></symbol>
<symbol id="i-survey" viewBox="0 0 24 24"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2h6v2M8 10h8M8 14h8M8 18h5"/></symbol>
<symbol id="i-warning" viewBox="0 0 24 24"><path d="M12 3 2.8 19h18.4L12 3Z"/><path d="M12 9v4M12 17h.01"/></symbol>
<symbol id="i-calendar" viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M8 3v4M16 3v4M3 10h18"/></symbol>
<symbol id="i-filter" viewBox="0 0 24 24"><path d="M4 5h16l-6 7v6l-4 2v-8L4 5Z"/></symbol>
<symbol id="i-external" viewBox="0 0 24 24"><path d="M14 4h6v6M20 4l-9 9"/><path d="M18 13v5a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h5"/></symbol>
"""


def icon(name: str, extra_class: str = "") -> Markup:
    classes = f"icon {extra_class}".strip()
    return Markup(
        f'<svg class="{classes}" aria-hidden="true" focusable="false">'
        f'<use href="#i-{name}"></use></svg>'
    )


def nav_url(item: dict) -> str:
    if item.get("endpoint"):
        return url_for(item["endpoint"])
    if item.get("slug") == "overview":
        return url_for("index")
    return url_for("topic_page", slug=item["slug"])


def is_active(item: dict) -> bool:
    if item.get("endpoint"):
        return request.endpoint == item["endpoint"]
    slug = item.get("slug")
    if slug == "overview" and request.endpoint == "index":
        return True
    return request.endpoint == "topic_page" and request.view_args.get("slug") == slug


BASE_TEMPLATE = r"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{% if page_title == project.title %}{{ project.title }}{% else %}{{ page_title }} · {{ project.title }}{% endif %}</title>
  <meta name="description" content="{{ page_description }}">
  <meta name="theme-color" content="#0b2545">
  <link rel="icon" href="data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'%3E%3Crect width='64' height='64' rx='16' fill='%230b2545'/%3E%3Cpath d='M48 14C32 14 16 21 16 36c0 7 5 12 12 12 16 0 20-18 20-34Z' fill='none' stroke='%239be7df' stroke-width='4' stroke-linejoin='round'/%3E%3Cpath d='M13 52c7-16 18-24 31-30' fill='none' stroke='%239be7df' stroke-width='4' stroke-linecap='round'/%3E%3C/svg%3E">
  {% if noindex %}<meta name="robots" content="noindex, nofollow">{% endif %}
  <style>{{ css|safe }}</style>
</head>
<body class="{{ body_class }}">
  <a class="skip-link" href="#main">Skip to content</a>
  <div class="app-shell">
    <div class="scrim" data-nav-close></div>
    <aside class="sidebar" id="site-nav" aria-label="Project navigation">
      <a class="brand" href="{{ url_for('index') }}">
        <span class="brand-mark">{{ icon('leaf') }}</span>
        <span>
          <strong>NutriLab</strong>
          <small>Chemistry IP · 2026–27</small>
        </span>
      </a>
      <div class="sidebar-project">Study of Nutrient Contents in Breakfast Items</div>
      <nav class="side-nav">
        {% for group in sidebar_groups %}
        <div class="nav-group">
          <p>{{ group.label }}</p>
          {% for item in group['items'] %}
          {% set active = is_active(item) %}
          <a class="nav-link{% if active %} active{% endif %}" href="{{ nav_url(item) }}"{% if active %} aria-current="page"{% endif %}>{{ item.label }}</a>
          {% endfor %}
        </div>
        {% endfor %}
      </nav>
      <div class="sidebar-foot">
        <a href="{{ url_for('topic_page', slug='project_info') }}">Project information</a>
        <p>Qualitative chemistry · Not medical or dietary advice.</p>
      </div>
    </aside>

    <div class="main-wrap">
      <header class="topbar">
        <button class="menu-btn" type="button" data-nav-toggle aria-controls="site-nav" aria-expanded="false">
          {{ icon('menu') }}<span>Menu</span>
        </button>
        <div class="topbar-context">
          <span>{{ eyebrow }}</span>
          <strong>{{ page_title }}</strong>
        </div>
        {% if show_survey_cta %}
        <a class="btn btn-small" href="{{ url_for('survey') }}">Take the survey {{ icon('arrow') }}</a>
        {% endif %}
      </header>

      <main id="main">{{ content|safe }}</main>

      <footer class="site-footer">
        <div class="footer-inner">
          <div class="footer-brand">
            <span class="footer-mark">{{ icon('leaf') }}</span>
            <div>
              <strong>{{ project.title }}</strong>
              <p>{{ project.footer_lede }}</p>
            </div>
          </div>
          <div class="footer-links">
            <a href="{{ url_for('survey') }}">Student survey</a>
            <a href="{{ url_for('survey_results') }}">Live results</a>
            <a href="{{ url_for('topic_page', slug='project_info') }}">Project information</a>
            <a href="{{ url_for('topic_page', slug='project_info') }}#credits">Image credits</a>
          </div>
          <div class="footer-meta">
            <p>{{ project.footer_meta }}</p>
            <p>Food photography via Wikimedia Commons · Charts require internet access.</p>
          </div>
        </div>
      </footer>
    </div>
  </div>
  <svg class="icon-sprite" aria-hidden="true" xmlns="http://www.w3.org/2000/svg">{{ icon_symbols|safe }}</svg>
  {% if needs_charts %}<script src="{{ chart_cdn }}"></script>{% endif %}
  <script>{{ base_js|safe }}</script>
  {% if live_js %}<script>{{ live_js|safe }}</script>{% endif %}
  {% if admin_js %}<script>{{ admin_js|safe }}</script>{% endif %}
</body>
</html>
"""


def render_page(
    content: str,
    *,
    page_title: str,
    page_description: str,
    eyebrow: str = "Chemistry IP",
    needs_charts: bool = False,
    admin: bool = False,
    noindex: bool = False,
    show_survey_cta: bool = True,
    live_js: str = "",
    admin_js: str = "",
) -> str:
    return render_template_string(
        BASE_TEMPLATE,
        content=Markup(content),
        page_title=page_title,
        page_description=page_description,
        eyebrow=eyebrow,
        body_class="admin-body" if admin else "",
        needs_charts=needs_charts,
        noindex=noindex,
        show_survey_cta=show_survey_cta,
        project=PROJECT,
        sidebar_groups=SIDEBAR_GROUPS,
        icon=icon,
        nav_url=nav_url,
        is_active=is_active,
        icon_symbols=Markup(ICON_SYMBOLS),
        css=CSS,
        base_js=BASE_JS,
        live_js=live_js,
        admin_js=admin_js,
        chart_cdn=CHART_CDN,
    )

# ---------------------------------------------------------------------------
# Project topic pages
# ---------------------------------------------------------------------------
TOPIC_TEMPLATE = r"""
{% macro level_chip(label, level) -%}
<span class="level-chip level-{{ level }}">{{ label }}</span>
{%- endmacro %}

{% macro chart_card(key, title, tall=False) -%}
<article class="chart-card" data-chart-card="{{ key }}">
  <div class="chart-head">
    <h3>{{ title }}</h3>
    <span>Live</span>
  </div>
  <div class="chart-canvas{% if tall %} tall{% endif %}">
    <canvas id="home-{{ key }}" data-chart-canvas aria-label="{{ title }} chart" role="img"></canvas>
  </div>
  <div class="chart-fallback" data-chart-fallback></div>
</article>
{%- endmacro %}

<div class="page">
{% for block in topic.blocks %}
  {% if block.type == 'hero' %}
  <section class="hero reveal">
    <div class="hero-copy">
      <div class="kicker">{{ block.eyebrow }}</div>
      <h1>{{ block.title_prefix }} <span class="accent">{{ block.title_accent }}</span></h1>
      <p class="hero-subtitle">{{ block.subtitle }}</p>
      <p class="hero-lede">{{ block.lede }}</p>
      <div class="hero-actions">
        <a class="btn btn-light" href="{{ url_for('survey') }}">Take the student survey {{ icon('survey') }}</a>
        <a class="btn btn-ghost" href="#project-map">Explore the investigation {{ icon('arrow') }}</a>
      </div>
      <div class="hero-metrics">
        {% for metric in block.metrics %}
        <div class="hero-metric">
          <strong>{{ metric.value }}</strong>
          <span>{{ metric.label }}</span>
        </div>
        {% endfor %}
      </div>
    </div>
    <div class="hero-media">
      {% if block.image %}
      <img src="{{ block.image }}" alt="{{ block.image_alt }}" fetchpriority="high" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
      {% else %}
      <div class="media-fallback"></div>
      {% endif %}
      <div class="hero-ticker">
        {% for item in block.ticker %}<span>{{ item }}</span>{% endfor %}
        <small>{{ block.ticker_note }}</small>
      </div>
    </div>
  </section>

  {% elif block.type == 'prose' %}
  <section class="section reveal">
    <div class="split">
      <div class="prose">
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lead">{{ block.lead }}</p>
        {% for paragraph in block.paragraphs %}<p>{{ paragraph }}</p>{% endfor %}
        {% if block.note %}<div class="note info">{{ block.note }}</div>{% endif %}
      </div>
      <aside class="card flat" style="background:#fbfdff;">
        <div class="kicker">At a glance</div>
        <h3>Four samples. Eight tests.</h3>
        <p class="muted small">A qualitative bench comparison of familiar breakfasts, from starch and sugars to calcium, iron and pH.</p>
        <div class="metric-grid" style="grid-template-columns:1fr 1fr;margin-top:16px;">
          <div class="metric"><span>Samples</span><strong>04</strong></div>
          <div class="metric"><span>Tests</span><strong>08</strong></div>
          <div class="metric"><span>Nutrient families</span><strong>03</strong></div>
          <div class="metric"><span>Mineral targets</span><strong>02</strong></div>
        </div>
        <div class="note" style="margin-top:16px;">The tests identify presence and relative response — not exact nutrient quantities.</div>
      </aside>
    </div>
  </section>

  {% elif block.type == 'aim' %}
  <section class="section reveal">
    <div class="card aim-card">
      <div class="aim-text">
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p>{{ block.text }}</p>
      </div>
      <ol class="objectives">
        {% for objective in block.objectives %}<li>{{ objective }}</li>{% endfor %}
      </ol>
    </div>
  </section>

  {% elif block.type == 'samples' %}
  <section class="section reveal" id="samples">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="sample-grid">
      {% for sample in block['items'] %}
      <article class="card sample-card">
        <div class="sample-media">
          {% if sample.image %}
          <img src="{{ sample.image }}" alt="{{ sample.alt }}" loading="lazy" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
          {% endif %}
          <span class="sample-index">{{ sample.index }}</span>
        </div>
        <div class="sample-body">
          <span class="sample-tag">{{ sample.tag }}</span>
          <h3>{{ sample.name }}</h3>
          <p>{{ sample.description }}</p>
          <div class="finding-row">
            {% for finding in sample.findings %}{{ level_chip(finding.label, finding.level) }}{% endfor %}
          </div>
          <details>
            <summary>Open sample dossier</summary>
            <div class="detail-copy">
              <p>{{ sample.detail }}</p>
              <h4>Nutrition context</h4>
              <ul>{% for note in sample.nutrition %}<li>{{ note }}</li>{% endfor %}</ul>
              <h4>Recorded observations</h4>
              <ul>{% for observation in sample.observations %}<li>{{ observation }}</li>{% endfor %}</ul>
            </div>
          </details>
        </div>
      </article>
      {% endfor %}
    </div>
  </section>

  {% elif block.type == 'survey_insights' %}
  <section class="insights-band reveal" data-home-insights>
    <div class="insights-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p>{{ block.lede }}</p>
      </div>
      <a class="text-link" href="{{ url_for('survey_results') }}">See all results {{ icon('arrow') }}</a>
    </div>
    <div class="empty-state" data-empty-state>No survey responses yet. Be the first to <a href="{{ url_for('survey') }}">take the survey</a>.</div>
    <div class="summary-grid">
      <div class="summary-card"><span>Responses</span><strong data-summary="total" data-scope="home">—</strong><small>Aggregate count</small></div>
      <div class="summary-card"><span>Most preferred</span><strong data-summary="preferred" data-scope="home" style="font-size:1.05rem;">—</strong><small>Top breakfast choice</small></div>
      <div class="summary-card"><span>Breakfast habit</span><strong data-summary="frequency" data-scope="home" style="font-size:1.05rem;">—</strong><small>Most common frequency</small></div>
      <div class="summary-card"><span>Last response</span><strong data-summary="last" data-scope="home" style="font-size:.9rem;">—</strong><small>Local time</small></div>
    </div>
    <div class="chart-grid three">
      {{ chart_card('class_distribution', 'Class Distribution') }}
      {{ chart_card('preferred_breakfast', 'Preferred Breakfast') }}
      {{ chart_card('breakfast_frequency', 'Breakfast Frequency') }}
    </div>
    <div class="chart-note">
      <span><i class="live-dot"></i> Live aggregate data</span>
      <span>Updated <span data-generated-at>—</span></span>
      <button class="btn btn-small btn-ghost" type="button" data-refresh-home>{{ icon('refresh') }} Refresh</button>
      <span>Public charts never show names, emails or comments.</span>
    </div>
    <p class="field-error" data-chart-error hidden></p>
  </section>

  {% elif block.type == 'topics' %}
  <section class="section reveal" id="project-map">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="grid grid-3">
      {% for item in block['items'] %}
      <a class="card topic-card" href="{{ url_for('topic_page', slug=item.slug) }}">
        <span class="topic-icon">{{ icon(item.icon) }}</span>
        <h3>{{ item.label }}</h3>
        <p>{{ item.description }}</p>
        <span class="text-link">Open section {{ icon('arrow') }}</span>
      </a>
      {% endfor %}
    </div>
  </section>

  {% elif block.type == 'visual' %}
  <section class="section reveal">
    <div class="visual-band{% if block.reverse %} reverse{% endif %}">
      <div class="visual-media">
        {% if block.image %}
        <img src="{{ block.image }}" alt="{{ block.image_alt }}" loading="lazy" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
        {% endif %}
        {% if block.caption %}<div class="visual-caption">{{ block.caption }}</div>{% endif %}
      </div>
      <div class="visual-copy">
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p>{{ block.text }}</p>
        {% if block.note %}<div class="note info">{{ block.note }}</div>{% endif %}
      </div>
    </div>
  </section>

  {% elif block.type == 'gallery' %}
  <section class="section reveal">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="image-gallery">
      {% for item in block['items'] %}
      <figure class="gallery-card">
        {% if item.image %}
        <img src="{{ item.image }}" alt="{{ item.alt }}" loading="lazy" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
        {% endif %}
        <figcaption class="gallery-caption">
          <strong>{{ item.name }}</strong>
          <small>{{ item.tag }}</small>
        </figcaption>
      </figure>
      {% endfor %}
    </div>
  </section>

  {% elif block.type == 'callout' %}
  <section class="section reveal">
    <div class="callout {{ block.tone }}">
      <span class="callout-icon">{{ icon('alert' if block.tone == 'warning' else 'info') }}</span>
      <div>
        <h3>{{ block.title }}</h3>
        <p>{{ block.text }}</p>
      </div>
    </div>
  </section>

  {% elif block.type == 'nutrients' %}
  <section class="section reveal">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="grid grid-4">
      {% for item in block['items'] %}
      <article class="nutrient-card tone-{{ item.tone }}">
        <span class="nutrient-icon">{{ icon(item.icon) }}</span>
        <h3>{{ item.name }}</h3>
        <p>{{ item.description }}</p>
      </article>
      {% endfor %}
    </div>
  </section>

  {% elif block.type == 'materials' %}
  <section class="section reveal">
    <div class="card" data-list>
      <div class="section-head">
        <div>
          <div class="kicker">{{ block.eyebrow }}</div>
          <h2>{{ block.title }}</h2>
          <p class="lede">{{ block.lede }}</p>
        </div>
      </div>
      <div class="list-toolbar">
        <label class="search-box">
          {{ icon('search') }}
          <input type="search" placeholder="Search this list…" aria-label="Search materials" data-list-search>
        </label>
        <span class="small muted">{{ block['items']|length }} items</span>
      </div>
      <div class="item-grid">
        {% for item in block['items'] %}
        <div class="item-card" data-search="{{ item.name|lower }} {{ item.formula|default('')|lower }} {{ item.note|default('')|lower }}">
          <span class="item-num">{{ '%02d'|format(loop.index) }}</span>
          <div>
            <strong>{{ item.name }}</strong>
            {% if item.formula %}<small class="mono">{{ item.formula }}</small>{% endif %}
            {% if item.hazard %}<span class="hazard hazard-{{ item.hazard_class }}">{{ item.hazard }}</span>{% endif %}
            {% if item.note %}<small>{{ item.note }}</small>{% endif %}
          </div>
        </div>
        {% endfor %}
      </div>
      <p class="small muted" data-list-empty hidden>No matching items in this list.</p>
      <div class="note warning hazard-note">{{ block.hazard_note }}</div>
    </div>
  </section>

  {% elif block.type == 'prep' %}
  <section class="section reveal">
    <div class="card prep-card">
      <div class="prep-media">
        {% if block.image %}
        <img src="{{ block.image }}" alt="{{ block.image_alt }}" loading="lazy" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
        {% endif %}
        <div class="prep-media-caption">{{ block.caption }}</div>
      </div>
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
        <div class="note info">
          <strong>{{ block.callout_title }}</strong><br>{{ block.callout_text }}
        </div>
        <ol class="steps">
          {% for step in block['steps'] %}
          <li><strong>{{ step.title }}</strong><span>{{ step.text }}</span></li>
          {% endfor %}
        </ol>
      </div>
    </div>
  </section>

  {% elif block.type == 'tests' %}
  <section class="section reveal">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="test-grid">
      {% for test in block['items'] %}
      <article class="card test-card">
        <div class="test-head">
          <span class="test-icon">{{ icon(test.icon) }}</span>
          <div>
            <h3>{{ test.name }}</h3>
            <p class="test-target">{{ test.target }}</p>
          </div>
          {% if test.danger %}<span class="danger-pill">Requires care</span>{% endif %}
        </div>
        <details>
          <summary>Materials, method &amp; chemistry</summary>
          <div class="test-detail">
            <section><h4>Materials</h4><p>{{ test.materials }}</p></section>
            <section><h4>Procedure</h4><p>{{ test.procedure }}</p></section>
            <section><h4>Positive observation</h4><p>{{ test.observation }}</p></section>
            <section><h4>Chemical basis</h4><p>{{ test.basis }}</p></section>
            <section><h4>Safety</h4><p>{{ test.safety }}</p></section>
          </div>
        </details>
      </article>
      {% endfor %}
    </div>
  </section>

  {% elif block.type == 'results' %}
  <section class="section reveal">
    <div class="card">
      <div class="section-head">
        <div>
          <div class="kicker">{{ block.eyebrow }}</div>
          <h2>{{ block.title }}</h2>
          <p class="lede">{{ block.lede }}</p>
        </div>
      </div>
      <div class="filter-row" role="group" aria-label="Filter observation table">
        {% for filter in block.filters %}
        <button class="filter-chip{% if filter.value == 'all' %} active{% endif %}" type="button" data-results-filter="{{ filter.value }}" aria-pressed="{{ 'true' if filter.value == 'all' else 'false' }}">{{ filter.label }}</button>
        {% endfor %}
      </div>
      <div class="table-wrap">
        <table class="results-table">
          <thead>
            <tr>{% for column in block['columns'] %}<th scope="col">{{ column }}</th>{% endfor %}</tr>
          </thead>
          <tbody>
            {% for row in block.rows %}
            <tr data-result-row data-category="{{ row.category }}">
              <th scope="row">{{ row.test }}</th>
              {% for value in row['values'] %}<td>{{ level_chip(value.label, value.level) }}</td>{% endfor %}
            </tr>
            {% endfor %}
          </tbody>
        </table>
      </div>
      <div class="legend">
        {% for item in block.legend %}<span><i class="lg-{{ item.level }}"></i>{{ item.label }}</span>{% endfor %}
      </div>
      <div class="note info" style="margin-top:16px;">{{ block.note }}</div>
    </div>
  </section>

  {% elif block.type == 'conclusions' %}
  <section class="section reveal">
    <div class="section-head">
      <div>
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <p class="lede">{{ block.lede }}</p>
      </div>
    </div>
    <div class="conclusion-grid">
      {% for item in block['items'] %}
      <article class="card conclusion-card">
        <span class="sample-tag">{{ item.index }} · {{ item.tag }}</span>
        <h3>{{ item.name }}</h3>
        <div class="finding-row">{% for finding in item.findings %}{{ level_chip(finding.label, finding.level) }}{% endfor %}</div>
        <p>{{ item.text }}</p>
      </article>
      {% endfor %}
    </div>
    <div class="card overall">
      <div class="kicker" style="color:#8cded8;">Overall comparison</div>
      <h2>{{ block.overall.title }}</h2>
      <p>{{ block.overall.text }}</p>
      <div class="overall-grid">
        {% for item in block.overall['items'] %}
        <div class="overall-item">
          <small>{{ item.label }}</small>
          <strong>{{ item.strong }}</strong>
          <span>{{ item.text }}</span>
        </div>
        {% endfor %}
      </div>
      <div class="note" style="background:rgba(255,255,255,.08);border-color:#8cded8;color:#cfe7f5;margin-top:18px;">{{ block.note }}</div>
    </div>
  </section>

  {% elif block.type == 'safety' %}
  <section class="section reveal">
    <div class="safety-banner">
      <span class="callout-icon">{{ icon('alert') }}</span>
      <div>
        <h3>{{ block.banner_title }}</h3>
        <p>{{ block.banner_text }}</p>
      </div>
    </div>
    <div class="safety-grid">
      {% for item in block.precautions %}
      <article class="card safety-card">
        <span class="nutrient-icon">{{ icon(item.icon) }}</span>
        <h3>{{ item.title }}</h3>
        <p>{{ item.text }}</p>
      </article>
      {% endfor %}
    </div>
    <div class="card" style="margin-top:16px;">
      <div class="kicker">Personal protective equipment</div>
      <h2 style="font-size:1.5rem;">Dress for the bench.</h2>
      <div class="ppe-row">
        {% for item in block.ppe %}<span class="ppe-item">{{ icon(item.icon) }}{{ item.label }}</span>{% endfor %}
      </div>
      <div class="note warning" style="margin-top:16px;">{{ block.ppe_note }}</div>
    </div>
  </section>

  {% elif block.type == 'project_info' %}
  <section class="section reveal">
    <div class="split">
      <div class="card">
        <div class="kicker">{{ block.eyebrow }}</div>
        <h2>{{ block.title }}</h2>
        <dl class="project-details">
          {% for item in block['details'] %}
          <div><dt>{{ item.label }}</dt><dd>{{ item.value }}</dd></div>
          {% endfor %}
        </dl>
        <div class="note info" style="margin-top:18px;">
          <strong>{{ block.source_title }}</strong><br>{{ block.source_text }}
          <div class="small" style="margin-top:8px;">Source document: {{ block.source_pdf_label }}</div>
        </div>
      </div>
      <div class="card" id="credits">
        <div class="kicker">Attribution</div>
        <h2 style="font-size:1.7rem;">{{ block.credits_title }}</h2>
        <ul class="credit-list">
          {% for credit in block.credits %}
          <li>
            <span class="credit-label">{{ credit.label }}</span>
            <div>
              <a href="{{ credit.url }}" target="_blank" rel="noopener noreferrer">{{ credit.title }}</a>
              <small>{{ credit.author }} · {{ credit.licence }}</small>
            </div>
          </li>
          {% endfor %}
        </ul>
        <p class="small muted" style="margin-top:16px;">External photographs are loaded from Wikimedia Commons. If an image is unavailable, the page shows a styled fallback.</p>
      </div>
    </div>
    <div class="card" id="development" style="margin-top:16px;">
      <div class="kicker">Project development</div>
      <h2 style="font-size:1.5rem;">Built and maintained by</h2>
      <dl class="project-details">
        <div>
          <dt>Website</dt>
          <dd>This site is developed by <a href="https://techorigin.online" target="_blank" rel="noopener noreferrer">techorigin.online</a></dd>
        </div>
        <div>
          <dt>Developer</dt>
          <dd>KAMLESHVAR · <a href="https://kamleshvar.online" target="_blank" rel="noopener noreferrer">kamleshvar.online</a></dd>
        </div>
      </dl>
    </div>
  </section>
  {% endif %}
{% endfor %}
</div>
"""


def render_topic_page(slug: str) -> str:
    topic = TOPICS[slug]
    content = render_template_string(
        TOPIC_TEMPLATE,
        topic=topic,
        icon=icon,
        project=PROJECT,
    )
    if slug == "overview":
        page_title = PROJECT["title"]
        eyebrow = PROJECT["eyebrow"]
        description = PROJECT["subtitle"] + " — four breakfasts, eight qualitative tests, and a live student survey."
    else:
        page_title = topic["title"]
        eyebrow = topic["eyebrow"]
        description = topic["lede"]
    return render_page(
        content,
        page_title=page_title,
        page_description=description,
        eyebrow=eyebrow,
        needs_charts=(slug == "overview"),
        show_survey_cta=True,
        live_js=LIVE_JS if slug == "overview" else "",
    )

# ---------------------------------------------------------------------------
# Student survey
# ---------------------------------------------------------------------------
SURVEY_TEMPLATE = r"""
{% if submitted %}
<div class="page narrow">
  <section class="card success-card reveal">
    <div class="success-mark">{{ icon('check') }}</div>
    <div class="kicker" style="justify-content:center;">Response recorded</div>
    <h1>Thank you for adding your voice.</h1>
    <p>Your answers have been saved. The public results update from aggregate responses only, so your name, email and comments are never shown in the charts.</p>
    <div class="success-actions">
      <a class="btn" href="{{ url_for('survey_results') }}">View live results {{ icon('chart') }}</a>
      <a class="btn btn-light" href="{{ url_for('index') }}">Return to the project {{ icon('arrow') }}</a>
    </div>
  </section>
</div>
{% else %}
<div class="page">
  <div class="survey-hero">
    <section class="card reveal">
      <div class="kicker">Live student survey</div>
      <h1>What does breakfast mean to you?</h1>
      <p class="lede">Sixteen short questions about your breakfast habits, nutrition awareness and the chemistry you have explored in this project. It takes about three minutes.</p>
      <div class="survey-meta">
        <span class="pill"><i class="dot"></i> Anonymous public results</span>
        <span class="pill">{{ icon('clock') }} 3 minutes</span>
        <span class="pill">{{ icon('survey') }} 16 questions</span>
      </div>
    </section>
    <aside class="card survey-side reveal">
      <div class="survey-side-media">
        <img src="https://commons.wikimedia.org/wiki/Special:FilePath/Healthy_breakfast_%28Unsplash%29.jpg?width=1400" alt="A breakfast table with several morning foods" loading="lazy">
      </div>
      <div class="kicker" style="color:#8cded8;">Before you begin</div>
      <h2>Your answers, handled carefully.</h2>
      <p>Names and email addresses are stored only for the project team. Public charts show aggregate counts — never individual responses.</p>
      <ul>
        <li>{{ icon('check') }} Q1–Q15 are required; the final comment is optional.</li>
        <li>{{ icon('check') }} Your name is limited to 80 characters.</li>
        <li>{{ icon('check') }} Comments are limited to 500 characters.</li>
        <li>{{ icon('check') }} Please do not include phone numbers, addresses or other sensitive information.</li>
      </ul>
    </aside>
  </div>

  <div class="progress-shell">
    <div class="progress-top">
      <span>Survey progress</span>
      <span data-progress-label>0 of 16 required answers</span>
    </div>
    <div class="progress-track" role="progressbar" aria-label="Survey progress" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0">
      <div class="progress-fill" data-progress-fill></div>
    </div>
  </div>

  {% if errors.get('_form') %}<div class="form-error" role="alert">{{ errors['_form'] }}</div>{% endif %}

  <form class="survey-form" method="post" action="{{ url_for('survey_submit') }}" data-survey-form>
    <input type="hidden" name="form_token" value="{{ form_token }}">

    {% for section in survey_sections %}
    <section class="card form-section reveal">
      <div class="form-section-head">
        <span class="section-number">{{ '%02d'|format(loop.index) }}</span>
        <div>
          <h2>{{ section.title }}</h2>
          <p>{{ section.description }}</p>
        </div>
      </div>

      {% for question in section.questions %}
      {% set value = form_data.get(question.field, '') %}
      {% set error = errors.get(question.field) %}
      <div class="question{% if error %} has-error{% endif %}"{% if question.required %} data-required-question{% endif %}>
        <div class="question-label">
          <span class="question-number">{{ '%02d'|format(question.number) }}</span>
          <span>{{ question.label }}{% if question.required %}<span class="required-mark" aria-hidden="true"> *</span>{% endif %}</span>
        </div>
        <p class="question-help">{{ question.help }}</p>
        {% if question.project_note %}<div class="project-note">{{ question.project_note }}</div>{% endif %}

        <div class="field-control">
          {% if question.type == 'radio' %}
          <div class="option-grid" role="radiogroup" aria-label="{{ question.label }}">
            {% for option in question['options'] %}
            <label class="option">
              <input type="radio" name="{{ question.field }}" value="{{ option.value }}"{% if value == option.value %} checked{% endif %}{% if question.required %} required{% endif %}{% if error %} aria-invalid="true" aria-describedby="{{ question.field }}-error"{% endif %}>
              <span>{{ option.label }}</span>
            </label>
            {% endfor %}
          </div>
          {% elif question.type == 'image-radio' %}
          <div class="image-options" role="radiogroup" aria-label="{{ question.label }}">
            {% for option in question['options'] %}
            <label class="image-option">
              <input type="radio" name="{{ question.field }}" value="{{ option.value }}"{% if value == option.value %} checked{% endif %}{% if question.required %} required{% endif %}{% if error %} aria-invalid="true" aria-describedby="{{ question.field }}-error"{% endif %}>
              <span class="image-option-media{% if not option.image %} media-fallback{% endif %}">
                {% if option.image %}
                <img src="{{ option.image }}" alt="{{ option.alt }}" loading="lazy" onerror="this.parentElement.classList.add('media-fallback');this.remove();">
                {% else %}
                <span style="position:absolute;inset:0;display:grid;place-items:center;color:#7d9ab1;">{{ icon('leaf') }}</span>
                {% endif %}
              </span>
              <span class="image-option-label">{{ option.label }}</span>
            </label>
            {% endfor %}
          </div>
          {% elif question.type == 'textarea' %}
          <textarea class="textarea" id="{{ question.field }}" name="{{ question.field }}" maxlength="{{ question.maxlength }}" placeholder="{{ question.placeholder }}"{% if error %} aria-invalid="true" aria-describedby="{{ question.field }}-error"{% endif %}>{{ value|e }}</textarea>
          <div class="small muted" style="text-align:right;margin-top:4px;"><span data-comment-count>0 / 500</span></div>
          {% else %}
          <input class="text-input" id="{{ question.field }}" type="{{ 'email' if question.type == 'email' else 'text' }}" name="{{ question.field }}" value="{{ value|e }}" maxlength="{{ question.maxlength }}" autocomplete="{{ question.autocomplete }}" placeholder="{{ question.placeholder }}"{% if question.required %} required{% endif %}{% if error %} aria-invalid="true" aria-describedby="{{ question.field }}-error"{% endif %}>
          {% endif %}
          {% if error %}<span class="field-error" id="{{ question.field }}-error">{{ error }}</span>{% endif %}
        </div>
      </div>
      {% endfor %}
    </section>
    {% endfor %}

    <section class="card form-section reveal">
      <div class="consent-box">
        <label class="consent-label">
          <input type="checkbox" name="consent" data-consent required{% if form_data.get('consent') == 'on' %} checked{% endif %}>
          <span>I consent to my name and email address being stored securely for this Chemistry IP survey. I understand that public results show aggregate counts only.</span>
        </label>
        {% if errors.get('consent') %}<span class="field-error">{{ errors['consent'] }}</span>{% endif %}
      </div>
      <div class="form-actions" style="margin-top:16px;">
        <button class="btn btn-teal" type="submit" data-submit>Submit response {{ icon('arrow') }}</button>
        <small>Your response is saved only after you submit this form. Required questions are marked with an asterisk.</small>
      </div>
    </section>
  </form>
</div>
{% endif %}
"""


def render_survey_page(
    *,
    submitted: bool = False,
    errors: dict[str, str] | None = None,
    form_data: dict[str, str] | None = None,
    status: int = 200,
):
    token = issue_form_token()

    content = render_template_string(
        SURVEY_TEMPLATE,
        submitted=submitted,
        errors=errors or {},
        form_data=form_data or {},
        form_token=token,
        survey_sections=SURVEY_SECTIONS,
        icon=icon,
        project=PROJECT,
    )
    html = render_page(
        content,
        page_title="Student Survey",
        page_description="Share your breakfast habits and nutrition awareness in the Chemistry IP student survey.",
        eyebrow="Live student survey",
        show_survey_cta=False,
    )
    return html, status

# ---------------------------------------------------------------------------
# Live public results and the unlisted admin dashboard
# ---------------------------------------------------------------------------
RESULTS_TEMPLATE = r"""
<div class="page" data-results-page>
  <div class="results-page-head">
    <div>
      <div class="kicker">Live aggregate results</div>
      <h1>Survey Results</h1>
      <p class="lede">Every chart below is built from saved survey responses. Public results contain aggregate counts only — never names, emails, comments or individual records.</p>
    </div>
    <div class="results-actions">
      <a class="btn btn-light" href="{{ url_for('survey') }}">Take the survey {{ icon('survey') }}</a>
      <button class="btn btn-ghost" type="button" data-refresh-results>{{ icon('refresh') }} Refresh</button>
    </div>
  </div>

  <div class="empty-state" data-empty-state>No survey responses yet. <a href="{{ url_for('survey') }}">Be the first to take the survey</a> — these charts will populate automatically.</div>

  <div class="summary-grid">
    <div class="summary-card"><span>Total responses</span><strong data-summary="total" data-scope="results">—</strong><small>Saved submissions</small></div>
    <div class="summary-card"><span>Classes represented</span><strong data-summary="classes" data-scope="results">—</strong><small>Different class groups</small></div>
    <div class="summary-card"><span>Most preferred</span><strong data-summary="preferred" data-scope="results" style="font-size:1.04rem;">—</strong><small>Top breakfast choice</small></div>
    <div class="summary-card"><span>Most common habit</span><strong data-summary="frequency" data-scope="results" style="font-size:1.04rem;">—</strong><small>Breakfast frequency</small></div>
  </div>

  <div class="chart-grid" style="margin-top:14px;">
    {% for chart in chart_defs %}
    <article class="chart-card" data-chart-card="{{ chart.key }}">
      <div class="chart-head">
        <h3>{{ chart.title }}</h3>
        <span>Live</span>
      </div>
      <div class="chart-canvas{% if chart.key in ['preferred_breakfast', 'learning_method', 'interesting_topic', 'perceived_grease_spot'] %} tall{% endif %}">
        <canvas id="results-{{ chart.key }}" data-chart-canvas aria-label="{{ chart.title }} chart" role="img"></canvas>
      </div>
      <div class="chart-fallback" data-chart-fallback></div>
    </article>
    {% endfor %}
  </div>

  <div class="chart-note">
    <span><i class="live-dot"></i> Live aggregate data</span>
    <span>Updated <span data-generated-at>—</span></span>
    <span class="loading-line" data-loading hidden><span class="spinner"></span> Refreshing</span>
    <span>Auto-refreshes every 60 seconds while this tab is visible.</span>
  </div>
  <p class="field-error" data-chart-error hidden></p>
  <div class="note info" style="margin-top:16px;">If Chart.js cannot load, the same figures are shown as readable fallback tables. Charts require internet access for the pinned Chart.js CDN.</div>
  <noscript><div class="noscript-note">JavaScript is required for the live charts. The underlying survey data remains safely stored in the database.</div></noscript>
</div>
"""


def render_results_page() -> str:
    content = render_template_string(
        RESULTS_TEMPLATE,
        chart_defs=CHART_DEFS,
        icon=icon,
        project=PROJECT,
    )
    return render_page(
        content,
        page_title="Survey Results",
        page_description="Live aggregate survey results for the breakfast nutrient Chemistry IP.",
        eyebrow="Live aggregate results",
        needs_charts=True,
        show_survey_cta=True,
        live_js=LIVE_JS,
    )


ADMIN_TEMPLATE = r"""
<div class="page" data-admin>
  <div class="admin-warning" role="note">
    {{ icon('warning') }}
    <div>
      <strong>Unlisted response dashboard</strong>
      <p>This dashboard has no password. Anyone who knows this URL may access survey responses.</p>
    </div>
  </div>

  <div class="results-page-head">
    <div>
      <div class="kicker">Response dashboard</div>
      <h1>Survey Admin</h1>
      <p class="lede">Read-only access to every submitted response, with filters, pagination and CSV/XLSX export. This page is not linked from the public site.</p>
    </div>
    <div class="results-actions">
      <button class="btn btn-ghost" type="button" data-refresh-admin>{{ icon('refresh') }} Refresh</button>
    </div>
  </div>

  <div class="summary-grid admin-summary-grid">
    <div class="summary-card"><span>Total responses</span><strong data-summary="total" data-scope="admin">—</strong><small>All saved rows</small></div>
    <div class="summary-card"><span>Classes</span><strong data-summary="classes" data-scope="admin">—</strong><small>Represented</small></div>
    <div class="summary-card"><span>Most preferred</span><strong data-summary="preferred" data-scope="admin" style="font-size:.98rem;">—</strong><small>Breakfast choice</small></div>
    <div class="summary-card"><span>Most common</span><strong data-summary="frequency" data-scope="admin" style="font-size:.98rem;">—</strong><small>Frequency</small></div>
    <div class="summary-card"><span>Last response</span><strong data-summary="last" data-scope="admin" style="font-size:.82rem;">—</strong><small>Local time</small></div>
  </div>

  <div class="admin-section-title">
    <div>
      <h2>Aggregate charts</h2>
      <p>The same public aggregates, shown for convenience.</p>
    </div>
    <div class="admin-toolbar" style="margin:0; justify-content:flex-end;">
      <span class="small muted">Updated <span data-generated-at>—</span></span>
      <button type="button" class="btn btn-small btn-light" data-export-charts-pdf>{{ icon('download') }} Charts PDF</button>
      <button type="button" class="btn btn-small btn-light" data-export-charts-docx>{{ icon('download') }} Charts Word</button>
    </div>
  </div>
  <div class="chart-grid admin-chart-grid">
    {% for chart in chart_defs %}
    <article class="chart-card" data-chart-card="{{ chart.key }}">
      <div class="chart-head">
        <h3>{{ chart.title }}</h3>
        <span>Live</span>
      </div>
      <div class="chart-canvas">
        <canvas id="admin-{{ chart.key }}" data-chart-canvas aria-label="{{ chart.title }} chart" role="img"></canvas>
      </div>
      <div class="chart-fallback" data-chart-fallback></div>
    </article>
    {% endfor %}
  </div>

  <div class="admin-section-title">
    <div>
      <h2>Individual responses</h2>
      <p><span data-admin-filtered-count>—</span> matching response(s) in the current filter.</p>
    </div>
    <div class="admin-toolbar" style="margin:0;">
      <a class="btn btn-small btn-light" href="{{ url_for('export_csv') }}" data-export-csv>{{ icon('download') }} CSV</a>
      <a class="btn btn-small btn-light" href="{{ url_for('export_pdf') }}" data-export-pdf>{{ icon('download') }} PDF</a>
      <a class="btn btn-small btn-light" href="{{ url_for('export_docx') }}" data-export-docx>{{ icon('download') }} Word</a>
    </div>
  </div>

  <form class="card admin-filters" data-admin-filters>
    <label class="search-field">Search
      <input class="text-input" type="search" name="q" placeholder="Name, email, answer or comment" autocomplete="off">
    </label>
    <label>Class
      <select class="select-input" name="class_level">
        <option value="">All classes</option>
        {% for option in class_options %}<option value="{{ option }}">{{ option }}</option>{% endfor %}
      </select>
    </label>
    <label>Preferred breakfast
      <select class="select-input" name="preferred_breakfast">
        <option value="">All breakfasts</option>
        {% for option in preferred_options %}<option value="{{ option }}">{{ option }}</option>{% endfor %}
      </select>
    </label>
    <label>Frequency
      <select class="select-input" name="breakfast_frequency">
        <option value="">All frequencies</option>
        {% for option in frequency_options %}<option value="{{ option }}">{{ option }}</option>{% endfor %}
      </select>
    </label>
    <label>From date
      <input class="text-input" type="date" name="date_from">
    </label>
    <label>To date
      <input class="text-input" type="date" name="date_to">
    </label>
    <label>Sort
      <select class="select-input" name="sort">
        <option value="newest">Newest first</option>
        <option value="oldest">Oldest first</option>
      </select>
    </label>
    <label>Rows per page
      <select class="select-input" name="per_page">
        <option value="10">10</option>
        <option value="25">25</option>
        <option value="50">50</option>
        <option value="100">100</option>
      </select>
    </label>
    <div class="admin-filters-actions">
      <button class="btn btn-small" type="submit">{{ icon('filter') }} Apply filters</button>
      <button class="btn btn-small btn-ghost" type="reset">Reset</button>
    </div>
  </form>

  <div class="admin-toolbar">
    <span class="admin-loading" data-admin-loading><span class="spinner"></span> Loading responses</span>
    <span class="small muted">Click a name to open the full response.</span>
  </div>
  <div class="admin-error" data-admin-error role="alert"></div>

  <div class="table-wrap">
    <table class="admin-table">
      <thead>
        <tr>
          <th scope="col">Submitted</th>
          <th scope="col">Name</th>
          <th scope="col">Class</th>
          <th scope="col">Email</th>
          <th scope="col">Frequency</th>
          <th scope="col">Preferred</th>
          <th scope="col">Main reason</th>
          <th scope="col">Eating time</th>
          <th scope="col">Nutrient</th>
          <th scope="col">Familiarity</th>
          <th scope="col">Iodine awareness</th>
          <th scope="col">Grease spot</th>
          <th scope="col">Iodine knowledge</th>
          <th scope="col">Learning method</th>
          <th scope="col">Influence</th>
          <th scope="col">Interesting topic</th>
          <th scope="col">Comment</th>
        </tr>
      </thead>
      <tbody data-admin-table-body>
        <tr><td colspan="17" style="padding:24px 14px;color:#8290a4;">Loading survey responses…</td></tr>
      </tbody>
    </table>
  </div>

  <div class="pagination-bar">
    <span data-pagination-info>Loading pagination…</span>
    <div class="pagination-actions">
      <button class="btn btn-small btn-light" type="button" data-page-prev>Previous</button>
      <button class="btn btn-small btn-light" type="button" data-page-next>Next</button>
    </div>
  </div>

  <dialog class="admin-dialog" data-detail-dialog aria-labelledby="detail-title">
    <div class="dialog-head">
      <div>
        <h2 id="detail-title">Full response</h2>
        <p>Personal data — visible only on this unlisted dashboard.</p>
      </div>
      <button class="dialog-close" type="button" data-dialog-close aria-label="Close response details">{{ icon('close') }}</button>
    </div>
    <div class="dialog-body" data-detail-body></div>
  </dialog>

  <div class="note warning" style="margin-top:18px;">
    This page is read-only and intentionally password-free for project review. Do not share the URL beyond the intended audience. Names, emails and comments are personal data.
  </div>
</div>
"""


def render_admin_page() -> str:
    content = render_template_string(
        ADMIN_TEMPLATE,
        chart_defs=CHART_DEFS,
        class_options=CLASS_OPTIONS,
        preferred_options=PREFERRED_OPTIONS,
        frequency_options=FREQUENCY_OPTIONS,
        icon=icon,
        project=PROJECT,
    )
    return render_page(
        content,
        page_title="Survey Admin",
        page_description="Unlisted read-only survey response dashboard.",
        eyebrow="Response dashboard",
        needs_charts=True,
        admin=True,
        noindex=True,
        show_survey_cta=False,
        admin_js=ADMIN_JS,
    )

# ---------------------------------------------------------------------------
# Request handling, routes and security headers
# ---------------------------------------------------------------------------
ERROR_TEMPLATE = r"""
<div class="page narrow">
  <section class="card success-card reveal">
    <div class="success-mark" style="background:var(--orange-soft);color:var(--orange);">{{ icon('alert') }}</div>
    <div class="kicker" style="justify-content:center;">{{ code }}</div>
    <h1>{{ title }}</h1>
    <p>{{ message }}</p>
    <div class="success-actions">
      <a class="btn" href="{{ url_for('index') }}">Return to the project {{ icon('arrow') }}</a>
      <a class="btn btn-light" href="{{ url_for('survey_results') }}">View live results {{ icon('chart') }}</a>
    </div>
  </section>
</div>
"""


def render_error_page(code: str, title: str, message: str, status: int):
    content = render_template_string(
        ERROR_TEMPLATE,
        code=code,
        title=title,
        message=message,
        icon=icon,
        project=PROJECT,
    )
    return (
        render_page(
            content,
            page_title=title,
            page_description=message,
            eyebrow=f"Error {code}",
            show_survey_cta=True,
        ),
        status,
    )


@app.after_request
def security_headers(response: Response):
    response.headers.setdefault("X-Content-Type-Options", "nosniff")
    response.headers.setdefault("X-Frame-Options", "SAMEORIGIN")
    response.headers.setdefault("Referrer-Policy", "strict-origin-when-cross-origin")
    response.headers.setdefault(
        "Permissions-Policy", "geolocation=(), microphone=(), camera=(), payment=()"
    )
    response.headers.setdefault(
        "Content-Security-Policy",
        "default-src 'self'; "
        "img-src 'self' data: https://commons.wikimedia.org https://upload.wikimedia.org https://thumb.wikimedia.org; "
        "style-src 'self' 'unsafe-inline'; "
        "script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        "font-src 'self' data:; "
        "connect-src 'self'; "
        "object-src 'none'; "
        "base-uri 'self'; "
        "form-action 'self'; "
        "frame-ancestors 'self'",
    )
    if request.path.startswith("/admin") or request.path.startswith("/api/admin"):
        response.headers["Cache-Control"] = "no-store"
        response.headers["X-Robots-Tag"] = "noindex, nofollow"
    return response


@app.get("/")
def index():
    return render_topic_page("overview")


@app.get("/topic/<slug>")
def topic_page(slug: str):
    if slug not in TOPICS:
        abort(404)
    return render_topic_page(slug)


@app.get("/survey")
def survey():
    return render_survey_page(submitted=request.args.get("submitted") == "1")


@app.post("/survey/submit")
@limiter.limit("12 per hour; 1 per 15 seconds", methods=["POST"])
def survey_submit():
    posted_token = request.form.get("form_token", "")
    expected_token = session.pop("survey_form_token", None)
    form_data = request.form.to_dict()
    errors: dict[str, str] = {}

    if (
        not expected_token
        or not posted_token
        or not hmac.compare_digest(posted_token, expected_token)
    ):
        errors["_form"] = (
            "This form session has expired. Please review your answers and submit again."
        )
        return render_survey_page(errors=errors, form_data=form_data, status=400)

    cleaned, errors = validate_survey(request.form)
    if errors:
        return render_survey_page(errors=errors, form_data=form_data, status=400)

    allowed, rate_message = check_session_rate_limit()
    if not allowed:
        errors["_form"] = rate_message or "Please wait before submitting again."
        return render_survey_page(errors=errors, form_data=form_data, status=429)

    signature = submission_signature(cleaned)
    if is_duplicate_in_session(signature):
        errors["_form"] = "This exact response was already submitted from this browser."
        return render_survey_page(errors=errors, form_data=form_data, status=429)

    submitted_at = utcnow().strftime("%Y-%m-%d %H:%M:%S")
    values = []
    for field in RESPONSE_COLUMNS:
        if field == "submitted_at":
            values.append(submitted_at)
        elif field == "consent_given":
            values.append(1)
        else:
            values.append(cleaned.get(field, ""))

    connection = get_db()
    try:
        placeholders = ", ".join("?" for _ in RESPONSE_COLUMNS)
        connection.execute(
            "INSERT INTO survey_responses (" + ", ".join(RESPONSE_COLUMNS) + ") "
            "VALUES (" + placeholders + ")",
            tuple(values),
        )
        connection.commit()
    except sqlite3.Error:
        connection.rollback()
        app.logger.exception("Survey submission could not be saved")
        errors["_form"] = (
            "Your response could not be saved because of a database error. Please try again."
        )
        return render_survey_page(errors=errors, form_data=form_data, status=500)
    finally:
        connection.close()

    record_session_submission(signature)
    return redirect(url_for("survey", submitted=1), code=303)


@app.get("/survey/results")
def survey_results():
    return render_results_page()


@app.get("/api/survey/summary")
def api_survey_summary():
    try:
        payload = build_summary()
    except sqlite3.Error:
        app.logger.exception("Could not build survey summary")
        return jsonify({"error": "Survey analytics are temporarily unavailable."}), 500
    response = jsonify(payload)
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/admin")
def admin():
    return render_admin_page()


@app.get("/api/admin/responses")
def api_admin_responses():
    filters = parse_filters(request.args)
    try:
        rows, total, page, pages = fetch_admin_page(filters)
    except sqlite3.Error:
        app.logger.exception("Could not load admin responses")
        return jsonify({"error": "Survey responses are temporarily unavailable."}), 500

    filters["page"] = page
    payload = {
        "items": [serialize_response(row) for row in rows],
        "pagination": {
            "page": page,
            "per_page": filters["per_page"],
            "total": total,
            "pages": pages,
            "has_prev": page > 1,
            "has_next": page < pages,
        },
        "filters": filters_payload(filters),
    }
    response = jsonify(payload)
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/admin/export.csv")
def export_csv():
    filters = parse_filters(request.args)
    try:
        data = build_csv(filters)
    except sqlite3.Error:
        app.logger.exception("Could not export survey responses as CSV")
        return "Survey responses are temporarily unavailable.", 500
    response = send_file(
        io.BytesIO(data),
        mimetype="text/csv; charset=utf-8",
        as_attachment=True,
        download_name="chemistry_ip_survey_responses.csv",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/admin/export.docx")
def export_docx():
    filters = parse_filters(request.args)
    try:
        output = build_docx(filters)
    except sqlite3.Error:
        app.logger.exception("Could not export survey responses as Word document")
        return "Survey responses are temporarily unavailable.", 500
    response = send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="chemistry_ip_survey_responses.docx",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/admin/export.pdf")
def export_pdf():
    filters = parse_filters(request.args)
    try:
        output = build_pdf(filters)
    except sqlite3.Error:
        app.logger.exception("Could not export survey responses as PDF")
        return "Survey responses are temporarily unavailable.", 500
    response = send_file(
        output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="chemistry_ip_survey_responses.pdf",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@app.post("/admin/export.charts.pdf")
def export_charts_pdf():
    images = decode_chart_images(request.get_json(silent=True))
    try:
        output = build_chart_report_pdf(images)
    except sqlite3.Error:
        app.logger.exception("Could not export aggregate charts as PDF")
        return "Survey responses are temporarily unavailable.", 500
    response = send_file(
        output,
        mimetype="application/pdf",
        as_attachment=True,
        download_name="chemistry_ip_aggregate_charts.pdf",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@app.post("/admin/export.charts.docx")
def export_charts_docx():
    images = decode_chart_images(request.get_json(silent=True))
    try:
        output = build_chart_report_docx(images)
    except sqlite3.Error:
        app.logger.exception("Could not export aggregate charts as Word document")
        return "Survey responses are temporarily unavailable.", 500
    response = send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="chemistry_ip_aggregate_charts.docx",
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@app.errorhandler(404)
def not_found(error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "Not found."}), 404
    return render_error_page(
        "404",
        "This page is not on the bench.",
        "The page you asked for does not exist. Use the project navigation to return to the investigation.",
        404,
    )


@app.errorhandler(429)
def too_many_requests(error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "Too many requests. Please try again later."}), 429
    if request.path == "/survey/submit" and request.method == "POST":
        return render_survey_page(
            errors={
                "_form": "Too many submission attempts. Please wait a few minutes and try again."
            },
            form_data=request.form.to_dict(),
            status=429,
        )
    return render_error_page(
        "429",
        "A short pause is needed.",
        "Too many requests were made from this connection. Please wait a moment and try again.",
        429,
    )


@app.errorhandler(500)
def server_error(error):
    if request.path.startswith("/api/"):
        return jsonify({"error": "Something went wrong."}), 500
    return render_error_page(
        "500",
        "Something went wrong.",
        "The server could not complete that request. Please try again in a moment.",
        500,
    )


if __name__ == "__main__":
    app.run(
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5555")),
        debug=os.environ.get("FLASK_DEBUG") == "1",
    )

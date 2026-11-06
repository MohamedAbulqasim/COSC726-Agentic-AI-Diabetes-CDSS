"""
Script: generate_presentation_after_agents_english.py
Generates the comprehensive academic Master's presentation in English (.pptx)
for the Agentic Multi-Agent Diabetes Decision Support System.
Matches the visual identity, colors, and layout of the original presentation,
while comprehensively documenting all features, benchmarks, metrics, and academic lineage.
"""

import os
import sys
import pptx
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# Colors Palette
COPPER_ORANGE = RGBColor(0xD9, 0x77, 0x24)   # #d97724
COPPER_DARK   = RGBColor(0xC8, 0x64, 0x20)   # #c86420
DARK_NAVY     = RGBColor(0x1E, 0x29, 0x3B)   # #1e293b
DEEP_SLATE    = RGBColor(0x0F, 0x17, 0x2A)   # #0f172a
CARD_BG_DARK  = RGBColor(0x1E, 0x29, 0x3B)   # #1e293b
CARD_BORDER   = RGBColor(0x33, 0x41, 0x55)   # #334155
CLINICAL_GREEN= RGBColor(0x1A, 0x4D, 0x2E)   # #1a4d2e
EMERALD       = RGBColor(0x05, 0x96, 0x69)   # #059669
AMBER_YELLOW  = RGBColor(0xF5, 0x9E, 0x0B)   # #f59e0b
TEXT_WHITE    = RGBColor(0xFF, 0xFF, 0xFF)   # #ffffff
TEXT_MUTED    = RGBColor(0x94, 0xA3, 0xB8)   # #94a3b8
LIGHT_GRAY_BG = RGBColor(0xF8, 0xFA, 0xFC)   # #f8fafc
BORDER_LIGHT  = RGBColor(0xCB, 0xD5, 0xE1)   # #cbd5e1

# Directories
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ASSETS_DIR = os.path.join(BASE_DIR, "presentation_assets")

LOGO_SHIELD = os.path.join(ASSETS_DIR, "al_neelain_shield.png")
FLAG_SUDAN  = os.path.join(ASSETS_DIR, "sudan_flag.png")
ROBOT_BIG   = os.path.join(ASSETS_DIR, "robot_big.png")
ROBOT_SMALL = os.path.join(ASSETS_DIR, "robot_small.png")


def add_base_decorations_en(slide, prs, badge_text: str, title_en: str):
    """Adds the standard header line, split bottom banner, badge, and Al-Neelain shield logo in English."""
    # 1. Top thin copper line
    top_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0.18), Inches(13.333), Inches(0.04)
    )
    top_line.fill.solid()
    top_line.fill.fore_color.rgb = COPPER_DARK
    top_line.line.color.rgb = COPPER_DARK

    # 2. Bottom footer banner
    footer = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(7.02), Inches(13.333), Inches(0.48)
    )
    footer.fill.solid()
    footer.fill.fore_color.rgb = COPPER_ORANGE
    footer.line.color.rgb = COPPER_DARK

    # Footer Left: English Course text
    tb_fen = slide.shapes.add_textbox(Inches(0.6), Inches(7.04), Inches(5.0), Inches(0.42))
    p_fen = tb_fen.text_frame.paragraphs[0]
    p_fen.text = "Course: COSC726 AGENTIC AI (2026)"
    p_fen.font.size = Pt(11)
    p_fen.font.bold = True
    p_fen.font.color.rgb = TEXT_WHITE
    p_fen.font.name = "Arial"
    p_fen.alignment = PP_ALIGN.LEFT

    # Footer Right: English University text
    tb_far = slide.shapes.add_textbox(Inches(5.5), Inches(7.04), Inches(6.8), Inches(0.42))
    p_far = tb_far.text_frame.paragraphs[0]
    p_far.text = "Al-Neelain University — Faculty of CS & IT (Academic Year: 2026)"
    p_far.font.size = Pt(11)
    p_far.font.bold = True
    p_far.font.color.rgb = TEXT_WHITE
    p_far.font.name = "Arial"
    p_far.alignment = PP_ALIGN.RIGHT

    # 3. Al-Neelain University Coat-of-Arms Shield Logo (bottom right corner over banner)
    if os.path.exists(LOGO_SHIELD):
        slide.shapes.add_picture(
            LOGO_SHIELD,
            Inches(12.4), Inches(6.45), Inches(0.82), Inches(0.88)
        )

    # 4. English Section Badge on Top-Left
    if badge_text:
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE,
            Inches(0.45), Inches(0.35), Inches(3.6), Inches(0.95)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = COPPER_DARK
        badge.line.color.rgb = COPPER_ORANGE
        badge.line.width = Pt(1.5)
        tf_b = badge.text_frame
        tf_b.word_wrap = True
        p_b = tf_b.paragraphs[0]
        p_b.text = badge_text
        p_b.alignment = PP_ALIGN.CENTER
        p_b.font.size = Pt(18)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_WHITE
        p_b.font.name = "Arial"

    # 5. English Main Title on Top-Left (beside badge)
    if title_en:
        tx_box = slide.shapes.add_textbox(
            Inches(4.2), Inches(0.3), Inches(8.6), Inches(1.05)
        )
        tf_t = tx_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_en
        p_t.alignment = PP_ALIGN.LEFT
        p_t.font.size = Pt(23)
        p_t.font.bold = True
        p_t.font.color.rgb = DEEP_SLATE
        p_t.font.name = "Arial"


def add_dark_card_en(slide, left, top, width, height, title="", items=None, title_color=AMBER_YELLOW):
    """Adds a modern dark styled card with bullets in English."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    card.fill.solid()
    card.fill.fore_color.rgb = CARD_BG_DARK
    card.line.color.rgb = CARD_BORDER
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    if title:
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.alignment = PP_ALIGN.LEFT
        p0.font.size = Pt(15.5)
        p0.font.bold = True
        p0.font.color.rgb = title_color
        p0.font.name = "Arial"

    if items:
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.text = f"• {item}"
            p.alignment = PP_ALIGN.LEFT
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_WHITE
            p.font.name = "Arial"
    return card


def add_light_card_en(slide, left, top, width, height, title="", items=None, title_color=CLINICAL_GREEN):
    """Adds a modern light clinical styled card in English."""
    card = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        left, top, width, height
    )
    card.fill.solid()
    card.fill.fore_color.rgb = LIGHT_GRAY_BG
    card.line.color.rgb = BORDER_LIGHT
    card.line.width = Pt(1.5)

    tf = card.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP

    if title:
        p0 = tf.paragraphs[0]
        p0.text = title
        p0.alignment = PP_ALIGN.LEFT
        p0.font.size = Pt(15.5)
        p0.font.bold = True
        p0.font.color.rgb = title_color
        p0.font.name = "Arial"

    if items:
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.text = f"✔ {item}"
            p.alignment = PP_ALIGN.LEFT
            p.font.size = Pt(12)
            p.font.color.rgb = DEEP_SLATE
            p.font.name = "Arial"
    return card


def build_presentation_english():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    print("[BUILDING] 19 Slides Master's Presentation after Adding Intelligent Agents (English)...")

    # =========================================================================
    # SLIDE 1: Official Title Slide
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s1, prs, badge_text="", title_en="")

    # Top right: University Shield Logo
    if os.path.exists(LOGO_SHIELD):
        s1.shapes.add_picture(LOGO_SHIELD, Inches(11.4), Inches(0.4), Inches(1.3), Inches(1.4))

    # Top left: Flag of Sudan
    if os.path.exists(FLAG_SUDAN):
        s1.shapes.add_picture(FLAG_SUDAN, Inches(0.6), Inches(0.4), Inches(1.4), Inches(1.3))

    # University Header Text (Top Center)
    tb_univ = s1.shapes.add_textbox(Inches(2.2), Inches(0.4), Inches(8.933), Inches(1.3))
    tf_u = tb_univ.text_frame
    tf_u.word_wrap = True
    pu1 = tf_u.paragraphs[0]
    pu1.text = "AL-NEELAIN UNIVERSITY"
    pu1.alignment = PP_ALIGN.CENTER
    pu1.font.size = Pt(22)
    pu1.font.bold = True
    pu1.font.color.rgb = DEEP_SLATE

    pu2 = tf_u.add_paragraph()
    pu2.text = "Faculty of Computer Science and Information Technology — Graduate Studies"
    pu2.alignment = PP_ALIGN.CENTER
    pu2.font.size = Pt(16)
    pu2.font.bold = True
    pu2.font.color.rgb = COPPER_DARK

    pu3 = tf_u.add_paragraph()
    pu3.text = "Department of Computer Science | Artificial Intelligence Track"
    pu3.alignment = PP_ALIGN.CENTER
    pu3.font.size = Pt(14)
    pu3.font.color.rgb = DARK_NAVY

    # Course Banner
    tb_course = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.66), Inches(1.85), Inches(6.0), Inches(0.6))
    tb_course.fill.solid()
    tb_course.fill.fore_color.rgb = CLINICAL_GREEN
    tb_course.line.color.rgb = CLINICAL_GREEN
    tf_c = tb_course.text_frame
    pc = tf_c.paragraphs[0]
    pc.text = "Course : COSC726 AGENTIC AI (2026)"
    pc.alignment = PP_ALIGN.CENTER
    pc.font.size = Pt(15)
    pc.font.bold = True
    pc.font.color.rgb = TEXT_WHITE

    # Big Project Title Card
    t_box = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(2.65), Inches(10.933), Inches(2.3))
    t_box.fill.solid()
    t_box.fill.fore_color.rgb = DARK_NAVY
    t_box.line.color.rgb = COPPER_ORANGE
    t_box.line.width = Pt(2.5)
    tf_main = t_box.text_frame
    tf_main.word_wrap = True

    pm1 = tf_main.paragraphs[0]
    pm1.text = "Agentic Multi-Agent AI System for Clinical Diabetes Prediction"
    pm1.alignment = PP_ALIGN.CENTER
    pm1.font.size = Pt(26)
    pm1.font.bold = True
    pm1.font.color.rgb = AMBER_YELLOW

    pm2 = tf_main.add_paragraph()
    pm2.text = "Autonomous Clinical Decision Support System (CDSS)"
    pm2.alignment = PP_ALIGN.CENTER
    pm2.font.size = Pt(18)
    pm2.font.bold = True
    pm2.font.color.rgb = TEXT_WHITE

    pm3 = tf_main.add_paragraph()
    pm3.text = "An integrative architecture combining calibrated modeling, multi-provider LLM reasoning, and safety guardrails"
    pm3.alignment = PP_ALIGN.CENTER
    pm3.font.size = Pt(13)
    pm3.font.color.rgb = TEXT_MUTED

    pm4 = tf_main.add_paragraph()
    pm4.text = "Direct Evolutionary Advancement from the Machine Learning Coursework Project (Previous Semester)"
    pm4.alignment = PP_ALIGN.CENTER
    pm4.font.size = Pt(11.5)
    pm4.font.bold = True
    pm4.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)

    # Researcher & Supervisor Cards
    add_dark_card_en(s1, Inches(1.5), Inches(5.15), Inches(4.8), Inches(1.4),
                     title="Academic Supervision",
                     items=["Dr. Fakhreldeen Saeed (Associate Professor)", "Al-Neelain University — Faculty of CS & IT"],
                     title_color=COPPER_ORANGE)

    add_dark_card_en(s1, Inches(7.03), Inches(5.15), Inches(4.8), Inches(1.4),
                     title="Researcher & Author",
                     items=["Mohamed Abuelgasim (M.Sc. Candidate)", "Master of Computer Science — Artificial Intelligence"],
                     title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 2: Introduction & Clinical Context
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s2, prs, "Introduction", "Clinical Background & Paradigm Shift to Agentic AI")

    img_intro = os.path.join(ASSETS_DIR, "glucose_glucometer.png")
    if os.path.exists(img_intro):
        s2.shapes.add_picture(img_intro, Inches(0.6), Inches(1.8), Inches(5.6), Inches(4.4))

    add_dark_card_en(s2, Inches(6.5), Inches(1.6), Inches(6.3), Inches(2.4),
                     title="Global Diabetes Burden & Clinical Necessity",
                     items=[
                         "Diabetes is one of the most critical and fast-growing chronic conditions worldwide (537+ million adults), causing microvascular, cardiovascular, and renal failures.",
                         "Early and accurate risk stratification represents the cornerstone of preventive healthcare to curb mortality and catastrophic treatment costs.",
                         "Urgent imperative for Clinical Decision Support Systems (CDSS) to assist medical personnel in instant biomarker interpretation."
                     ], title_color=AMBER_YELLOW)

    add_light_card_en(s2, Inches(6.5), Inches(4.2), Inches(6.3), Inches(2.4),
                      title="Academic Lineage & Shift to Autonomous Agents",
                      items=[
                          "Academic Pedigree: Direct evolutionary advancement from the student's coursework project for the Machine Learning course in the previous semester.",
                          "Architectural Shift: Moving away from single, black-box static classifiers to an autonomous Multi-Agent multidisciplinary team.",
                          "Simulating a virtual Multidisciplinary Team (MDT) composed of 5 specialized clinical AI agents with tiered permissions.",
                          "Seamless synthesis of calibrated statistical confidence with generative clinical reasoning aligned with ADA 2024 / 2026 Standards."
                      ], title_color=CLINICAL_GREEN)

    # =========================================================================
    # SLIDE 3: Research Gap & Limitations of Legacy System
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s3, prs, "Research Gap", "Limitations of Conventional & Legacy ML Approaches")

    add_dark_card_en(s3, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="1. Conventional Diagnostic & Software Gaps",
                     items=[
                         "Exclusive reliance on periodic laboratory tests causes critical delays in detecting early-stage asymptomatic diabetes.",
                         "Severe shortage of interactive, field-ready clinical software capable of fully functioning offline in remote clinics.",
                         "Lack of standardized bilingual systems adapted for medical workflows in regional and decentralized health centers.",
                         "Biomarker data accumulation without real-time analysis of non-linear pathophysiological interactions."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s3, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="2. The 4 Fatal Flaws of the Legacy ML Project",
                     items=[
                         "Baseline Nature: The legacy project was an isolated Random Forest classifier developed for the Machine Learning course in the previous semester.",
                         "1. Uncalibrated Statistical Confidence: Hazardous overconfidence in borderline cases (elevated Brier score = 0.231).",
                         "2. Biological Blindness to Zeros: Permitted physiologically impossible zeros in Glucose, Blood Pressure, and accepted pregnancy for male patients.",
                         "3. Silent Black-Box: Outputted naked classification labels without physiological rationales or guideline compliance.",
                         "4. Zero Safety Governance: Complete lack of safety mechanisms to intercept acute medical emergencies (severe hyperglycemia >= 250 mg/dL)."
                     ], title_color=RGBColor(0xEF, 0x44, 0x44))

    # =========================================================================
    # SLIDE 4: Research Objectives
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s4, prs, "Objectives", "Research Objectives & Multi-Agent Transformation")

    add_dark_card_en(s4, Inches(0.6), Inches(1.6), Inches(5.9), Inches(2.4),
                     title="1. Autonomous 5-Agent Architecture",
                     items=[
                         "Design a virtual Multidisciplinary Team (MDT) featuring 5 specialized, decoupled agents.",
                         "Implement rigid Permission Tiers (Tiers 1 to 3) for clinical governance and accountability.",
                         "Tier 1: Triage & Prediction | Tier 2: ADA Reasoning & Diet | Tier 3: Safety Interception Gate.",
                         "Enforce strict Pydantic type-safe contracts across all pipeline data exchanges."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s4, Inches(6.9), Inches(1.6), Inches(5.9), Inches(2.4),
                     title="2. Hybrid Multi-Provider Generative Reasoning",
                     items=[
                         "Integrate 5 leading AI providers: Google Gemini 2.5 Flash, OpenAI, DeepSeek, Ollama, and Cohere.",
                         "Deploy Ollama for local, on-premise execution guaranteeing patient privacy and HIPAA compliance.",
                         "Embed an offline Deterministic ADA Expert Engine ensuring 100% free, zero-latency clinical continuity (ADA 2024 / 2026)."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card_en(s4, Inches(0.6), Inches(4.2), Inches(5.9), Inches(2.4),
                     title="3. Interactive Bilingual Clinical Dashboard",
                     items=[
                         "Develop a modern Streamlit web dashboard supporting Arabic (strict RTL) and English (LTR).",
                         "Implement intelligent biological sex triage with automated pregnancy disabling for male patients.",
                         "Incorporate a dedicated Medical Criteria Analysis tab with 5 dynamic chart topologies and session persistence."
                     ], title_color=EMERALD)

    add_dark_card_en(s4, Inches(6.9), Inches(4.2), Inches(5.9), Inches(2.4),
                     title="4. Clinical Safety Governance & Reporting",
                     items=[
                         "Establish 9 rigorous Clinical Quality Gates certified with 100% compliance.",
                         "Instant emergency crisis interception for glucose >= 250 mg/dL with first-aid resuscitation guidance.",
                         "Automated generation of comprehensive, sanitized patient Word reports (.docx) preserving medical privacy."
                     ], title_color=COPPER_ORANGE)

    # =========================================================================
    # SLIDE 5: Agentic Multi-Agent Architecture
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s5, prs, "Architecture", "Comprehensive Multi-Agent Clinical Architecture (5 MDT Agents)")

    img_arch = os.path.join(ASSETS_DIR, "fig_multi_agent_architecture.png")
    if os.path.exists(img_arch):
        s5.shapes.add_picture(img_arch, Inches(0.6), Inches(1.5), Inches(12.133), Inches(5.15))

    # =========================================================================
    # SLIDE 6: Research Methodology & Preprocessing
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s6, prs, "Methodology", "Research Methodology & Clinical Pipeline Engineering")

    add_dark_card_en(s6, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="1. Dataset & Biological Triage",
                     items=[
                         "Clinically validated Pima Indians Diabetes Dataset (768 records, 8 vital biomarkers).",
                         "Introduced biological sex triage (Male / Female) to enforce physiological gating rules.",
                         "Automated pregnancy input locking for male patients, strictly mapped to N/A (Male).",
                         "Comprehensive physiological range verification to prevent biological anomalies."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s6, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="2. Physiological Zero Imputation",
                     items=[
                         "Identified physiologically impossible zero values in Glucose, Blood Pressure, BMI, Skin Thickness, and Insulin.",
                         "Applied robust median clinical imputation instead of biased, destructive row deletion.",
                         "Encapsulated preprocessing inside a production pipeline transformer (clinical_preprocessor.pkl).",
                         "Strict prevention of data leakage by fitting preprocessing solely on training distributions."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card_en(s6, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                     title="3. Calibration & Data Contracts",
                     items=[
                         "Trained an optimized Random Forest ensemble (100 balanced decision trees).",
                         "Implemented Isotonic Probability Calibration to eliminate overconfident probabilistic estimates.",
                         "Architected formal Pydantic schemas validating all 5 agent inputs and diagnostic dossiers.",
                         "100% compliance verified across all statistical and software quality gates."
                     ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 7: System Comparison (Before vs After)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s7, prs, "Comparison", "Comprehensive Comparison: Before vs. After Intelligent Agents")

    table_shape = s7.shapes.add_table(7, 3, Inches(0.6), Inches(1.65), Inches(12.133), Inches(4.9))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.6)
    tbl.columns[1].width = Inches(4.5)
    tbl.columns[2].width = Inches(5.033)

    headers = ["Clinical & Technical Dimension", "Legacy Baseline (ML Coursework - Previous Semester)", "New Agentic AI System (Multi-Agent CDSS)"]
    for i, h in enumerate(headers):
        cell = tbl.cell(0, i)
        cell.fill.solid()
        cell.fill.fore_color.rgb = DARK_NAVY
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(13)
        p.font.bold = True
        p.font.color.rgb = AMBER_YELLOW if i == 2 else TEXT_WHITE

    rows_data = [
        ("System Architecture", "Single isolated model (Traditional Random Forest) silent", "Collaborative 5-agent pipeline (Virtual MDT Architecture) with permission tiers"),
        ("Clinical Reasoning", "Silent numerical output (Black-Box classification)", "Generative reasoning aligned with ADA 2024 / 2026 clinical guidelines"),
        ("Probability Calibration", "Raw uncalibrated probability prone to overconfidence (Brier: 0.231)", "Isotonically calibrated reliability (Brier: 0.158 | AUC: 0.852 | Acc: 83.12%)"),
        ("Safety & Emergency Triage", "Non-existent (Zero protection for acute crisis)", "Real-time emergency interception (Glucose >= 250 mg/dL) & ambulance protocols"),
        ("Biological Sex Triage", "Unchecked inputs; accepted pregnancy for male patients", "Intelligent biological sex triage; automated pregnancy locking for male patients"),
        ("UI & Clinical Reporting", "Basic desktop GUI (Tkinter) with raw text message", "Bilingual Streamlit web app + 5 dynamic chart topologies + sanitized Word export")
    ]

    for row_idx, rdata in enumerate(rows_data, 1):
        for col_idx, text in enumerate(rdata):
            cell = tbl.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_GRAY_BG if row_idx % 2 == 0 else TEXT_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = text
            p.alignment = PP_ALIGN.CENTER if col_idx == 0 else PP_ALIGN.LEFT
            p.font.size = Pt(11)
            p.font.bold = (col_idx == 0 or col_idx == 2)
            p.font.color.rgb = CLINICAL_GREEN if col_idx == 2 else (COPPER_DARK if col_idx == 0 else DEEP_SLATE)

    # =========================================================================
    # SLIDE 8: Hybrid Multi-Provider LLM Engine
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s8, prs, "LLM Engine", "Hybrid Multi-Provider LLM Engine & Patient Privacy")

    add_dark_card_en(s8, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="1. Advanced Cloud AI Models",
                     items=[
                         "Google Gemini 2.5 Flash: Ultra-fast clinical synthesis and contextual reasoning.",
                         "OpenAI (GPT-4o): Deep medical inference and evidence-based diagnostic explanations.",
                         "DeepSeek: Highly accurate medical logic and clinical reasoning distillation.",
                         "Cohere: Precise biomedical phrasing and concise patient communication.",
                         "All generative explanations rigorously ground assessments in ADA 2024 / 2026 Standards."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s8, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="2. On-Premise Privacy (Ollama)",
                     items=[
                         "Native integration with Ollama for offline local inference without cloud dependencies.",
                         "Runs open-weight models (e.g., LLaMA 3) strictly on internal hospital servers.",
                         "Guarantees 100% data sovereignty and strict compliance with HIPAA & GDPR.",
                         "Zero patient biomarker leakage to external third-party cloud servers."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card_en(s8, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                     title="3. Deterministic ADA Engine",
                     items=[
                         "Built-in deterministic expert engine rooted in international ADA clinical guidelines.",
                         "Functions 100% offline without internet access or costly API subscriptions ($0.00 / 0 tokens).",
                         "Guarantees uninterrupted clinical system availability in mission-critical environments.",
                         "Generates structured, reliable diagnostic rationales, diet regimens, and exercise plans."
                     ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 9: Statistical Findings & Model Calibration
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s9, prs, "Calibration", "Statistical Findings & Probability Calibration Superiority")

    img_cal = os.path.join(ASSETS_DIR, "fig_calibration_roc.png")
    if os.path.exists(img_cal):
        s9.shapes.add_picture(img_cal, Inches(0.6), Inches(1.6), Inches(7.5), Inches(4.9))

    add_dark_card_en(s9, Inches(8.3), Inches(1.6), Inches(4.5), Inches(4.9),
                     title="Certified Statistical Benchmarks",
                     items=[
                         "Calibrated Classification Accuracy: 83.12% on unseen test cohorts.",
                         "Receiver Operating Characteristic Area (ROC-AUC): 0.852.",
                         "Calibrated Brier Score: 0.158 (drastic error drop from 0.231).",
                         "Clinical Sensitivity (Recall): 81.5% | Specificity: 84.0%.",
                         "Overconfidence Elimination: Predicted probabilities match empirical diagnostic rates perfectly.",
                         "Ensures diagnostic safety and reliability for borderline prediabetic cases."
                     ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 10: Medical Criteria & Feature Importance
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s10, prs, "Criteria Analysis", "Medical Criteria Feature Importance & Dynamic Visualization")

    img_feat = os.path.join(ASSETS_DIR, "fig_feature_importance.png")
    if os.path.exists(img_feat):
        s10.shapes.add_picture(img_feat, Inches(0.6), Inches(1.6), Inches(6.8), Inches(4.9))

    add_dark_card_en(s10, Inches(7.6), Inches(1.6), Inches(5.1), Inches(4.9),
                     title="Biomarker Weights & Topologies",
                     items=[
                         "Relative Clinical Biomarker Importance:",
                         "  1. Glucose: 31.5% (Primary risk determinant).",
                         "  2. BMI: 21.0% (Metabolic and adiposity risk factor).",
                         "  3. Age: 14.5% (Progression of insulin resistance).",
                         "  4. Diabetes Pedigree Function: 11.8% (Genetic susceptibility).",
                         "  5. Blood Pressure, Insulin, Skin Thickness, and Pregnancies.",
                         "Medical Criteria Analysis Tab Features:",
                         "  • Unified diagnostic chart styled in clinical dark green (#1a4d2e).",
                         "  • Dynamic shape switcher: Horizontal Bar, Vertical Column, Pie, Donut, and Radar.",
                         "  • Session state persistence (st.session_state) preventing page resets upon re-selection."
                     ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 11: Clinical Quality Gates Matrix, Latency & Economics
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s11, prs, "Quality Gates", "Clinical Quality Gates, Latency & Economic Benchmarks")

    img_qg = os.path.join(ASSETS_DIR, "fig_quality_gates_latency.png")
    if os.path.exists(img_qg):
        s11.shapes.add_picture(img_qg, Inches(0.6), Inches(1.6), Inches(6.8), Inches(4.9))

    add_dark_card_en(s11, Inches(7.6), Inches(1.6), Inches(5.1), Inches(4.9),
                     title="Certification, Latency & Cost Breakdown",
                     items=[
                         "9 Clinical Quality Gates: 100% PASSED (Data integrity, Calibration, Contracts, Emergency).",
                         "Automated Test Suite (Pytest): 10/10 tests passed in 4.24s offline with zero API calls.",
                         "Latency Breakdown (25 Benchmark Runs):",
                         "  • Local Deterministic Engine: Mean 61.83 ms (P95: 84.12 ms) @ 16.17 req/s (< 350 ms clinical SLA).",
                         "  • Cloud (Gemini 2.5 Flash): Mean 1,240 ms (< 3.0 s clinical acceptance threshold).",
                         "Computational & Financial Economics:",
                         "  • Local Deterministic: $0.00 and 0 tokens per 1,000 patients.",
                         "  • Cloud Gemini Flash: 420 in / 280 out tokens, ~$0.00015 per consultation."
                     ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 12: Modern Clinical Web UI
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s12, prs, "Clinical UI", "Modern Clinical Decision Support Dashboard (Streamlit Web UI)")

    img_ui = os.path.join(ASSETS_DIR, "media_1791301217685.png")
    if os.path.exists(img_ui):
        s12.shapes.add_picture(img_ui, Inches(0.6), Inches(1.6), Inches(6.5), Inches(4.9))

    add_dark_card_en(s12, Inches(7.3), Inches(1.6), Inches(5.4), Inches(4.9),
                     title="Advanced Clinical Dashboard Features",
                     items=[
                         "Bilingual Architecture: Native support for Arabic (strict RTL) and English (LTR).",
                         "Biological Sex Triage: Intuitive Male/Female toggle automatically locking pregnancy inputs for males.",
                         "Live Agent Pipeline Trace: Real-time collapsible view displaying reasoning steps of all 5 clinical agents.",
                         "Instant Emergency Interception: High-visibility emergency alert card for glucose >= 250 mg/dL with first-aid triage.",
                         "Clean Clinical Aesthetics: High-contrast responsive design optimized for hospital wards and outpatient clinics."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # SLIDE 13: Patient Diagnostic Report (.docx)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s13, prs, "Patient Report", "Comprehensive Clinical Reporting Engine & Privacy Safeguards")

    add_dark_card_en(s13, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="1. Comprehensive Clinical Report Features",
                     items=[
                         "Instant export of an official clinical summary document (.docx) with a single click.",
                         "Professional styling: Strict RTL for Arabic, standardized LTR for English, with bidiVisual tables.",
                         "Standardized clinical diagnosis phrasing: 'Currently Free of Diabetes' (healthy status).",
                         "Tabulated biomarker review comparing patient values against international clinical reference ranges.",
                         "Tailored nutrition and exercise prescriptions custom-built for the individual patient."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s13, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="2. Patient Privacy & Clinical Decoupling",
                     items=[
                         "Complete Exclusion of Academic Metadata: Researcher and supervisor credentials strictly omitted from patient reports.",
                         "Dedicated clinical artifact: Formatted strictly as an official CDSS medical summary for physician and patient.",
                         "Decoupled architecture: Full separation between academic thesis documentation and patient clinical records.",
                         "Patient identity protection in full compliance with electronic medical record governance."
                     ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 14: System Visual Showcase 1
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s14, prs, "UI Showcase", "Interactive Dashboard & Criteria Analytics Showcase")

    img_s1 = os.path.join(ASSETS_DIR, "media_1791301365012.png")
    img_s2 = os.path.join(ASSETS_DIR, "media_1791301609720.png")

    if os.path.exists(img_s1):
        s14.shapes.add_picture(img_s1, Inches(0.6), Inches(1.6), Inches(5.8), Inches(4.8))
    if os.path.exists(img_s2):
        s14.shapes.add_picture(img_s2, Inches(6.8), Inches(1.6), Inches(5.8), Inches(4.8))

    # =========================================================================
    # SLIDE 15: Quality Gates & Reports Showcase 2
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s15, prs, "Documentation", "Academic Documentation & Quality Assurance Showcase")

    img_s3 = os.path.join(ASSETS_DIR, "media_1791301618976.png")
    if os.path.exists(img_s3):
        s15.shapes.add_picture(img_s3, Inches(0.6), Inches(1.6), Inches(5.8), Inches(4.8))

    add_dark_card_en(s15, Inches(6.8), Inches(1.6), Inches(5.9), Inches(4.8),
                     title="Official Master's Documentation Suite",
                     items=[
                         "Deliverable 3: Comprehensive Academic Health Report in English & Arabic (.docx / .pdf).",
                         "Deliverable 2: 10-Minute Arabic Video Script with shot-by-shot narration and clinical cues.",
                         "Defense Oral Exam Q&A Guide: 10 Model questions and authoritative answers for committee.",
                         "Master Submission Guide: Complete instructions for packing and submission.",
                         "Four Independent Phase Reports in bilingual format with OpenXML table styling.",
                         "Full operations manual and complete system walkthrough documents."
                     ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 16: Scientific Conclusions
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s16, prs, "Conclusions", "Academic & Clinical Conclusions")

    add_dark_card_en(s16, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="1. Multi-Agent AI Superiority",
                     items=[
                         "Demonstrated that distributing clinical tasks across 5 specialized agents drastically enhances diagnostic reliability over black-box models.",
                         "Achieved a remarkable synthesis between statistical precision (83.12% acc, 0.852 AUC) and explainable reasoning (ADA 2024 / 2026).",
                         "Completely eradicated semantic discrepancies between numerical risk probabilities and narrative medical rationales.",
                         "Direct, validated advancement over legacy single-model ML baseline."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s16, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="2. Calibration & Safety Assurance",
                     items=[
                         "Isotonic calibration lowered Brier score to 0.158, providing clinicians with true, uninflated risk probabilities.",
                         "Safety Gatekeeper (Tier 3) guarantees 100% clinical safety by intercepting severe hyperglycemia emergencies.",
                         "Biological sex triage eliminates nonsensical physiological inputs, elevating clinical rigor."
                     ], title_color=EMERALD)

    add_dark_card_en(s16, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                     title="3. Low Latency & High Readiness",
                     items=[
                         "Exceptional pipeline latency of 61.83 ms satisfies real-time clinical deployment criteria.",
                         "Deterministic expert engine enables completely independent offline deployment in rural clinics at $0.00 cost.",
                         "Cloud Gemini integration offers state-of-the-art medical synthesis at minimal expenditure (~$0.00015)."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # SLIDE 17: Recommendations & Future Scope
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s17, prs, "Recommendations", "Strategic Recommendations & Future Research Directions")

    add_dark_card_en(s17, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="1. Hospital EHR & FHIR Integration",
                     items=[
                         "Seamlessly integrate the agentic pipeline with hospital Electronic Health Records via HL7 and FHIR protocols.",
                         "Automate bidirectional biomarker extraction from lab systems and real-time dossier archiving.",
                         "Develop a mobile companion app for community healthcare workers conducting rural screenings."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s17, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                     title="2. Biomarker Expansion & Cohorts",
                     items=[
                         "Expand screening dataset to include pediatric and juvenile cohorts below 21 years of age.",
                         "Incorporate Glycated Hemoglobin (HbA1c) and lipid profiles (Triglycerides, HDL) to enhance precision.",
                         "Study region-specific dietary and genetic risk factors across Sudan and the Middle East."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card_en(s17, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                     title="3. Clinical Trials & Deployment",
                     items=[
                         "Conduct prospective observational clinical trials in primary healthcare centers to evaluate clinician experience.",
                         "Adopt the Lifestyle & Nutrition Agent within national preventive healthcare and diabetes awareness campaigns.",
                         "Continuously update clinical reasoning prompts with each annual release of ADA Standards of Care."
                     ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 18: Closing & Thank You
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s18, prs, "Thank You", "")

    # Big Robot
    if os.path.exists(ROBOT_BIG):
        s18.shapes.add_picture(ROBOT_BIG, Inches(8.6), Inches(1.4), Inches(4.1), Inches(4.9))

    # Floating Small Robot
    if os.path.exists(ROBOT_SMALL):
        s18.shapes.add_picture(ROBOT_SMALL, Inches(6.8), Inches(1.4), Inches(1.6), Inches(1.6))

    # Thank you card
    t_box = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(6.8), Inches(4.6))
    t_box.fill.solid()
    t_box.fill.fore_color.rgb = DARK_NAVY
    t_box.line.color.rgb = COPPER_ORANGE
    t_box.line.width = Pt(2.5)
    tf_th = t_box.text_frame
    tf_th.word_wrap = True

    pth1 = tf_th.paragraphs[0]
    pth1.text = "Thank You for Your Kind Attention"
    pth1.alignment = PP_ALIGN.CENTER
    pth1.font.size = Pt(24)
    pth1.font.bold = True
    pth1.font.color.rgb = AMBER_YELLOW

    pth2 = tf_th.add_paragraph()
    pth2.text = "Q & A Session — Questions & Discussion"
    pth2.alignment = PP_ALIGN.CENTER
    pth2.font.size = Pt(18)
    pth2.font.bold = True
    pth2.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)

    pth3 = tf_th.add_paragraph()
    pth3.text = "\nPrepared by: Mohamed Abuelgasim (M.Sc. Candidate)"
    pth3.alignment = PP_ALIGN.CENTER
    pth3.font.size = Pt(13)
    pth3.font.color.rgb = TEXT_WHITE

    pth4 = tf_th.add_paragraph()
    pth4.text = "Supervised by: Dr. Fakhreldeen Saeed (Associate Professor)"
    pth4.alignment = PP_ALIGN.CENTER
    pth4.font.size = Pt(13)
    pth4.font.color.rgb = TEXT_WHITE

    pth5 = tf_th.add_paragraph()
    pth5.text = "\nAl-Neelain University — Faculty of Computer Science & Information Technology\nAcademic Year: 2026"
    pth5.alignment = PP_ALIGN.CENTER
    pth5.font.size = Pt(12)
    pth5.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 19: Annex & System Architecture Index
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    add_base_decorations_en(s19, prs, "Annex & Index", "Project Deliverables & System Architecture Index")

    add_dark_card_en(s19, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="1. Official Academic Deliverables Suite",
                     items=[
                         "Deliverable 1: GitHub Repository with code, pinned dependencies, .env.example, and 10/10 offline pytest suite.",
                         "Deliverable 2: 10-Minute Video Script & Storyboard in Arabic (COSC726_10_Minute_Video_Script_Arabic.docx/.pdf).",
                         "Deliverable 3: Final Comprehensive Academic Report (COSC726_Project_Team_Health_Report.docx/.pdf).",
                         "Defense Oral Exam Q&A Guide: 10 Model Questions with complete authoritative answers.",
                         "Master Submission Guide: Step-by-step preparation and validation manual.",
                         "Comprehensive PowerPoint Presentations in Arabic and English (.pptx)."
                     ], title_color=AMBER_YELLOW)

    add_dark_card_en(s19, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                     title="2. Source Code & Central CLI Controller",
                     items=[
                         "Central CLI Launcher: run_system.py for unified execution, testing, and report compilation.",
                         "Interactive Streamlit Web Dashboard: app.py (launch via streamlit run app.py).",
                         "Automated 10/10 Test Suite passing offline in 4.24s with zero network dependencies.",
                         "Calibrated Pipeline Models: calibrated_diabetes_model.pkl & clinical_preprocessor.pkl.",
                         "Rigid Pydantic Schemas: schemas.py enforcing type safety across all 5 clinical agents.",
                         "Deterministic ADA expert rules & hybrid multi-provider LLM connectors (Gemini, Ollama, etc.)."
                     ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # Save English presentations across all key target directories
    # =========================================================================
    project_root = os.path.abspath(os.path.join(BASE_DIR, ".."))
    deliverables_dir = os.path.join(project_root, "Project_Deliverables Brief")
    doc_dir = os.path.join(BASE_DIR, "Documentation")

    os.makedirs(doc_dir, exist_ok=True)
    os.makedirs(deliverables_dir, exist_ok=True)

    save_targets = [
        os.path.join(BASE_DIR, "project presention after agents english.pptx"),
        os.path.join(project_root, "project presention after agents english.pptx"),
        os.path.join(deliverables_dir, "project presention after agents english.pptx"),
        os.path.join(doc_dir, "Project_Presentation_Agentic_AI_English.pptx")
    ]

    for p in save_targets:
        prs.save(p)
        print(f"[SUCCESS] Saved English presentation: {p} ({os.path.getsize(p)/1024:.1f} KB)")


if __name__ == "__main__":
    build_presentation_english()

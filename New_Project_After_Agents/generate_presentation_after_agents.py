"""
Script: generate_presentation_after_agents.py
Generates the comprehensive academic Master's presentation in Arabic (.pptx)
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


def set_rtl(paragraph):
    """Sets paragraph bidirectional RTL attribute in PowerPoint OpenXML."""
    pPr = paragraph._p.get_or_add_pPr()
    pPr.set('rtl', '1')


def add_base_decorations(slide, prs, badge_text_en: str, title_ar: str):
    """Adds the standard header line, split bottom banner, badge, and Al-Neelain shield logo."""
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
    tb_fen = slide.shapes.add_textbox(Inches(0.6), Inches(7.04), Inches(4.5), Inches(0.42))
    p_fen = tb_fen.text_frame.paragraphs[0]
    p_fen.text = "Course: COSC726 AGENTIC AI"
    p_fen.font.size = Pt(11)
    p_fen.font.bold = True
    p_fen.font.color.rgb = TEXT_WHITE
    p_fen.font.name = "Arial"
    p_fen.alignment = PP_ALIGN.LEFT

    # Footer Right: Arabic University text
    tb_far = slide.shapes.add_textbox(Inches(5.0), Inches(7.04), Inches(7.2), Inches(0.42))
    p_far = tb_far.text_frame.paragraphs[0]
    p_far.text = "جامعة النيلين - كلية علوم الحاسوب وتقانة المعلومات (2026 م)"
    p_far.font.size = Pt(11)
    p_far.font.bold = True
    p_far.font.color.rgb = TEXT_WHITE
    p_far.font.name = "Arial"
    p_far.alignment = PP_ALIGN.RIGHT
    set_rtl(p_far)

    # 3. Al-Neelain University Coat-of-Arms Shield Logo (bottom right corner over banner)
    if os.path.exists(LOGO_SHIELD):
        slide.shapes.add_picture(
            LOGO_SHIELD,
            Inches(12.4), Inches(6.45), Inches(0.82), Inches(0.88)
        )

    # 4. English Section Badge on Top-Left
    if badge_text_en:
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
        p_b.text = badge_text_en
        p_b.alignment = PP_ALIGN.CENTER
        p_b.font.size = Pt(18)
        p_b.font.bold = True
        p_b.font.color.rgb = TEXT_WHITE
        p_b.font.name = "Arial"

    # 5. Arabic Main Title on Top-Right
    if title_ar:
        tx_box = slide.shapes.add_textbox(
            Inches(4.2), Inches(0.3), Inches(8.6), Inches(1.05)
        )
        tf_t = tx_box.text_frame
        tf_t.word_wrap = True
        p_t = tf_t.paragraphs[0]
        p_t.text = title_ar
        p_t.alignment = PP_ALIGN.RIGHT
        p_t.font.size = Pt(25)
        p_t.font.bold = True
        p_t.font.color.rgb = DEEP_SLATE
        p_t.font.name = "Arial"
        set_rtl(p_t)


def add_dark_card(slide, left, top, width, height, title="", items=None, title_color=AMBER_YELLOW):
    """Adds a modern dark styled card with bullets."""
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
        p0.alignment = PP_ALIGN.RIGHT
        p0.font.size = Pt(15.5)
        p0.font.bold = True
        p0.font.color.rgb = title_color
        p0.font.name = "Arial"
        set_rtl(p0)

    if items:
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.text = f"• {item}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(12)
            p.font.color.rgb = TEXT_WHITE
            p.font.name = "Arial"
            set_rtl(p)
    return card


def add_light_card(slide, left, top, width, height, title="", items=None, title_color=CLINICAL_GREEN):
    """Adds a modern light clinical styled card."""
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
        p0.alignment = PP_ALIGN.RIGHT
        p0.font.size = Pt(15.5)
        p0.font.bold = True
        p0.font.color.rgb = title_color
        p0.font.name = "Arial"
        set_rtl(p0)

    if items:
        for idx, item in enumerate(items):
            p = tf.add_paragraph() if (title or idx > 0) else tf.paragraphs[0]
            p.text = f"✔ {item}"
            p.alignment = PP_ALIGN.RIGHT
            p.font.size = Pt(12)
            p.font.color.rgb = DEEP_SLATE
            p.font.name = "Arial"
            set_rtl(p)
    return card


def build_presentation():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    print("[BUILDING] 19 Slides Master's Presentation after Adding Intelligent Agents (Arabic)...")

    # =========================================================================
    # SLIDE 1: Title Slide (الغلاف الرسمي)
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s1, prs, badge_text_en="", title_ar="")

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
    pu1.text = "جامعة النيلين"
    pu1.alignment = PP_ALIGN.CENTER
    pu1.font.size = Pt(22)
    pu1.font.bold = True
    pu1.font.color.rgb = DEEP_SLATE
    set_rtl(pu1)

    pu2 = tf_u.add_paragraph()
    pu2.text = "كلية علوم الحاسوب وتقانة المعلومات - كلية الدراسات العليا"
    pu2.alignment = PP_ALIGN.CENTER
    pu2.font.size = Pt(17)
    pu2.font.bold = True
    pu2.font.color.rgb = COPPER_DARK
    set_rtl(pu2)

    pu3 = tf_u.add_paragraph()
    pu3.text = "قسم علوم الحاسوب | مسار الذكاء الاصطناعي (AI Track)"
    pu3.alignment = PP_ALIGN.CENTER
    pu3.font.size = Pt(14)
    pu3.font.color.rgb = DARK_NAVY
    set_rtl(pu3)

    # Course Banner
    tb_course = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.66), Inches(1.85), Inches(6.0), Inches(0.6))
    tb_course.fill.solid()
    tb_course.fill.fore_color.rgb = CLINICAL_GREEN
    tb_course.line.color.rgb = CLINICAL_GREEN
    tf_c = tb_course.text_frame
    pc = tf_c.paragraphs[0]
    pc.text = "مادة ( Course ) : COSC726 AGENTIC AI (2026)"
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
    pm1.text = "نظام التشخيص التنبئي لمرض السكري بالوكلاء المتعددين"
    pm1.alignment = PP_ALIGN.CENTER
    pm1.font.size = Pt(28)
    pm1.font.bold = True
    pm1.font.color.rgb = AMBER_YELLOW
    set_rtl(pm1)

    pm2 = tf_main.add_paragraph()
    pm2.text = "Agentic Multi-Agent AI System for Clinical Decision Support"
    pm2.alignment = PP_ALIGN.CENTER
    pm2.font.size = Pt(19)
    pm2.font.bold = True
    pm2.font.color.rgb = TEXT_WHITE

    pm3 = tf_main.add_paragraph()
    pm3.text = "هندسة تشخيصية تكاملية تجمع بين النمذجة المعايرة والتعليل السريري التوليدي وبوابات الأمان الصارمة"
    pm3.alignment = PP_ALIGN.CENTER
    pm3.font.size = Pt(13)
    pm3.font.color.rgb = TEXT_MUTED
    set_rtl(pm3)

    pm4 = tf_main.add_paragraph()
    pm4.text = "تطور نوعي مباشر لمشروع مادة تعلم الآلة (Machine Learning) المنجز في الفصل الدراسي السابق"
    pm4.alignment = PP_ALIGN.CENTER
    pm4.font.size = Pt(11.5)
    pm4.font.bold = True
    pm4.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)
    set_rtl(pm4)

    # Researcher & Supervisor Cards
    add_dark_card(s1, Inches(1.5), Inches(5.15), Inches(4.8), Inches(1.4),
                  title="إشراف الأستاذ المشارك",
                  items=["د. فخرالدين سعيد (Dr. Fakhreldeen Saeed)", "جامعة النيلين - Al-Neelain University"],
                  title_color=COPPER_ORANGE)

    add_dark_card(s1, Inches(7.03), Inches(5.15), Inches(4.8), Inches(1.4),
                  title="إعداد الباحث",
                  items=["محمد ابوالقاسم (Mohammed Abuelgasim)", "ماجستير علوم الحاسوب - الذكاء الاصطناعي"],
                  title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 2: Introduction & Clinical Context (المقدمة والسياق السريري للمشروع)
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s2, prs, "Introduction", "المقدمة والسياق السريري للمشروع")

    img_intro = os.path.join(ASSETS_DIR, "media_1791301217685.png")
    if os.path.exists(img_intro):
        s2.shapes.add_picture(img_intro, Inches(0.6), Inches(1.6), Inches(5.6), Inches(4.9))

    add_dark_card(s2, Inches(6.5), Inches(1.6), Inches(6.3), Inches(2.4),
                  title="تحديات مرض السكري والرعاية الصحية",
                  items=[
                      "يعد السكري من أسرع الأمراض المزمنة انتشاراً (أكثر من 537 مليون مصاب عالمياً)، ويؤدي لمضاعفات كلوية وقلبية وبصرية خطيرة.",
                      "التشخيص المبكر والدقيق يمثل حجر الزاوية في الرعاية الوقائية لتقليل الوفيات وتكاليف العلاج الباهظة.",
                      "الحاجة الملحة لأنظمة دعم القرار السريري الذكية (CDSS) لمساندة الأطباء في التحليل الفوري للمؤشرات الحيوية."
                  ], title_color=AMBER_YELLOW)

    add_light_card(s2, Inches(6.5), Inches(4.2), Inches(6.3), Inches(2.4),
                   title="السياق الأكاديمي والنقلة النوعية للوكلاء (Agentic AI)",
                   items=[
                       "الأصل الأكاديمي: المشروع يعد تطويراً وتحويلاً شاملاً لمشروع مادة تعلم الآلة (Machine Learning) المنجز في الفصل الدراسي السابق.",
                       "التحول المعماري: الانتقال من النموذج الإحصائي الفردي الصامت إلى معمارية وكلاء متفاعلة ذاتية الحكم.",
                       "محاكاة فريق الاستشارة الطبية متعدد التخصصات (Virtual MDT) المكون من 5 وكلاء متخصصين بصلاحيات متدرجة.",
                       "دمج المعايرة الإحصائية الدقيقة مع التعليل الطبي التوليدي المستند لمعايير ADA 2024 / 2026."
                   ], title_color=CLINICAL_GREEN)

    # =========================================================================
    # SLIDE 3: Research Gap (فجوة البحث ومحدودية الأنظمة السابقة)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s3, prs, "Research Gap", "فجوة البحث ومحدودية الأنظمة السابقة")

    add_dark_card(s3, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="أولاً: قصور التشخيص الطبي والتقني الميداني",
                  items=[
                      "الاعتماد الحصري على الفحوصات المختبرية المنعزلة، مما يؤخر الكشف عن المرض في مراحله المبكرة دون أعراض.",
                      "نقص حاد في البرمجيات التفاعلية السريرية التي تدعم العمل الميداني المستقل تماماً دون إنترنت في المراكز الصحية النائية.",
                      "غياب الأدوات الطبية المعربة ذات التنسيق الصحيح (RTL) التي تلائم بيئات العمل في المراكز الصحية الإقليمية والمحلية.",
                      "تراكم البيانات الحيوية دون تحليل فوري للعلاقات الفسيولوجية المعقدة وغير الخطية بين المؤشرات."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s3, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="ثانياً: العيوب الأربعة القاتلة لمشروع تعلم الآلة السابق",
                  items=[
                      "طبيعة الأساس السابق: كان مجرد مصنف فردي لتعلم الآلة (Random Forest معزول) من الفصل السابق.",
                      "1. انعدام المعايرة الاحتمالية: الوقوع في فخ الثقة المفرطة الزائفة في الحالات الحدية (Brier Score = 0.231).",
                      "2. العمى البيولوجي للأصفار: قبول أصفار مستحيلة بيولوجياً في الجلوكوز والضغط وقبول الحمل لمرضى ذكور.",
                      "3. غياب التعليل السريري (الصندوق الأسود): تقديم نسب رقمية صامتة دون أي تفسير فسيولوجي أو بروتوكول علاجي.",
                      "4. انعدام الأمان السريري: عدم وجود أي بوابة حوكمة لاعتراض طوارئ السكر الحاد الخطرة (>= 250 mg/dL)."
                  ], title_color=RGBColor(0xEF, 0x44, 0x44))

    # =========================================================================
    # SLIDE 4: Research Objectives (أهداف البحث والتحول المعماري الشامل)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s4, prs, "Objectives", "أهداف البحث والتحول المعماري الشامل")

    add_dark_card(s4, Inches(6.9), Inches(1.6), Inches(5.9), Inches(2.4),
                  title="أولاً: بناء معمارية وكلاء أذكياء متعددين",
                  items=[
                      "تصميم فريق طبي استشاري افتراضي يضم 5 وكلاء متخصصين ومستقلين.",
                      "توزيع مستويات الصلاحيات السريرية الصارمة (Permission Tiers 1-3).",
                      "المستوى 1: الفرز والتنبؤ | المستوى 2: التعليل ونمط الحياة | المستوى 3: بوابات الأمان والحوكمة.",
                      "فرض عقود بيانات صارمة باستخدام مكتبة Pydantic لضمان تكامل خط الأنابيب ومنع الهلوسة."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s4, Inches(0.6), Inches(1.6), Inches(5.9), Inches(2.4),
                  title="ثانياً: هندسة تعليل سريري توليدي هجين",
                  items=[
                      "ربط التشخيص مع 5 مزودين للذكاء الاصطناعي: Ollama, Gemini, OpenAI, DeepSeek, Cohere.",
                      "دعم التشغيل المحلي عبر Ollama لحماية خصوصية بيانات المرضى داخل المستشفيات (امتثال HIPAA/GDPR).",
                      "بناء محرك حتمي خبير مستقل للعمل دون إنترنت وبشكل مجاني 100% وفق معايير ADA 2024 / 2026."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card(s4, Inches(6.9), Inches(4.2), Inches(5.9), Inches(2.4),
                  title="ثالثاً: تطوير واجهة سريرية تفاعلية ثنائية اللغة",
                  items=[
                      "واجهة ويب Streamlit عصرية تدعم العربية (RTL كامل أصيل) والإنجليزية (LTR).",
                      "فرز الجنس البيولوجي الذكي مع تعطيل وقفل آلي لحقل الحمل للمرضى الذكور لضمان المنطق الفسيولوجي.",
                      "تبويب تفاعلي لتحليل المعايير الطبية مع 5 أشكال بيانية ديناميكية واستقرار تام لحالة الجلسة."
                  ], title_color=EMERALD)

    add_dark_card(s4, Inches(0.6), Inches(4.2), Inches(5.9), Inches(2.4),
                  title="رابعاً: حوكمة الأمان وبوابات الجودة والتقارير الطبية",
                  items=[
                      "تأسيس 9 بوابات جودة سريرية مجازة بالكامل بنسبة 100%، وسرعة استجابة فائقة (61.83 مللي ثانية).",
                      "اعتراض فوري لحالات الطوارئ الحرجة (سكر الدم >= 250 mg/dL) وتفعيل بروتوكول استدعاء الإسعاف.",
                      "توليد تقارير طبية رسمية للمرضى بصيغة Word خالية تماماً من البيانات الأكاديمية لحماية الخصوصية."
                  ], title_color=COPPER_ORANGE)

    # =========================================================================
    # SLIDE 5: Agentic Multi-Agent Architecture (معمارية الوكلاء الأذكياء المتعددين)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s5, prs, "Architecture", "معمارية الوكلاء الأذكياء المتعددين (5 MDT Agents)")

    img_arch = os.path.join(ASSETS_DIR, "fig_multi_agent_architecture.png")
    if os.path.exists(img_arch):
        s5.shapes.add_picture(img_arch, Inches(0.6), Inches(1.5), Inches(12.133), Inches(5.15))

    # =========================================================================
    # SLIDE 6: Research Methodology & Preprocessing (منهجية البحث ومعالجة البيانات)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s6, prs, "Methodology", "منهجية البحث ومعالجة البيانات السريرية")

    add_dark_card(s6, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                  title="أولاً: قاعدة البيانات والفرز البيولوجي",
                  items=[
                      "قاعدة بيانات Pima Indians المعتمدة طبياً (768 سجلاً، 8 معايير سريرية حيوية).",
                      "إضافة متغير الجنس البيولوجي (ذكر/أنثى) لضبط القيود الفسيولوجية الصارمة.",
                      "قفل وتعطيل حقل مرات الحمل تلقائياً للمرضى الذكور وتعيينه كـ 'غير متاح (ذكر)'.",
                      "فحص القيم المتطرفة وضمان سلامة النطاقات الحيوية سريرياً ومنع الشذوذ الإحصائي."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s6, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثانياً: المعالجة الفسيولوجية للأصفار",
                  items=[
                      "اكتشاف الأصفار المستحيلة بيولوجياً في الجلوكوز، الضغط، كتلة الجسم، وسماكة الجلد والأنسولين.",
                      "تطبيق المعالجة بالوسيط الحسابي الفسيولوجي الموجه بدلاً من الحذف العشوائي المشوه للبيانات.",
                      "تضمين محول المعالجة في خط أنابيب سريري موحد (clinical_preprocessor.pkl).",
                      "منع تسرب البيانات الصارم (Data Leakage) عبر قصر تدريب المعالج على بيانات التدريب فقط."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card(s6, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثالثاً: المعايرة الإحصائية وعقود البيانات",
                  items=[
                      "تدريب مصنف الغابات العشوائية المحسن (100 شجرة قرار، وعمق متوازن للتحكم بالتعقيد).",
                      "تطبيق المعايرة الإيزوتونية (Isotonic Regression) لضبط الاحتماليات الإحصائية وحذف الثقة المفرطة.",
                      "بناء عقود Pydantic صارمة لجميع مدخلات ومخرجات الوكلاء الخمسة لضمان تكامل التدفق البرمجي.",
                      "اجتياز بوابات الجودة الإحصائية والبرمجية بنسبة 100% واكتمال التحقق."
                  ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 7: System Comparison (Before vs After)
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s7, prs, "Comparison", "مقارنة شاملة: النظام قبل وبعد إضافة الوكلاء الأذكياء")

    table_shape = s7.shapes.add_table(7, 3, Inches(0.6), Inches(1.65), Inches(12.133), Inches(4.9))
    tbl = table_shape.table
    tbl.columns[0].width = Inches(2.7)
    tbl.columns[1].width = Inches(4.6)
    tbl.columns[2].width = Inches(4.833)

    headers = ["المعيار السريري والتقني", "النظام السابق (مشروع مادة تعلم الآلة - الفصل السابق)", "النظام الجديد بعد الوكلاء الأذكياء (Agentic AI)"]
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
        set_rtl(p)

    rows_data = [
        ("معمارية النظام", "نموذج فردي تقليدي معزول (Single Random Forest) صامت", "معمارية خماسية متعددة الوكلاء (5 Specialized MDT Agents) بصلاحيات متدرجة"),
        ("التعليل السريري المفسر", "تصنيف صامت وأرقام مجردة دون تعليل (Black-Box ML)", "تعليل سريري توليدي هجين وفق معايير ADA 2024 / 2026 الدولية"),
        ("معايرة الاحتمالية", "احتمالية خام غير معايرة تعاني من الثقة المفرطة (Brier: 0.231)", "معايرة إيزوتونية موثوقة (Brier: 0.158 | AUC: 0.852 | Acc: 83.12%)"),
        ("الأمان واعتراض الطوارئ", "منعدم تماماً (لا توجد أي حماية لحالات الخطر الحرج)", "اعتراض لحظي للطوارئ (سكر >= 250 mg/dL) وتفعيل بروتوكول الإسعاف الفوري"),
        ("الفرز البيولوجي للجنس", "إدخال عام دون تمييز، وقبول الحمل لمرضى ذكور", "فرز جنسي ذكي مع قفل وتعطيل آلي لحقل الحمل للمرضى الذكور لمنع التناقض"),
        ("الواجهة وتصدير التقارير", "تطبيق سطح مكتب قديم (Tkinter) بنتيجة نصية مقتضبة", "منصة ويب سحابية ثنائية اللغة + 5 أشكال بيانية ديناميكية + تقارير Word مطهرة")
    ]

    for row_idx, rdata in enumerate(rows_data, 1):
        for col_idx, text in enumerate(rdata):
            cell = tbl.cell(row_idx, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_GRAY_BG if row_idx % 2 == 0 else TEXT_WHITE
            p = cell.text_frame.paragraphs[0]
            p.text = text
            p.alignment = PP_ALIGN.CENTER if col_idx == 0 else PP_ALIGN.RIGHT
            p.font.size = Pt(11)
            p.font.bold = (col_idx == 0 or col_idx == 2)
            p.font.color.rgb = CLINICAL_GREEN if col_idx == 2 else (COPPER_DARK if col_idx == 0 else DEEP_SLATE)
            set_rtl(p)

    # =========================================================================
    # SLIDE 8: Hybrid Multi-Provider LLM Engine (المحرك الهجين والخصوصية)
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s8, prs, "LLM Engine", "محرك الذكاء الاصطناعي التوليدي الهجين والخصوصية")

    add_dark_card(s8, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                  title="أولاً: نماذج السحابة المتقدمة",
                  items=[
                      "Google Gemini 2.5 Flash: استدعاء سحابي فائق السرعة واستدلال سياقي دقيق.",
                      "OpenAI (GPT-4o): تعليل طبي متقدم وتحليل عميق للحالات المعقدة.",
                      "DeepSeek: استدلال منطقي وطبي عالي الدقة.",
                      "Cohere: صياغة سريرية وتلخيص لغوي دقيق.",
                      "توليد تعليلات تشخيصية مبنية على أحدث معايير الجمعية الأمريكية للسكري ADA 2024 / 2026."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s8, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثانياً: التشغيل المحلي وحماية الخصوصية",
                  items=[
                      "تكامل كامل مع محرك Ollama للتشغيل المحلي التام دون اتصال بالسحابة.",
                      "تشغيل نماذج مثل LLaMA 3 محلياً داخل خوادم المستشفى دون إرسال بيانات للخارج.",
                      "الامتثال لمعايير الخصوصية الطبية العالمية HIPAA و GDPR.",
                      "حماية سرية البيانات الحيوية للمرضى بنسبة 100% ومنع تسرب السجلات."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card(s8, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثالثاً: المحرك السريري الحتمي المستقل",
                  items=[
                      "محرك خبير حتمي مدمج مستند لقواعد ADA الدولية الصارمة.",
                      "يعمل محلياً وأوفلاين دون إنترنت وبشكل مجاني 100% وتكلفة 0.00$ وصفر توكنز.",
                      "يضمن استمرارية عمل النظام في غياب الشبكة أو انقطاع الخدمات السحابية في الأرياف.",
                      "توليد تشخيص دقيق وتوصيات علاجية وخطط تغذية ولياقة موثوقة معيارياً."
                  ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 9: Statistical Findings & Model Calibration (النتائج والمعايرة)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s9, prs, "Calibration", "النتائج الإحصائية وتفوق المعايرة الاحتمالية")

    img_cal = os.path.join(ASSETS_DIR, "fig_calibration_roc.png")
    if os.path.exists(img_cal):
        s9.shapes.add_picture(img_cal, Inches(0.6), Inches(1.6), Inches(7.5), Inches(4.9))

    add_dark_card(s9, Inches(8.3), Inches(1.6), Inches(4.5), Inches(4.9),
                  title="المؤشرات الإحصائية المعتمدة",
                  items=[
                      "دقة النموذج المعاير: 83.12% (Calibrated Accuracy).",
                      "المساحة تحت منحنى ROC تبلغ: 0.852 (ROC-AUC).",
                      "مقياس برير المعاير: 0.158 (انخفاض هائل ومثالي من 0.231).",
                      "الحساسية السريرية (Sensitivity): 81.5%.",
                      "النوعية السريرية (Specificity): 84.0%.",
                      "منع الثقة المفرطة: تطابق الاحتمال التنبئي مع التكرار الإكلينيكي الحقيقي للإصابة.",
                      "تأمين الحالات الحرجة وتفادي التشخيص الخاطئ في الفئات الحدية المقاربة للإصابة."
                  ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 10: Medical Criteria & Feature Importance (تحليل المعايير)
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s10, prs, "Criteria Analysis", "تحليل المعايير الطبية والتصور البياني الديناميكي")

    img_feat = os.path.join(ASSETS_DIR, "fig_feature_importance.png")
    if os.path.exists(img_feat):
        s10.shapes.add_picture(img_feat, Inches(0.6), Inches(1.6), Inches(6.8), Inches(4.9))

    add_dark_card(s10, Inches(7.6), Inches(1.6), Inches(5.1), Inches(4.9),
                  title="الأهمية السريرية والتبديل الديناميكي",
                  items=[
                      "الأوزان النسبية للمؤشرات الحيوية:",
                      "  1. الجلوكوز (Glucose): 31.5% (المحدد الأكبر لخطورة الإصابة).",
                      "  2. مؤشر كتلة الجسم (BMI): 21.0% (عامل الخطر الأيضي الأول).",
                      "  3. العمر (Age): 14.5% (تزايد مقاومة الأنسولين الفسيولوجية).",
                      "  4. وظيفة السلالة الوراثية: 11.8% (التأثير الجيني العائلي).",
                      "  5. ضغط الدم، الأنسولين، سماكة الجلد، والحمل.",
                      "تبويب تحليل المعايير في الواجهة:",
                      "  • مخطط بياني موحد بلون أخضر سريري (#1a4d2e).",
                      "  • إمكانية التبديل بين 5 أشكال: أفقي، عمودي، دائري، دونات، وراداري.",
                      "  • ثبات الجلسة (Session State) ومنع إعادة تحميل الصفحة عند التبديل."
                  ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 11: Clinical Quality Gates & Latency Benchmarks (بوابات الجودة والسرعة)
    # =========================================================================
    s11 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s11, prs, "Quality Gates", "منظومة بوابات الجودة والتحقق السريري وتحليل الأداء")

    img_qg = os.path.join(ASSETS_DIR, "fig_quality_gates_latency.png")
    if os.path.exists(img_qg):
        s11.shapes.add_picture(img_qg, Inches(0.6), Inches(1.6), Inches(6.8), Inches(4.9))

    add_dark_card(s11, Inches(7.6), Inches(1.6), Inches(5.1), Inches(4.9),
                  title="شهادة الجودة وتحليل الأداء والتكلفة",
                  items=[
                      "بوابات الجودة السريرية (9 Quality Gates): مجازة 100% بنجاح.",
                      "الاختبارات الآلية (Automated Pytest): اجتياز 10 من أصل 10 اختبارات في 4.24 ثانية أوفلاين تماماً.",
                      "تفكيك أزمنة الاستجابة (Latency Breakdown):",
                      "  • المحرك الحتمي المحلي: متوسط 61.83 ms (الشريحة 95: 84.12 ms) بمعدل 16.17 طلب/ثانية (تحت عتبة القبول < 350 ms).",
                      "  • السحابة (Gemini 2.5 Flash): متوسط 1,240 ms (تحت عتبة القبول < 3.0 s).",
                      "التكلفة الحسابية والمالية:",
                      "  • المحلي الحتمي: 0.00$ وصفر توكنز لكل 1,000 مريض.",
                      "  • السحابي (Gemini Flash): 420 توكن إدخال و 280 توكن إخراج، بتكلفة ~0.00015$ للاستشارة."
                  ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 12: Modern Clinical Web UI (الواجهة السريرية التفاعلية)
    # =========================================================================
    s12 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s12, prs, "Clinical UI", "الواجهة السريرية التفاعلية الحديثة (Streamlit Dashboard)")

    img_ui = os.path.join(ASSETS_DIR, "media_1791301217685.png")
    if os.path.exists(img_ui):
        s12.shapes.add_picture(img_ui, Inches(0.6), Inches(1.6), Inches(6.5), Inches(4.9))

    add_dark_card(s12, Inches(7.3), Inches(1.6), Inches(5.4), Inches(4.9),
                  title="خصائص ومميزات الواجهة المطورة",
                  items=[
                      "تعريب كامل ودعم أصيل للاتجاه من اليمين لليسار بالعربية (RTL)، واليسار لليمين بالإنجليزية (LTR).",
                      "فرز الجنس البيولوجي (ذكر/أنثى) مع تعطيل وقفل آلي لحقل الحمل للمرضى الذكور وتحديدها كـ 'غير متاح (ذكر)'.",
                      "متتبع حي لمسار تفكير الوكلاء مع إمكانية طي وعرض كل وكيل وتفاصيل استنتاجاته السريرية بدقة.",
                      "اعتراض لحظي لحالات الطوارئ مع بطاقة تحذير حمراء وإرشادات إسعاف فوري عاجل للسكري >= 250 mg/dL.",
                      "تصميم سريري احترافي ذو تباين بصري عالٍ ملائم للمستشفيات والعيادات الخارجية."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # SLIDE 13: Patient Diagnostic Report (.docx) (التقارير الطبية والخصوصية)
    # =========================================================================
    s13 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s13, prs, "Patient Report", "نظام التقارير الطبية الشاملة وحماية الخصوصية")

    add_dark_card(s13, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="أولاً: مواصفات التقرير السريري الشامل",
                  items=[
                      "تصدير فوري لتقرير طبي شامل بصيغة Word (.docx) بضغطة زر واحدة.",
                      "تنسيق RTL صارم باللغة العربية: اتجاه الفقرات، العناوين، والجداول (w:bidiVisual).",
                      "اعتماد صياغة التشخيص السليم: 'الحالة سليمة من السكري حالياً' بدقة ووضوح.",
                      "جدول المؤشرات الحيوية مع المديات الطبيعية المعتمدة دولياً والمقارنة السريرية.",
                      "خطة التغذية والتمارين الرياضية المفصلة والمخصصة لاحتياجات المريض الفردية."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s13, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="ثانياً: خصوصية المريض والنزاهة السريرية",
                  items=[
                      "حذف الترويسة الأكاديمية تماماً: حذف بيانات الباحث والمشرف الأكاديمية من تقرير المريض عند تحميله.",
                      "تحويل التقرير إلى وثيقة طبية رسمية خالصة صادرة عن نظام دعم القرار السريري (CDSS).",
                      "فصل تام بين التقارير الأكاديمية للمشروع وتقارير المرضى السريرية الصادرة للمستشفيات.",
                      "حماية خصوصية وهوية المريض ومطابقة معايير السجلات الطبية الإلكترونية العالمية."
                  ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 14: System Visual Showcase 1 (معرض شاشات النظام)
    # =========================================================================
    s14 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s14, prs, "UI Showcase", "معرض شاشات النظام السريري وتبويب تحليل المعايير")

    img_s1 = os.path.join(ASSETS_DIR, "media_1791301365012.png")
    img_s2 = os.path.join(ASSETS_DIR, "media_1791301609720.png")

    if os.path.exists(img_s1):
        s14.shapes.add_picture(img_s1, Inches(0.6), Inches(1.6), Inches(5.8), Inches(4.8))
    if os.path.exists(img_s2):
        s14.shapes.add_picture(img_s2, Inches(6.8), Inches(1.6), Inches(5.8), Inches(4.8))

    # =========================================================================
    # SLIDE 15: Quality Gates & Reports Showcase 2 (معرض التقارير الأكاديمية)
    # =========================================================================
    s15 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s15, prs, "Documentation", "معرض التقارير الأكاديمية وبوابات الجودة المولدة")

    img_s3 = os.path.join(ASSETS_DIR, "media_1791301618976.png")
    if os.path.exists(img_s3):
        s15.shapes.add_picture(img_s3, Inches(0.6), Inches(1.6), Inches(5.8), Inches(4.8))

    add_dark_card(s15, Inches(6.8), Inches(1.6), Inches(5.9), Inches(4.8),
                  title="المستندات والتقارير الرسمية المولدة للمشروع",
                  items=[
                      "تقرير المشروع الأكاديمي الشامل (المخرج 3) باللغتين العربية والإنجليزية (.docx / .pdf).",
                      "سيناريو العرض المرئي الكامل (المخرج 2) باللغة العربية مع المشاهد وتوجيهات الإلقاء.",
                      "دليل أسئلة الدفاع العشرة النموذجية مع الإجابات المحكمة لأسئلة المناقشين والمشرف.",
                      "دليل التحضير للتسليم النهائي المنظم والشامل وفق متطلبات مادة COSC726.",
                      "تقارير المراحل الأربع المستقلة ثنائية اللغة وأدلة التشغيل ومخطط الإنجاز.",
                      "تطبيق اتجاه الكتابة والجداول (RTL) في كافة المستندات العربية المعتمدة."
                  ], title_color=AMBER_YELLOW)

    # =========================================================================
    # SLIDE 16: Scientific Conclusions (الاستنتاجات العلمية)
    # =========================================================================
    s16 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s16, prs, "Conclusions", "الاستنتاجات الأكاديمية والسريرية للمشروع")

    add_dark_card(s16, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                  title="أولاً: تفوق معمارية الوكلاء الذكية",
                  items=[
                      "أثبتت النتائج أن توزيع المهام السريرية على 5 وكلاء متخصصين يرفع موثوقية التشخيص ويزيل قيود النماذج الصامتة.",
                      "تحقيق توازن فريد بين الدقة الإحصائية (83.12%) والتعليل المفسر وفق معايير ADA 2024 / 2026.",
                      "القضاء التام على أخطاء عدم الاتساق الدلالي بين النسب الرقمية والشروح الطبية التوليدية.",
                      "تطور نوعي راسخ لمشروع تعلم الآلة السابق."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s16, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثانياً: موثوقية المعايرة وحوكمة الأمان",
                  items=[
                      "نجحت المعايرة الإيزوتونية في خفض مقياس برير إلى 0.158، مما يمنح الأطباء احتمالية حقيقية غير مضللة.",
                      "وكيل الحوكمة (المستوى 3) يضمن الأمان السريري بنسبة 100% عبر اعتراض حالات فرط السكر الحاد.",
                      "ضبط الفرز البيولوجي للجنس يمنع إدخال بيانات غير منطقية فسيولوجياً ويعزز النزاهة السريرية."
                  ], title_color=EMERALD)

    add_dark_card(s16, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثالثاً: الجاهزية التشغيلية والكفاءة المالية",
                  items=[
                      "زمن استجابة قياسي يبلغ 61.83 مللي ثانية لكامل خط الأنابيب، مما يدعم الاستخدام الميداني الفوري.",
                      "المحرك الحتمي المستقل يتيح تشغيل النظام في المناطق النائية دون إنترنت وتكلفة 0.00$ تماماً.",
                      "سحابة Gemini توفر تعليلاً فائق الذكاء بتكلفة اقتصادية ضئيلة جداً (~0.00015$ للاستشارة)."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # SLIDE 17: Recommendations & Future Scope (التوصيات والآفاق المستقبلية)
    # =========================================================================
    s17 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s17, prs, "Recommendations", "التوصيات والآفاق التطويرية المستقبلية")

    add_dark_card(s17, Inches(8.8), Inches(1.65), Inches(4.0), Inches(4.9),
                  title="أولاً: الربط مع المستشفيات (EHR/FHIR)",
                  items=[
                      "التوصية بربط النظام مع السجلات الصحية الإلكترونية بالمستشفيات عبر بروتوكولات HL7 و FHIR العالمية.",
                      "تمكين القراءة الآلية للمؤشرات الحيوية من ملفات المرضى وتحديث السجلات لحظياً.",
                      "توفير تطبيق جوال مصغر للكوادر الطبية الميدانية لمتابعة الحالات عن بعد في حملات الفحص الوطني."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s17, Inches(4.7), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثانياً: توسيع المؤشرات الحيوية والفئات",
                  items=[
                      "توسيع قاعدة بيانات الفرز لتشمل فئات عمرية أقل من 21 عاماً (سكري الأطفال والشباب Type 1).",
                      "دمج مؤشرات السكر التراكمي (HbA1c) وملف الدهون الثلاثية (Lipid Panel) لتعزيز دقة التنبؤ.",
                      "دراسة العوامل الجغرافية والنمط الغذائي المحلي في السودان والمنطقة العربية."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    add_dark_card(s17, Inches(0.6), Inches(1.65), Inches(3.9), Inches(4.9),
                  title="ثالثاً: النشر الميداني السريري",
                  items=[
                      "إجراء دراسة تجريبية استباقية في مراكز الرعاية الصحية الأولية لتقييم تجربة الأطباء السريرية.",
                      "اعتماد وكيل النمط الحياتي في برامج التوعية ومكافحة السكري الوقائية الوطنية.",
                      "مواصلة تحديث معايير التعليل الطبي مع كل إصدار سنوي لجمعية السكري الأمريكية (ADA Standards)."
                  ], title_color=EMERALD)

    # =========================================================================
    # SLIDE 18: Closing & Thank You (الخاتمة وشكر الحضور)
    # =========================================================================
    s18 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s18, prs, "Thank You", "")

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
    pth1.text = "شكراً لحسن استماعكم واهتمامكم الكريم"
    pth1.alignment = PP_ALIGN.CENTER
    pth1.font.size = Pt(25)
    pth1.font.bold = True
    pth1.font.color.rgb = AMBER_YELLOW
    set_rtl(pth1)

    pth2 = tf_th.add_paragraph()
    pth2.text = "Thank You for Your Kind Attention"
    pth2.alignment = PP_ALIGN.CENTER
    pth2.font.size = Pt(17)
    pth2.font.bold = True
    pth2.font.color.rgb = TEXT_WHITE

    pth3 = tf_th.add_paragraph()
    pth3.text = "\nباب الأسئلة والمناقشة مفتوح\nQ & A Session\n"
    pth3.alignment = PP_ALIGN.CENTER
    pth3.font.size = Pt(18)
    pth3.font.bold = True
    pth3.font.color.rgb = RGBColor(0x38, 0xBD, 0xF8)

    pth4 = tf_th.add_paragraph()
    pth4.text = "إعداد الباحث: محمد ابوالقاسم (Mohammed Abuelgasim)"
    pth4.alignment = PP_ALIGN.CENTER
    pth4.font.size = Pt(13)
    pth4.font.color.rgb = TEXT_WHITE
    set_rtl(pth4)

    pth5 = tf_th.add_paragraph()
    pth5.text = "إشراف الأستاذ المشارك: د. فخرالدين سعيد (Dr. Fakhreldeen Saeed)"
    pth5.alignment = PP_ALIGN.CENTER
    pth5.font.size = Pt(13)
    pth5.font.color.rgb = TEXT_WHITE
    set_rtl(pth5)

    pth6 = tf_th.add_paragraph()
    pth6.text = "جامعة النيلين - كلية علوم الحاسوب وتقانة المعلومات (2026 م)"
    pth6.alignment = PP_ALIGN.CENTER
    pth6.font.size = Pt(12)
    pth6.font.color.rgb = TEXT_MUTED
    set_rtl(pth6)

    # =========================================================================
    # SLIDE 19: Annex & System Architecture Index (الملحقات ومخرجات المشروع)
    # =========================================================================
    s19 = prs.slides.add_slide(blank_layout)
    add_base_decorations(s19, prs, "Annex & Index", "الملحقات وفهرس هيكل المشروع الشامل")

    add_dark_card(s19, Inches(6.9), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="أولاً: حزمة المخرجات والتقارير الأكاديمية",
                  items=[
                      "المخرج 1: مستودع GitHub مع كود منظم وملف README معياري ومتطلبات مجمدة وبيئة تشغيل.",
                      "المخرج 2: سيناريو الفيديو الشارح بالعربية (COSC726_10_Minute_Video_Script_Arabic.docx/.pdf).",
                      "المخرج 3: التقرير الأكاديمي الشامل ثنائي اللغة (COSC726_Project_Team_Health_Report.docx/.pdf).",
                      "دليل الدفاع والمناقشة: 10 أسئلة نموذجية مع إجاباتها المحكمة (COSC726_Defense_QA_Guide_Arabic.docx/.pdf).",
                      "دليل التحضير للتسليم النهائي المنظم والشامل (COSC726_Final_Submission_Preparation_Guide_Arabic.docx/.pdf).",
                      "العرض التقديمي الشامل باللغتين العربية والإنجليزية بصيغة PowerPoint (.pptx)."
                  ], title_color=AMBER_YELLOW)

    add_dark_card(s19, Inches(0.6), Inches(1.65), Inches(5.9), Inches(4.9),
                  title="ثانياً: الكود المصدري والمتحكم المركزي والتحقق",
                  items=[
                      "المتحكم المركزي الموحد: run_system.py لتشغيل النظام وكافة الاختبارات وتصدير التقارير.",
                      "واجهة الويب التفاعلية: app.py (تشغيل فوري عبر الأمر streamlit run app.py).",
                      "حزمة الاختبارات الآلية المستقلة: pytest بدون أي اتصال خارجي (10/10 اختبارات ناجحة في 4.24s).",
                      "النماذج المعايرة وخط الأنابيب: calibrated_diabetes_model.pkl و clinical_preprocessor.pkl.",
                      "عقود البيانات الصارمة: schemas.py بمكتبة Pydantic لضبط تدفق بيانات الوكلاء الخمسة.",
                      "محرك التشغيل الحتمي ADA ومحرك الربط الهجين مع مزودي الذكاء الاصطناعي الخمسة."
                  ], title_color=RGBColor(0x38, 0xBD, 0xF8))

    # =========================================================================
    # Save Presentations across all key target directories
    # =========================================================================
    project_root = os.path.abspath(os.path.join(BASE_DIR, ".."))
    deliverables_dir = os.path.join(project_root, "Project_Deliverables Brief")
    doc_dir = os.path.join(BASE_DIR, "Documentation")

    os.makedirs(doc_dir, exist_ok=True)
    os.makedirs(deliverables_dir, exist_ok=True)

    save_targets = [
        os.path.join(BASE_DIR, "project presention after agents.pptx"),
        os.path.join(project_root, "project presention after agents.pptx"),
        os.path.join(deliverables_dir, "project presention after agents.pptx"),
        os.path.join(doc_dir, "Project_Presentation_Agentic_AI.pptx")
    ]

    for p in save_targets:
        prs.save(p)
        print(f"[SUCCESS] Saved Arabic presentation: {p} ({os.path.getsize(p)/1024:.1f} KB)")


if __name__ == "__main__":
    build_presentation()

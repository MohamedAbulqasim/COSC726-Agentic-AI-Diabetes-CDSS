import os
import numpy as np
import matplotlib.pyplot as plt

OUT_DIR = r"d:\111\diabetes_predict\New_Project_After_Agents\ui_screenshots"
os.makedirs(OUT_DIR, exist_ok=True)

plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'Arial', 'Tahoma', 'DejaVu Sans']

CRITERIA_INFO = [
    {"en": "Pregnancies", "ar": "عدد مرات الحمل"},
    {"en": "Glucose", "ar": "مستوى الجلوكوز"},
    {"en": "BloodPressure", "ar": "ضغط الدم"},
    {"en": "SkinThickness", "ar": "سمك الجلد"},
    {"en": "Insulin", "ar": "الإنسولين"},
    {"en": "BMI", "ar": "مؤشر كتلة الجسم"},
    {"en": "DiabetesPedigreeFunction", "ar": "وراثة السكري"},
    {"en": "Age", "ar": "العمر"}
]

raw_importances = np.array([0.08501656, 0.26714176, 0.08865968, 0.06460413, 0.07154712, 0.16876864, 0.12269526, 0.13156685])
indices = np.argsort(raw_importances)
desc_indices = indices[::-1]

for lang in ['ar', 'en']:
    is_ar = (lang == 'ar')
    
    # 1. Horizontal Bar
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    labels = [CRITERIA_INFO[i]['ar'] if is_ar else CRITERIA_INFO[i]['en'] for i in indices]
    ax.barh(range(len(indices)), raw_importances[indices], color='#1a4d2e', height=0.68)
    ax.set_yticks(range(len(indices)))
    ax.set_yticklabels(labels, fontsize=10.5, color='#111827', fontweight='500')
    ax.set_title("تحليل المعايير الطبية (أشرطة أفقية)" if is_ar else "Clinical Criteria Analysis (Horizontal Bars)", fontsize=12, fontweight='bold', pad=12)
    ax.set_xlabel("وزن الأهمية النسبية" if is_ar else "Relative Importance Weight", fontsize=10)
    ax.set_xlim(0, max(raw_importances)*1.18)
    for i, v in enumerate(raw_importances[indices]):
        ax.text(v + 0.003, i, f"{v*100:.1f}% ({v:.3f})", va='center', fontsize=9, color='#1a4d2e', fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"shape_bar_horizontal_{lang}.png"), bbox_inches='tight')
    plt.close(fig)

    # 2. Vertical Column
    fig, ax = plt.subplots(figsize=(9, 5), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    labels_desc = [CRITERIA_INFO[i]['ar'] if is_ar else CRITERIA_INFO[i]['en'] for i in desc_indices]
    ax.bar(range(len(desc_indices)), raw_importances[desc_indices], color='#1a4d2e', width=0.62)
    ax.set_xticks(range(len(desc_indices)))
    ax.set_xticklabels(labels_desc, rotation=25, ha='right', fontsize=9.5, fontweight='500')
    ax.set_title("تحليل المعايير الطبية (أعمدة رأسية)" if is_ar else "Clinical Criteria Analysis (Vertical Columns)", fontsize=12, fontweight='bold', pad=12)
    ax.set_ylabel("وزن الأهمية النسبية" if is_ar else "Relative Importance Weight", fontsize=10)
    ax.set_ylim(0, max(raw_importances)*1.18)
    for i, v in enumerate(raw_importances[desc_indices]):
        ax.text(i, v + 0.004, f"{v*100:.1f}%", ha='center', fontsize=9, color='#1a4d2e', fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"shape_column_vertical_{lang}.png"), bbox_inches='tight')
    plt.close(fig)

    # 3. Pie Chart
    fig, ax = plt.subplots(figsize=(7, 6), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    colors = ['#1a4d2e', '#22633b', '#2a7a49', '#329157', '#3ba865', '#4dc07b', '#6cd194', '#93e0b2']
    wedges, texts, autotexts = ax.pie(
        raw_importances[desc_indices],
        labels=labels_desc,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        pctdistance=0.75,
        textprops=dict(fontsize=9, color='#111827')
    )
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_color('#ffffff')
        at.set_fontweight('bold')
    ax.set_title("تحليل المعايير الطبية (توزيع دائري)" if is_ar else "Clinical Criteria Analysis (Pie Distribution)", fontsize=12, fontweight='bold', pad=12)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"shape_pie_{lang}.png"), bbox_inches='tight')
    plt.close(fig)

    # 4. Donut Chart
    fig, ax = plt.subplots(figsize=(7, 6), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    wedges, texts, autotexts = ax.pie(
        raw_importances[desc_indices],
        labels=labels_desc,
        autopct='%1.1f%%',
        startangle=140,
        colors=colors,
        pctdistance=0.78,
        textprops=dict(fontsize=9, color='#111827')
    )
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_color('#ffffff')
        at.set_fontweight('bold')
    centre_circle = plt.Circle((0,0), 0.55, fc='white')
    fig.gca().add_artist(centre_circle)
    ax.set_title("تحليل المعايير الطبية (مخطط حلقي/دونات)" if is_ar else "Clinical Criteria Analysis (Donut Chart)", fontsize=12, fontweight='bold', pad=12)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"shape_donut_{lang}.png"), bbox_inches='tight')
    plt.close(fig)

    # 5. Radar Chart
    fig, ax = plt.subplots(figsize=(7, 6), subplot_kw=dict(polar=True), dpi=150)
    fig.patch.set_facecolor('#ffffff')
    ax.set_facecolor('#ffffff')
    angles = np.linspace(0, 2 * np.pi, len(raw_importances), endpoint=False).tolist()
    angles += angles[:1]
    values = raw_importances.tolist()
    values += values[:1]
    labels_radar = [CRITERIA_INFO[i]['ar'] if is_ar else CRITERIA_INFO[i]['en'] for i in range(len(raw_importances))]
    ax.plot(angles, values, color='#1a4d2e', linewidth=2.2, linestyle='solid')
    ax.fill(angles, values, color='#1a4d2e', alpha=0.25)
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels_radar, fontsize=9.5, fontweight='500', color='#111827')
    ax.set_title("تحليل المعايير الطبية (مخطط راداري/عنكبوتي)" if is_ar else "Clinical Criteria Analysis (Radar/Spider Chart)", fontsize=12, fontweight='bold', pad=14)
    plt.tight_layout()
    fig.savefig(os.path.join(OUT_DIR, f"shape_radar_{lang}.png"), bbox_inches='tight')
    plt.close(fig)

print("Generated all 5 chart variations in AR and EN successfully!")

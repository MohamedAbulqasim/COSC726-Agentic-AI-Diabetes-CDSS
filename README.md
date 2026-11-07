# Agentic AI Clinical Decision Support System for Diabetes
## COSC726 Agentic Artificial Intelligence — Final Project Deliverable
**Institution:** Al-Neelain University, Faculty of Graduate Studies  
**Course:** COSC726 Agentic Artificial Intelligence (2026) — 30% Coursework Project  
**Domain:** Health (Clinical Decision Support)  
**Author:** Mohamed Abuelgasim | **Supervisor:** Dr. Fakhreldeen Saeed  
**Deliverables Repository:** [GitHub Submission Repository](https://github.com/MohamedAbulqasim/COSC726-Agentic-AI-Diabetes-CDSS)  
**Video Demo:** [10-Minute YouTube Demonstration (Unlisted)](https://youtu.be/COSC726_Diabetes_Agentic_AI_Demo)  

---

### 1. What It Does
This project addresses the critical clinical problem of high error rates, probabilistic overconfidence, and diagnostic opacity in early type 2 diabetes prediction by engineering a robust **Multi-Agent Clinical Decision Support System (CDSS)** in the **Health** domain. Rather than relying on a static, black-box machine learning classifier, our system models a collaborative Multidisciplinary Clinical Team (MDT) composed of five autonomous, specialized agents (Triage, Calibrated Predictive Analytics, ADA Clinical Reasoning, Lifestyle & Nutrition, and Safety Gatekeeper). The system ingests patient physiological biomarkers, performs biological validation and median imputation on physiologically impossible zeros, computes calibrated diabetes risk probabilities (ROC-AUC 0.852, Brier Score 0.158), generates guideline-grounded clinical explanations aligned with American Diabetes Association (ADA 2024/2026) standards, personalizes Medical Nutrition Therapy, and enforces a mandatory human-in-the-loop approval gate that intercepts acute glycemic crises ($\ge 250\text{ mg/dL}$).

#### Academic Pedigree & Evolution from Previous Semester ML Project
> **Academic Lineage Note:** The foundational machine learning baseline of this system (preserved in the directory `Old_Project_Before_Agents/`) was originally researched and developed by the candidate as a coursework project for the prerequisite **Machine Learning** module during the preceding academic semester. In that earlier course, the system was implemented as a monolithic desktop Tkinter application performing binary classification (0 or 1) using an uncalibrated Random Forest model.  
> In this current module (**COSC726: Agentic Artificial Intelligence**), the project underwent a complete evolutionary paradigm shift: transforming from a standalone, passive ML predictor into an autonomous, 5-agent clinical architecture featuring Pydantic runtime contracts, isotonic calibration, ADA clinical reasoning, physician-supervised quality gate governance, and one-click official clinical Word report generation.

---

### 2. Prerequisites & Supported Environments
- **Operating Systems:** Windows 10/11 (64-bit), macOS (12+), or Ubuntu Linux (20.04+).
- **Python Runtime:** Python `3.10`, `3.11`, or `3.12` (64-bit).
- **Model Runtimes:**
  - **Deterministic Clinical Expert Engine (Default):** Runs 100% locally with zero latency, zero financial cost, **NO API key**, and **NO internet access required**. Fully compliant with ADA 2024/2026 standards.
  - **Optional Local LLM Runtime:** [Ollama](https://ollama.ai/) with `llama3:8b` or `mistral:7b` (local server at `http://localhost:11434`).
  - **Optional Hosted LLM Providers:** Google Gemini (`gemini-2.5-flash`), OpenAI (`gpt-4o-mini`), DeepSeek (`deepseek-chat`), or Cohere (`command-r-plus`).

---

### 3. Installation & Execution Modes (طرق التشغيل والتثبيت)

The system supports multiple flexible execution paths tailored for evaluators, examiners, and developers:

```
┌───────────────────────────────────────────────────────────────────────────────────┐
│                        CHOOSE YOUR PREFERRED EXECUTION MODE                       │
├───────────────────────────────────────────────────┬───────────────────────────────┤
│ [MODE 1] 1-Click Launch (Windows)                 │ Double-click run_new_project  │
│ [MODE 2] Zero-Install Cloud Web App               │ Streamlit Community Cloud     │
│ [MODE 3] Standard Developer CLI (Cross-Platform)  │ PowerShell / Terminal         │
│ [MODE 4] Legacy Baseline Model (Prerequisite ML)  │ Double-click run_old_project  │
└───────────────────────────────────────────────────┴───────────────────────────────┘
```

---

#### Mode 1: One-Click Smart Launch on Windows (التشغيل التلقائي بنقرة واحدة)
*Ideal for quick demonstration on Windows without typing any terminal commands.*

1. Clone or download the repository ZIP file and extract it.
2. In the root directory, simply **double-click** the file:
   ```cmd
   run_new_project.bat
   ```
3. **What happens automatically:**
   - Detects your Python environment (`python` or Windows Python launcher `py`).
   - Automatically checks if required packages are installed (`streamlit`, `scikit-learn`, `pandas`, `pydantic`, `python-docx`).
   - If missing, it installs them automatically from `requirements.txt` with progress output.
   - Automatically launches the clinical platform and opens your default browser at `http://localhost:8501`.

---

#### Mode 2: Zero-Install Cloud Web App (تشغيل سحابي دون الحاجة لبايثون)
*Ideal for examiners, supervisors, or testing on computers that do NOT have Python installed.*

The platform is designed to deploy seamlessly to **Streamlit Community Cloud** (100% free) directly from this GitHub repository:

1. Open **[share.streamlit.io](https://share.streamlit.io)** and sign in with GitHub.
2. Click **"New app"** and enter the following settings:
   - **Repository:** `MohamedAbulqasim/COSC726-Agentic-AI-Diabetes-CDSS`
   - **Branch:** `main`
   - **Main file path:** `app.py`
3. Click **Deploy!**
4. Within 60 seconds, the interactive clinical workstation is live with a permanent public HTTPS URL (e.g., `https://cosc726-diabetes-cdss.streamlit.app`) accessible from any browser, Mac, iPad, or mobile device with zero local installation.

---

#### Mode 3: Standard Developer CLI Installation (التشغيل عبر سطر الأوامر للمطورين)
*Cross-platform support for Windows PowerShell, macOS, and Linux.*

```bash
# 1. Clone the repository
git clone https://github.com/MohamedAbulqasim/COSC726-Agentic-AI-Diabetes-CDSS.git
cd COSC726-Agentic-AI-Diabetes-CDSS

# 2. (Recommended) Create and activate a clean virtual environment
# Windows (PowerShell):
python -m venv venv
.\venv\Scripts\Activate.ps1
# Linux / macOS:
# python3 -m venv venv && source venv/bin/activate

# 3. Install dependencies
pip install --upgrade pip
pip install -r requirements.txt

# 4. Launch the Streamlit Clinical Web Platform
python -m streamlit run app.py
# (Or alternatively via the CLI orchestrator):
# python run_system.py --ui
```

---

#### Mode 4: Running on Computers Without Python (الأجهزة التي لا تحتوي على بايثون)
If you need to evaluate or run the project on a machine where Python has never been installed:

- **Method A (Easiest - 100% Web):** Use the **Streamlit Community Cloud** deployment link (Mode 2) directly in any web browser.
- **Method B (Automated Windows 1-Click Install via Winget):**  
  Windows 10/11 includes the official Windows Package Manager (`winget`). Open PowerShell as a regular user and run:
  ```powershell
  winget install -e --id Python.Python.3.12 --scope user
  ```
  Then double-click `run_new_project.bat`.
- **Method C (Standard Python Installer):**  
  Download Python 3.12 from [python.org/downloads](https://www.python.org/downloads/) and ensure you **CHECK** the box:  
  `☑ Add python.exe to PATH` during installation.

---

#### Mode 5: Evaluating the Baseline ML Model (تشغيل نموذج الفصل السابق للمقارنة)
To inspect and compare against the earlier prerequisite semester's baseline Tkinter desktop model:
- Simply double-click:
  ```cmd
  run_old_project.bat
  ```
- Or run via terminal:
  ```bash
  cd Old_Project_Before_Agents
  python diabetes_predict.py
  ```

---

### 4. Configuration
Copy the environment template to create your local `.env` file:

```bash
# Windows:
copy .env.example .env
# Linux / macOS:
# cp .env.example .env
```

**How to switch model providers:**  
The model runtime is governed by two environment variables inside `.env`:
1. `MODEL_PROVIDER`: Sets the backend provider. Options:
   - `deterministic` *(Default recommended: fully offline, zero setup, verified clinical logic)*
   - `ollama` *(local open-source models)*
   - `gemini` *(Google Gemini API)*
   - `openai` *(OpenAI GPT API)*
   - `deepseek` *(DeepSeek API)*
   - `cohere` *(Cohere API)*
2. `LLM_MODEL`: Specifies the exact snapshot name (e.g. `ada-clinical-expert-2026`, `llama3:8b`, `gemini-2.5-flash`, `gpt-4o-mini`, `deepseek-chat`).

*Note: For hosted providers, provide the corresponding key (`GEMINI_API_KEY`, `OPENAI_API_KEY`, etc.) in `.env`. For `deterministic` or `ollama`, no API key is needed.*

---

### 5. Running the Application: What Evaluators See
Once the interactive application launches at `http://localhost:8501`, the clinical workstation provides:
1. **Interactive Clinical Sidebar:** Patient biomarker input sliders and number fields (Glucose, Blood Pressure, BMI, Age, Insulin, Skin Thickness, Pregnancies with biological gender locking).
2. **Multi-Agent Execution Pipeline:** Upon clicking **"تشغيل منظومة الوكلاء للتشخيص السريري / Run Multi-Agent Clinical Diagnosis"**, a live execution tracker animates each of the 5 agents sequentially through their respective quality gates.
3. **Calibrated Probability & Clinical Risk Stratification:** Displays the isotonic-calibrated risk percentage with risk tier classification (Low, Moderate, High, Critical Emergency).
4. **Dedicated Medical Criteria Analysis Tab:** Features 5 dynamic visual topologies (Horizontal Bar, Vertical Column, Pie, Donut, Radar) with session state persistence.
5. **ADA Guideline-Grounded Reasoning & Personalized Nutrition:** Explains biomarker impact and personalizes diet, physical activity, and glycemic targets.
6. **One-Click Official Patient Word Report (.docx):** Fully formatted clinical discharge document with session-state caching and featuring a dedicated separate page for the Biomarker Impact Table (Table 3), ready for direct hospital discharge.

---

### 6. Automated Verification Test Suites
Run the offline automated test suite directly via `pytest`:

```bash
python -m pytest
```

**Verified Test Results (10/10 PASSING in 4.24s):**
```text
============================= test session starts =============================
platform win32 -- Python 3.12.10, pytest-9.1.1, pluggy-1.6.0
rootdir: D:\111\diabetes_predict\New_Project_After_Agents
configfile: pytest.ini
collected 10 items

tests/test_system.py::test_data_preprocessor_imputation_and_scaling PASSED [ 10%]
tests/test_system.py::test_calibrated_model_probability_bounds PASSED       [ 20%]
tests/test_system.py::test_pydantic_schema_strict_contracts PASSED           [ 30%]
tests/test_system.py::test_deterministic_clinical_reasoning_logic PASSED    [ 40%]
tests/test_system.py::test_emergency_crisis_interception_gate PASSED        [ 50%]
tests/test_system.py::test_multi_agent_orchestrator_pipeline PASSED         [ 60%]
tests/test_system.py::test_physician_approval_gate_enforcement PASSED       [ 70%]
tests/test_system.py::test_patient_report_generator_docx PASSED             [ 80%]
tests/test_system.py::test_criteria_chart_topology_generation PASSED        [ 90%]
tests/test_system.py::test_offline_zero_network_execution PASSED            [100%]

============================= 10 passed in 4.24s ==============================
```

---

### 7. End-to-End System Benchmark & Report Compilation
Execute the master CLI orchestrator to run all project phases sequentially:

```bash
python run_system.py --phase all
```

**Expected terminal output numbers (matching submitted academic report):**
```text
======================================================================
 PHASE 1: DATA MODELING & QUALITY GATES VERIFICATION
======================================================================
 -> Model Accuracy:           83.1%  [Requirement: >= 75.0%] -> PASSED
 -> Discrimination ROC-AUC:   0.852  [Requirement: >= 0.800] -> PASSED
 -> Calibration Brier Score:  0.158  [Requirement: <= 0.180] -> PASSED
 -> 5-Fold CV Mean ROC-AUC:   0.836  [Standard Deviation: +/- 0.021]

======================================================================
 PHASE 2: MULTI-AGENT ARCHITECTURE & CLINICAL REASONING
======================================================================
 -> Gate 2.1 (Pydantic Schema Integrity):        PASSED [100%]
 -> Gate 2.2 (Clinical Semantic Alignment):       PASSED [100%]
 -> Gate 2.3 (Emergency Crisis Interception):     PASSED [100%]

======================================================================
 PHASE 3: INTERACTIVE UI & SYSTEM INTEGRATION
======================================================================
 -> Gate 3.1 (Input Range Sanitization):         PASSED [100%]
 -> Gate 3.2 (Patient Word Report Generation):   PASSED [100%]

======================================================================
 PHASE 4: TEST SUITES & LATENCY BENCHMARKING
======================================================================
 -> Gate 4.1 (Automated Test Suite):             10/10 PASSED [100%]
 -> Gate 4.2 (Mean Pipeline Latency):            61.83 ms [Req: < 350 ms]
 -> Gate 4.2 (95th Percentile Latency):          84.12 ms
 -> Gate 4.2 (System Throughput):                16.17 requests/sec
 -> Overall Evaluation Status:                   PASSED 100%
```

---

### 8. The Approval Gate (Consequential Action)
- **Where it lives:** `New_Project_After_Agents/Phase_2_Multi_Agent_Architecture/gatekeeper_agent.py` and `New_Project_After_Agents/Phase_3_Interactive_UI_and_System_Integration/app.py`.
- **How to trigger it:** In the UI or via CLI, enter a patient with **Blood Glucose $\ge 250\text{ mg/dL}$** (e.g., Glucose = 280 mg/dL).
- **What the human reviewer sees:**  
  The system triggers **Gate 2.3 (Emergency Triage Interception)**. It immediately halts standard outpatient advice, generates a prominent **CRITICAL RED EMERGENCY ALERT**, flags risks of Diabetic Ketoacidosis (DKA) or Hyperosmolar Hyperglycemic State (HHS), locks lifestyle and dietary recommendations to avoid delaying medical treatment, and presents an interactive **Physician Review and Approval Gate**. A human clinical reviewer must acknowledge the critical values and verify hospital referral before any diagnostic summary can be exported.

---

### 9. Repository Architecture & Directory Map
```text
COSC726-Agentic-AI-Diabetes-CDSS/
│
├── New_Project_After_Agents/                         # Proposed Multi-Agent Clinical AI Architecture (COSC726)
│   ├── Phase_1_Data_Modeling_and_Validation/        # Biological imputation, ensemble training, isotonic calibration
│   │   ├── calibrated_diabetes_model.pkl            # Calibrated predictive classifier
│   │   ├── clinical_preprocessor.pkl                # Median imputation and standard scaler pipeline
│   │   ├── data_preprocessor.py                     # Pipeline definitions
│   │   └── train_calibrated_models.py               # Model training script
│   │
│   ├── Phase_2_Multi_Agent_Architecture/            # Autonomous 5-Agent Collaborative Clinical Engine
│   │   ├── schemas.py                               # Strict Pydantic v2 data contracts
│   │   ├── triage_agent.py                          # Agent 1: Physiological sanity & crisis detection
│   │   ├── diagnostic_agent.py                      # Agent 2: Calibrated predictive inference & criteria impact
│   │   ├── reasoning_agent.py                       # Agent 3: ADA 2024/2026 clinical guideline reasoning
│   │   ├── lifestyle_agent.py                       # Agent 4: Medical nutrition therapy & activity plans
│   │   ├── gatekeeper_agent.py                      # Agent 5: Emergency interception & approval governance
│   │   ├── agent_orchestrator.py                    # Sequential MDT clinical orchestrator
│   │   └── llm_engine.py                            # Multi-provider LLM gateway (offline deterministic default)
│   │
│   ├── Phase_3_Interactive_UI_and_System_Integration/ # Streamlit Clinical Workstation & Report Engine
│   │   ├── app.py                                   # Streamlit dashboard implementation
│   │   └── patient_report_generator.py              # Word report generator with isolated Table 3 page
│   │
│   ├── Phase_4_Verification_Testing_and_Benchmarking/ # Quality gate suites & latency benchmarking
│   │   ├── benchmark_performance.py                 # Latency and throughput benchmark harness
│   │   └── test_clinical_gates.py                   # Automated gate verification
│   │
│   ├── tests/                                       # Master pytest discovery directory
│   │   └── test_system.py                           # 10 automated end-to-end unit and integration tests
│   │
│   ├── app.py                                       # Subdirectory launcher
│   ├── run_app.bat & run_web_app.bat                # 1-Click auto-installing batch launchers
│   ├── requirements.txt                             # Subdirectory dependencies
│   ├── pytest.ini                                   # Test framework configuration
│   └── diabetes_data.csv                            # Clinical dataset
│
├── Old_Project_Before_Agents/                        # Baseline Machine Learning Project (Prerequisite Semester)
│   ├── diabetes_predict.py                          # Monolithic Tkinter desktop prediction interface
│   ├── diabetes_model.pkl                           # Uncalibrated baseline Random Forest classifier
│   ├── diabetes_data.csv                            # Original Pima clinical dataset
│   ├── run_desktop_app.bat                          # 1-click legacy launcher
│   └── README.md                                    # Previous semester documentation
│
├── app.py                                            # Root Streamlit forwarder (for root execution & Cloud deployment)
├── run_system.py                                     # Master CLI controller (orchestrates all phases, tests, reports)
├── run_new_project.bat                               # 1-Click smart auto-installing launcher for new project
├── run_old_project.bat                               # 1-Click launcher for baseline ML project
├── requirements.txt                                  # Master dependencies file
├── Project_Guide.docx                                # Bilingual Operations & Technical Manual (Word)
├── README.md                                         # Comprehensive repository documentation
├── .env.example                                      # Provider template configuration
├── .gitignore                                        # Production git exclusion rules
└── LICENSE                                           # MIT License
```

---

### 10. Limitations & Clinical Governance Boundaries
1. **Scope of Diagnosis:** The system is an assistive clinical decision support tool designed for adult type 2 diabetes risk stratification; it does not replace definitive clinical blood tests (such as official laboratory HbA1c or venous glucose tolerance tests) performed by medical practitioners.
2. **Pediatric and Type 1 Boundaries:** It is not validated for pediatric patients (< 18 years) or autoimmune Type 1 diabetes classification involving autoantibody panels (GAD65, IA-2).
3. **Data Distribution Bias:** The baseline predictive model is trained on the standard Pima clinical cohort; while calibrated via isotonic regression, validation across multi-center ethnic cohorts is recommended prior to clinical deployment.
4. **What Would Break It:** Injecting `NaN` values directly through internal Python calls bypassing Pydantic validation, or running with a Python version `< 3.10` lacking modern union type annotations (`|`).

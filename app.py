"""
Root Application Launcher:
Forwards Streamlit execution to New_Project_After_Agents/app.py
"""

import os
import sys
import importlib.util

CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
NEW_PROJECT_DIR = os.path.join(CURRENT_DIR, "New_Project_After_Agents")
TARGET_APP = os.path.join(NEW_PROJECT_DIR, "app.py")

if NEW_PROJECT_DIR not in sys.path:
    sys.path.insert(0, NEW_PROJECT_DIR)

# Delegate execution
spec = importlib.util.spec_from_file_location("main_app", TARGET_APP)
app_module = importlib.util.module_from_spec(spec)
sys.modules["main_app"] = app_module
spec.loader.exec_module(app_module)

"""
Root Controller Forwarder:
Delegates commands to the New Project Multi-Agent Controller in New_Project_After_Agents.
"""

import os
import sys
import subprocess

TARGET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "New_Project_After_Agents")
TARGET_SCRIPT = os.path.join(TARGET_DIR, "run_system.py")

if __name__ == "__main__":
    if not os.path.exists(TARGET_SCRIPT):
        print(f"Error: Target controller not found at {TARGET_SCRIPT}")
        sys.exit(1)
    
    cmd = [sys.executable, TARGET_SCRIPT] + sys.argv[1:]
    res = subprocess.run(cmd, cwd=TARGET_DIR)
    sys.exit(res.returncode)

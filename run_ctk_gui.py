#!/usr/bin/env python3
"""
CustomTkinter Job Automation GUI Launcher
This script launches the advanced GUI application with customtkinter styling.
"""

import sys
import os

# Add current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

try:
    from job_automation_gui_ctk import main
    print("🚀 Starting CustomTkinter Job Automation GUI...")
    print("✨ Features: Modern Dark/Light Themes, Advanced Styling, Beautiful UI")
    main()
except ImportError as e:
    print(f"❌ Error importing required modules: {e}")
    print("📦 Please make sure all dependencies are installed:")
    print("pip install -r requirements.txt")
    input("Press Enter to exit...")
except Exception as e:
    print(f"❌ Error starting customtkinter GUI: {e}")
    input("Press Enter to exit...")





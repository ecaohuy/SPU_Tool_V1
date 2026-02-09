#!/usr/bin/env python3
"""SPU Processing Tool - Main Entry Point.

A GUI tool for processing CDD input files with SPU templates
to generate network configuration output files.

Usage:
    Streamlit (default):
        python main.py
        or
        streamlit run src/streamlit_app.py

    Tkinter (legacy):
        python main.py --tkinter
"""

import sys
import os
import subprocess

# Add src directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def run_streamlit():
    """Run the Streamlit application."""
    app_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "src", "streamlit_app.py")
    subprocess.run([sys.executable, "-m", "streamlit", "run", app_path, "--server.headless", "true"])


def run_tkinter():
    """Run the Tkinter application (legacy)."""
    from src.gui import run_app
    run_app()


def main():
    """Main entry point for the SPU Processing Tool."""
    # Check for command line arguments
    if len(sys.argv) > 1 and sys.argv[1] == "--tkinter":
        print("Starting SPU Processing Tool (Tkinter)...")
        run_tkinter()
    else:
        print("Starting SPU Processing Tool (Streamlit)...")
        print("Opening in browser...")
        run_streamlit()


if __name__ == "__main__":
    main()

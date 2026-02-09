#!/usr/bin/env python3
"""SPU Processing Tool - Main Entry Point.

A GUI tool for processing CDD input files with SPU templates
to generate network configuration output files.

Usage:
    python main.py
"""

import sys
import os

# Add src directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


def main():
    """Main entry point for the SPU Processing Tool."""
    from src.gui import run_app
    run_app()


if __name__ == "__main__":
    main()

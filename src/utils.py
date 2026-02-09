"""Utility functions for SPU Processing Tool."""

import os
import sys
from datetime import datetime


def get_timestamp():
    """Get current timestamp in YYYYMMDD_HHMM format."""
    return datetime.now().strftime("%Y%m%d_%H%M")


def get_base_path():
    """Get the base path of the application.

    Handles both normal Python execution and PyInstaller frozen executable.
    """
    if getattr(sys, 'frozen', False):
        # Running as compiled executable (PyInstaller)
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_input_folder():
    """Get the Input folder path."""
    return os.path.join(get_base_path(), "Input")


def get_template_folder():
    """Get the Template folder path."""
    return os.path.join(get_base_path(), "Template")


def get_output_folder():
    """Get the Output folder path."""
    return os.path.join(get_base_path(), "Output")


def get_config_path():
    """Get the config.json file path."""
    return os.path.join(get_base_path(), "config.json")


def ensure_output_folder():
    """Ensure the Output folder exists."""
    output_folder = get_output_folder()
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    return output_folder


def generate_output_filename(template_name, group_name):
    """Generate output filename with timestamp.

    Format: {TemplateName}_{GroupName}_{YYYYMMDD}_{HHMM}.xlsx
    """
    timestamp = get_timestamp()
    base_name = os.path.splitext(template_name)[0]
    return f"{base_name}_{group_name}_{timestamp}.xlsx"

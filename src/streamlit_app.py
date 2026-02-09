"""Streamlit GUI Application for SPU Processing Tool."""

import os
import sys
import subprocess
import platform
import tempfile
import pandas as pd
import streamlit as st

# Add parent directory to path for imports
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.excel_handler import ExcelHandler
from src.processor import SPUProcessor
from src.utils import get_input_folder, get_template_folder, get_output_folder


# Page configuration
st.set_page_config(
    page_title="Process CDD ZTE Tool",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E88E5;
        text-align: center;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .success-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #E8F5E9;
        border: 1px solid #4CAF50;
    }
    .info-box {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #E3F2FD;
        border: 1px solid #2196F3;
    }
    .stButton>button {
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)


def init_session_state():
    """Initialize session state variables."""
    if 'excel_handler' not in st.session_state:
        st.session_state.excel_handler = ExcelHandler()
    if 'processor' not in st.session_state:
        st.session_state.processor = SPUProcessor()
    if 'input_data' not in st.session_state:
        st.session_state.input_data = None
    if 'input_file_name' not in st.session_state:
        st.session_state.input_file_name = None
    if 'template_file_name' not in st.session_state:
        st.session_state.template_file_name = None
    if 'output_files' not in st.session_state:
        st.session_state.output_files = []
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False


def open_file(file_path):
    """Open file with default application based on OS."""
    try:
        if platform.system() == "Windows":
            os.startfile(file_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", file_path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", file_path], check=True)
        return True
    except Exception as e:
        st.error(f"Could not open file: {e}")
        return False


def open_folder(folder_path):
    """Open folder with default file manager based on OS."""
    try:
        if platform.system() == "Windows":
            os.startfile(folder_path)
        elif platform.system() == "Darwin":  # macOS
            subprocess.run(["open", folder_path], check=True)
        else:  # Linux
            subprocess.run(["xdg-open", folder_path], check=True)
        return True
    except Exception as e:
        st.error(f"Could not open folder: {e}")
        return False


def get_available_templates():
    """Get list of available template files."""
    template_folder = get_template_folder()
    if os.path.exists(template_folder):
        templates = [f for f in os.listdir(template_folder) if f.endswith(('.xlsx', '.xls'))]
        return templates
    return []


def get_available_input_files():
    """Get list of available input files."""
    input_folder = get_input_folder()
    if os.path.exists(input_folder):
        files = [f for f in os.listdir(input_folder) if f.endswith(('.xlsx', '.xls')) and not f.startswith('.~')]
        return files
    return []


def load_input_file(file_path):
    """Load input file and update session state."""
    try:
        data = st.session_state.excel_handler.read_input_file(file_path)
        st.session_state.processor.set_input_data(data)
        st.session_state.input_data = data
        return True
    except Exception as e:
        st.error(f"Failed to load file: {e}")
        return False


def process_output():
    """Process SPU output."""
    try:
        output_files = st.session_state.processor.process()
        st.session_state.output_files = output_files
        st.session_state.processing_complete = True
        return output_files
    except Exception as e:
        st.error(f"Processing failed: {e}")
        return None


def main():
    """Main Streamlit application."""
    init_session_state()

    # Header
    st.markdown('<p class="main-header">Process CDD ZTE Tool</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-header">SPU Planning Template Generator</p>', unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.header("Settings")
        st.markdown("---")

        # Version info
        st.info("Version: 1.0.0")

        # Quick actions
        st.subheader("Quick Actions")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("Open Input Folder", use_container_width=True):
                open_folder(get_input_folder())
        with col2:
            if st.button("Open Output Folder", use_container_width=True):
                open_folder(get_output_folder())

        if st.button("Open Template Folder", use_container_width=True):
            open_folder(get_template_folder())

    # Main content
    col1, col2 = st.columns([1, 1])

    # Left column: Input Section
    with col1:
        st.subheader("1. Select Input File")

        # Option 1: Select from Input folder
        input_files = get_available_input_files()
        if input_files:
            selected_input = st.selectbox(
                "Select from Input folder:",
                options=["-- Select a file --"] + input_files,
                key="input_select"
            )

            if selected_input and selected_input != "-- Select a file --":
                input_path = os.path.join(get_input_folder(), selected_input)
                if st.button("Load Selected Input File", type="primary", use_container_width=True):
                    with st.spinner("Loading input file..."):
                        if load_input_file(input_path):
                            st.session_state.input_file_name = selected_input
                            st.success(f"Loaded: {selected_input}")
        else:
            st.warning("No input files found in Input folder")

        st.markdown("---")

        # Option 2: Upload file
        st.markdown("**Or upload a file:**")
        uploaded_file = st.file_uploader(
            "Upload CDD Input File",
            type=['xlsx', 'xls'],
            key="input_uploader"
        )

        if uploaded_file:
            # Save uploaded file temporarily
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                tmp.write(uploaded_file.getvalue())
                tmp_path = tmp.name

            if st.button("Load Uploaded File", use_container_width=True):
                with st.spinner("Loading uploaded file..."):
                    if load_input_file(tmp_path):
                        st.session_state.input_file_name = uploaded_file.name
                        st.success(f"Loaded: {uploaded_file.name}")

    # Right column: Template & Process Section
    with col2:
        st.subheader("2. Select Template")

        templates = get_available_templates()
        if templates:
            selected_template = st.selectbox(
                "Select SPU Template:",
                options=["-- Select a template --"] + templates,
                key="template_select"
            )

            if selected_template and selected_template != "-- Select a template --":
                template_path = os.path.join(get_template_folder(), selected_template)
                if st.button("Load Template", use_container_width=True):
                    try:
                        st.session_state.processor.set_template(template_path)
                        st.session_state.template_file_name = selected_template
                        st.success(f"Template loaded: {selected_template}")
                    except Exception as e:
                        st.error(f"Failed to load template: {e}")
        else:
            st.warning("No template files found in Template folder")

        st.markdown("---")

        # Process Section
        st.subheader("3. Process Output")

        # Status display
        status_col1, status_col2 = st.columns(2)
        with status_col1:
            if st.session_state.input_file_name:
                st.success(f"Input: {st.session_state.input_file_name}")
            else:
                st.warning("Input: Not loaded")

        with status_col2:
            if st.session_state.template_file_name:
                st.success(f"Template: {st.session_state.template_file_name}")
            else:
                st.warning("Template: Not loaded")

        # Process button
        can_process = (st.session_state.input_data is not None and
                       st.session_state.template_file_name is not None)

        if st.button("Process SPU Output", type="primary", disabled=not can_process, use_container_width=True):
            with st.spinner("Processing... Please wait"):
                progress_bar = st.progress(0)
                output_files = process_output()
                progress_bar.progress(100)

                if output_files:
                    st.success("Processing complete!")

                    # Show output files
                    for output_file in output_files:
                        st.markdown(f"**Output:** `{os.path.basename(output_file)}`")

                        col_a, col_b = st.columns(2)
                        with col_a:
                            if st.button(f"Open File", key=f"open_{output_file}", use_container_width=True):
                                open_file(output_file)
                        with col_b:
                            if st.button(f"Open Folder", key=f"folder_{output_file}", use_container_width=True):
                                open_folder(os.path.dirname(output_file))

    # Data Preview Section
    st.markdown("---")
    st.subheader("Data Preview")

    if st.session_state.input_data:
        # Create tabs for each sheet
        sheet_names = list(st.session_state.input_data.keys())

        if sheet_names:
            tabs = st.tabs(sheet_names)

            for tab, sheet_name in zip(tabs, sheet_names):
                with tab:
                    df = st.session_state.input_data[sheet_name]
                    if not df.empty:
                        st.dataframe(
                            df.head(500),
                            use_container_width=True,
                            height=400
                        )
                        st.caption(f"Showing first 500 rows of {len(df)} total rows")
                    else:
                        st.info(f"No data in {sheet_name} sheet")
    else:
        st.info("Load an input file to preview data")

    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center; color: #666; font-size: 0.9rem;'>
            Process CDD ZTE Tool | Built with Streamlit
        </div>
        """,
        unsafe_allow_html=True
    )


if __name__ == "__main__":
    main()

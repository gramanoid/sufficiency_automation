#!/bin/bash
# Haleon Budget Sync Tool Launcher

cd "$(dirname "$0")"

# Activate virtual environment
source .venv/bin/activate

# Run Streamlit app
echo "Starting Haleon Budget Sync Tool..."
echo "Open http://localhost:8501 in your browser"
echo ""
streamlit run streamlit_app.py

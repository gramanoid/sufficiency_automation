#!/usr/bin/env python3
"""
Haleon Budget Sufficiency Sync Tool
Streamlit app to sync PPT with Excel data
"""

import streamlit as st
import tempfile
from pathlib import Path
from datetime import datetime

# Import the core sync function
import sys
sys.path.insert(0, str(Path(__file__).parent / 'scripts'))
from update_ppt_from_excel import update_ppt_from_excel, MARKET_ROW_RANGES


def inject_custom_css():
    """Inject custom CSS for dark mode styling"""
    st.markdown("""
    <style>
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        max-width: 900px;
    }

    /* Gradient title */
    .gradient-title {
        background: linear-gradient(90deg, #FF6B9D 0%, #FFB347 50%, #FFEB3B 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.8rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        line-height: 1.2;
    }

    .subtitle {
        color: #888;
        font-size: 1.1rem;
        margin-bottom: 2rem;
    }

    /* Info card */
    .info-card {
        background: #1A1D24;
        border-radius: 12px;
        padding: 1.5rem 2rem;
        margin: 1.5rem 0;
    }

    .card-header {
        color: #4ADE80;
        font-size: 0.85rem;
        font-weight: 600;
        letter-spacing: 2px;
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .card-header::before {
        content: '';
        width: 8px;
        height: 8px;
        background: #4ADE80;
        border-radius: 50%;
    }

    /* Step list */
    .step-list {
        list-style: none;
        padding: 0;
        margin: 0;
    }

    .step-item {
        display: flex;
        align-items: center;
        gap: 1rem;
        margin-bottom: 1rem;
        color: #E5E5E5;
        font-size: 1rem;
    }

    .step-num {
        color: #666;
        font-size: 0.9rem;
        min-width: 20px;
    }

    /* Colored badges */
    .badge {
        display: inline-block;
        padding: 0.35rem 0.75rem;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 500;
    }

    .badge-pink {
        background: rgba(255, 107, 157, 0.15);
        color: #FF6B9D;
    }

    .badge-teal {
        background: rgba(45, 212, 191, 0.15);
        color: #2DD4BF;
    }

    .badge-yellow {
        background: rgba(250, 204, 21, 0.15);
        color: #FACC15;
    }

    .badge-purple {
        background: rgba(168, 85, 247, 0.15);
        color: #A855F7;
    }

    .badge-green {
        background: transparent;
        border: 1px solid #4ADE80;
        color: #4ADE80;
    }

    .badge-blue-filled {
        background: rgba(59, 130, 246, 0.8);
        color: white;
    }

    .badge-pink-filled {
        background: rgba(236, 72, 153, 0.8);
        color: white;
    }

    /* Status badges row */
    .status-badges {
        display: flex;
        gap: 1rem;
        margin: 1.5rem 0;
        flex-wrap: wrap;
    }

    .status-badge {
        padding: 0.5rem 1rem;
        border-radius: 8px;
        font-size: 0.9rem;
        font-weight: 500;
    }

    /* File uploader styling */
    [data-testid="stFileUploader"] {
        background: #1A1D24;
        border-radius: 12px;
        padding: 1rem;
    }

    [data-testid="stFileUploader"] > div {
        border-color: #333 !important;
        border-radius: 8px !important;
    }

    [data-testid="stFileUploader"] label {
        color: #E5E5E5 !important;
    }

    /* Primary button */
    .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #FF6B9D 0%, #FF8E53 100%);
        border: none;
        font-weight: 600;
        font-size: 1rem;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        transition: all 0.3s ease;
    }

    .stButton > button[kind="primary"]:hover {
        opacity: 0.9;
        transform: translateY(-1px);
    }

    /* Success/Info alerts */
    [data-testid="stAlert"] {
        background: #1A1D24;
        border-radius: 8px;
    }

    /* Metrics */
    [data-testid="stMetric"] {
        background: #1A1D24;
        padding: 1rem;
        border-radius: 8px;
    }

    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #4ADE80;
    }

    /* Expander */
    [data-testid="stExpander"] {
        background: #1A1D24;
        border: 1px solid #333;
        border-radius: 8px;
    }

    /* Download buttons */
    .stDownloadButton > button {
        background: #1A1D24;
        border: 1px solid #333;
        border-radius: 8px;
        color: #E5E5E5;
    }

    .stDownloadButton > button:hover {
        border-color: #FF6B9D;
        color: #FF6B9D;
    }

    /* Divider */
    hr {
        border-color: #333 !important;
        margin: 2rem 0 !important;
    }

    /* Footer */
    .footer {
        text-align: center;
        padding: 2rem 1rem;
        color: #666;
        font-size: 0.85rem;
    }
    </style>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="Haleon Budget Sync",
        page_icon="📊",
        layout="centered"
    )

    inject_custom_css()

    # Gradient title
    st.markdown("""
    <div class="gradient-title">Haleon MEA Budget<br>Sufficiency Sync</div>
    <div class="subtitle">Transform Excel budget data into PowerPoint presentations</div>
    """, unsafe_allow_html=True)

    # How to use card
    st.markdown("""
    <div class="info-card">
        <div class="card-header">HOW TO USE</div>
        <div class="step-list">
            <div class="step-item">
                <span class="step-num">1.</span>
                Upload your <span class="badge badge-teal">Excel File</span> with 2026 Sufficiency data
            </div>
            <div class="step-item">
                <span class="step-num">2.</span>
                Upload your <span class="badge badge-yellow">PowerPoint File</span> to update
            </div>
            <div class="step-item">
                <span class="step-num">3.</span>
                Click <span class="badge badge-pink">Sync Data</span> and download your updated deck
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Status badges
    markets = list(MARKET_ROW_RANGES.keys())
    st.markdown(f"""
    <div class="status-badges">
        <span class="status-badge badge-green">✓ {len(markets)} Markets</span>
        <span class="status-badge badge-blue-filled">📊 3 Table Types</span>
        <span class="status-badge badge-pink-filled">⚡ Auto-Labeling ON</span>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # File uploaders
    st.markdown("**Upload Excel File**")
    excel_file = st.file_uploader(
        "Excel file with budget data",
        type=['xlsx', 'xlsm'],
        key='excel',
        label_visibility="collapsed"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name}")

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown("**Upload PowerPoint File**")
    ppt_file = st.file_uploader(
        "PPT file to update",
        type=['pptx'],
        key='ppt',
        label_visibility="collapsed"
    )
    if ppt_file:
        st.success(f"✓ {ppt_file.name}")

    st.markdown("<br>", unsafe_allow_html=True)

    # Sync button
    if excel_file and ppt_file:
        if st.button("🔄 Sync Data", type="primary", use_container_width=True):
            with st.spinner("Syncing data..."):
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        tmpdir = Path(tmpdir)
                        excel_path = tmpdir / excel_file.name
                        ppt_path = tmpdir / ppt_file.name

                        with open(excel_path, 'wb') as f:
                            f.write(excel_file.getvalue())
                        with open(ppt_path, 'wb') as f:
                            f.write(ppt_file.getvalue())

                        result = update_ppt_from_excel(
                            ppt_path=ppt_path,
                            excel_path=excel_path,
                            output_dir=tmpdir
                        )

                        if result['success']:
                            st.success(f"✅ Sync complete! Updated {result['cells_updated']} cells")

                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Cells", result['cells_updated'])
                            with col2:
                                st.metric("Warnings", len(result['warnings']))
                            with col3:
                                st.metric("Changes", len(result['changes']))

                            if result['warnings']:
                                with st.expander(f"⚠️ Warnings ({len(result['warnings'])})"):
                                    for w in result['warnings']:
                                        st.warning(w)

                            if result['changes']:
                                grand_totals = [c for c in result['changes'] if c.get('type') == 'grand_total']
                                brand_by_market = [c for c in result['changes'] if c.get('type') == 'brand_by_market']
                                brand_detail = [c for c in result['changes'] if c.get('type') == 'brand_detail']

                                with st.expander("📝 Changes Summary"):
                                    if grand_totals:
                                        st.markdown("**Grand Totals (Slide 3)**")
                                        for c in grand_totals:
                                            st.markdown(f"- {c['field'].replace('_', ' ').title()}: {c['new_value']}")

                                    if brand_by_market:
                                        st.markdown(f"**Brand-by-Market** ({len(brand_by_market)} updates)")

                                    if brand_detail:
                                        st.markdown(f"**Brand Detail** ({len(brand_detail)} updates)")

                            with open(result['output_ppt'], 'rb') as f:
                                output_data = f.read()
                            output_filename = Path(result['output_ppt']).name

                            st.download_button(
                                label="📥 Download Updated PPT",
                                data=output_data,
                                file_name=output_filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                type="primary",
                                use_container_width=True
                            )

                            with open(result['backup_ppt'], 'rb') as f:
                                backup_data = f.read()
                            backup_filename = Path(result['backup_ppt']).name

                            st.download_button(
                                label="📦 Download Backup",
                                data=backup_data,
                                file_name=backup_filename,
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        else:
                            st.error("Sync failed")
                except Exception as e:
                    st.error(f"Error: {str(e)}")
                    st.exception(e)
    else:
        st.info("👆 Upload both files to enable sync")

    # Footer
    st.markdown(f"""
    <div class="footer">
        Haleon MEA Budget Sufficiency Sync • v1.0.0
    </div>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()

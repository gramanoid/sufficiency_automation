#!/usr/bin/env python3
"""
Haleon Data Sync Tool
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
    """Inject custom CSS for dark mode styling with animations"""
    st.markdown("""
    <style>
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Animations */
    @keyframes fadeInUp {
        from {
            opacity: 0;
            transform: translateY(20px);
        }
        to {
            opacity: 1;
            transform: translateY(0);
        }
    }

    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }

    @keyframes pulse {
        0%, 100% { opacity: 1; }
        50% { opacity: 0.7; }
    }

    @keyframes glow {
        0%, 100% { box-shadow: 0 0 5px rgba(74, 222, 128, 0.3); }
        50% { box-shadow: 0 0 20px rgba(74, 222, 128, 0.6); }
    }

    @keyframes slideIn {
        from {
            opacity: 0;
            transform: translateX(-20px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }

    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        max-width: 900px;
    }

    /* Gradient title with animation */
    .gradient-title {
        background: linear-gradient(90deg, #FF6B9D 0%, #FFB347 25%, #FFEB3B 50%, #FFB347 75%, #FF6B9D 100%);
        background-size: 200% auto;
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-size: 2.8rem;
        font-weight: 700;
        margin-bottom: 0.5rem;
        line-height: 1.2;
        animation: gradientShift 4s ease infinite, fadeInUp 0.8s ease-out;
    }

    .subtitle {
        color: #888;
        font-size: 1.1rem;
        margin-bottom: 2rem;
        animation: fadeInUp 0.8s ease-out 0.2s both;
    }

    /* Info card with animation */
    .info-card {
        background: #1A1D24;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        margin: 1rem 0 0.75rem 0;
        animation: fadeInUp 0.8s ease-out 0.3s both;
        transition: transform 0.3s ease, box-shadow 0.3s ease;
    }

    .info-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.3);
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
        animation: pulse 2s ease-in-out infinite;
    }

    /* Step list with staggered animation */
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
        animation: slideIn 0.5s ease-out both;
    }

    .step-item:nth-child(1) { animation-delay: 0.4s; }
    .step-item:nth-child(2) { animation-delay: 0.5s; }
    .step-item:nth-child(3) { animation-delay: 0.6s; }

    .step-num {
        color: #666;
        font-size: 0.9rem;
        min-width: 20px;
    }

    /* Colored badges with hover effects */
    .badge {
        display: inline-block;
        padding: 0.35rem 0.75rem;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 500;
        transition: all 0.3s ease;
    }

    .badge:hover {
        transform: scale(1.05);
    }

    .badge-pink {
        background: rgba(255, 107, 157, 0.15);
        color: #FF6B9D;
    }

    .badge-pink:hover {
        background: rgba(255, 107, 157, 0.3);
        box-shadow: 0 0 15px rgba(255, 107, 157, 0.3);
    }

    .badge-teal {
        background: rgba(45, 212, 191, 0.15);
        color: #2DD4BF;
    }

    .badge-teal:hover {
        background: rgba(45, 212, 191, 0.3);
        box-shadow: 0 0 15px rgba(45, 212, 191, 0.3);
    }

    .badge-yellow {
        background: rgba(250, 204, 21, 0.15);
        color: #FACC15;
    }

    .badge-yellow:hover {
        background: rgba(250, 204, 21, 0.3);
        box-shadow: 0 0 15px rgba(250, 204, 21, 0.3);
    }

    .badge-purple {
        background: rgba(168, 85, 247, 0.15);
        color: #A855F7;
    }

    .badge-green {
        background: transparent;
        border: 1px solid #4ADE80;
        color: #4ADE80;
        animation: glow 3s ease-in-out infinite;
    }

    .badge-blue-filled {
        background: rgba(59, 130, 246, 0.8);
        color: white;
        transition: all 0.3s ease;
    }

    .badge-blue-filled:hover {
        background: rgba(59, 130, 246, 1);
        transform: scale(1.05);
    }

    .badge-pink-filled {
        background: rgba(236, 72, 153, 0.8);
        color: white;
        transition: all 0.3s ease;
    }

    .badge-pink-filled:hover {
        background: rgba(236, 72, 153, 1);
        transform: scale(1.05);
    }

    /* Status badges row */
    .status-badges {
        display: flex;
        gap: 0.75rem;
        margin: 1rem 0 0 0;
        flex-wrap: wrap;
        animation: fadeInUp 0.8s ease-out 0.5s both;
    }

    .status-badge {
        padding: 0.4rem 0.85rem;
        border-radius: 8px;
        font-size: 0.85rem;
        font-weight: 500;
        cursor: default;
        transition: all 0.3s ease;
    }

    .status-badge:hover {
        transform: translateY(-2px);
    }

    /* How it works section */
    .how-it-works {
        background: linear-gradient(135deg, #1A1D24 0%, #252830 100%);
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        margin: 0.5rem 0 0.75rem 0;
        border: 1px solid #333;
        animation: fadeInUp 0.8s ease-out 0.4s both;
    }

    .how-it-works:hover {
        border-color: #4ADE80;
    }

    .how-header {
        color: #A855F7;
        font-size: 0.85rem;
        font-weight: 600;
        letter-spacing: 2px;
        margin-bottom: 1.25rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }

    .how-header::before {
        content: '⚙️';
        font-size: 1rem;
    }

    .feature-grid {
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
        gap: 1rem;
        margin-top: 1rem;
    }

    .feature-item {
        background: rgba(255, 255, 255, 0.03);
        border-radius: 8px;
        padding: 1rem;
        transition: all 0.3s ease;
        border: 1px solid transparent;
    }

    .feature-item:hover {
        background: rgba(255, 255, 255, 0.06);
        border-color: #333;
        transform: translateY(-2px);
    }

    .feature-icon {
        font-size: 1.5rem;
        margin-bottom: 0.5rem;
    }

    .feature-title {
        color: #E5E5E5;
        font-weight: 600;
        font-size: 0.95rem;
        margin-bottom: 0.25rem;
    }

    .feature-desc {
        color: #888;
        font-size: 0.85rem;
        line-height: 1.4;
    }

    /* File uploader styling */
    [data-testid="stFileUploader"] {
        background: #1A1D24;
        border-radius: 12px;
        padding: 1rem;
        animation: fadeInUp 0.8s ease-out 0.6s both;
        transition: all 0.3s ease;
    }

    [data-testid="stFileUploader"]:hover {
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.2);
    }

    [data-testid="stFileUploader"] > div {
        border-color: #333 !important;
        border-radius: 8px !important;
        transition: border-color 0.3s ease;
    }

    [data-testid="stFileUploader"] > div:hover {
        border-color: #4ADE80 !important;
    }

    [data-testid="stFileUploader"] label {
        color: #E5E5E5 !important;
    }

    /* Primary button with animation */
    .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #FF6B9D 0%, #FF8E53 50%, #FF6B9D 100%);
        background-size: 200% auto;
        border: none;
        font-weight: 600;
        font-size: 1rem;
        padding: 0.75rem 2rem;
        border-radius: 8px;
        transition: all 0.3s ease;
        animation: fadeInUp 0.8s ease-out 0.7s both;
    }

    .stButton > button[kind="primary"]:hover {
        background-position: right center;
        transform: translateY(-2px);
        box-shadow: 0 8px 25px rgba(255, 107, 157, 0.4);
    }

    /* Success/Info alerts */
    [data-testid="stAlert"] {
        background: #1A1D24;
        border-radius: 8px;
        animation: fadeInUp 0.5s ease-out;
    }

    /* Metrics */
    [data-testid="stMetric"] {
        background: #1A1D24;
        padding: 1rem;
        border-radius: 8px;
        transition: all 0.3s ease;
    }

    [data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    }

    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #4ADE80;
    }

    /* Expander */
    [data-testid="stExpander"] {
        background: #1A1D24;
        border: 1px solid #333;
        border-radius: 8px;
        transition: all 0.3s ease;
    }

    [data-testid="stExpander"]:hover {
        border-color: #4ADE80;
    }

    /* Download buttons */
    .stDownloadButton > button {
        background: #1A1D24;
        border: 1px solid #333;
        border-radius: 8px;
        color: #E5E5E5;
        transition: all 0.3s ease;
    }

    .stDownloadButton > button:hover {
        border-color: #FF6B9D;
        color: #FF6B9D;
        transform: translateY(-2px);
        box-shadow: 0 4px 15px rgba(255, 107, 157, 0.2);
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
        animation: fadeInUp 0.8s ease-out 0.8s both;
    }

    /* Spinner animation enhancement */
    .stSpinner > div {
        border-color: #FF6B9D transparent transparent transparent !important;
    }
    </style>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="Haleon Data Sync",
        page_icon="📊",
        layout="centered"
    )

    inject_custom_css()

    # Gradient title
    st.markdown("""
    <div class="gradient-title">Haleon MEA Data<br>Sufficiency Sync</div>
    <div class="subtitle">Transform Excel data into PowerPoint presentations</div>
    """, unsafe_allow_html=True)

    # How to use card with status badges
    markets = list(MARKET_ROW_RANGES.keys())
    st.markdown(f"""
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
        <div class="status-badges">
            <span class="status-badge badge-green">✓ {len(markets)} Markets</span>
            <span class="status-badge badge-blue-filled">📊 3 Table Types</span>
            <span class="status-badge badge-pink-filled">⚡ Auto-Labeling ON</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # How it works explanation
    st.markdown("""
    <div class="how-it-works">
        <div class="how-header">HOW IT WORKS</div>
        <div class="feature-grid">
            <div class="feature-item">
                <div class="feature-icon">📊</div>
                <div class="feature-title">Grand Totals</div>
                <div class="feature-desc">Updates slide 3 with overall summary figures</div>
            </div>
            <div class="feature-item">
                <div class="feature-icon">🏷️</div>
                <div class="feature-title">Brand-by-Market</div>
                <div class="feature-desc">Syncs slides 15-18 with brand performance per market</div>
            </div>
            <div class="feature-item">
                <div class="feature-icon">📋</div>
                <div class="feature-title">Brand Details</div>
                <div class="feature-desc">Updates market-specific detail tables (slides 22+)</div>
            </div>
            <div class="feature-item">
                <div class="feature-icon">✨</div>
                <div class="feature-title">Auto-Labeling</div>
                <div class="feature-desc">Adds "SYNCED" badges to all updated slides</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # File uploaders
    st.markdown("**Upload Excel File**")
    excel_file = st.file_uploader(
        "Excel file with data",
        type=['xlsx', 'xlsm'],
        key='excel',
        label_visibility="collapsed"
    )
    if excel_file:
        st.success(f"✓ {excel_file.name}")

    st.markdown("**Upload PowerPoint File**")
    ppt_file = st.file_uploader(
        "PPT file to update",
        type=['pptx'],
        key='ppt',
        label_visibility="collapsed"
    )
    if ppt_file:
        st.success(f"✓ {ppt_file.name}")

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
        Haleon MEA Data Sufficiency Sync • v1.0.0
    </div>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()

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

# Haleon brand colors
HALEON_TEAL = "#00857C"
HALEON_TEAL_LIGHT = "#E6F4F3"
HALEON_TEAL_DARK = "#006B64"
HALEON_ORANGE = "#FF6B35"

def inject_custom_css():
    """Inject custom CSS for Haleon branding"""
    st.markdown("""
    <style>
    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        max-width: 1200px;
    }
    
    /* Subtle background pattern */
    .main {
        background: linear-gradient(180deg, #F8FAFA 0%, #FFFFFF 100%);
    }
    
    /* Header styling */
    .hero-header {
        background: linear-gradient(135deg, #00857C 0%, #006B64 50%, #005550 100%);
        padding: 2.5rem 2rem;
        border-radius: 20px;
        margin-bottom: 2rem;
        box-shadow: 0 8px 32px rgba(0, 133, 124, 0.25);
        position: relative;
        overflow: hidden;
    }
    
    .hero-header::before {
        content: '';
        position: absolute;
        top: -50%;
        right: -30%;
        width: 80%;
        height: 200%;
        background: radial-gradient(circle, rgba(255,255,255,0.08) 0%, transparent 50%);
        pointer-events: none;
        animation: shimmer 8s ease-in-out infinite;
    }
    
    @keyframes shimmer {
        0%, 100% { transform: translateX(0) translateY(0); }
        50% { transform: translateX(10%) translateY(-5%); }
    }
    
    .hero-header h1 {
        color: white !important;
        font-size: 2.1rem !important;
        font-weight: 700 !important;
        margin-bottom: 0.5rem !important;
        position: relative;
    }
    
    .hero-header p {
        color: rgba(255, 255, 255, 0.9);
        font-size: 1.05rem;
        margin: 0;
        position: relative;
    }
    
    /* Step cards container */
    .step-cards {
        display: flex;
        align-items: stretch;
        gap: 1rem;
        margin-bottom: 1.5rem;
    }
    
    /* Upload cards */
    .upload-card {
        background: white;
        border: 2px solid #E6F4F3;
        border-radius: 16px;
        padding: 1.5rem;
        transition: all 0.3s ease;
        box-shadow: 0 2px 12px rgba(0, 0, 0, 0.04);
        position: relative;
    }
    
    .upload-card:hover {
        border-color: #00857C;
        box-shadow: 0 8px 24px rgba(0, 133, 124, 0.15);
        transform: translateY(-2px);
    }
    
    .upload-card .step-badge {
        position: absolute;
        top: -12px;
        left: 20px;
        background: linear-gradient(135deg, #00857C 0%, #006B64 100%);
        color: white;
        width: 28px;
        height: 28px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-weight: 700;
        font-size: 0.9rem;
        box-shadow: 0 2px 8px rgba(0, 133, 124, 0.3);
    }
    
    .upload-card h3 {
        color: #00857C !important;
        font-size: 1.2rem !important;
        font-weight: 600 !important;
        margin-bottom: 0.25rem !important;
        margin-top: 0.5rem !important;
    }
    
    .upload-card .file-type {
        display: inline-block;
        background: #E6F4F3;
        color: #006B64;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-top: 0.5rem;
    }
    
    /* Flow arrow */
    .flow-arrow {
        display: flex;
        align-items: center;
        justify-content: center;
        color: #00857C;
        font-size: 1.5rem;
        padding: 0 0.5rem;
    }
    
    /* File uploader styling */
    [data-testid="stFileUploader"] {
        border-radius: 12px;
    }
    
    [data-testid="stFileUploader"] > div {
        border-color: #E6F4F3 !important;
        border-radius: 12px !important;
    }
    
    [data-testid="stFileUploader"] > div:hover {
        border-color: #00857C !important;
    }
    
    /* Primary button */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #00857C 0%, #006B64 100%);
        border: none;
        font-weight: 600;
        font-size: 1.1rem;
        padding: 0.875rem 2rem;
        border-radius: 12px;
        transition: all 0.3s ease;
        box-shadow: 0 4px 12px rgba(0, 133, 124, 0.2);
    }
    
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #006B64 0%, #005550 100%);
        box-shadow: 0 6px 20px rgba(0, 133, 124, 0.35);
        transform: translateY(-2px);
    }
    
    /* Info box */
    [data-testid="stAlert"] {
        background-color: #E6F4F3;
        border-left: 4px solid #00857C;
        border-radius: 12px;
        padding: 1rem 1.25rem;
    }
    
    /* Success message */
    [data-testid="stAlert"][data-baseweb="notification"] {
        border-radius: 12px;
    }
    
    /* Metrics */
    [data-testid="stMetric"] {
        background: white;
        padding: 1.25rem;
        border-radius: 12px;
        border: 1px solid #E6F4F3;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    
    [data-testid="stMetric"] [data-testid="stMetricValue"] {
        color: #00857C;
    }
    
    /* Expander */
    [data-testid="stExpander"] {
        border: 1px solid #E6F4F3;
        border-radius: 12px;
        background: white;
        box-shadow: 0 2px 8px rgba(0, 0, 0, 0.04);
    }
    
    [data-testid="stExpander"] summary {
        font-weight: 600;
    }
    
    /* Footer */
    .footer {
        text-align: center;
        padding: 1.75rem 1.5rem;
        color: #666;
        border-top: 2px solid #E6F4F3;
        margin-top: 2.5rem;
        background: linear-gradient(180deg, #FAFBFB 0%, #FFFFFF 100%);
        border-radius: 16px 16px 0 0;
    }
    
    /* Divider */
    hr {
        border-color: #E6F4F3 !important;
        margin: 2rem 0 !important;
    }
    
    /* Download buttons */
    .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 500;
        transition: all 0.2s ease;
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-1px);
    }
    
    /* Success alerts */
    [data-testid="stAlert"] {
        animation: slideIn 0.3s ease-out;
    }
    
    @keyframes slideIn {
        from { opacity: 0; transform: translateY(-10px); }
        to { opacity: 1; transform: translateY(0); }
    }
    
    /* Hide Streamlit branding */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)


def main():
    st.set_page_config(
        page_title="Haleon Budget Sync",
        page_icon="📊",
        layout="wide"
    )
    
    inject_custom_css()

    # Hero header
    st.markdown("""
    <div class="hero-header">
        <div style="display: inline-block; background: rgba(255,255,255,0.2); padding: 0.35rem 0.85rem; border-radius: 20px; font-size: 0.75rem; color: white; margin-bottom: 1rem; font-weight: 500; letter-spacing: 0.5px;">
            ✨ 2026 Budget Planning Tool
        </div>
        <h1>📊 Haleon MEA Budget Sufficiency Sync</h1>
        <p>Sync PowerPoint tables with Excel data while preserving all formatting</p>
        <div style="display: flex; gap: 1.5rem; margin-top: 1.25rem;">
            <div style="display: flex; align-items: center; gap: 0.4rem; color: rgba(255,255,255,0.85); font-size: 0.85rem;">
                <span style="font-size: 1rem;">🎯</span> 10 Markets
            </div>
            <div style="display: flex; align-items: center; gap: 0.4rem; color: rgba(255,255,255,0.85); font-size: 0.85rem;">
                <span style="font-size: 1rem;">📄</span> 14 Slides
            </div>
            <div style="display: flex; align-items: center; gap: 0.4rem; color: rgba(255,255,255,0.85); font-size: 0.85rem;">
                <span style="font-size: 1rem;">⚡</span> Instant Sync
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # File inputs
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("""
        <div class="upload-card">
            <div class="step-badge">1</div>
            <h3>📗 Excel File</h3>
            <p style="color: #666; margin: 0.25rem 0 0.75rem 0; font-size: 0.9rem;">Source data with budget values</p>
            <span class="file-type">XLSX / XLSM</span>
        </div>
        """, unsafe_allow_html=True)
        excel_file = st.file_uploader(
            "Upload Excel file with budget data",
            type=['xlsx', 'xlsm'],
            key='excel',
            label_visibility="collapsed"
        )
        if excel_file:
            st.success(f"✓ {excel_file.name}")

    with col2:
        st.markdown("""
        <div class="upload-card">
            <div class="step-badge">2</div>
            <h3>📙 PowerPoint File</h3>
            <p style="color: #666; margin: 0.25rem 0 0.75rem 0; font-size: 0.9rem;">Target presentation to update</p>
            <span class="file-type">PPTX</span>
        </div>
        """, unsafe_allow_html=True)
        ppt_file = st.file_uploader(
            "Upload PPT file to update",
            type=['pptx'],
            key='ppt',
            label_visibility="collapsed"
        )
        if ppt_file:
            st.success(f"✓ {ppt_file.name}")

    st.markdown("<br>", unsafe_allow_html=True)

    # Step 3: Sync section
    st.markdown("""
    <div style="text-align: center; margin: 1rem 0;">
        <div style="display: inline-flex; align-items: center; gap: 0.5rem; color: #00857C; font-size: 0.9rem; font-weight: 600;">
            <span style="background: linear-gradient(135deg, #00857C 0%, #006B64 100%); color: white; width: 24px; height: 24px; border-radius: 50%; display: inline-flex; align-items: center; justify-content: center; font-size: 0.8rem;">3</span>
            Click to sync
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Sync button
    if excel_file and ppt_file:
        if st.button("🔄 Sync Excel → PPT", type="primary", use_container_width=True):
            with st.spinner("Syncing data..."):
                try:
                    # Create temp directory for processing
                    with tempfile.TemporaryDirectory() as tmpdir:
                        tmpdir = Path(tmpdir)

                        # Save uploaded files
                        excel_path = tmpdir / excel_file.name
                        ppt_path = tmpdir / ppt_file.name

                        with open(excel_path, 'wb') as f:
                            f.write(excel_file.getvalue())

                        with open(ppt_path, 'wb') as f:
                            f.write(ppt_file.getvalue())

                        # Run sync
                        result = update_ppt_from_excel(
                            ppt_path=ppt_path,
                            excel_path=excel_path,
                            output_dir=tmpdir
                        )

                        if result['success']:
                            st.success(f"✅ Sync complete! Updated {result['cells_updated']} cells")

                            # Show summary
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Cells Updated", result['cells_updated'])
                            with col2:
                                st.metric("Warnings", len(result['warnings']))
                            with col3:
                                st.metric("Changes", len(result['changes']))

                            # Warnings
                            if result['warnings']:
                                with st.expander(f"⚠️ Warnings ({len(result['warnings'])})"):
                                    for w in result['warnings']:
                                        st.warning(w)

                            # Changes breakdown
                            if result['changes']:
                                # Group by type
                                grand_totals = [c for c in result['changes'] if c.get('type') == 'grand_total']
                                brand_by_market = [c for c in result['changes'] if c.get('type') == 'brand_by_market']
                                brand_detail = [c for c in result['changes'] if c.get('type') == 'brand_detail']
                                
                                with st.expander("📝 Changes Summary"):
                                    # Grand totals
                                    if grand_totals:
                                        st.markdown("**Grand Totals (Slide 3)**")
                                        for c in grand_totals:
                                            st.markdown(f"- {c['field'].replace('_', ' ').title()}: {c['new_value']}")
                                    
                                    # Brand-by-market tables
                                    if brand_by_market:
                                        st.markdown(f"**Brand-by-Market Tables** ({len(brand_by_market)} updates)")
                                        brands_updated = set(c['brand'] for c in brand_by_market)
                                        st.caption(f"Brands: {', '.join(sorted(brands_updated))}")
                                    
                                    # Brand detail tables
                                    if brand_detail:
                                        st.markdown(f"**Brand Detail Tables** ({len(brand_detail)} updates)")
                                        changes_by_market = {}
                                        for c in brand_detail:
                                            market = c.get('market', 'Unknown')
                                            if market not in changes_by_market:
                                                changes_by_market[market] = []
                                            changes_by_market[market].append(c)

                                        for market, changes in sorted(changes_by_market.items()):
                                            brands = set(c['brand'] for c in changes)
                                            st.markdown(f"- **{market}**: {len(changes)} changes ({', '.join(sorted(brands))})")

                            # Download button
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

                            # Also offer backup
                            with open(result['backup_ppt'], 'rb') as f:
                                backup_data = f.read()

                            backup_filename = Path(result['backup_ppt']).name

                            st.download_button(
                                label="📦 Download Backup (Original PPT)",
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
        st.info("👆 Upload both Excel and PowerPoint files to enable sync")

    # Info section
    st.divider()
    with st.expander("ℹ️ How it works", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 📍 Slides Updated
            | Slide | Content |
            |-------|---------|
            | **3** | Grand totals (Budget & Sufficient) |
            | **15-18** | Brand summary by market |
            | **22+** | Brand detail tables |
            
            ### 🔄 Data Synced
            - Budget amounts (2026 Budget, Sufficient, Gap)
            - Percentages (AWA, CON, PUR, TV, Digital, Others)
            - Campaign counts (Long, Short campaigns)
            """)
        
        with col2:
            st.markdown("""
            ### ✨ What's Preserved
            - All PPT formatting (fonts, colors, borders)
            - Cell sizes and layouts
            - Non-table content
            - Slide structure
            
            ### 🌍 Markets Covered
            """)
            # Market badges
            markets = list(MARKET_ROW_RANGES.keys())
            market_html = ' '.join([f'<span style="display: inline-block; background: #E6F4F3; color: #006B64; padding: 0.25rem 0.6rem; border-radius: 15px; font-size: 0.75rem; font-weight: 500; margin: 0.15rem;">{m}</span>' for m in markets])
            st.markdown(f'<div style="margin-top: 0.5rem;">{market_html}</div>', unsafe_allow_html=True)

    # Footer
    st.markdown(f"""
    <div class="footer">
        <div style="display: flex; justify-content: center; align-items: center; gap: 2rem; flex-wrap: wrap;">
            <div>
                <strong style="color: #00857C;">Haleon MEA Budget Sufficiency Sync</strong>
            </div>
            <div style="color: #999; font-size: 0.8rem;">
                v1.0.0 • {datetime.now().strftime('%Y')}
            </div>
        </div>
        <div style="margin-top: 0.75rem; font-size: 0.75rem; color: #aaa;">
            Preserves all PowerPoint formatting while syncing Excel data
        </div>
    </div>
    """, unsafe_allow_html=True)


if __name__ == '__main__':
    main()

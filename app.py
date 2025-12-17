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


def main():
    st.set_page_config(
        page_title="Haleon Budget Sync",
        page_icon="📊",
        layout="wide"
    )

    st.title("📊 Haleon MEA Budget Sufficiency Sync")
    st.markdown("Sync PowerPoint tables with Excel data while preserving formatting.")

    st.divider()

    # File inputs
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📗 Excel File (Source)")
        excel_file = st.file_uploader(
            "Upload Excel file with budget data",
            type=['xlsx', 'xlsm'],
            key='excel'
        )
        if excel_file:
            st.success(f"✓ {excel_file.name}")

    with col2:
        st.subheader("📙 PowerPoint File (Target)")
        ppt_file = st.file_uploader(
            "Upload PPT file to update",
            type=['pptx'],
            key='ppt'
        )
        if ppt_file:
            st.success(f"✓ {ppt_file.name}")

    st.divider()

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
    with st.expander("ℹ️ How it works"):
        st.markdown("""
        ### Slides Updated
        - **Slide 3**: Grand total boxes (Briefed Budget, 100% Sufficient)
        - **Slides 15-18**: Brand-by-market summary tables (Sensodyne, Parodontax, Panadol, Centrum)
        - **Slides 22+**: Brand detail tables per market
        
        ### Data Flow
        1. **Excel** contains the source budget data (2026 Sufficiency sheet)
        2. **PowerPoint** contains the presentation with data tables
        3. This tool updates PPT table cells to match Excel values

        ### What gets synced
        - Budget amounts (2026 Budget, Sufficient, Gap)
        - Percentages (Gap %, AWA, CON, PUR, TV, Digital, Others, Long %)
        - Campaign counts (Long Campaigns, Short Campaigns)

        ### What's preserved
        - All PPT formatting (fonts, colors, borders, cell sizes)
        - Slide layouts and structure
        - Non-table content

        ### Markets covered
        """)
        markets = list(MARKET_ROW_RANGES.keys())
        cols = st.columns(5)
        for i, market in enumerate(markets):
            cols[i % 5].markdown(f"- {market}")

    # Footer
    st.divider()
    st.caption(f"Haleon MEA Budget Sufficiency Sync Tool | {datetime.now().strftime('%Y-%m-%d')}")


if __name__ == '__main__':
    main()

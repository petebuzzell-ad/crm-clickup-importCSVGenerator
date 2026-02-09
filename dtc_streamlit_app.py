#!/usr/bin/env python3
"""
DTC Calendar to ClickUp CSV Converter - Streamlit Web App

Converts DTC Calendar Excel files for Arcadia Digital brands (PB, TGW) into ClickUp-importable CSV format.

================================================================================
INTELLECTUAL PROPERTY NOTICE
================================================================================

This tool is part of Arcadia Digital's proprietary operational infrastructure.

This material is confidential and may not be shared, reproduced, or used 
outside of Arcadia Digital without explicit authorization.

Unauthorized distribution, modification, or use of this software is prohibited.

© 2026 Arcadia Digital. All rights reserved.

================================================================================
"""

import streamlit as st
import tempfile
import os
from pathlib import Path
from dtc_to_clickup import DTCtoClickUpConverter

# Page config
st.set_page_config(
    page_title="DTC → ClickUp Converter",
    page_icon="📋",
    layout="centered"
)

# Header
st.title("📋 DTC Calendar → ClickUp CSV Converter")
st.markdown("Convert your DTC Calendar Excel files into ClickUp-importable CSV format")
st.markdown("---")

# Instructions
with st.expander("📖 How to Use"):
    st.markdown("""
    **Step 1:** Upload your DTC Calendar Excel file (contains Wk# sheets with email briefs)
    
    **Step 2:** Select your brand (PB or TGW)
    
    **Step 3:** Click "Convert to ClickUp CSV"
    
    **Step 4:** Download the generated CSV file
    
    **What this tool does:**
    - Extracts email brief tasks from weekly sheets (Wk6, Wk7, Wk8, etc.)
    - Converts campaign data into ClickUp-compatible CSV format
    - Includes task names, descriptions, dates, priorities, and tags
    - Handles both PB and TGW calendar formats
    
    **What gets extracted:**
    - Campaign Type & Name
    - Email Overview & Copy Requirements
    - Send Date & Time
    - Priority (based on campaign type)
    - Featured Products & URLs
    - DAM Assets & Hero Images
    - Landing Page URLs
    - SMS briefs (when applicable)
    """)

# File upload
uploaded_file = st.file_uploader(
    "Upload DTC Calendar Excel File",
    type=['xlsx', 'xls'],
    help="Upload your DTC Calendar Excel file containing weekly email brief sheets"
)

# Brand selection
brand = st.radio(
    "Select Brand",
    options=["PB", "TGW"],
    horizontal=True,
    help="Choose the brand for this calendar (determines tagging and formatting)"
)

# Week selection (only show after file is uploaded)
selected_weeks = None
if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    # Preview available weeks by loading workbook temporarily
    with tempfile.TemporaryDirectory() as temp_dir:
        # Save uploaded file to temp location
        input_path = Path(temp_dir) / uploaded_file.name
        with open(input_path, 'wb') as f:
            f.write(uploaded_file.getbuffer())
        
        # Create converter just to get available weeks
        try:
            preview_converter = DTCtoClickUpConverter(
                excel_file=str(input_path),
                brand=brand,
                output_file=str(Path(temp_dir) / "temp.csv")
            )
            
            if preview_converter.load_workbook_safe():
                available_weeks = preview_converter.get_available_weekly_sheets()
                
                if available_weeks:
                    st.markdown("### Select Weeks to Import")
                    st.markdown("**Tip:** Only select new weeks to avoid creating duplicate tasks in ClickUp")
                    
                    # Option to select all or specific weeks
                    import_option = st.radio(
                        "Import option:",
                        options=["Import all weeks", "Select specific weeks"],
                        horizontal=True,
                        help="Choose whether to import all weeks or select specific ones"
                    )
                    
                    if import_option == "Select specific weeks":
                        selected_weeks = st.multiselect(
                            "Choose weeks to import:",
                            options=available_weeks,
                            default=available_weeks,
                            help="Select one or more weeks. Deselect weeks that are already in ClickUp to avoid duplicates."
                        )
                        
                        if not selected_weeks:
                            st.warning("⚠️ Please select at least one week to import")
                    else:
                        selected_weeks = None  # None means import all
                        st.info(f"📋 Will import all {len(available_weeks)} weeks: {', '.join(available_weeks)}")
                else:
                    st.error("❌ No weekly sheets found in this file (looking for sheets named Wk6, Wk7, etc.)")
        except Exception as e:
            st.error(f"❌ Error reading Excel file: {str(e)}")

# Convert button
if uploaded_file is not None:
    # Only show button if weeks are selected (or all weeks option chosen)
    can_proceed = True
    
    # If user chose "Select specific weeks" but didn't select any, block conversion
    if selected_weeks is not None and len(selected_weeks) == 0:
        can_proceed = False
        st.warning("⚠️ Please select at least one week to import")
    
    if can_proceed:
        if st.button("🚀 Convert to ClickUp CSV", type="primary"):
            
            with st.spinner("Converting your DTC Calendar..."):
                
                # Create temporary directory for processing
                with tempfile.TemporaryDirectory() as temp_dir:
                    
                    # Save uploaded file to temp location
                    input_path = Path(temp_dir) / uploaded_file.name
                    with open(input_path, 'wb') as f:
                        f.write(uploaded_file.getbuffer())
                    
                    # Define output path
                    output_filename = f"{brand}_ClickUp_Import.csv"
                    output_path = Path(temp_dir) / output_filename
                    
                    # Run conversion
                    try:
                        converter = DTCtoClickUpConverter(
                            excel_file=str(input_path),
                            brand=brand,
                            output_file=str(output_path)
                        )
                        
                        # Pass selected_weeks to convert method
                        success = converter.convert(selected_weeks=selected_weeks)
                        
                        if success:
                            # Display summary
                            st.success("✅ Conversion Complete!")
                            
                            # Show stats in columns
                            col1, col2, col3 = st.columns(3)
                            with col1:
                                st.metric("Email Briefs", converter.stats['campaign_tasks'])
                            with col2:
                                st.metric("Sheets Processed", converter.stats['sheets_processed'])
                            with col3:
                                st.metric("Total Tasks", len(converter.tasks))
                            
                            # Read the generated CSV for download
                            with open(output_path, 'rb') as f:
                                csv_data = f.read()
                            
                            # Download button
                            st.download_button(
                                label="⬇️ Download ClickUp CSV",
                                data=csv_data,
                                file_name=output_filename,
                                mime="text/csv",
                                type="primary"
                            )
                            
                            # Show preview of tasks
                            with st.expander("👀 Preview Tasks"):
                                st.markdown(f"**First 5 tasks from {len(converter.tasks)} total:**")
                                for i, task in enumerate(converter.tasks[:5], 1):
                                    st.markdown(f"**{i}. {task['Task Name']}**")
                                    st.text(f"Due: {task['Due Date']} | Priority: {task['Priority']}")
                                    st.text(f"Tags: {task['Tags']}")
                                    # Show truncated description directly (no nested expander)
                                    desc = task['Task Description']
                                    if len(desc) > 200:
                                        st.text(desc[:200] + "...")
                                    else:
                                        st.text(desc)
                                    st.markdown("---")
                        
                        else:
                            st.error("❌ Conversion failed. Check your Excel file format.")
                            st.info("Make sure your file contains weekly sheets (Wk6, Wk7, etc.) with email brief data.")
                    
                    except Exception as e:
                        st.error(f"❌ Error during conversion: {str(e)}")
                        st.info("If this error persists, contact your Arcadia Digital team.")

else:
    st.info("👆 Upload an Excel file to get started")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.9em;'>
    <p><strong>Built for Arcadia Digital | DTC Operations Team</strong></p>
    <p>Questions? Contact your team lead or Arcadia AI support</p>
    <p style='margin-top: 1em; font-size: 0.85em; color: #999;'>
        <strong>Proprietary & Confidential</strong><br>
        This tool is part of Arcadia Digital's proprietary operational infrastructure.<br>
        Unauthorized distribution, modification, or use is prohibited.<br>
        © 2026 Arcadia Digital. All rights reserved.
    </p>
</div>
""", unsafe_allow_html=True)

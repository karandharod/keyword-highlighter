import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import re
import streamlit as st

# Function to process keywords and update the data
def add_keywords_to_columns(data_file, keyword_file, output_file):
    # Load the data and keyword files
    data_df = pd.read_excel(data_file)
    keywords_df = pd.read_excel(keyword_file)
    
    # Extract keywords as a list of lowercase strings
    keywords = keywords_df.iloc[:, 0].dropna().str.strip().str.lower().tolist()
    
    # Add new columns for each keyword at the end of the data file
    keyword_columns = [f"Keyword: {kw}" for kw in keywords]
    for col_name in keyword_columns:
        data_df[col_name] = "No"  # Initialize with "No"
    
    # Check each cell for keywords and mark "Yes" in keyword columns
    for row in range(len(data_df)):  # Process each row in the DataFrame
        for col in data_df.columns[:-len(keywords)]:  # Loop through original columns
            cell_value = str(data_df.at[row, col]).lower() if pd.notnull(data_df.at[row, col]) else ""
            for idx, keyword in enumerate(keywords):
                if re.search(rf"\b{re.escape(keyword)}\b", cell_value):
                    data_df.at[row, f"Keyword: {keyword}"] = "Yes"

    # Save the updated file
    data_df.to_excel(output_file, index=False)
    return output_file

# Streamlit UI
st.title("Keyword Column Updater Tool")

# File uploaders
data_file = st.file_uploader("Upload your data file (Excel format)", type=["xlsx"], key="data_file")
keyword_file = st.file_uploader("Upload your keyword file (Excel format)", type=["xlsx"], key="keyword_file")

if data_file and keyword_file:
    with st.spinner("Processing your files..."):
        output_file = "updated_data.xlsx"  # Define output file name
        result_path = add_keywords_to_columns(data_file, keyword_file, output_file)
    st.success("File processed successfully!")

    # Provide the file for download
    with open(result_path, "rb") as file:
        st.download_button(label="Download Processed File", data=file, file_name="updated_data.xlsx")


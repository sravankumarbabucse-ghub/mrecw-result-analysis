import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="MRECW Result Analysis", layout="wide")

# --- Categorization Logic Functions ---

def categorize_daily(score):
    if str(score).strip().upper() == "AB": return "AB"
    try:
        score = float(score)
        if score == 300: return "300"
        elif 200 <= score < 300: return "299-200"
        elif 100 <= score < 200: return "200-100"
        elif 10 <= score < 100: return "100-10"
        elif score == 0: return "0"
        else: return "Others"
    except: return "Others"

def categorize_monday(score):
    if str(score).strip().upper() == "AB": return "AB"
    try:
        score = float(score)
        if score == 600: return "600"
        elif 400 <= score < 600: return "400-599"
        elif 200 <= score < 400: return "200-399"
        elif 10 <= score < 200: return "10-199"
        elif score == 0: return "0"
        else: return "Others"
    except: return "AB"

def categorize_wednesday(grade):
    try:
        grade = float(grade)
        if grade >= 1500: return ">1500"
        elif 1000 <= grade < 1500: return "1000-1500"
        elif 500 <= grade < 1000: return "500-999"
        elif 100 <= grade < 500: return "100-499"
        elif 1 <= grade < 100: return "1-99"
        else: return "0"
    except: return "0"

# --- Main App Interface ---

st.title("📊 MALLA REDDY ENGINEERING COLLEGE FOR WOMEN")
st.markdown("### Code Chef Result Analysis Portal")
st.markdown("### DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING")

# Dropdown list placed after the Department heading
contest_type = st.selectbox(
    "Select Assessment Type:",
    ["Daily Assessment", "Monday Contest", "Wednesday Contest"],
    index=0
)

st.write("---") 
st.info(f"Please upload the Excel file for **{contest_type}** below.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # 1. Router logic
        if contest_type == "Daily Assessment":
            df["Category"] = df["User Score"].apply(categorize_daily)
            score_cols = ["300", "299-200", "200-100", "100-10", "0", "AB"]
            
        elif contest_type == "Monday Contest":
            df["Category"] = df["User Score"].apply(categorize_monday)
            score_cols = ["600", "400-599", "200-399", "10-199", "0", "AB"]
            
        elif contest_type == "Wednesday Contest":
            df["Category"] = df["Grade"].apply(categorize_wednesday)
            score_cols = [">1500", "1000-1500", "500-999", "100-499", "1-99", "0"]

        # 2. Process Table
        result = pd.crosstab(df["Section"], df["Category"])
        result = result.reindex(columns=score_cols, fill_value=0)
        
        # 3. Stats & Restoration of Remarks
        if contest_type == "Wednesday Contest":
            abs_counts = df[df['Attempted/Not'] == 'AB'].groupby('Section').size()
            result['Absentees'] = abs_counts.reindex(result.index, fill_value=0)
            result['Attended Count'] = result[score_cols].sum(axis=1)
        else:
            result['Absentees'] = result.get('AB', 0)
            numeric_cols = [c for c in score_cols if c != "AB"]
            result['Attended Count'] = result[numeric_cols].sum(axis=1)

        result["Total Strength"] = result["Attended Count"] + result["Absentees"]
        
        # RESTORED REMARKS: These are the "last 2 lines" in your report format
        result['Remark 1'] = result['0'].apply(lambda x: f"{x} Students got 0")
        result['Remark 2'] = result['Absentees'].apply(lambda x: f"{x} Students not written exam")
        
        # 4. Final Display Formatting
        result_final = result.reset_index()
        result_final.insert(0, 'S.No', range(1, 1 + len(result_final)))
        
        st.success(f"{contest_type} Analysis Complete!")
        st.dataframe(result_final, use_container_width=True)

        # 5. Excel Export
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_final.to_excel(writer, index=False, startrow=8, sheet_name='Analysis')
            workbook = writer.book
            worksheet = writer.sheets['Analysis']
            
            title_format = workbook.add_format({'align': 'center', 'bold': True, 'font_size': 11})
            header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3', 'align': 'center'})
            
            titles = [
                "MALLA REDDY ENGINEERING COLLEGE FOR WOMEN",
                "(Autonomous Institution-UGC, Govt. of India)",
                "Accredited by NAAC with ‘A+’ Grade | Programmes Accredited by NBA",
                "National Ranking by NIRF Innovation – Rank band(151-300), MHRD, Govt. of India",
                "Approved by AICTE, Affiliated to JNTUH, ISO 9001:2015 Certified Institution.",
                "Maisammaguda, Dhulapally, Secunderabad 500100.",
                "DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING",
                f"{contest_type.upper()} RESULT ANALYSIS DATA"
            ]
            
            # Merge header across all columns including Remarks
            last_col = len(result_final.columns) - 1
            for i, title in enumerate(titles):
                worksheet.merge_range(i, 0, i, last_col, title, title_format)

            for col_num, value in enumerate(result_final.columns.values):
                worksheet.write(8, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)

        st.download_button(
            label="📥 Download Formatted Report",
            data=output.getvalue(),
            file_name=f"{contest_type.replace(' ', '_')}_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}. Ensure columns match the {contest_type} format.")

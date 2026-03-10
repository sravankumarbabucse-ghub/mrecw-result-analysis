import streamlit as st
import pandas as pd
import io

# Set page config
st.set_page_config(page_title="MRECW Result Analysis", layout="wide")

def categorize_score(score):
    if str(score).strip().upper() == "AB":
        return "AB"
    try:
        score = float(score)
        if score == 300: return "300"
        elif 200 <= score < 300: return "299-200"
        elif 100 <= score < 200: return "200-100"
        elif 10 <= score < 100: return "100-10"
        elif score == 0: return "0"
        else: return "Others"
    except:
        return "Others"

st.title("MALLA REDDY ENGINEERING COLLEGE FOR WOMEN")
st.markdown("### DEPARTMENT OF COMPUTER SCIENCE AND ENINEERING")
st.markdown("📊 Code Chef Result Analysis Portal")
st.info("Upload your department's Excel file to generate the formatted Result Analysis Report.")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

if uploaded_file is not None:
    try:
        # Step 1: Process Data
        df = pd.read_excel(uploaded_file)
        df["Category"] = df["User Score"].apply(categorize_score)
        
        # Create Table
        result = pd.crosstab(df["Section"], df["Category"])
        cols = ["300", "299-200", "200-100", "100-10", "0", "AB"]
        result = result.reindex(columns=cols, fill_value=0)
        
        # Calculate totals
        result["Attended Count"] = result[["300", "299-200", "200-100", "100-10", "0"]].sum(axis=1)
        result["Absentees"] = result["AB"]
        result["Total Strength"] = result["Attended Count"] + result["Absentees"]
        
        # Add Remarks
        result["Remark 1"] = result["0"].apply(lambda x: f"{x} Students got 0" if x > 0 else "")
        result["Remark 2"] = result["Absentees"].apply(lambda x: f"{x} Students not written exam" if x > 0 else "")
        
        # Reset index for display
        result_final = result.reset_index()
        result_final.insert(0, 'S.No', range(1, 1 + len(result_final)))
        
        st.success("Analysis Complete!")
        st.dataframe(result_final, use_container_width=True)

        # Step 2: Excel Export with Formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_final.to_excel(writer, index=False, startrow=8, sheet_name='Sheet1')
            
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Formats
            title_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True, 'font_size': 12})
            header_format = workbook.add_format({'bold': True, 'border': 1, 'bg_color': '#D3D3D3', 'align': 'center'})
            
            # Titles
            titles = [
                "MALLA REDDY ENGINEERING COLLEGE FOR WOMEN",
                "(Autonomous Institution-UGC, Govt. of India)",
                "Accredited by NAAC with ‘A+’ Grade | Programmes Accredited by NBA",
                "National Ranking by NIRF Innovation – Rank band(151-300), MHRD, Govt. of India",
                "Approved by AICTE, Affiliated to JNTUH, ISO 9001:2015 Certified Institution.",
                "Maisammaguda, Dhulapally, Secunderabad 500100.",
                "DEPARTMENT OF COMPUTER SCIENCE AND ENGINEERING",
                "CODE CHEF RESULT ANALYSIS DATA"
            ]
            
            for i, title in enumerate(titles):
                worksheet.merge_range(i, 0, i, 11, title, title_format)

            # Column formatting
            for col_num, value in enumerate(result_final.columns.values):
                worksheet.write(8, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 15)

        st.download_button(
            label="📥 Download Formatted Report",
            data=output.getvalue(),
            file_name="Result_Analysis_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"Error: {e}. Please ensure the Excel file has 'Section' and 'User Score' columns.")



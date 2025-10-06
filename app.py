#app.py
 
 
 
import streamlit as st
import pandas as pd
import util
 
st.set_page_config(page_title="Project Summary Generator", layout="wide")
st.title("Project Summary Generator")
 
# Accept multiple file types
uploaded_file = st.file_uploader("Upload your data file (.xlsx, .csv, .txt)", type=["xlsx", "csv", "txt"])
 
if uploaded_file is not None:
    file_ready = False
    try:
        if uploaded_file.name.endswith('.xlsx'):
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            file_ready = True
        elif uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            file_ready = True
        elif uploaded_file.name.endswith('.txt'):
            df = pd.read_csv(uploaded_file, delimiter="\t")
            file_ready = True
        else:
            st.error("Unsupported file format.")
            st.stop()
    except Exception as e:
        st.error(f"Could not read the uploaded file: {e}")
        st.stop()
 
    if file_ready:
        required_cols = ['p_number', 'short_description', 'description', 'affected_customers', 'state', 'completion_code']
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.warning(f"Uploaded file is missing expected columns: {missing}. The app will attempt to continue but results may be incomplete.")
 
        if st.button("Generate"):
            # Add Download All Summaries button above the table
            combined_buf = util.create_combined_docx(df)
            st.download_button(label="Download All Summaries (.docx)",
                              data=combined_buf.getvalue(),
                              file_name="All_Project_Summaries.docx",
                              mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
 
            st.write("## Projects Table")
            import base64
            # Table header
            header_cols = st.columns([2, 4, 3, 2])
            header_cols[0].markdown("**Project Number**")
            header_cols[1].markdown("**Short Description**")
            header_cols[2].markdown("**Download Summary**")
            header_cols[3].markdown("**Preview**")
 
            if 'preview_states' not in st.session_state:
                st.session_state.preview_states = {}
 
            for idx, row in df.iterrows():
                title = row.get('short_description') or 'No title provided'
                pnum = row.get('p_number') or f'row-{idx}'
                doc_buf = util.create_docx(row)
                b64 = base64.b64encode(doc_buf.getvalue()).decode()
                row_cols = st.columns([2, 4, 3, 2])
                row_cols[0].write(pnum)
                row_cols[1].write(title)
                row_cols[2].markdown(f'<a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64}" download="Project_{pnum}.docx">Download Summary</a>', unsafe_allow_html=True)
                preview_key = f"preview_{pnum}_{idx}"
                preview_btn = row_cols[3].button("Preview", key=preview_key)
                if preview_btn:
                    st.session_state.preview_states[preview_key] = True
                if st.session_state.preview_states.get(preview_key, False):
                    st.markdown(f"**Preview for {pnum}:**")
                    st.write(util.generate_summary_llm(row))
 
            st.write("### Note")
            st.write("If any project summaries are missing or incomplete, please check the uploaded data for missing fields.")
            # st.write("For any issues, contact   [Your Contact Information].")      
            # st.write("Developed by [Your Name or Team].")  
            # st.write("Powered by Streamlit and Python.")
            st.write("© COFORGE. All rights reserved.")
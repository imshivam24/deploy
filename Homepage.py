import streamlit as st
import pandas as pd
import os
from utility import *

st.set_page_config(
    page_title="MKM Input File Generator and Solver",
    page_icon="👋",
)

# Title of the Streamlit app
st.title('MKM Input File Generator and Solver')

def main():
    if "uploaded_file" not in st.session_state:
        st.session_state["uploaded_file"] = ""

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")
    
    if uploaded_file:
        try:
            data = pd.read_excel(uploaded_file)
            st.write("Data Loaded Successfully!")
        except Exception as e:
            st.error(f"Error Loading Data: {str(e)}")
            return
        try:
            data1 = pd.read_excel(uploaded_file, sheet_name="Reactions")
            st.write("Reactions Loaded Successfully!")
            df1 = data1
            st.write("Reactions Preview:", df1.head())
        except Exception as e:
            st.error(f"Error reading Reactions sheet: {str(e)}")
            return
        
        try:
            data2 = pd.read_excel(uploaded_file, sheet_name="Local Environment")
            st.write("Local Environment Loaded Successfully!")
            df2 = data2
            st.write("Local Environment Preview:", df2.head())
        except Exception as e:
            st.error(f"Error reading Local Environment sheet: {str(e)}")
            return
        
        try:
            data3 = pd.read_excel(uploaded_file, sheet_name="Input-Output Species")
            st.write("Input-output Loaded Successfully!")
            df3 = data3
            st.write("Input-output Preview:", df3.head())
        except Exception as e:
            st.error(f"Error reading Input-output sheet: {str(e)}")
            return
        
        # Generate Input File Button
        if st.button("Generate MKM Input"):
            try:
                # Call the function to generate the input file
                input_file_path = inp_file_gen(uploaded_file)
                st.success("Input file generated successfully!")

                # After generating the file, allow the user to download it
                with open(input_file_path, "rb") as file:
                    st.download_button(
                        label="Download Generated MKM Input ",
                        data=file,
                        file_name="input_file.mkm",
                        mime="application/octet-stream"
                    )
            except Exception as e:
                st.error(f"Error generating input file: {str(e)}")

        input_file_path = "single_run/input_file.mkm"
        # Run Solver Button
        if st.button("Run Solver"):
            result_message, success = run_executable(input_file_path)
            if success:
                st.success(result_message)
                coverage()
            else:
                st.error(result_message)

if __name__ == "__main__":
    main()

import streamlit as st
import pandas as pd
import io
import zipfile
import base64
import openpyxl

def main():
    st.title("Multiple Excel Files Uploader")
    uploaded_files = st.file_uploader("Choose your Excel files", accept_multiple_files=True, type=['xlsx'])
    if uploaded_files:
        zip_file = zipfile.ZipFile("data_download.zip", mode="w")
        for uploaded_file in uploaded_files:
            bytes_data = uploaded_file.read()
            df = pd.read_excel(io.BytesIO(bytes_data),engine='openpyxl')
            st.write(df)
            zip_file.writestr(uploaded_file.name, bytes_data)
        zip_file.close()
        with open("data_download.zip", "rb") as f:
            bytes = f.read()
            b64 = base64.b64encode(bytes).decode()
        href = f'<a href="data:application/zip;base64,{b64}" download="data_download.zip">Download Zip File</a>'
        st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()

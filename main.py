import streamlit as st
import pandas as pd
import io
import zipfile
import base64
import openpyxl


def find_file(filename, uploaded_files):
    for file in uploaded_files:
        if filename in file.name:
            return file


def generate_reports(uploaded_files):
    
    if uploaded_files is not None:
        zip_file = zipfile.ZipFile("data_download.zip", mode="w")
        
        
        file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'                                       
        file_sugg_obj = find_file(file_name_sugg,uploaded_files)
        
        if file_sugg_obj is not None:
            bytes_data = file_sugg_obj.read()
            df1 = pd.read_excel(io.BytesIO(bytes_data),sheet_name='Live Contact List - Other',header=1, engine='openpyxl')
            st.write(df1)
            zip_file.writestr(file_sugg_obj.name, bytes_data)
            

        zip_file.close()
        
        
        
        #Generate download zip button
        with open("data_download.zip", "rb") as f:
            bytes = f.read()
            b64 = base64.b64encode(bytes).decode()
        href = f'<a href="data:application/zip;base64,{b64}" download="data_download.zip">Download Zip File</a>'
        st.markdown(href, unsafe_allow_html=True)






def main():
 
    #Title
    st.title("Multiple Excel Files Uploader")
    #Dropdown list
    options = ["QuarterlyReview", "Option 2", "Option 3"]
    selected_option = st.selectbox("Choose an option", options)
    #Uploader
    uploaded_files = st.file_uploader("Choose your files", accept_multiple_files=True, type=['xlsx'])
    #Execute Buttion
    button = st.button('Create Folder')

    if button:
        status = generate_reports(uploaded_files)
        st.write(status)
        
    
if __name__ == "__main__":
    main()

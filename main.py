import pandas as pd
import streamlit as st
import os
import openpyxl
import zipfile
import io
import base64

def find_file(filename, uploaded_files):
    for file in uploaded_files:
        if filename in file.name:
            return file

def generate_excel_download_link(df):
    towrite = io.BytesIO()
    df.to_excel(towrite, encoding="utf-8", index=False, header=True)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="data_download.xlsx">Download Excel File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_zip_download_link(zip_file):
    towrite = io.BytesIO()
    zip_file.save(towrite)
    towrite.seek(0)
    b64 = base64.b64encode(towrite.read()).decode()
    href = f'<a href="data:application/zip;base64,{b64}" download="data_download.zip">Download Zip File</a>'
    return st.markdown(href, unsafe_allow_html=True)

def generate_reports(uploaded_files):

    if uploaded_files is not None:     
        dfs = []
        
        for file in uploaded_files:    
            filename = file.name
            #Read source files, df, df1, df2, df3
            ##read df1 from suggestion
            #st.write(filename)
            
            file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'                                       
            file_path_sugg = find_file(file_name_sugg,uploaded_files)
            
            df1 = pd.read_excel(file_path_sugg,sheet_name='Live Contact List - Other',header=1,engine='openpyxl')           
            # df1['Role'] = df1['Role'].astype(str)
            # df1['Role'] = df1['Role'].apply(lambda x: x.split('/')).explode().reset_index(drop=True)                
            # df1['Role'] = df1['Role'].str.lstrip()
            # df1['Role'] = df1['Role'].str.rstrip()
            # ## debug_output:
            # #st.write(df1)          
            # Generate the output file name based on the input file name
            

            dfs.append(df1)
            # st.write(f'File name: {file.name}')           
            # generate_excel_download_link(df1)
            
            
        if len(dfs) > 0:
            zip_file = zipfile.ZipFile('data_download.zip', 'w')
            for i in range(len(dfs)):
                df = dfs[i]
                filename = f'data_{i}.xlsx'
                df.to_excel(filename, encoding="utf-8", index=False)
                zip_file.write(filename)
            zip_file.close()
            generate_zip_download_link(zip_file)
            
            
           
            # #backup
            # output_file_name = f"{file_path_sugg.name.split('.')[0]}_output.xlsx"
            # st.write(output_file_name)
            # #Convert the DataFrame to an Excel file and add it to the list of output files
            # output_file_contents = df1.to_excel(file_path_sugg,engine='openpyxl')
            # dfs.append((output_file_name, output_file_contents))
            
            # output_file_contents = io.BytesIO()
            # with pd.ExcelWriter(output_file_contents, engine='openpyxl') as writer:
            #     df1.to_excel(writer, index=False)
            # output_file_contents.seek(0)
            

            # ##backup
            
            # st.write(type(output_files))
            # st.write(len(output_files))
            # #df1.to_excel(path + '\\df1_debug.xlsx')
            # ## read df2 from suggestion
            # df2 = pd.read_excel(file_path_sugg,sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'])
            # df2 = df2.loc[~df2['Country/Region Name'].isna(),:]
            # df2['Code_Sub'] = df2['6 Digit Code'].str[:3]
            # df2 = df2.drop_duplicates(subset='Code_Sub',keep='last')
            # df2 = df2.drop(columns='6 Digit Code')
                
            # # Generate the output file name based on the input file name
            # output_file_name_df2 = f"{filename.split('.')[0]}_output.xlsx"
            # # Convert the DataFrame to an Excel file and add it to the list of output files
            # output_file_contents_df2 = df2.to_excel('df2_debug.xlsx')
            # output_files.append((output_file_name_df2, output_file_contents_df2))
                    
                    
            # if "name list" in filename:               
            #     df3 = pd.read_excel(file,sheet_name='名录（按组织）')
            #     df3 = df3.rename(columns={'电子邮件地址': 'Email_Source', '职务头衔':'Title'})
            #     def GetEmailAddress(x):
            #         return x.split('（')[0].strip(' ')
            #     df3.loc[:,'Email'] = df3['Email_Source'].astype(str).apply(lambda x: GetEmailAddress(x))
            #     df3 = pd.DataFrame(df3,columns=(['Email','Title']))
            #     df3 = df3.drop_duplicates(subset='Email')
            #     df3.loc[:,'Email_Upper'] = df3.loc[:,'Email']
            #     df3.loc[:,'Email_Upper'] = df3.apply(lambda x: x.str.upper())
            #     df3 = df3.drop(columns='Email')
            #     ##debug_output:
            #     #df3.to_excel(path + '\\df3_debug.xlsx')
            #     st.write(df3)            
        
        # #backup
        # # Create a zip file containing all the output files
        # zip_file_name = "output.zip"
        # with zipfile.ZipFile(zip_file_name, "w") as zip_file:
        #     for df in dfs:
        #         zip_file.writestr(df[0], df[1])
        # # Add a download button to allow the user to download the zip file
        # with open(zip_file_name, "rb") as f:
        #     zip_file_contents = f.read()
        # st.download_button(label="Download all output files as a zip", data=zip_file_contents, file_name=zip_file_name, mime="application/zip")
                
                

    # df1.to_excel(path + '\\df1_debug.xlsx')
    # df2.to_excel(path + '\\df2_debug.xlsx')
    # df3.to_excel(path + '\\df3_debug.xlsx')

def main():
    
    options = ["QuarterlyReview", "Option 2", "Option 3"]
    selected_option = st.selectbox("Choose an option", options)
    
    uploaded_files = st.file_uploader("Choose your files", accept_multiple_files=True, type=['xlsx'])
    
    #folder_name = st.text_input('Folder Name')
    
    button = st.button('Create Folder')
    # button1 = st.button('Generate Files')
    
    # if button:
    #     status = create_folder(selected_option, folder_name)
    #     st.write(status)
    if button:
        status = generate_reports(uploaded_files)
        st.write(status)

if __name__ == '__main__':
    main()

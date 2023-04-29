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
        
        #read df1
        file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'                                       
        file_sugg_obj = find_file(file_name_sugg,uploaded_files)      
        #read df1
        if file_sugg_obj is not None:
            bytes_data = file_sugg_obj.read()
            df1 = pd.read_excel(io.BytesIO(bytes_data),sheet_name='Live Contact List - Other',header=1, engine='openpyxl')
            df1['Role'] = df1['Role'].astype(str)
            df1['Role'] = df1['Role'].apply(lambda x: x.split('/')).explode().reset_index(drop=True)                
            df1['Role'] = df1['Role'].str.lstrip()
            df1['Role'] = df1['Role'].str.rstrip()
            st.write(df1)          
            output_df1_name = f"{file_sugg_obj.name.split('.')[0]}_df1.xlsx"          
            zip_file.writestr(output_df1_name, bytes_data)
            #read df2
            df2 = pd.read_excel(io.BytesIO(bytes_data),sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'],engine='openpyxl')
            df2 = df2.loc[~df2['Country/Region Name'].isna(),:]
            df2['Code_Sub'] = df2['6 Digit Code'].str[:3]
            df2 = df2.drop_duplicates(subset='Code_Sub',keep='last')
            df2 = df2.drop(columns='6 Digit Code')
            st.write(df2)
            output_df2_name = f"{file_sugg_obj.name.split('.')[0]}_df2.xlsx"
            zip_file.writestr(output_df2_name, bytes_data)
        #read df3
        file_name_nmlst = 'Name List.xlsx'
        file_nmlst_obj = find_file(file_name_nmlst,uploaded_files)    
        if file_nmlst_obj is not None:
            bytes_data = file_nmlst_obj.read()
            df3 = pd.read_excel(io.BytesIO(bytes_data),sheet_name='名录（按组织）',engine='openpyxl')
            df3 = df3.rename(columns={'电子邮件地址': 'Email_Source', '职务头衔':'Title'})
            def GetEmailAddress(x):
                return x.split('（')[0].strip(' ')
            df3.loc[:,'Email'] = df3['Email_Source'].astype(str).apply(lambda x: GetEmailAddress(x))
            df3 = pd.DataFrame(df3,columns=(['Email','Title']))
            df3 = df3.drop_duplicates(subset='Email')
            df3.loc[:,'Email_Upper'] = df3.loc[:,'Email']
            df3.loc[:,'Email_Upper'] = df3.apply(lambda x: x.str.upper())
            df3 = df3.drop(columns='Email')
            st.write(df3)
            output_df3_name = f"{file_sugg_obj.name.split('.')[0]}_df3.xlsx"
            zip_file.writestr(output_df3_name, bytes_data)
        #read access reports_multiple
        for file in uploaded_files:
            bytes_data = file.read()
            if file.endswith('.xlsx') and file.startswith('Quarterly Access Report'):
                df = pd.read_excel(bytes_data,dtype={'Study Environment Site Number': str},header=11,engine='openpyxl')
                # remove 'Unnamed' column
                for col, values in df.iteritems():
                    if 'Unnamed' in col:
                        df = df.drop(columns=col)
                #Method            
                def NoNeedReview(x, y):
                    if '@mdsol.com' in x:
                        return 'no need to review'
                    elif '@Medidata.com' in x:
                        return 'no need to review'
                    elif '@medidata.com' in x:
                        return 'no need to review'
                    elif '@3ds.com' in x:
                        return 'no need to review'
                    elif y == 'Medidata Internal Beigeneclinical_ebr':
                        return 'no need to review'  
                df['Assignment'] = df.apply(lambda x: NoNeedReview(x['Email'], x['Platform Role']),axis=1)
                df_row = df['Assignment'] != 'no need to review'
                df_flter = df.loc[df_row,:]
                df_flter = df_flter.drop(columns = ['Assignment'])
                st.write(df_flter)
                output_df_flter_name = f"{file_sugg_obj.name.split('.')[0]}_review01.xlsx"
                zip_file.writestr(output_df_flter_name, bytes_data)
                
        

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

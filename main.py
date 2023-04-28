# -*import os
import pandas as pd
import streamlit as st
import os
import openpyxl

# def create_folder(selected_option, folder_name):
#     path = os.path.join(os.getcwd(), folder_name, selected_option)
#     if os.path.isdir(path):
#         return 'Folder already exists'
#     else:
#         os.makedirs(path)
#         sub_dir = 'Source Files'
#         sub_folder = os.path.join(path, sub_dir)
#         os.mkdir(sub_folder)
#         return 'Folder created'

def find_file(filename, path):
    for root, dirs, files in os.walk(path):
        if filename in files:
            return os.path.join(root, filename)

def generate_reports(folder_name):
    
    path = os.path.join(os.getcwd(), folder_name)
    sub_dir_srcfiles = os.path.join(path, 'Source')
    
    if not os.path.isdir(sub_dir_srcfiles):
        return 'Source Files folder does not exist.'
    file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'
    file_path_sugg = find_file(file_name_sugg, sub_dir_srcfiles)
    if not file_path_sugg:
        return 'Suggestion file is not uploaded.'

    
    #Read source files, df, df1, df2, df3
    ##read df1 from suggestion
    df1 = pd.read_excel(file_path_sugg,sheet_name='Live Contact List - Other',header=1)
    df1 = df1.apply(lambda x: x.str.split('/').explode()).reset_index()
    df1['Role'] = df1['Role'].str.lstrip()
    df1['Role'] = df1['Role'].str.rstrip()
    ## debug_output:
    #df1.to_excel(path + '\\df1_debug.xlsx')
    ## read df2 from suggestion
    df2 = pd.read_excel(file_path_sugg,sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'])
    df2 = df2.loc[~df2['Country/Region Name'].isna(),:]
    df2['Code_Sub'] = df2['6 Digit Code'].str[:3]
    df2 = df2.drop_duplicates(subset='Code_Sub',keep='last')
    df2 = df2.drop(columns='6 Digit Code')
    #df2.to_excel(path + '\\df2.xlsx',index=False)
    #print(len(df2))
    ## debug_output:
    #df2.to_excel(path + '\\df2_debug.xlsx',index=False)  
    ##read df3
    file_name_nmlst = 'Name List.xlsx'
    file_path_nmlst = find_file(file_name_nmlst,sub_dir_srcfiles)                
    if file_path_nmlst == "":
        return 'name list file is not uploaded.'
        os._exit(0)
    df3 = pd.read_excel(file_path_nmlst,sheet_name='名录（按组织）')
    df3 = df3.rename(columns={'电子邮件地址': 'Email_Source', '职务头衔':'Title'})
    def GetEmailAddress(x):
        return x.split('（')[0].strip(' ')
    df3.loc[:,'Email'] = df3['Email_Source'].astype(str).apply(lambda x: GetEmailAddress(x))
    df3 = pd.DataFrame(df3,columns=(['Email','Title']))
    df3 = df3.drop_duplicates(subset='Email')
    df3.loc[:,'Email_Upper'] = df3.loc[:,'Email']
    df3.loc[:,'Email_Upper'] = df3.apply(lambda x: x.str.upper())
    df3 = df3.drop(columns='Email')
    ##debug_output:
    #df3.to_excel(path + '\\df3_debug.xlsx')

    df1.to_excel(path + '\\df1_debug.xlsx')
    df2.to_excel(path + '\\df2_debug.xlsx')
    df3.to_excel(path + '\\df3_debug.xlsx')

def main():

    folder_name = st.text_input('Folder Name')
    
    options = ["QuarterlyReview", "Option 2", "Option 3"]
    selected_option = st.selectbox("Choose an option", options)
    
    #button = st.button('Create Folder')
    button1 = st.button('Generate Files')
    
    # if button:
    #     status = create_folder(selected_option, folder_name)
    #     st.write(status)
    if button1 and selected_option == "QuarterlyReview":
        status = generate_reports(folder_name)
        st.write(status)

if __name__ == '__main__':
    main()

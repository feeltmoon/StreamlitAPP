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
        #read each access reports_multiple
        for file in uploaded_files:
            bytes_data = file.read()
            #if file.endswith('.xlsx') and file.startswith('Quarterly Access Report'):
            if "Quarterly Access Report" in file.name:
                df = pd.read_excel(bytes_data,dtype={'Study Environment Site Number': str},header=11,engine='openpyxl')
                # remove 'Unnamed' column
                for col, values in df.iteritems():
                    if 'Unnamed' in col:
                        df = df.drop(columns=col)
                # remove any empty rows
                df = df.dropna(how='all')
                #Method            
                def NoNeedReview(x, y):
                    if '@mdsol.com' in str(x):
                        return 'no need to review'
                    elif '@Medidata.com' in str(x):
                        return 'no need to review'
                    elif '@medidata.com' in str(x):
                        return 'no need to review'
                    elif '@3ds.com' in str(x):
                        return 'no need to review'
                    elif str(y) == 'Medidata Internal Beigeneclinical_ebr':
                        return 'no need to review'  
                df['Assignment'] = df.apply(lambda x: NoNeedReview(x['Email'], x['Platform Role']),axis=1)
                df_row = df['Assignment'] != 'no need to review'
                df_flter = df.loc[df_row,:]
                df_flter = df_flter.drop(columns = ['Assignment'])
                # Save the filtered data to a new Excel file
                writer_df_flter = pd.ExcelWriter('review_01.xlsx', engine='openpyxl')
                df_flter.to_excel(writer_df_flter, sheet_name='Sheet1', index=False)
                writer_df_flter.save()
                # Add the filtered data to the zip file
                output_df_flter_name = f"{file.name.split('.')[0]}_review01.xlsx"
                with open('review_01.xlsx', 'rb') as f:
                    data = f.read()
                    zip_file.writestr(output_df_flter_name, data)
                
                #*************************Review02**************************
                revw02 = df_flter.copy()
                revw02 = revw02.rename(columns={'Platform Role': 'Role'})
                revw02 = revw02.loc[:,['Email','Role','Environment']]
                revw02['ID'] = revw02['Email'] + "_" + revw02['Role']
                revw02['ID'] = revw02['ID'].str.upper()
                revw02_01 = df1.copy()
                revw02_01['ID'] = revw02_01['Email'] + "_" + revw02_01['Role']
                revw02_01['ID'] = revw02_01['ID'].str.upper()
                revw02_mrg = pd.merge(revw02_01,revw02,how='left',on=['ID'])
                revw02_mrg.loc[:,'Assignment'] = 'EDC Admin'     
                def NotFound(x):
                    if pd.isna(x):
                        return 'Fail - Contact EDC Admin'
                    elif ~pd.isna(x):
                        return 'Pass'
                revw02_mrg['Review Result/Comment/Action'] = revw02_mrg['Environment_y'].apply(lambda x: NotFound(x))  
                # 20220811 debugging
                # Update review02 failure reminder wording, unify to 'Fail - Contact EDC Admin'    
                revw02_mrg_row = (revw02_mrg['Role_x'].isin(['Read Only - All Sites','EDC Admin','Coder','Lab Entry','Lab Admin','Data PDF','Power User - SiM'])) & (revw02_mrg['Environment_y'].isna())
                revw02_mrg.loc[revw02_mrg_row, 'Review Result/Comment/Action'] = 'Fail - Contact EDC Admin'
                revw02_mrg = revw02_mrg.drop(['ID','Email_y','Role_y', 'Environment_y'],axis=1)
                revw02_mrg = revw02_mrg.rename(columns={'Email_x':'Email','Role_x':'Role','Environment_x':'Environment'})
                revw02_mrg_agg = revw02_mrg.groupby(['Email'])['Environment'].count().reset_index()
                revw02_mrg_agg = revw02_mrg_agg.rename(columns={'Environment': 'Email_Count'})
                revw02_mrg_flt = pd.merge(revw02_mrg,revw02_mrg_agg,how='left',on='Email')
                revw02_mrg_flt = revw02_mrg_flt.loc[revw02_mrg_flt['Email_Count'] == 2,:]
                def PassOrNot(x):
                    if x == 'Pass':
                        return 1
                    else:
                        return 0
                revw02_mrg_flt['Pass'] = revw02_mrg_flt['Review Result/Comment/Action'].apply(lambda x: PassOrNot(x))
                revw02_mrg_flt_agg = revw02_mrg_flt.groupby('Email')['Pass'].sum().reset_index()
                revw02_mrg_flt_agg = revw02_mrg_flt_agg.rename(columns={'Pass': 'Pass_Sum'})
                revw02_mrg_flt = revw02_mrg_flt.drop_duplicates(subset='Email')
                revw02_mrg_flt = pd.DataFrame(revw02_mrg_flt,columns=(['Email','Email_Count']))
                revw02_mrg = pd.merge(revw02_mrg,revw02_mrg_flt,how='left',on='Email')
                revw02_mrg = pd.merge(revw02_mrg,revw02_mrg_flt_agg,how='left',on='Email')
                revw02_mrg_row1 = (revw02_mrg['Email_Count'] == 2) & (revw02_mrg['Pass_Sum'] == 1)
                revw02_mrg.loc[revw02_mrg_row1, 'Review Result/Comment/Action'] = 'Pass'
                review01 = df.copy()
                review01['Email_ID'] = review01['Email']
                review01['Email_ID'] = review01['Email_ID'].str.upper()
                review01['Email_ID'] = review01['Email_ID'].str.lstrip()
                review01['Email_ID'] = review01['Email_ID'].str.rstrip()
                revw02_mrg_pass = revw02_mrg.loc[revw02_mrg['Review Result/Comment/Action'] == 'Pass',:]
                revw02_mrg_pass = revw02_mrg_pass.drop_duplicates(subset='Email')
                revw02_mrg_pass['Email_ID'] = revw02_mrg_pass['Email']
                revw02_mrg_pass['Email_ID'] = revw02_mrg_pass['Email_ID'].str.upper()
                revw02_mrg_pass['Email_ID'] = revw02_mrg_pass['Email_ID'].str.lstrip()
                revw02_mrg_pass['Email_ID'] = revw02_mrg_pass['Email_ID'].str.rstrip()
                revw02_mrg_pass = pd.DataFrame(revw02_mrg_pass,columns=(['Email','Assignment','Review Result/Comment/Action','Email_ID']))
                review01_mrg = pd.merge(review01,revw02_mrg_pass,how='left',on='Email_ID')
                review01_mrg = pd.DataFrame(review01_mrg,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email_x','Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment_x','Assignment_y','Review Result/Comment/Action']))
                review01_mrg = review01_mrg.rename(columns={'Email_x':'Email','Assignment_x':'Assignment_pri','Assignment_y':'Assignment_rev02','Review Result/Comment/Action':'Review02'})
                diff = revw02_mrg.loc[revw02_mrg['Review Result/Comment/Action'].str.contains('Fail'),:]                                                 
                new_cols = ['Client Division Scheme', 'Study', 'Phone #', 'Assignment Status', 'Location','Study Environment Site Number','Assignment_pri']
                diff = diff.reindex(diff.columns.union(new_cols), axis=1)
                diff = diff.rename(columns={'Assignment':'Assignment_rev02','Review Result/Comment/Action':'Review02','Role':'Platform Role'})
                diff = diff.drop(columns=['Email_Count','Pass_Sum'])                                                   
                diff = pd.DataFrame(diff,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment_pri','Assignment_rev02','Review02']))                                                   
                review01_mrg = review01_mrg.reset_index()                                                   
                concat = pd.concat([review01_mrg,diff],ignore_index=True,sort=False)                                                  
                concat = pd.DataFrame(concat,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email',
                                                       'Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number',
                                                       'Assignment_pri','Assignment_rev02','Review02']))
                
                
                # # debug output
                # # Save the data to a new Excel file
                # writer_concat = pd.ExcelWriter('concat.xlsx', engine='openpyxl')
                # concat.to_excel(writer_concat, index=False)
                # writer_concat.save()
                # # Add the filtered data to the zip file
                # output_concat_name = f"{file.name.split('.')[0]}_concat.xlsx"
                # with open('concat.xlsx', 'rb') as f:
                #     data = f.read()
                #     zip_file.writestr(output_concat_name, data)
                
                #**************************Check 03*********************************
                chk03 = df_flter.copy()
                chk03_01 = chk03.loc[chk03['Assignment Status'] != 'Active',:]
                # 20220809 modified:
                # required by Yun and Praveen
                def Reminder(x):
                    if x == 'Activation Expired':
                        return 'User did not activate their iMedidata account within the 45-day time frame, please request EDC Admin resending invitation mail to user.'
                    elif x == 'Activation Pending':
                        return 'please request EDC Admin resending invitation mail to user and inform the user to activate the account'
                    elif x == 'Activation Declined':
                        return 'User has declined the End-User License Agreement, please request EDC Admin resending invitation mail to user.'
                    elif x == 'Activation Email Delivered':
                        return 'User has not yet activated their iMedidata account. please remind user to activated their iMedidata account'
                    elif x == 'Activation Email Error' or x == 'Activation Email Failure' or x == 'Activation Email Send Failure' or x == 'Activation Email Delivery Failure':
                        return 'please request EDC Admin resending invitation mail to user or double check eMail ID with user'
                    elif x == 'Email Does Not Exist' or x == 'Activation Email Blocked':
                        return 'please double check with user'
                    elif x == 'eLearning Required':
                        return 'please remind user to complete the eLearning.'
                chk03_01.loc[:,'Review Result/Comment/Action'] = chk03_01['Assignment Status'].apply(lambda x: Reminder(x))
                
                if not chk03_01.empty:
                    def chk03_Classify(x, y, z):
                        if x == 'IxRS - Investigator' or x == 'Investigator' or x == 'IxRS - Sub-I' or x == 'Sub-I' or x == 'IxRS - Clinical Research Coordinator' or x == 'Clinical Research Coordinator' or x == 'Data Entry' or x == 'Test - IxRS - Investigator' or x == 'Test - Investigator' or x == 'Test - IxRS - Sub-I' or x == 'Test - Sub-I' or x == 'Test - IxRS - Clinical Research Coordinator' or x == 'Test - Clinical Research Coordinator':
                            return 'ACOM/PMA/CTA/Regional designee'
                        elif x == 'Read Only - Blinded' or x == 'Data Manager' or x == 'Safety' or x == 'Medical Monitor 1' or x == 'Medical Monitor 2' or x == 'Medical Monitor Blinded' or x == 'Test - Data Manager' or x == 'Test - Medical Monitor 1' or x == 'Test - Medical Monitor 2': 
                            return 'DM'
                        elif x == 'Read Only - All Sites' or x == 'EDC Admin' or x == 'Coder' or x == 'Lab Entry' or x == 'Data PDF' or x == 'Power User - SiM' or x == 'Outputs Standard' or x == 'Outputs - Blinded' or x == 'Output Locked': 
                            return 'EDC Admin'
                        elif x == 'Read Only' and y.__contains__('@beigene.com') and z == 'All Sites':
                            return 'DM, COM/COM designee'
                        elif x == 'Read Only' and (y.find('@beigene.com') == -1 or pd.isna(z)):
                            return 'COM/COM designee'
                        elif x == 'Acknowledger' or x == 'Clinical Research Associate' or x == 'Test - Clinical Research Associate':
                            return 'COM/COM designee'   
                
                    chk03_01.loc[:, 'Assignment'] = chk03_01.apply(lambda x: chk03_Classify(x['Platform Role'],x['Email'],x['Location']), axis=1)
                    chk03_01 = pd.DataFrame(chk03_01,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #',
                                                                'Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment',
                                                                'Review Result/Comment/Action']))
                elif chk03_01.empty:
                    chk03_01 = pd.DataFrame(chk03_01,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #',
                                                                'Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment',
                                                                'Review Result/Comment/Action']))
                    
                # get error count
                st.write(chk03_01)
                chk03_01_sumError = len(chk03_01)
                st.write(str(len(chk03_01)))
                
                # # Check03 Merge
                # # check03_merge into concat
                # chk03_01.fillna('99x083x', inplace = True)    
                # # 20220811 debugging
                # # add astype(str), otherwise procedure failure
                # #chk03_01.loc[:,'ID'] = chk03_01['Client Division Scheme'] + '_' + chk03_01['Study'] + '_' + chk03_01['Environment'] + '_' + chk03_01['First Name'] + '_' + chk03_01['Last Name'] + '_' + chk03_01['Email'] + '_' + chk03_01['Phone #'] + '_' + chk03_01['Platform Role']  + '_' + chk03_01['Assignment Status']  + '_' + chk03_01['Location']  + '_' + chk03_01['Study Environment Site Number']   
                # chk03_01['ID'] = chk03_01['Client Division Scheme'].astype(str) + '_' + chk03_01['Study'].astype(str) + '_' + chk03_01['Environment'].astype(str) + '_' + chk03_01['First Name'].astype(str) + '_' + chk03_01['Last Name'].astype(str) + '_' + chk03_01['Email'].astype(str) + '_' + chk03_01['Phone #'].astype(str) + '_' + chk03_01['Platform Role'].astype(str)  + '_' + chk03_01['Assignment Status'].astype(str)  + '_' + chk03_01['Location'].astype(str)  + '_' + chk03_01['Study Environment Site Number'].astype(str)                                                     
                # chk03_01 = chk03_01.rename(columns={'Assignment':'Assignment_chk03','Review Result/Comment/Action':'Check03'})
                # chk03_01 = pd.DataFrame(chk03_01,columns=(['ID','Assignment_chk03','Check03']))          
                # concat.fillna('99x083x', inplace = True)
                # # 20220811 debugging
                # # add astype(str), otherwise procedure failure
                # #concat['ID'] = concat['Client Division Scheme'] + '_' + concat['Study'] + '_' + concat['Environment'] + '_' + concat['First Name'] + '_' + concat['Last Name'] + '_' + concat['Email'] + '_' + concat['Phone #'] + '_' + concat['Platform Role']  + '_' + concat['Assignment Status']  + '_' + concat['Location']  + '_' + concat['Study Environment Site Number']
                # concat['ID'] = concat['Client Division Scheme'].astype(str) + '_' + concat['Study'].astype(str) + '_' + concat['Environment'].astype(str) + '_' + concat['First Name'].astype(str) + '_' + concat['Last Name'].astype(str) + '_' + concat['Email'].astype(str) + '_' + concat['Phone #'].astype(str) + '_' + concat['Platform Role'].astype(str)  + '_' + concat['Assignment Status'].astype(str)  + '_' + concat['Location'].astype(str)  + '_' + concat['Study Environment Site Number'].astype(str)
                # concat_chk03 = pd.merge(concat,chk03_01,how='left',on='ID')
                # #concat_chk03.to_excel(path + r'\concat_chk03.xlsx',index=False)
                
                           
                # # Create a writer for concat_chk03
                # writer_concat_chk03 = pd.ExcelWriter('concat_chk03.xlsx', engine='openpyxl')
                # concat_chk03.to_excel(writer_concat_chk03, index=False)                
                # writer_concat_chk03.save()                
                # # Add a new worksheet as checklist
                # workbook = openpyxl.load_workbook('concat_chk03.xlsx')
                # new_sheet = workbook.create_sheet('Checklist')
                # # Write in checklist
                # new_sheet['A1'] = 'Checkpoint'
                # new_sheet['A2'] = 'Check03'
                # new_sheet['B1'] = 'Description'
                # new_sheet['B2'] = '"Assignment status" column has been reviewed and proper reminders have been sent to EDC Admin or users (if any).'
                # new_sheet['C1'] = 'Pass/Fail'
                
                # def PassOrFail2(sumError, cell):
                #     if sumError == 0:
                #         new_sheet[cell] = "Pass"
                #     else:
                #         #sheet_schedule.write(cell, 'Fail' + '(' + str(sumErr) + ')')  
                #         new_sheet[cell] = "Fail" + "(" + str(sumError) + ")"
                
                # PassOrFail2(chk03_01_sumError, "C2")
                # # Save the changes to the file
                # workbook.save('concat_chk03.xlsx')                    
                # # Add the filtered data to the zip file
                # output_concat_chk03 = f"{file.name.split('.')[0]}_result.xlsx"
                # with open('concat_chk03.xlsx', 'rb') as f:
                #     data = f.read()
                #     zip_file.writestr(output_concat_chk03, data)
                    
                    
                    
                                                
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

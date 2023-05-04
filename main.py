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
            #Title
            st.title("Live Contact List - Other")
            st.write(df1)
            # # add to zip file
            # output_df1_name = f"{file_sugg_obj.name.split('.')[0]}_df1.xlsx"          
            # zip_file.writestr(output_df1_name, bytes_data)
            #read df2
            df2 = pd.read_excel(io.BytesIO(bytes_data),sheet_name='Country Codes',usecols=['Country/Region Name','6 Digit Code'],engine='openpyxl')
            df2 = df2.loc[~df2['Country/Region Name'].isna(),:]
            df2['Code_Sub'] = df2['6 Digit Code'].str[:3]
            df2 = df2.drop_duplicates(subset='Code_Sub',keep='last')
            df2 = df2.drop(columns='6 Digit Code')
            #Title
            st.title("Country Codes")
            st.write(df2)
            # # add to zip file
            # output_df2_name = f"{file_sugg_obj.name.split('.')[0]}_df2.xlsx"
            # zip_file.writestr(output_df2_name, bytes_data)
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
            #Title
            st.title("Name List")
            st.write(df3)
            # # add to zip file
            # output_df3_name = f"{file_sugg_obj.name.split('.')[0]}_df3.xlsx"
            # zip_file.writestr(output_df3_name, bytes_data)
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
                # # ADD TO ZIP FILE
                # # Save the filtered data to a new Excel file
                # writer_df_flter = pd.ExcelWriter('review_01.xlsx', engine='openpyxl')
                # df_flter.to_excel(writer_df_flter, sheet_name='Sheet1', index=False)
                # writer_df_flter.save()
                # # Add the filtered data to the zip file
                # output_df_flter_name = f"{file.name.split('.')[0]}_review01.xlsx"
                # with open('review_01.xlsx', 'rb') as f:
                #     data = f.read()
                #     zip_file.writestr(output_df_flter_name, data)
                
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
                #*************************Review02**************************
                
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
                
                #****************************************************Check 03*****************************************************************
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
                
                # GET ERROR SUM
                #st.write(chk03_01)
                chk03_01_sumError = len(chk03_01)
                #st.write(str(len(chk03_01)))
                
                # MERGE CHK03 INTO CONCAT
                chk03_01.fillna('99x083x', inplace = True)    
                # 20220811 debugging
                # add astype(str), otherwise procedure failure
                #chk03_01.loc[:,'ID'] = chk03_01['Client Division Scheme'] + '_' + chk03_01['Study'] + '_' + chk03_01['Environment'] + '_' + chk03_01['First Name'] + '_' + chk03_01['Last Name'] + '_' + chk03_01['Email'] + '_' + chk03_01['Phone #'] + '_' + chk03_01['Platform Role']  + '_' + chk03_01['Assignment Status']  + '_' + chk03_01['Location']  + '_' + chk03_01['Study Environment Site Number']   
                chk03_01['ID'] = chk03_01['Client Division Scheme'].astype(str) + '_' + chk03_01['Study'].astype(str) + '_' + chk03_01['Environment'].astype(str) + '_' + chk03_01['First Name'].astype(str) + '_' + chk03_01['Last Name'].astype(str) + '_' + chk03_01['Email'].astype(str) + '_' + chk03_01['Phone #'].astype(str) + '_' + chk03_01['Platform Role'].astype(str)  + '_' + chk03_01['Assignment Status'].astype(str)  + '_' + chk03_01['Location'].astype(str)  + '_' + chk03_01['Study Environment Site Number'].astype(str)                                                     
                chk03_01 = chk03_01.rename(columns={'Assignment':'Assignment_chk03','Review Result/Comment/Action':'Check03'})
                chk03_01 = pd.DataFrame(chk03_01,columns=(['ID','Assignment_chk03','Check03']))          
                concat.fillna('99x083x', inplace = True)
                # 20220811 debugging
                # add astype(str), otherwise procedure failure
                #concat['ID'] = concat['Client Division Scheme'] + '_' + concat['Study'] + '_' + concat['Environment'] + '_' + concat['First Name'] + '_' + concat['Last Name'] + '_' + concat['Email'] + '_' + concat['Phone #'] + '_' + concat['Platform Role']  + '_' + concat['Assignment Status']  + '_' + concat['Location']  + '_' + concat['Study Environment Site Number']
                concat['ID'] = concat['Client Division Scheme'].astype(str) + '_' + concat['Study'].astype(str) + '_' + concat['Environment'].astype(str) + '_' + concat['First Name'].astype(str) + '_' + concat['Last Name'].astype(str) + '_' + concat['Email'].astype(str) + '_' + concat['Phone #'].astype(str) + '_' + concat['Platform Role'].astype(str)  + '_' + concat['Assignment Status'].astype(str)  + '_' + concat['Location'].astype(str)  + '_' + concat['Study Environment Site Number'].astype(str)
                concat_chk03 = pd.merge(concat,chk03_01,how='left',on='ID')
                #concat_chk03.to_excel(path + r'\concat_chk03.xlsx',index=False)
                
                #****************************************************Check 03*****************************************************************
                
                
                #****************************************************Check 04*****************************************************************
                chk04 = df_flter.copy()
                chk04['Full_Name'] = chk04['First Name'] + ' ' + chk04['Last Name']
                chk04['Full_Name'] = chk04['Full_Name'].str.upper()
                chk04['Full_Name'] = chk04['Full_Name'].str.strip()
                chk04 = chk04.reindex(columns=['Client Division Scheme','Study','Environment','First Name','Last Name','Full_Name','Email','Phone #','Platform Role',
                                               'Assignment Status','Location','Study Environment Site Number'])
                agg = chk04.groupby(['Full_Name', 'Platform Role','Study Environment Site Number'])['Email'].count()
                agg = agg.reset_index()
                agg = pd.DataFrame(agg,columns=['Full_Name','Platform Role','Study Environment Site Number','Email'])
                chk04_mrg = pd.merge(chk04,agg,how='left',on=['Full_Name','Platform Role', 'Study Environment Site Number'])    
                chk04_mrg = chk04_mrg.loc[~chk04_mrg['Email_y'].isna(),:]
                chk04_msk = chk04_mrg['Email_y'] > 1
                chk04_mrg.loc[chk04_msk,'Review Result/Comment/Action'] = 'User has more than 1 EDC account in this study, please check and only keep one.'              
                # GET SUM ERROR
                chk04_sumError = chk04_mrg.loc[chk04_mrg['Review Result/Comment/Action'] == "User has more than 1 EDC account in this study, please check and only keep one.",:]
                len(chk04_sumError)           
                def chk04_Classify(x,y,z):
                    if x == 'IxRS - Investigator' or x == 'Investigator' or x == 'IxRS - Sub-I' or x == 'Sub-I' or x == 'IxRS - Clinical Research Coordinator' or x == 'Clinical Research Coordinator' or x == 'Data Entry':
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
                chk04_msk_01 = chk04_mrg['Email_y'] > 1
                chk04_mrg.loc[chk04_msk_01,'Assignment'] = chk04_mrg.apply(lambda x: chk04_Classify(x['Platform Role'],x['Email_x'],x['Location']), axis=1)
                chk04_mrg = chk04_mrg.drop(columns='Email_y')
                chk04_mrg = chk04_mrg.rename(columns={'Email_x':'Email'})
                #st.write(chk04_mrg)
                # # GET SUM ERROR
                # chk04_sumError = len(chk04_mrg)
                # MERGE chk04_mrg into concat
                chk04_mrg.fillna('99x083x', inplace = True)
                # 20220811 debugging
                # add astype(str), otherwise procedure failure
                #chk04_mrg['ID'] = chk04_mrg['Client Division Scheme'] + '_' + chk04_mrg['Study'] + '_' + chk04_mrg['Environment'] + '_' + chk04_mrg['First Name'] + '_' + chk04_mrg['Last Name'] + '_' + chk04_mrg['Email'] + '_' + chk04_mrg['Phone #'] + '_' + chk04_mrg['Platform Role']  + '_' + chk04_mrg['Assignment Status']  + '_' + chk04_mrg['Location']  + '_' + chk04_mrg['Study Environment Site Number']
                chk04_mrg['ID'] = chk04_mrg['Client Division Scheme'].astype(str) + '_' + chk04_mrg['Study'].astype(str) + '_' + chk04_mrg['Environment'].astype(str) + '_' + chk04_mrg['First Name'].astype(str) + '_' + chk04_mrg['Last Name'].astype(str) + '_' + chk04_mrg['Email'].astype(str) + '_' + chk04_mrg['Phone #'].astype(str) + '_' + chk04_mrg['Platform Role'].astype(str)  + '_' + chk04_mrg['Assignment Status'].astype(str)  + '_' + chk04_mrg['Location'].astype(str)  + '_' + chk04_mrg['Study Environment Site Number'].astype(str)  
                chk04_mrg = chk04_mrg.rename(columns={'Assignment':'Assignment_chk04','Review Result/Comment/Action':'Check04'})
                chk04_mrg = pd.DataFrame(chk04_mrg,columns=(['ID','Assignment_chk04','Check04']))         
                concat_chk03_04 = pd.merge(concat_chk03,chk04_mrg,how='left',on='ID')        
                # 20220810 debug:
                # praveen found issues: duplicate row in source report
                # remove duplicated rows
                concat_chk03_04 = concat_chk03_04.drop_duplicates()
                
                #****************************************************Check 04*****************************************************************
                
                
                #****************************************************Check 05*****************************************************************
                chk05 = df_flter.copy()
                chk05_row = (chk05['Environment'] == 'Production') & (chk05['Platform Role'].isin(['Clinical Research Coordinator','IxRS - Clinical Research Coordinator']))
                chk05 = chk05.loc[chk05_row,:]
                chk05 = chk05.loc[chk05['Location'] == 'All Sites',:]
                chk05['Review Result/Comment/Action'] = 'CRC user should not have "All Sites" privilege, please check and update.'
                def chk05_Classify(x):
                    if x == 'Clinical Research Coordinator' or x == 'IxRS - Clinical Research Coordinator' or x == 'Test - IxRS - Investigator' or x == 'Test - Investigator' or x == 'Test - IxRS - Sub-I' or x == 'Test - Sub-I' or x == 'Test - IxRS - Clinical Research Coordinator' or x == 'Test - Clinical Research Coordinator' or x == 'Test - Clinical Research Coordinator':
                        return 'COM/COM designee'    
                chk05['Assignment'] = chk05['Platform Role'].apply(lambda x: chk05_Classify(x))
                chk05 = pd.DataFrame(chk05,columns=['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #',
                                                    'Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment',
                                                    'Review Result/Comment/Action'])
                # GET SUM ERROR
                chk05_sumError = len(chk05)
                # MERGE check05 into concat_chk03_04_05
                chk05.fillna('99x083x', inplace = True)
                chk05['ID'] = chk05['Client Division Scheme'] + '_' + chk05['Study'] + '_' + chk05['Environment'] + '_' + chk05['First Name'] + '_' + chk05['Last Name'] + '_' + chk05['Email'] + '_' + chk05['Phone #'] + '_' + chk05['Platform Role']  + '_' + chk05['Assignment Status']  + '_' + chk05['Location']  + '_' + chk05['Study Environment Site Number']
                chk05 = chk05.rename(columns={'Assignment':'Assignment_chk05','Review Result/Comment/Action':'Check05'})
                chk05 = pd.DataFrame(chk05,columns=(['ID','Assignment_chk05','Check05']))
                concat_chk03_04_05 = pd.merge(concat_chk03_04,chk05,how='left',on='ID')
                #****************************************************Check 05*****************************************************************
                
                
                #****************************************************Check 06*****************************************************************
                chk06 = df_flter.copy()
                chk06_row = (chk06['Environment'] == 'Production') & (chk06['Platform Role'] == 'Data Entry')
                chk06 = chk06.loc[chk06_row,:]
                chk06['Review Result/Comment/Action'] = 'Please update to "Clinical Research Coordinator" or "IxRS-Clinical Research Coordinator" role.'
                def chk06_Classify(x):
                    if x == 'All Sites':
                        return 'COM/COM designee'
                    elif pd.isna(x):
                        return 'ACOM/PMA/CTA/Regional designee'       
                chk06['Assignment'] = chk06['Location'].apply(lambda x: chk06_Classify(x))
                # GET SUM ERROR
                chk06_sumError = len(chk06)
                # Check06 merge into concat_chk03_04_05
                chk06.fillna('99x083x', inplace = True)
                chk06['ID'] = chk06['Client Division Scheme'] + '_' + chk06['Study'] + '_' + chk06['Environment'] + '_' + chk06['First Name'] + '_' + chk06['Last Name'] + '_' + chk06['Email'] + '_' + chk06['Phone #'] + '_' + chk06['Platform Role']  + '_' + chk06['Assignment Status']  + '_' + chk06['Location']  + '_' + chk06['Study Environment Site Number']
                chk06 = chk06.rename(columns={'Assignment':'Assignment_chk06','Review Result/Comment/Action':'Check06'})
                chk06 = pd.DataFrame(chk06,columns=(['ID','Assignment_chk06','Check06']))
                concat_chk03_04_05_06 = pd.merge(concat_chk03_04_05,chk06,how='left',on='ID')
                #****************************************************Check 06*****************************************************************
                
                
                #****************************************************Check 07_01*****************************************************************
                chk07 = df_flter.copy()
                chk07_row = (~chk07['Study'].isin(['BGB-A317-209','BGB-900-102'])) & (chk07['Platform Role'].str.contains('Clinical Research Coordinator|Investigator|Sub-I'))
                chk07_01 = chk07.loc[chk07_row,:]
                chk07_01 = chk07_01.loc[~chk07_01['Platform Role'].str.contains('IXRS|Test - IxRS',case=False),:]
                def chk07_01_Classify(x):
                    if x == 'All Sites':
                        return 'COM/COM designee'
                    elif pd.isna(x):
                        return 'ACOM/PMA/CTA/Regional designee'   
                chk07_01['Assignment'] = chk07_01['Location'].apply(lambda x: chk07_01_Classify(x))
                def chk07_01_writer(x):
                    return 'Please update the role to IxRS - ' + x
                chk07_01.loc[:,'Review Result/Comment/Action'] = chk07_01['Platform Role'].apply(lambda x:chk07_01_writer(x))
                # GET SUM ERROR
                chk07_01_sumError = len(chk07_01)
                #****************************************************Check 07_01*****************************************************************
                
                
                #****************************************************Check 07_02*****************************************************************
                chk07_row1 = (chk07['Study'].isin(['BGB-A317-209','BGB-900-102'])) & (chk07['Platform Role'].str.contains('Clinical Research Coordinator|Investigator|Sub-I',case=False))
                chk07_02 = chk07.loc[chk07_row1,:]
                chk07_02 = chk07_02.loc[chk07_02['Platform Role'].str.contains('IXRS|Test - IxRS',case=False),:]
                def chk07_02_Classify(x):
                    if x == 'All Sites':
                        return 'COM/COM designee'
                    elif pd.isna(x):
                        return 'ACOM/PMA/CTA/Regional designee'        
                chk07_02['Assignment'] = chk07_02['Location'].apply(lambda x: chk07_02_Classify(x))
                def chk07_02_writer(x):
                    if x[0:4] == 'IxRS':
                        return 'Please update the role to ' + x[7:]
                    elif x[0:4] == 'Test':
                        return 'Please update the role to Test - ' + x[14:]    
                chk07_02.loc[:,'Review Result/Comment/Action'] = chk07_02['Platform Role'].apply(lambda x:chk07_02_writer(x))
                # GET SUM ERROR
                chk07_02_sumError = len(chk07_02)
                #****************************************************Check 07_02*****************************************************************
                
                
                #****************************************************Merge Check_07_01 and 02*****************************************************************
                # Check07_01/02 merge
                chk07_01.fillna('99x083x', inplace = True)
                # 20220811 debugging
                # add astype(str), otherwise procedure failure
                #chk07_01['ID'] = chk07_01['Client Division Scheme'] + '_' + chk07_01['Study'] + '_' + chk07_01['Environment'] + '_' + chk07_01['First Name'] + '_' + chk07_01['Last Name'] + '_' + chk07_01['Email'] + '_' + chk07_01['Phone #'] + '_' + chk07_01['Platform Role']  + '_' + chk07_01['Assignment Status']  + '_' + chk07_01['Location']  + '_' + chk07_01['Study Environment Site Number']
                chk07_01['ID'] = chk07_01['Client Division Scheme'].astype(str) + '_' + chk07_01['Study'].astype(str) + '_' + chk07_01['Environment'].astype(str) + '_' + chk07_01['First Name'].astype(str) + '_' + chk07_01['Last Name'].astype(str) + '_' + chk07_01['Email'].astype(str) + '_' + chk07_01['Phone #'].astype(str) + '_' + chk07_01['Platform Role'].astype(str)  + '_' + chk07_01['Assignment Status'].astype(str)  + '_' + chk07_01['Location'].astype(str)  + '_' + chk07_01['Study Environment Site Number'].astype(str)               
                chk07_01 = chk07_01.rename(columns={'Assignment':'Assignment_chk07','Review Result/Comment/Action':'Check07'})
                chk07_01 = pd.DataFrame(chk07_01,columns=(['ID','Assignment_chk07','Check07']))
                chk07_02.fillna('99x083x', inplace = True)
                # 20220811 debugging
                # add astype(str), otherwise procedure failure
                #chk07_02['ID'] = chk07_02['Client Division Scheme'] + '_' + chk07_02['Study'] + '_' + chk07_02['Environment'] + '_' + chk07_02['First Name'] + '_' + chk07_02['Last Name'] + '_' + chk07_02['Email'] + '_' + chk07_02['Phone #'] + '_' + chk07_02['Platform Role']  + '_' + chk07_02['Assignment Status']  + '_' + chk07_02['Location']  + '_' + chk07_02['Study Environment Site Number']
                chk07_02['ID'] = chk07_02['Client Division Scheme'].astype(str) + '_' + chk07_02['Study'].astype(str) + '_' + chk07_02['Environment'].astype(str) + '_' + chk07_02['First Name'].astype(str) + '_' + chk07_02['Last Name'].astype(str) + '_' + chk07_02['Email'].astype(str) + '_' + chk07_02['Phone #'].astype(str) + '_' + chk07_02['Platform Role'].astype(str)  + '_' + chk07_02['Assignment Status'].astype(str)  + '_' + chk07_02['Location'].astype(str)  + '_' + chk07_02['Study Environment Site Number'].astype(str)
                chk07_02 = chk07_02.rename(columns={'Assignment':'Assignment_chk07','Review Result/Comment/Action':'Check07'})
                chk07_02 = pd.DataFrame(chk07_02,columns=(['ID','Assignment_chk07','Check07']))        
                if ~chk07_01.empty & chk07_02.empty:
                    concat_chk03_04_05_06_07 = pd.merge(concat_chk03_04_05_06,chk07_01,how='left',on='ID')
                elif chk07_01.empty & ~chk07_02.empty:
                    concat_chk03_04_05_06_07 = pd.merge(concat_chk03_04_05_06,chk07_02,how='left',on='ID')
                elif chk07_01.empty & chk07_02.empty:
                    concat_chk03_04_05_06_07 = pd.merge(concat_chk03_04_05_06,chk07_01,how='left',on='ID')
                #****************************************************Merge Check_07_01 and 02*****************************************************************
                
                
                #****************************************************Check 08*****************************************************************
                chk08 = df_flter.copy()
                chk08_row = (chk08['Environment'] == 'Production') & (chk08['Platform Role'].str.contains('Power User|Study Developer|Test - Power User|EDC LAB Admin',case=False))
                chk08 = chk08.loc[chk08_row,:]
                chk08 = chk08.loc[chk08['Platform Role'] != 'Power User - SiM',:]
                chk08['Assignment'] = 'EDC Admin'
                chk08['Review Result/Comment/Action'] = 'This role is not allowed in the Prod environment. Please contact the EDC Admin to revoke them.'
                # GET SUM ERROR
                chk08_sumError = len(chk08)
                # check08 merge
                chk08.fillna('99x083x', inplace = True)
                chk08['ID'] = chk08['Client Division Scheme'] + '_' + chk08['Study'] + '_' + chk08['Environment'] + '_' + chk08['First Name'] + '_' + chk08['Last Name'] + '_' + chk08['Email'] + '_' + chk08['Phone #'] + '_' + chk08['Platform Role']  + '_' + chk08['Assignment Status']  + '_' + chk08['Location']  + '_' + chk08['Study Environment Site Number']
                chk08 = chk08.rename(columns={'Assignment':'Assignment_chk08','Review Result/Comment/Action':'Check08'})
                chk08 = pd.DataFrame(chk08,columns=(['ID','Assignment_chk08','Check08']))
                concat_chk03_04_05_06_07_08 = pd.merge(concat_chk03_04_05_06_07,chk08,how='left',on='ID')    
                #****************************************************Check 08*****************************************************************
                
                
                #****************************************************Check 09*****************************************************************
                chk09 = df_flter.copy()
                ### SW 20220713 modified:
                ### Hi Wei, 不好意思我又来了，因为最近PRA正在改邮箱域名，新增了两个，所以咱们的checkpoint 10也要相应改一下
                chk09_row = (chk09['Environment'] == 'Production') & (chk09['Platform Role'].str.contains('Data Manager|Test - Medical Monitor 1|Test - Medical Monitor 2')) & (~chk09['Email'].str.contains('@beigene.com|@prahs.com|@praintl.com|@iconplc.com'))
                chk09 = chk09.loc[chk09_row,:]
                chk09['Assignment'] = 'DM'
                chk09['Review Result/Comment/Action'] = 'DM role incorrectly assigned, please check.'
                # GET SUM ERROR
                chk09_sumError = len(chk09)
                # Check09 merge
                chk09.fillna('99x083x', inplace = True)
                chk09['ID'] = chk09['Client Division Scheme'] + '_' + chk09['Study'] + '_' + chk09['Environment'] + '_' + chk09['First Name'] + '_' + chk09['Last Name'] + '_' + chk09['Email'] + '_' + chk09['Phone #'] + '_' + chk09['Platform Role']  + '_' + chk09['Assignment Status']  + '_' + chk09['Location']  + '_' + chk09['Study Environment Site Number']
                chk09 = chk09.rename(columns={'Assignment':'Assignment_chk10','Review Result/Comment/Action':'Check10'})
                chk09 = pd.DataFrame(chk09,columns=(['ID','Assignment_chk10','Check10']))
                concat_chk03_04_05_06_07_08_09 = pd.merge(concat_chk03_04_05_06_07_08,chk09,how='left',on='ID')
                #****************************************************Check 09*****************************************************************
                
                
                # ===============Merge Country into Concat================
                def get_code(x):
                    if x != '99x083x':
                        return x[:3]
                # 20220811 debugging:
                # add astype(str)
                concat_chk03_04_05_06_07_08_09['6DigitCode'] = concat_chk03_04_05_06_07_08_09['Study Environment Site Number'].astype(str).apply(lambda x: get_code(x))   
                concat_chk03_04_05_06_07_08_09_country = pd.merge(concat_chk03_04_05_06_07_08_09,df2,how='left',left_on='6DigitCode',right_on='Code_Sub')
                concat_chk03_04_05_06_07_08_09_country = concat_chk03_04_05_06_07_08_09_country.drop(columns=['6DigitCode','Code_Sub'])
                concat_chk03_04_05_06_07_08_09_country = pd.DataFrame(concat_chk03_04_05_06_07_08_09_country,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment_pri','Assignment_rev02','Country/Region Name','Review02','ID','Assignment_chk03','Check03','Assignment_chk04','Check04','Assignment_chk05','Check05','Assignment_chk06','Check06','Assignment_chk07','Check07','Assignment_chk08','Check08','Assignment_chk10','Check10']))
                # 20220825 debugging:                                                                                                       
                # for below studies BGB-290-103, BGB-3111-212, BGB-3111-302, BGB-3111-304, mapping country/regrion is not needed
                if (concat_chk03_04_05_06_07_08_09_country.loc[2,'Study'] == 'BGB-290-103' or concat_chk03_04_05_06_07_08_09_country.loc[2,'Study'] == 'BGB-3111-212' or concat_chk03_04_05_06_07_08_09_country.loc[2,'Study'] == 'BGB-3111-302' or concat_chk03_04_05_06_07_08_09_country.loc[2,'Study'] == 'BGB-3111-304'):
                    #print('study is among BGB-290-103, BGB-3111-212, BGB-3111-302, BGB-3111-304')
                    concat_chk03_04_05_06_07_08_09_country = concat_chk03_04_05_06_07_08_09_country.drop(columns=['Country/Region Name'])
                    #print('study is among BGB-290-103, BGB-3111-212, BGB-3111-302, BGB-3111-304 drop columns')
                    concat_chk03_04_05_06_07_08_09_country['Country/Region Name'] = None
                    #print('study is among BGB-290-103, BGB-3111-212, BGB-3111-302, BGB-3111-304 add new column')
                    concat_chk03_04_05_06_07_08_09_country = pd.DataFrame(concat_chk03_04_05_06_07_08_09_country,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment_pri','Assignment_rev02','Country/Region Name','Review02','ID','Assignment_chk03','Check03','Assignment_chk04','Check04','Assignment_chk05','Check05','Assignment_chk06','Check06','Assignment_chk07','Check07','Assignment_chk08','Check08','Assignment_chk10','Check10']))
                    #print('study is among BGB-290-103, BGB-3111-212, BGB-3111-302, BGB-3111-304 reorder column')                                    
                # ===============Merge Country into Concat================
                
                
                # ===============Final Clean================
                concat_final = concat_chk03_04_05_06_07_08_09_country.fillna('99x083x')
                def Combine(a,b,c,d,e,f,g,h,i):
                    list = [a,b,c,d,e,f,g,h,i]    
                    val = ''
                    for i in list:
                        if i != '99x083x' and i != val.strip():
                            val += i + ' '
                    return val
                concat_final['Assignment'] = concat_final.apply(lambda x: Combine(x['Assignment_pri'],x['Assignment_rev02'],x['Assignment_chk03'],x['Assignment_chk04'],x['Assignment_chk05'],x['Assignment_chk06'],x['Assignment_chk07'],x['Assignment_chk08'],x['Assignment_chk10']),axis=1)
                concat_final['Assignment'] = concat_final['Assignment'].str.lstrip()
                concat_final['Assignment'] = concat_final['Assignment'].str.rstrip()
                concat_final = concat_final.drop(columns=['Assignment_pri','Assignment_rev02','Assignment_chk03','Assignment_chk04','Assignment_chk05','Assignment_chk06','Assignment_chk07','Assignment_chk08','Assignment_chk10','ID'])
                concat_final = pd.DataFrame(concat_final,columns=(['Client Division Scheme','Study','Environment','First Name','Last Name','Email','Title','Phone #','Platform Role','Assignment Status','Location','Study Environment Site Number','Assignment','Country/Region Name','Review02','Check03','Check04','Check05','Check06','Check07','Check08','Check10']))
                concat_final.loc[concat_final['Review02'] == 'Pass','Review02'] = ''
                def chk03_pass_role_classify(x,y,z):
                    if x == 'IxRS - Investigator' or x == 'Investigator' or x == 'IxRS - Sub-I' or x == 'Sub-I' or x == 'IxRS - Clinical Research Coordinator' or x == 'Clinical Research Coordinator' or x == 'Data Entry' or x == 'Test - IxRS - Investigator' or x == 'Test - Investigator' or x == 'Test - IxRS - Sub-I' or x == 'Test - Sub-I' or x == 'Test - IxRS - Clinical Research Coordinator' or x == 'Test - Clinical Research Coordinator':
                        return 'ACOM/PMA/CTA/Regional designee'
                    elif x == 'Read Only - Blinded' or x == 'Data Manager' or x == 'Safety' or x == 'Medical Monitor 1' or x == 'Medical Monitor 2' or x == 'Medical Monitor Blinded' or x == 'Test - Data Manager' or x == 'Test - Medical Monitor 1' or x == 'Test - Medical Monitor 2': 
                        return 'DM'
                    elif x == 'Read Only - All Sites' or x == 'EDC Admin' or x == 'Coder' or x == 'Lab Entry' or x == 'Data PDF' or x == 'Power User - SiM' or x == 'Outputs Standard' or x == 'Outputs - Blinded' or x == 'Output Locked': 
                        return 'EDC Admin'
                    elif x == 'Read Only' and y.__contains__('@beigene.com') and z == 'All Sites':
                        return 'DM, COM/COM designee'
                    #20220606: below elif code will confict with above one, for all cells equaling to 'Read Only', they can only enter either elif expression, because the conditions in above and below elif are all containing the same variables 
                    #20220606: that's why adding another method chk03_pass_role_classify1() below, using this method to output again for conflicted conditions
                    #20220606: note that pd.isna() is not taking effect, has to use =='99x083x' means equal to empty
                    elif x == 'Read Only' and (y.find('@beigene.com') == -1 or z == '99x083x'):
                        return 'COM/COM designee'
                    elif x == 'Acknowledger' or x == 'Clinical Research Associate' or x == 'Test - Clinical Research Associate':
                        return 'COM/COM designee'    
                mask = concat_final['Assignment Status'] == 'Active'
                concat_final.loc[mask,'Assig_chk03P'] = concat_final.apply(lambda x: chk03_pass_role_classify(x['Platform Role'],x['Email'],x['Location']), axis=1)
                #20220606 debug:
                # wei: Hi Yun，问下哈，你看这种算不算ok的
                # wei: 角色是readonly，location是空的
                # wei: assignment 没有分给任何role
                # yun: 其实应该assign，给COM/COM designee
                # filter out conflicted conditon-plat form Role = Read Only, entering assignment role as Com/COM designee for conditions either email endding with @beigene.com is not found OR platform role is empty
                def chk03_pass_role_classify1(x,y,z):   
                    #if x == 'Read Only' and (y.find('@beigene.com') == -1 or pd.isna(z)):
                    if y.find('@beigene.com') == -1 or z == '99x083x':
                        return 'COM/COM designee'  
                    else:
                        return x
                mask1 = (concat_final['Assignment Status'] == 'Active') & (concat_final['Platform Role'] == 'Read Only')
                concat_final.loc[mask1,'Assig_chk03P'] = concat_final.apply(lambda x: chk03_pass_role_classify1(x['Assig_chk03P'],x['Email'],x['Location']), axis=1)   
                #20220606 debug check output:
                #concat_final.to_excel(savePath + r'\concat_final_debug_20220606.xlsx')
                concat_final = concat_final.fillna('99x083x')
                # 01Jun2022 Debug check output:
                # concat_final.to_excel(savePath + r'\concat_final_debug.xlsx')
                def Combine1(x,y,z):        
                    if x != 'no need to review' and x != '99x083x' and y != '99x083x' and x.__contains__(y):
                        return x
                    # 20220811 debug
                    # Assignment Role is blank
                    #if x != 'no need to review' and x != '99x083x' and y == 'EDC Admin':
                        #return y
                    if x != 'no need to review' and x != '99x083x' and y != '99x083x' and (z.__contains__('Clinical System Implementation') or z.__contains__('Clinical Systems Specialist Implementation') or z.__contains__('GSDS Operational') or z.__contains__('Systems Validation') or z.startswith('IT')):
                        return x
                    #01Jun2022: Debug combined role still exists (EDC_Admin + DM, COM/COM designee)
                    elif x != 'no need to review' and x != '99x083x' and x == 'EDC Admin' and y != '99x083x':
                        return x
                    elif x != 'no need to review' and x != '99x083x' and y != '99x083x' and ~x.__contains__(y):
                        return x + '\n' + y
                    else:
                        return x
                concat_final['Assignment'] = concat_final.apply(lambda x: Combine1(x['Assignment'],x['Assig_chk03P'],x['Title']),axis=1)
                concat_final['Assignment'] = concat_final['Assignment'].str.lstrip()
                concat_final['Assignment'] = concat_final['Assignment'].str.rstrip()
                concat_final = concat_final.drop(columns='Assig_chk03P')
                concat_final.loc[concat_final['Assignment'].str.contains('no need to review'),'Country/Region Name'] = ''
                def concat_final_writer(x,y):
                    list = ['MEDS Reporter - IM','MEDS Reporter - SM','Read Only - SiM','CTMS Admin','CTMS - Read Only','COM - SiM','Clinical Research Associate - SiM','COM','CTA','ACOM - SiM','Admin - SM/IM','Power User - SiM (temporarily keep for CSI team members)']
                    if x in list:
                        return 'Pending: To Be Revoked/Replaced'
                    else:
                        return y
                concat_final['Assignment'] = concat_final.apply(lambda x: concat_final_writer(x['Platform Role'], x['Assignment']),axis=1)
                concat_final = concat_final.replace('99x083x','')      
                # ===============Final Clean================
                
                
                
                
                
                
                # ------------------------------------Test for creating checklist------------------------------------
                # Create a writer for concat_final
                writer_concat_final = pd.ExcelWriter('concat_final.xlsx', engine='openpyxl')
                concat_final.to_excel(writer_concat_final, index=False)                
                writer_concat_final.save()                
                # Add a new worksheet as checklist
                workbook = openpyxl.load_workbook('concat_final.xlsx')
                new_sheet = workbook.create_sheet('Checklist')
                # Move the second worksheet to the first position
                workbook.move_sheet(new_sheet, offset=-1)             
                # Write in checklist
                new_sheet['A1'] = 'Checkpoint'
                new_sheet['A2'] = 'Check03'
                new_sheet['A3'] = 'Check04'
                new_sheet['A4'] = 'Check05'
                new_sheet['A5'] = 'Check06'
                new_sheet['A6'] = 'Check07'
                new_sheet['A7'] = 'Check08'
                new_sheet['A8'] = 'Check10'
                new_sheet['B1'] = 'Description'
                new_sheet['B2'] = '"Assignment status" column has been reviewed and proper reminders have been sent to EDC Admin or users (if any).'
                new_sheet['B3'] = 'Any user would not have more than one EDC accounts registered in one study environment unless specific reason provided.'
                new_sheet['B4'] = '"Clinical Research Coordinator" or "IxRS-Clinical Research Coordinator" users can not have "All Sites" access.'
                new_sheet['B5'] = '"Data Entry" role cannot be used in study Prod environment, if any, it should be replaced by "Clinical Research Coordinator" or "IxRS-Clinical Research Coordinator" role.'
                new_sheet['B6'] = '01:If non-IRT study, PI, Sub-I and CRC should use corresponding "Investigator", "Sub-I" and "Clinical Research Coordinator" roles. 02:If IRT is used in study, PI, Sub-I and CRC should use corresponding "IxRS - Investigator", "IxRS - Sub-I" and "IxRS - Clinical Research Coordinator" roles.'
                new_sheet['B7'] = 'No user has "Power User", "Test - Power User", "Study Developer" or "EDC LAB Admin" role in study Prod environment.'
                new_sheet['B8'] = '"Data Manager" role can only be assigned to Beigene Global Data Management team in-house and out-sourced Data Managers.'                
                new_sheet['C1'] = 'Pass/Fail'
                new_sheet['C2'] = ''
                new_sheet['C3'] = ''
                new_sheet['C4'] = ''
                new_sheet['C5'] = ''
                new_sheet['C6'] = ''
                new_sheet['C7'] = ''
                new_sheet['C8'] = ''
                              
                def PassOrFail2(sumError, cell):
                    if sumError == 0:
                        new_sheet[cell] = "Pass"
                    else:
                        #sheet_schedule.write(cell, 'Fail' + '(' + str(sumErr) + ')')  
                        new_sheet[cell] = "Fail" + "(" + str(sumError) + ")"
                
                def PassOrFail1(cell, sumErr1, sumErr2):
                    if sumErr1 is None and sumErr2 is None:
                        new_sheet[cell] = "Pass"
                    elif sumErr1 is not None and sumErr2 is None:
                        new_sheet[cell] = "Fail" + "(" + str(sumErr1) + ")"
                    elif sumErr1 is None and sumErr2 is not None:  
                        new_sheet[cell] = "Fail" + "(" + str(sumErr2) + ")"
                
                PassOrFail2(chk03_01_sumError, "C2")
                PassOrFail2(chk04_sumError, "C3")
                PassOrFail2(chk05_sumError, "C4")
                PassOrFail2(chk06_sumError, "C5")
                PassOrFail1('C6', chk07_01_sumError, chk07_02_sumError)
                PassOrFail2(chk08_sumError, "C7")
                PassOrFail2(chk09_sumError, "C8")              
                # Save the changes to the file
                workbook.save('concat_final.xlsx')                    
                # Add the filtered data to the zip file
                concat_final = f"{file.name.split('.')[0]}_result.xlsx"
                with open('concat_final.xlsx', 'rb') as f:
                    data = f.read()
                    zip_file.writestr(concat_final, data)
                # ------------------------------------Test for creating checklist------------------------------------
                    
                    
                                                
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

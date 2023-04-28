# -*import os
import pandas as pd
import streamlit as st
import os
import openpyxl

def create_folder(selected_option, folder_name):
    path = os.path.join(os.getcwd(), folder_name, selected_option)
    if os.path.isdir(path):
        return 'Folder already exists'
    else:
        os.makedirs(path)
        sub_dir = 'Source Files'
        sub_folder = os.path.join(path, sub_dir)
        os.mkdir(sub_folder)
        return 'Folder created'

def find_file(filename, path):
    for root, dirs, files in os.walk(path):
        if filename in files:
            return os.path.join(root, filename)

def generate_reports(fruit, folder_name):
    path = os.path.join(os.getcwd(), folder_name, fruit)
    sub_dir_srcfiles = os.path.join(path, 'Source Files')
    if not os.path.isdir(sub_dir_srcfiles):
        return 'Source Files folder does not exist.'
    file_name_sugg = 'Medidata Rave EDC Roles Assignment and Quarterly Review Suggestions.xlsx'
    file_path_sugg = find_file(file_name_sugg, sub_dir_srcfiles)
    if not file_path_sugg:
        return 'Suggestion file is not uploaded.'
    df1 = pd.read_excel(file_path_sugg, sheet_name='Live Contact List - Other', header=1)
    # Do something with df1 here

def main():

    folder_name = st.text_input('Folder Name')
    
    options = ["QuarterlyReview", "Option 2", "Option 3"]
    selected_option = st.selectbox("Choose an option", options)
    
    button = st.button('Create Folder')
    button1 = st.button('Generate Files')
    
    if button:
        status = create_folder(selected_option, folder_name)
        st.write(status)
    elif button1:
        status = generate_reports(selected_option, folder_name)
        st.write(status)

if __name__ == '__main__':
    main()

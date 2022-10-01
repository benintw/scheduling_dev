'''
Last modified: Tues Sep 13, 2022 @ 21:11
Last modified: Tues Sep 20, 2022 @ 00:11
Last modified: Mon. Sep 26, 2022 @ 21:59
Last modified: Sat. Oct 1 , 2022 @ 15:06

'''

import os
from io import BytesIO
import streamlit as st
import time
from docx import Document
import pandas as pd
import xlsxwriter
from myFunctions import *



def main():

    st.markdown("# Intelligentsia")
    st.write("-"*100)
    password = st.sidebar.text_input("Enter Password", type='password')

    # if password == st.secrets[user_pw]:
    #     st.write("ok")
    if password == 'worldpeace':

        st.markdown("### 1. Upload Files")

        st.markdown("##### I acknowledge that:")

        condition1 = st.checkbox("* .xlsx .xls .docx are the only acceptable file-types")

        col_1, col_2 = st.columns(2)

        condition2 = col_1.checkbox("* all WORD.doc files are converted to WORD.docx files")

        need_help = col_2.button("How to convert to .docx?")
        try:
            if need_help:
                st.image("that.gif")
        except:
            print("testing")


        if condition1 and condition2:
            st.write("Great ! ")

            uploaded_files = st.file_uploader('Upload the files',type=["xlsx","xls","docx"] ,accept_multiple_files=True)
            st.write("-"*100)
            
   
            mega_team = get_different_teams_to_list(uploaded_files)


            st.markdown(f"#### {len(uploaded_files)} Files Uploaded:")
            # Start of Excel 
            if uploaded_files:
                st.markdown("###### EXCEL:")
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.endswith(".xlsx") or uploaded_file.name.endswith(".xls"):

                        st.write(uploaded_file.name)
                st.write(" ")
                st.markdown("###### WORD:")
                for uploaded_file in uploaded_files:
                    if uploaded_file.name.endswith(".docx"):

                        st.write(uploaded_file.name)

                for uploaded_file in uploaded_files:
                    if uploaded_file.name.split(".")[-1] in ["xlsx"]:
                        if "年" in uploaded_file.name:
                            YEAR = uploaded_file.name.split("年")[0]
                            MONTH = uploaded_file.name.split("月")[0].split("年")[1]
                
                day_of_week = get_day_of_week(mega_team["management_teams"])

                st.write("-"*80)
                st.markdown("### 2. Select Date")
                DAY = st.selectbox(
                    '',
                    (range(1,len(day_of_week))))


                get_my_excel_timetable(mega_team, DAY, MONTH, YEAR)
                

    elif password == "":
        st.write("Enter Password")

    elif password == "benchen":

        st.write("Entered Developer Mode... ")

        uploaded_files = st.file_uploader('Upload the files',type=["xlsx","xls","docx"] ,accept_multiple_files=True)
        st.write("-"*100)
        
        mega_team = get_different_teams_to_list(uploaded_files)


        st.markdown(f"#### {len(uploaded_files)} Files Uploaded:")
        # Start of Excel 
        if uploaded_files:
            st.markdown("###### EXCEL:")
            for uploaded_file in uploaded_files:
                if uploaded_file.name.endswith(".xlsx") or uploaded_file.name.endswith(".xls"):

                    st.write(uploaded_file.name)
            st.write(" ")
            st.markdown("###### WORD:")
            for uploaded_file in uploaded_files:
                if uploaded_file.name.endswith(".docx"):

                    st.write(uploaded_file.name)

            for uploaded_file in uploaded_files:
                if uploaded_file.name.split(".")[-1] in ["xlsx"]:
                    if "年" in uploaded_file.name:
                        YEAR = uploaded_file.name.split("年")[0]
                        MONTH = uploaded_file.name.split("月")[0].split("年")[1]
            
            day_of_week = get_day_of_week(mega_team["management_teams"])

            st.write("-"*80)

            count = 0
            for day in range(1,len(day_of_week)):
                get_my_excel_timetable(mega_team, day, MONTH, YEAR)
                count +=1

            st.write(f"Last file stops at file number: {count}")
            st.write(f"Number of successful excel output: {count}")

    else:
        try:
            st.image("haha2.png")
        except:
            print("Password incorrect")


if __name__ == "__main__":

    main()


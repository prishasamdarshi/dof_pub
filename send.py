"""
Create spreadsheets to be imported to CRM
"""
import os
import sys
import shutil
import numpy as np
import pandas as pd
import openpyxl as xl
from idclass import DocumentID, Letters
from openpyxl.utils.dataframe import dataframe_to_rows


###################################
# Set up templates and dataframes #
###################################

# Read and save new modified report data frame
def setup(filepath):
    df = pd.read_excel(filepath, sheet_name="Sheet1", dtype="object")
    df = df.set_index("Doc ID")
    return df


# Get absolute path to resource
def resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


# Copy case, contact, and email templates and store them in same folder as lag report
def copy_templates(filepath):
    path_list = filepath.rsplit("/", 1)
    filename_cases = path_list[0] + "/cases.xlsx"
    filename_contacts = path_list[0] + "/contacts.xlsx"
    filename_emails = path_list[0] + "/emails.xlsx"
    shutil.copy(resource_path("template_cases.xlsx"), filename_cases)
    shutil.copy(resource_path("template_contacts.xlsx"), filename_contacts)
    shutil.copy(resource_path("template_emails.xlsx"), filename_emails)
    return filename_cases, filename_contacts, filename_emails


# Open templates dataframes
def open_df(wb_name, ws_name):
    df = pd.read_excel(wb_name, sheet_name=ws_name, dtype="object")
    return df


################################################################
# Add customized information for documents to each spreadsheet #
################################################################

# Input case information from modified sheet
def create_case(df, title, customer, description):
    case_info = [np.nan, np.nan, np.nan, title, customer, "Correspondence", "General Correspondence",
                 "City Register - Manhattan", "Email", description, "In Progress"]
    df.loc[len(df.index)] = case_info
    return df


# Fill contact sheet with information from case dataframe
def create_contact(df, customer):
    contact_info = [np.nan, np.nan, np.nan, customer, customer]
    df.loc[len(df.index)] = contact_info
    return df


# Fill email sheet with information from case dataframe
def create_email(df, title, customer, template_name):
    email_info = [np.nan, np.nan, np.nan, "Email", "Outgoing", title, "City Register - Manhattan", customer,
                  template_name, "Sent"]
    df.loc[len(df.index)] = email_info
    return df


# Iterate through modified sheet to call create functions on each document
def add_info(df_main, df_cases, df_contacts, df_emails):
    id_list = list(df_main.index.values)
    for i in id_list:
        id = DocumentID(i)
        if df_main.at[i, "Rejected"] == 1 and df_main.at[i, "Search Manually"] == 0:
            amt = df_main.at[i, "Amount Paid for Entire Transaction (shown under first RPTT return in transaction)"]
            reasons = df_main.at[i, "Rejection Reason(s)"]
            rej_letter = Letters(i, amt, reasons)
            description = rej_letter.get_rej_description()
            title = "Document ID " + str(i)
            customer = df_main.at[i, "Customer Contact Email"]
            template_name = "Rejection_Letter"
            df_cases = create_case(df_cases, title, customer, description)
            df_contacts = create_contact(df_contacts, customer)
            df_emails = create_email(df_emails, title, customer, template_name)
        elif df_main.at[i, "Not Submitted"] == 1 and df_main.at[i, "Search Manually"] == 0:
            amt = df_main.at[i, "Amount Paid for Entire Transaction (shown under first RPTT return in transaction)"]
            year = id.get_date(i).year
            nosub_letter = Letters(i, amt, year)
            description = nosub_letter.get_nosub_description()
            title = "Document ID " + str(i)
            customer = df_main.at[i, "Customer Contact Email"]
            template_name = "Not_Submitted_Letter"
            df_cases = create_case(df_cases, title, customer, description)
            df_contacts = create_contact(df_contacts, customer)
            df_emails = create_email(df_emails, title, customer, template_name)
        else:
            continue
    return df_cases, df_contacts, df_emails


# Open template copies, load data from dataframes
def copy_df(df, wb_name, ws_name):
    wb = xl.load_workbook(wb_name)
    ws = wb[ws_name]
    ws.insert_rows(2, df.shape[0])
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)
    wb.save(wb_name)


######################
# Main (for testing) #
######################
'''
if __name__ == "__main__":
    filepath = input("Complete Filepath: ")
    df_main = setup(filepath)
    wb_names = copy_templates(filepath)
    df_cases = open_df(wb_names[0], "Case")
    df_contacts = open_df(wb_names[1], "Contact")
    df_emails = open_df(wb_names[2], "Email")
    df_list = add_info(df_main, df_cases, df_contacts, df_emails)
    copy_df(df_list[0], wb_names[0], "Case")
    copy_df(df_list[1], wb_names[1], "Contact")
    copy_df(df_list[2], wb_names[2], "Email")
'''


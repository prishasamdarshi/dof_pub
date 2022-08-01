"""
Search Document IDs in ACRIS and output modified spreadsheet with necessary information
"""


import constant
import pandas as pd
from tkinter import Tk
from idclass import DocumentID
from datetime import timedelta
from pywinauto import application, findwindows


##############################################################
# Open and output modified df copy of lag report spreadsheet #
##############################################################

# Find original sheet and set up new sheet
def setup(filepath):
    # Parse filepath for filename, path, and date
    path_list = filepath.rsplit("/", 1)
    name_list = path_list[1].rsplit("of ", 1)
    date_str = name_list[1].rsplit(".", 1)
    report_date = DocumentID.get_date(date_str[0])
    # Read and make reduced column duplicate of lag report
    df_full = pd.read_excel(filepath, sheet_name="Lagging RPTTs Recent", dtype="object")
    df_red = df_full[["Doc ID", "Doc Type", "Nbr Intakes", "Current Doc Status",
                      "Amount Paid for Entire Transaction (shown under first RPTT return in transaction)",
                      "Customer Contact Email"]]
    df_final = df_red.set_index("Doc ID")
    df_final.index.astype(str)
    add_col(df_final)
    return df_final, report_date, path_list[0]


# Filter document IDs from before lag report quarter
def itr(df, doc_dlg, report_date):
    id_list = list(df.index.values)
    for i in id_list:
        id = DocumentID(i)
        # Use trid in ACRIS search, use last if there are multiple results
        tr_id = id.get_trid()
        last = id.get_last()
        try:
            past_case(df, report_date, id.docID)
            search_docs(df, report_date, doc_dlg, id.docID, tr_id, last)
        except Exception:
            search_manually(df, id.docID)
            pass


# Add new columns to dataframe
def add_col(df):
    df.insert(len(df.columns), "Recorded", 0)
    df.insert(len(df.columns), "Not Submitted", 0)
    df.insert(len(df.columns), "Rejected", 0)
    df.insert(len(df.columns), "Rejection Reason(s)", 0)
    df.insert(len(df.columns), "Processing", 0)
    df.insert(len(df.columns), "Search Manually", 0)


# Mark "recorded" column with boolean true
def recorded(df, doc_id):
    df.at[doc_id, "Recorded"] = 1


# Mark "not submitted" column with boolean true
def not_submitted(df, doc_id):
    df.at[doc_id, "Not Submitted"] = 1


# Mark "rejected" column with boolean true
def rejected(df, doc_id):
    df.at[doc_id, "Rejected"] = 1


# Add rejection reasons to "rejection reason(s)" column
def add_reason(df, doc_id, reason):
    df.at[doc_id, "Rejection Reason(s)"] = reason


# Mark "search manually" column with boolean true
def search_manually(df, doc_id):
    df.at[doc_id, "Search Manually"] = 1


# Mark "processing" column with boolean true
def processing(df, doc_id):
    df.at[doc_id, "Processing"] = 1


# Mark "past case" column with boolean true
def past_case(df, report_date, doc_id):
    max_delta = timedelta(weeks=40)
    rec_date = DocumentID.get_date(str(doc_id))
    if (rec_date + max_delta) < report_date:
        search_manually(df, doc_id)
        return


#####################################################
# Connect to ACRIS, perform search, record response #
#####################################################

# Connect to running instance of ACRIS
def connect_search():
    acris = application.Application(backend="uia").connect(path=r"C:\Program Files (x86)\ACRIS\Admin\ACRIS.BackOffice"
                                                                r".Admin.exe")
    findwindows.find_window(title="ACRIS Document View and Print")
    doc_dlg = acris.window(title="ACRIS Document View and Print")
    return doc_dlg


# Search by transaction ID
def search_docs(df, report_date, doc_dlg, doc_id, trID, last):
    doc_dlg.Clear.click()
    doc_dlg["Transaction NumberEdit"].set_text(trID)
    doc_dlg.type_keys("{ENTER}")
    error_dlg = doc_dlg.child_window(title="Error Message")
    grid_dlg = doc_dlg.child_window(title="DataGridView", auto_id="dgvReason")
    main_dlg = doc_dlg.child_window(title="DataGridView", auto_id="dgvMain")
    if error_dlg.exists():
        check_error(df, report_date, doc_id, error_dlg)
        error_dlg.OKButton.click()
    elif main_dlg.exists():
        row = nav_rows(main_dlg, last)
        check_main(df, doc_id, row)
        if grid_dlg.exists():
            check_rej(df, doc_id, grid_dlg)
            past_case(df, report_date, doc_id)
    else:
        search_manually(df, doc_id)


# Go to correct row in main_dlg
def nav_rows(main_dlg, last):
    row_int = int(last) - 1
    row_str = "Row " + str(row_int)
    row = main_dlg.child_window(title=row_str, control_type="Custom")
    row.exists(timeout=10)
    # Scroll down while the row is not visible
    while not row.is_visible():
        scroll = main_dlg.child_window(title="Vertical Scroll Bar", control_type="ScrollBar")
        down = scroll.child_window(title="Line down", control_type="Button")
        down.click_input()
    row.click_input(coords=(0, 0))
    return row


# Perform necessary steps if error message window opens
def check_error(df, report_date, doc_id, error_dlg):
    error_dlg.click_input(button="right")
    # Copies error message onto clipboard in order to read
    error_dlg.type_keys("{DOWN}{DOWN}{DOWN}{ENTER}")
    error_msg = Tk().clipboard_get()
    if constant.RECORDED in error_msg:
        recorded(df, doc_id)
    elif constant.NOT_SUBMITTED in error_msg and df.at[doc_id, "Nbr Intakes"] == 0:
        not_submitted(df, doc_id)
        past_case(df, report_date, doc_id)
    elif constant.NOT_SUBMITTED in error_msg and df.at[doc_id, "Nbr Intakes"] != 0:
        not_submitted(df, doc_id)
        search_manually(df, doc_id)
    elif constant.NOT_SCANNED in error_msg:
        not_submitted(df, doc_id)
        past_case(df, report_date, doc_id)


# Perform necessary steps if main grid dialog says accepted or in process
def check_main(df, doc_id, row):
    row.type_keys("^c")
    if constant.ACCEPTED in Tk().clipboard_get():
        recorded(df, doc_id)
    elif constant.IN_PROCESS in Tk().clipboard_get():
        processing(df, doc_id)
    elif constant.REJECTED in Tk().clipboard_get():
        rejected(df, doc_id)


# Perform necessary steps if rejection reasons open
def check_rej(df, doc_id, grid_dlg):
    rej_table = grid_dlg.descendants(control_type="Edit") or grid_dlg.descendants(control_type="Data Item")
    reasons_str = ""
    for r in rej_table:
        if "Reason" in str(r):
            while not r.is_visible():
                scroll = grid_dlg.child_window(title="Vertical Scroll Bar", control_type="ScrollBar")
                down = scroll.child_window(title="Line down", control_type="Button")
                down.click_input()
            r.click_input(coords=(0, 0))
            r.type_keys("^c")
            reason = Tk().clipboard_get() + "\n"
            reasons_str = reasons_str + reason
    add_reason(df, doc_id, reasons_str)


######################
# Main (for testing) #
######################
'''
if __name__ == "__main__":
    filepath = input("Complete Filepath: ")
    result_list = setup(filepath)
    df = result_list[0]
    date = result_list[1]
    path = result_list[2]
    doc_dlg = connect_search()
    itr(df, doc_dlg, date)
    writer = pd.ExcelWriter(path + "\\Modified Report as of " + str(date) + ".xlsx")
    df.to_excel(writer)
    writer.save()
'''
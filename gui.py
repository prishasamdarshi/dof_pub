"""
User interface for document processing
"""

import read
import send
import pandas as pd
import tkinter as tk
from threading import *
from pywinauto import application, findwindows
from tkinter import filedialog, font, messagebox


#########################
# Class for root window #
#########################

class MainApp(tk.Frame):
    def __init__(self, root):
        self.root = root
        tk.Frame.__init__(self, self.root)
        self.configure_dialog()
        self.create_widgets()

    # Set up main window
    def configure_dialog(self):
        self.root.configure(background="white")
        self.root.title("Unprocessed Revenue Handling")
        self.root.geometry("500x380")

    # Call widget-creating functions
    def create_widgets(self):
        self.create_header()
        self.create_buttons()

    # Header image and labels
    def create_header(self):
        font_header = font.Font(family="Helvetica", weight="bold", size=35)
        frame_header = tk.LabelFrame(self.root, bd=0, background="white")
        label_header = tk.Label(frame_header, text="Welcome", font=font_header, foreground="#085aad",
                                background="white")
        label_question = tk.Label(frame_header, text="What do you want to do today?", font=("Helvetica", 12),
                                  background="white")
        label_header.pack()
        label_question.pack(pady=8)
        frame_header.pack(side="top", fill="x", pady=20)

    # Buttons for child windows
    def create_buttons(self):
        font_buttons = font.Font(family="Helvetica", size=12)
        button_acris = tk.Button(self.root, text="🔍 Search in ACRIS", height=3, width=45, bg="white",
                                 activebackground="#f0f0f0", command=self.open_child_acris)
        button_acris["font"] = font_buttons
        button_acris.pack(pady=10)
        button_crm = tk.Button(self.root, text="📝 Create spreadsheets for CRM",  height=3, width=45, bg="white",
                               activebackground="#f0f0f0", command=self.open_child_crm)
        button_crm["font"] = font_buttons
        button_crm.pack(pady=15)

    # Command for first button
    def open_child_acris(self):
        self.root.withdraw()
        AcrisChild(self.root)

    # Command for second button
    def open_child_crm(self):
        self.root.withdraw()
        CrmChild(self.root)


#################################
# Class for ACRIS search window #
#################################

class AcrisChild(tk.Frame):
    def __init__(self, root):
        super().__init__()
        self.root = root
        self.child = tk.Toplevel(root)
        tk.Frame.__init__(self, self.child)
        self.configure_dialog()
        self.create_widgets()

    # Set up child window
    def configure_dialog(self):
        self.child.configure(background="white")
        self.child.title("ACRIS Search")
        self.child.geometry("600x360")

    # Call widget-creating functions
    def create_widgets(self):
        self.create_header()
        self.create_footer()
        self.create_instructions()
        self.create_buttons()

    # Header image and labels
    def create_header(self):
        font_header = font.Font(family="Helvetica", weight="bold", size=25)
        frame_header = tk.LabelFrame(self.child, bd=0, background="white")
        label_header = tk.Label(frame_header, text="🔍 Search in ACRIS", font=font_header, foreground="#085aad",
                                background="white")
        label_header.pack()
        frame_header.pack(side="top", fill="x", pady=5)

    # Footer buttons
    def create_footer(self):
        font_buttons = font.Font(family="Helvetica", size=10)
        frame_footer = tk.Frame(self.child, bd=0, background="white")
        frame_footer.pack(side="bottom", fill="x", pady=20)
        button_return = tk.Button(frame_footer, text="\u2190 Back", height=1, width=10, bg="white", font=font_buttons,
                                  activebackground="#f0f0f0", command=self.go_back)
        button_continue = tk.Button(frame_footer, text="Continue \u2192", height=1, width=10, bg="white",
                                    font=font_buttons, activebackground="#f0f0f0", command=self.go_forward)
        button_return.pack(side="left", padx=20)
        button_continue.pack(side="right", padx=20)

    # Instructions
    def create_instructions(self):
        font_instructions = font.Font(family="Helvetica", size=9)
        instruction0 = "Before beginning make sure you do the following:"
        instruction1 = "1. Open ACRIS, log on, and navigate to the window titled \"Document View and Print\" " \
                       "\n(Administration > Document View and Print)"
        instruction2 = "2. Ensure the appropriate spreadsheet is downloaded onto your computer, the name of the\n" \
                       "spreadsheet includes \"as of {date}\" where the date is in the format yyyymmdd, and the " \
                       "worksheet you \nwould like to input for processing is named \"Lagging RPTTs Recent\" "
        frame_instruction = tk.LabelFrame(self.child, bd=0, background="#f0f0f0", labelanchor="nw")
        frame_instruction.pack(padx=20, pady=10)
        label_instruction0 = tk.Label(frame_instruction, anchor="w", justify="left", width=80, text=instruction0,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction1 = tk.Label(frame_instruction, anchor="w", width=80, justify="left", text=instruction1,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction2 = tk.Label(frame_instruction, anchor="w", width=80, justify="left", text=instruction2,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction0.pack()
        label_instruction1.pack()
        label_instruction2.pack()

    # Attach file button
    def create_buttons(self):
        font_buttons = font.Font(family="Helvetica", size=12)
        button_files = tk.Button(self.child, text="📁 Open lag report spreadsheet", height=2, width=40, bg="white",
                                 font=font_buttons, activebackground="#f0f0f0", command=self.upload_thread)
        button_files.pack(pady=30)

    # Command for back button
    def go_back(self):
        self.child.withdraw()
        root_dlg = tk.Toplevel()
        MainApp(root_dlg)

    # Command for continue button
    def go_forward(self):
        self.child.withdraw()
        CrmChild(self.root)

    # Command for open file button
    def upload_thread(self):
        filepath = filedialog.askopenfilename()
        Thread(target=self.main, args=[filepath]).start()

    # Begin "read" program from here
    @staticmethod
    def main(filepath):
        # Check the file name and worksheet name
        try:
            result_list = read.setup(filepath)
            doc_dlg = read.connect_search()
        except ValueError:
            messagebox.showerror("Error: File Name", "Please open the correct xlsx file and make sure the file and "
                                                     "worksheet are both named according to the instructions.")
        except application.ProcessNotFoundError:
            messagebox.showerror("Error: Application Not Connected", "Please open ACRIS and navigate to the correct "
                                                                     "window")
        except findwindows.WindowNotFoundError:
            messagebox.showerror("Error: Window Not Found", "Please navigate to the \"Document View and Print\" window")
        else:
            df = result_list[0]
            date = result_list[1]
            path = result_list[2]
            read.itr(df, doc_dlg, date)
            writer = pd.ExcelWriter(path + "/Modified Report as of " + str(date) + ".xlsx")
            df.to_excel(writer)
            writer.save()
        messagebox.showinfo("Success!", "The program is complete. Your modified lag report has been added to the same "
                                        "folder as the original report.")


#############################################
# Class for CRM spreadsheet creation window #
#############################################

class CrmChild(tk.Frame):
    def __init__(self, root):
        super().__init__()
        self.child = tk.Toplevel(root)
        tk.Frame.__init__(self, self.child)
        self.configure_dialog()
        self.create_widgets()

    # Set up child window
    def configure_dialog(self):
        self.child.configure(background="white")
        self.child.title("CRM Spreadsheets")
        self.child.geometry("600x360")

    # Call widget-creating functions
    def create_widgets(self):
        self.create_header()
        self.create_footer()
        self.create_instructions()
        self.create_buttons()

    # Header image and labels
    def create_header(self):
        font_header = font.Font(family="Helvetica", weight="bold", size=25)
        frame_header = tk.LabelFrame(self.child, bd=0, background="white")
        label_header = tk.Label(frame_header, text="📝 Create Spreadsheets for CRM", font=font_header,
                                foreground="#085aad", background="white")
        label_header.pack()
        frame_header.pack(side="top", fill="x", pady=5)

    # Footer back button
    def create_footer(self):
        font_buttons = font.Font(family="Helvetica", size=10)
        frame_footer = tk.LabelFrame(self.child, bd=0, background="white")
        button_return = tk.Button(frame_footer, text="\u2190 Back", height=1, width=8, bg="white", font=font_buttons,
                                  activebackground="#f0f0f0", command=self.go_back)
        button_return.pack(side="left", padx=20)
        frame_footer.pack(side="bottom", fill="x", pady=20)

    # Instructions
    def create_instructions(self):
        font_instructions = font.Font(family="Helvetica", size=9)
        instruction0 = "Before beginning make sure you do the following:"
        instruction1 = "1. Run the \"Search in ACRIS\" on the previous page"
        instruction2 = "2. Ensure the spreadsheet outputted by the aforementioned program (NOT the original) is \n" \
                       "downloaded onto your computer, the name of the new spreadsheet is \"Modified Report as of\" " \
                       "\n+ {date}, and the worksheet you would like to input for processing is named \"Sheet1\" "
        frame_instruction = tk.LabelFrame(self.child, bd=0, background="#f0f0f0", labelanchor="nw")
        frame_instruction.pack(padx=20, pady=10)
        label_instruction0 = tk.Label(frame_instruction, anchor="w", justify="left", width=80, text=instruction0,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction1 = tk.Label(frame_instruction, anchor="w", width=80, justify="left", text=instruction1,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction2 = tk.Label(frame_instruction, anchor="w", width=80, justify="left", text=instruction2,
                                      font=font_instructions, bg="#f0f0f0")
        label_instruction0.pack()
        label_instruction1.pack()
        label_instruction2.pack()

    # Attach file button
    def create_buttons(self):
        font_buttons = font.Font(family="Helvetica", size=12)
        button_files = tk.Button(self.child, text="📁 Open modified lag report spreadsheet", height=2, width=40,
                                 bg="white", font=font_buttons, activebackground="#f0f0f0", command=self.upload_thread)
        button_files.pack(pady=30)

    # Command for back button
    def go_back(self):
        self.child.withdraw()
        root_dlg = tk.Toplevel()
        MainApp(root_dlg)

    # Command for open file button
    def upload_thread(self):
        filepath = filedialog.askopenfilename()
        Thread(target=self.main, args=[filepath]).start()

    # Begin "send" program from here
    @staticmethod
    def main(filepath):
        df_main = send.setup(filepath)
        wb_names = send.copy_templates(filepath)
        df_cases = send.open_df(wb_names[0], "Case")
        df_contacts = send.open_df(wb_names[1], "Contact")
        df_emails = send.open_df(wb_names[2], "Email")
        df_list = send.add_info(df_main, df_cases, df_contacts, df_emails)
        send.copy_df(df_list[0], wb_names[0], "Case")
        send.copy_df(df_list[1], wb_names[1], "Contact")
        send.copy_df(df_list[2], wb_names[2], "Email")
        messagebox.showinfo("Success!", "The program is complete. Your newly created spreadsheets have been added to "
                                        "the same folder as the original report and are ready for import to CRM.")


######################
# Main (for testing) #
######################
'''
if __name__ == "__main__":
    root = tk.Tk()
    parent = MainApp(root)
    parent.mainloop()
'''
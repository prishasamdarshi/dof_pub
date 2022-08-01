"""
Main Program (runs gui.py)
"""

import gui
import tkinter as tk


if __name__ == "__main__":
    parent_dlg = tk.Tk()
    parent = gui.MainApp(parent_dlg)
    parent.mainloop()

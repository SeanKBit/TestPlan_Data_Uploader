# TestPlan_Data_Uploader.py
# 7/9/2021
# Sean Bittner
#
# "Universal" application for populating data to Warpdrive Test Plans. (previously Wormhole.py)
# This is the frontend GUI for users to update a config file for path to data, and row/column formatting values.
# This config is then used to parse appropriate data to a scanned (via barcode) test plan.
# A .vbs program interacts with the config file and handles the backend warp transfer tasks.
# If warpdrive API's get updated to work with Python3, that logic can all be handle here if desired.
#
# Dependencies:
# 1. CONFIG_PATH .txt file must exist as the directory below
# 2. VBS_PATH .vbs program must exist as the directory below


import tkinter as tk
import subprocess
import sys
from tkinter.filedialog import askopenfilename

CONFIG_PATH = r"C:\Users\Public\TestPlan_Data_Uploader\Resource\config.txt"
VBS_PATH = r"C:\Users\Public\TestPlan_Data_Uploader\Resource\WarpWorker.vbs"
ICON_PATH = r"C:\Users\Public\TestPlan_Data_Uploader\Resource\SpaceX-X-Black.png"


# (re)sets / writes to the config file with user entered values within entry fields
def set_config():
    int_data = ent_data.get()
    int_var = ent_variable.get()
    int_omit = ent_header.get()
    str_path = ent_path.get()
    str_path = str_path.replace("/", "\\")

    with open(CONFIG_PATH, 'w') as write_file:
        write_file.writelines(int_data + '\n')
        write_file.writelines(int_var + '\n')
        write_file.writelines(int_omit + '\n')
        write_file.writelines(str_path)
        write_file.close()

        read_config()


# reads in data from required config file and inserts results in entry fields
def read_config():
    with open(CONFIG_PATH, 'r') as read_file:
        int_data = read_file.readline()
        int_var = read_file.readline()
        int_omit = read_file.readline()
        str_path = read_file.readline()
        read_file.close()

    ent_data.delete(0, tk.END)
    ent_data.insert(0, int_data.rstrip())
    ent_variable.delete(0, tk.END)
    ent_variable.insert(0, int_var.rstrip())
    ent_header.delete(0, tk.END)
    ent_header.insert(0, int_omit.rstrip())
    ent_path.delete(0, tk.END)
    ent_path.insert(0, str_path)


def upload_data():
    subprocess.Popen("wscript  " + VBS_PATH)


# uses tkinter.filedialog function to open explorer to get path
def get_path():
    filepath = askopenfilename()
    if not filepath:
        return
    ent_path.delete(0, tk.END)
    ent_path.insert(0, filepath)


def on_enter(e):
    e.widget["background"] = "#c9c9c9"


def on_leave(e):
    e.widget["background"] = "SystemButtonFace"


def on_closing():
    window.destroy()
    sys.exit()


# main
window = tk.Tk()
window.protocol("WM_DELETE_WINDOW", on_closing)
window.title("Universal Test Plan Data Uploader")
window.rowconfigure(0, minsize=300, weight=1)
window.rowconfigure(1, minsize=100, weight=1)
window.columnconfigure(1, minsize=1100, weight=1)

p1 = tk.PhotoImage(file=ICON_PATH)
window.iconphoto(False, p1)


fr_upload = tk.Frame(window, background="white", relief=tk.RIDGE, borderwidth=10, padx=8, pady=8)
fr_entries = tk.Frame(window, relief=tk.SUNKEN, borderwidth=10, padx=10, pady=8)
fr_entries.grid(row=0, column=1, sticky="nsew")
fr_upload.grid(row=1, column=1)

lbl_data = tk.Label(fr_entries, text="Data column number: ", font=("Copperplate Gothic Light", 11))
lbl_data.grid(row=0, column=0, sticky="e", padx=8, pady=10)
ent_data = tk.Entry(fr_entries, width=10)
ent_data.grid(row=0, column=1, sticky="w", padx=8, pady=10)

lbl_variable = tk.Label(fr_entries, text="Variable column number: ", font=("Copperplate Gothic Light", 11))
lbl_variable.grid(row=1, column=0, sticky="e", padx=8, pady=10)
ent_variable = tk.Entry(fr_entries, width=10)
ent_variable.grid(row=1, column=1, sticky="w", padx=8, pady=10)

lbl_header = tk.Label(fr_entries, text="Number of header rows to omit: ", font=("Copperplate Gothic Light", 11))
lbl_header.grid(row=2, column=0, sticky="e", padx=8, pady=10)
ent_header = tk.Entry(fr_entries, width=10)
ent_header.grid(row=2, column=1, sticky="w", padx=8, pady=10)

lbl_path = tk.Label(fr_entries, text="Path to result file: ", font=("Copperplate Gothic Light", 11))
lbl_path.grid(row=3, column=0, sticky="e", padx=8, pady=10)
ent_path = tk.Entry(fr_entries, width=100)
ent_path.grid(row=3, column=1, sticky="w", padx=8, pady=10)

btn_browse = tk.Button(fr_entries, command=get_path, text="Browse...", width=10)
btn_browse.grid(row=3, column=2, sticky="w", padx=8, pady=10)
btn_browse.bind("<Enter>", on_enter)
btn_browse.bind("<Leave>", on_leave)

lbl_config = tk.Label(fr_entries, text="Push after above fields are edited: ", font=("Copperplate Gothic Light", 11))
lbl_config.grid(row=4, column=0, sticky="e", padx=8, pady=10)
btn_config = tk.Button(
    fr_entries,
    text="*Set Config File*",
    command=set_config,
    foreground="red",
    activebackground="grey",
    font=("Times New Roman", 12, "bold"),
    width=13,
    height=1,
    relief=tk.RAISED
)
btn_config.grid(row=4, column=1, sticky="w", padx=8, pady=10)
btn_config.bind("<Enter>", on_enter)
btn_config.bind("<Leave>", on_leave)


btn_upload = tk.Button(
    fr_upload,
    text=">Go For Launch<",
    command=upload_data,
    activebackground="white",
    foreground="#165700",
    background="#dbdbdb",
    font=("Showcard Gothic", 13, "italic"),
    width=16,
    height=2,
    relief=tk.RAISED
)
btn_upload.grid(row=0, column=0, sticky="nsew", padx=10, pady=12)
btn_upload.bind("<Enter>", on_enter)
btn_upload.bind("<Leave>", on_leave)

read_config()

window.mainloop()

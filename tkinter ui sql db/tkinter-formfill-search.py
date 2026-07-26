import tkinter
from tkinter import ttk
from tkinter import messagebox
import openpyxl
import os
import sqlite3

# Database setup
# enter database
def enter_data():
    accepted = accept_var.get()

    if accepted == "Accepted":
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()

        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()

            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numsemesters_spinbox.get()

            filepath = "formdataEX.xlsx"

            if not os.path.exists(filepath):
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Age", "Nationality",
                           "# Courses", "# Semesters", "Registration status"]
                sheet.append(heading)
                try:
                    workbook.save(filepath)
                except PermissionError:
                    tkinter.messagebox.showerror("Error", "Please close the Excel file before submitting!")
                    return

            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality, numcourses,
                          numsemesters, registration_status])
            workbook.save(filepath)

            conn = sqlite3.connect('datatkf.db')
            table_create_query = '''CREATE TABLE IF NOT EXISTS Student_Data 
                            (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT, 
                            registration_status TEXT, num_courses INT, num_semesters INT)
                            '''
            conn.execute(table_create_query)

            data_insert_query = '''INSERT INTO Student_Data (firstname, lastname, title, 
                    age, nationality, registration_status, num_courses, num_semesters) VALUES 
                    (?, ?, ?, ?, ?, ?, ?, ?)'''
            data_insert_tuple = (firstname, lastname, title, age, nationality, registration_status, numcourses, numsemesters)
            cursor = conn.cursor()
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()
            conn.close()

            tkinter.messagebox.showinfo(title="Success", message="Data submitted successfully!")
            load_data()

        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title="Error", message="You have not accepted the terms")

# tab2 view of search
def load_data(search_query=""):
    for row in tree.get_children():
        tree.delete(row)

    if not os.path.exists('datatkf.db'):
        return

    conn = sqlite3.connect('datatkf.db')
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS Student_Data 
                    (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT, 
                    registration_status TEXT, num_courses INT, num_semesters INT)''')

    if search_query.strip():
        query = '''SELECT firstname, lastname, title, age, nationality, 
                          num_courses, num_semesters, registration_status 
                   FROM Student_Data 
                   WHERE firstname LIKE ? OR lastname LIKE ? OR nationality LIKE ?'''
        param = f"%{search_query.strip()}%"
        cursor.execute(query, (param, param, param))
    else:
        cursor.execute('''SELECT firstname, lastname, title, age, nationality, 
                                 num_courses, num_semesters, registration_status 
                          FROM Student_Data''')

    rows = cursor.fetchall()
    for row in rows:
        tree.insert("", "end", values=row)

    conn.close()

#function to search data
def search_data():
    query = search_entry.get()
    load_data(query)

#function to clear search
def clear_search():
    search_entry.delete(0, tkinter.END)
    load_data()


# App UI

window = tkinter.Tk()
window.title("Data Entry & Record Management")

# Fixed compact size for the container (forces horizontal scrolling in Tab 2)
window.geometry("545x420")
window.resizable(False, False)

notebook = ttk.Notebook(window)
notebook.pack(fill="both", expand=True)

tab1 = ttk.Frame(notebook)
tab2 = ttk.Frame(notebook)

notebook.add(tab1, text="New Entry")
notebook.add(tab2, text="View Records")


# App UI - Tab1

frame = tkinter.Frame(tab1)
frame.pack(padx=10, pady=10)

user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=10)

first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)

first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)

title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["", "Mr.", "Ms.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)

age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame,
                                    values=["Africa", "Antarctica", "Asia", "Europe", "North America", "Oceania",
                                            "South America"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

registered_label = tkinter.Label(courses_frame, text="Registration Status")
reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered",
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not registered")

registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

numcourses_label = tkinter.Label(courses_frame, text="# Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)

numsemesters_label = tkinter.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to='infinity')
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)

for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.",
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

button = tkinter.Button(frame, text="Enter data", command=enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)


# App UI - Tab2

search_frame = tkinter.Frame(tab2)
search_frame.pack(fill="x", padx=10, pady=10)

search_label = tkinter.Label(search_frame, text="Search:")
search_label.pack(side="left", padx=5)

search_entry = tkinter.Entry(search_frame)
search_entry.pack(side="left", fill="x", expand=True, padx=5)
search_entry.bind("<Return>", lambda event: search_data())

search_button = tkinter.Button(search_frame, text="Search", command=search_data)
search_button.pack(side="left", padx=5)

clear_button = tkinter.Button(search_frame, text="Clear", command=clear_search)
clear_button.pack(side="left", padx=5)

# Treeview  Frame
tree_frame = tkinter.Frame(tab2)
tree_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

# Scrollbars
tree_scroll_y = ttk.Scrollbar(tree_frame, orient="vertical")
tree_scroll_y.pack(side="right", fill="y")

tree_scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal")
tree_scroll_x.pack(side="bottom", fill="x")

columns = ("First Name", "Last Name", "Title", "Age", "Nationality", "Courses", "Semesters", "Status")
#datashow view
tree = ttk.Treeview(
    tree_frame,
    columns=columns,
    show="headings",
    selectmode="browse",
    yscrollcommand=tree_scroll_y.set,
    xscrollcommand=tree_scroll_x.set
)

tree_scroll_y.config(command=tree.yview)
tree_scroll_x.config(command=tree.xview)

# Width of data elements
column_widths = {
    "First Name": 120,
    "Last Name": 120,
    "Title": 60,
    "Age": 60,
    "Nationality": 120,
    "Courses": 90,
    "Semesters": 90,
    "Status": 110
}

for col, width in column_widths.items():
    tree.heading(col, text=col)
    tree.column(col, width=width, minwidth=width, stretch=False, anchor="center")

tree.pack(fill="both", expand=True)

load_data()

window.mainloop()
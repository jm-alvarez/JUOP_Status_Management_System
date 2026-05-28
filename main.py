import tkinter
import datetime
import tkcalendar
from tkinter import messagebox

def enter_data():
    date = date_entry.get()
    juop_partner = juop_partner_entry.get()
    inspection_report_date = inspection_report_date_entry.get()
    coverage_area = coverage_area_entry.get()
    number_of_poles = number_of_poles_entry.get()
    number_of_downguys = number_of_downguys_entry.get()
    remarks = remarks_entry.get()
    
    print("Date:", date, "\nJUOP Partner:", juop_partner, "\nInspection Report Date:", inspection_report_date, "\nCoverage Area:", coverage_area, "\nNo. of Poles:", number_of_poles, "\nNo. of Downguys:", number_of_downguys, "\nRemarks:", remarks)

window = tkinter.Tk()
window.geometry("500x400")
window.title("JUOP Status Management System")

frame = tkinter.Frame(window)
frame.pack()

add_data_frame = tkinter.LabelFrame(frame, text = "Add Data")
add_data_frame.grid(row = 0, column = 0)

date_label = tkinter.Label(add_data_frame, text = "Date")
date_label.grid(row = 0, column = 0)
date_entry = tkcalendar.DateEntry(add_data_frame, date_pattern = "mm/dd/yyyy",)
date_entry.grid(row = 1, column = 0)

juop_partner_label = tkinter.Label(add_data_frame, text = "JUOP Partner")
juop_partner_label.grid(row = 0, column = 1)
juop_partner_entry = tkinter.Entry(add_data_frame, width = 50)
juop_partner_entry.grid(row = 1, column = 1)

inspection_report_date_label = tkinter.Label(add_data_frame, text = "Inspection Report Date")
inspection_report_date_label.grid(row = 3, column = 0)
inspection_report_date_entry = tkcalendar.DateEntry(add_data_frame, date_pattern = "mm/dd/yyyy")
inspection_report_date_entry.grid(row = 4, column = 0)

coverage_area_label = tkinter.Label(add_data_frame, text = "Coverage Area")
coverage_area_label.grid(row = 3, column = 1)
coverage_area_entry = tkinter.Entry(add_data_frame, width = 50)
coverage_area_entry.grid(row = 4, column = 1)

number_of_poles_label = tkinter.Label(add_data_frame, text = "No. of Poles")
number_of_poles_label.grid(row = 5, column = 0)
number_of_poles_entry = tkinter.Spinbox(add_data_frame, from_ = 0, to = 1000, width = 10, validate = "key", validatecommand = (add_data_frame.register(lambda s: s.isdigit()), "%P"))
number_of_poles_entry.grid(row = 6, column = 0)

number_of_downguys_label = tkinter.Label(add_data_frame, text = "No. of Downguys")
number_of_downguys_label.grid(row = 5, column = 1)
number_of_downguys_entry = tkinter.Spinbox(add_data_frame, from_ = 0, to = 1000, width = 10, validate = "key", validatecommand = (add_data_frame.register(lambda s: s.isdigit()), "%P"))
number_of_downguys_entry.grid(row = 6, column = 1)

remarks_frame = tkinter.LabelFrame(add_data_frame, text = "Remarks")
remarks_frame.grid(row = 7, column = 0, columnspan = 2, sticky = "news", padx = 20, pady = 20)
gm_label = tkinter.Label(remarks_frame, text = "GM-")
gm_label.grid(row = 0, column = 0)
#remarks_entry = tkinter.Entry(remarks_frame, width = 5, validate = "key", validatecommand = (remarks_frame.register(lambda s: s.isdigit() and len(s) <= 4), "%P"))
remarks_entry = tkinter.Entry(remarks_frame, width = 10)
remarks_entry.grid(row = 0, column = 1)

buttons_frame = tkinter.LabelFrame(frame, text = "Actions")
buttons_frame.grid(row = 2, column = 0, columnspan = 2, sticky = "news")
add_button = tkinter.Button(buttons_frame, text = "Add Data", command = enter_data)
add_button.grid(row = 0, column = 0)
update_button = tkinter.Button(buttons_frame, text = "Update Data")
update_button.grid(row = 0, column = 1)
delete_button = tkinter.Button(buttons_frame, text = "Delete Data")
delete_button.grid(row = 0, column = 2)

for widget in add_data_frame.winfo_children():
    widget.grid_configure(padx = 5, pady = 5)

for widget in buttons_frame.winfo_children():
    widget.grid_configure(padx = 20, pady = 20)

window.mainloop()
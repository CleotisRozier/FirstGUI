#This is the Excel Output version of the mainInputGUI.
#Everything is exported into excel sheet, path must be modified (line 40)
#Created this following along with YouTube video series.
#Link: https://youtube.com/playlist?list=PLs3IFJPw3G9LQfo8GYbnPRaD8DKHj9aOJ&si=ipJZ6f54LpjBSiJL

import tkinter
from tkinter import ttk
from tkinter import messagebox
import os
import openpyxl                      #allows output to excel documents

#-------variables----------
#for the enter button at the bottom of the screen
#will retrieve all info and verifies registration status from registration status checkbox
def enter_data():
    #User Agreement Acceptance Validator
    #User must accept to complete
    accepted = accept_var.get()    
    if accepted=="Accepted":
        firstname = first_name_entry.get()
        lastname = last_name_entry.get()
        #Proceed if name requireme nts met
        if firstname and lastname:
            title = title_combobox.get()
            age = age_spinbox.get()
            nationality = nationality_combobox.get()
            
            #Courses info
            registration_status = reg_status_var.get()
            numcourses = numcourses_spinbox.get()
            numsemesters = numcourses_spinbox.get()
            
            #Print output retrieved    
            print("First name: ", firstname, "Last name: ", lastname)
            print("Title: ", title, "Age: ", age, "Nationality: ", nationality)
            print("# Courses: ", numcourses, "# Semesters: ", numsemesters)
            print("Registration status", registration_status)
            print("------------------------------------------")
            
            filepath = "PATH ~/data.xlsx"                                                     #update filepath
            
            if not os.path.exists(filepath):                                               #Added openpyxl code
                workbook = openpyxl.Workbook()
                sheet = workbook.active
                heading = ["First Name", "Last Name", "Title", "Nationality", 
                           "# Courses", "#Semesters", "Registration status"]
                sheet.append(heading)
                workbook.save(filepath)
            workbook = openpyxl.load_workbook(filepath)
            sheet = workbook.active
            sheet.append([firstname, lastname, title, age, nationality,
                          numcourses, numsemesters, registration_status])
            workbook.save(filepath)
            
            
            
            
        else:
            tkinter.messagebox.showwarning(title="Error", message="First name and last name are required.")
    else:
        tkinter.messagebox.showwarning(title= "Error", message="You have to accept terms and conditions to proceed.")

##########################################
#-----------Build the canvas-------------#
##########################################

#The window and frame are separated by the border
#Building the window (the whole window)
window = tkinter.Tk()
window.title("Data Entry Form")
#Inside the window, inside the border
frame  = tkinter.Frame(window)
frame.pack()

#Subdividing the frame
#Will be using 3 frames: User Information, an unnamed middle frame and the bottom registration status frame

################################################
#------------Main GUI Interface----------------#
################################################

#Top Frame
#Create frame User Information
user_info_frame = tkinter.LabelFrame(frame, text="User Information")
user_info_frame.grid(row=0, column=0, padx=20, pady=20)

#First Last name data
first_name_label = tkinter.Label(user_info_frame, text="First Name")
first_name_label.grid(row=0, column=0)
last_name_label = tkinter.Label(user_info_frame, text="Last Name")
last_name_label.grid(row=0, column=1)
#First Last name data entry fields
first_name_entry = tkinter.Entry(user_info_frame)
last_name_entry = tkinter.Entry(user_info_frame)
first_name_entry.grid(row=1, column=0)
last_name_entry.grid(row=1, column=1)
#User picks Surname
title_label = tkinter.Label(user_info_frame, text="Title")
title_combobox = ttk.Combobox(user_info_frame, values=["Mr.", "Mrs.", "Dr."])
title_label.grid(row=0, column=2)
title_combobox.grid(row=1, column=2)
#Middle Frame
#User picks age
age_label = tkinter.Label(user_info_frame, text="Age")
age_spinbox = tkinter.Spinbox(user_info_frame, from_=18, to=110)
age_label.grid(row=2, column=0)
age_spinbox.grid(row=3, column=0)

#User picks nationality
nationality_label = tkinter.Label(user_info_frame, text="Nationality")
nationality_combobox = ttk.Combobox(user_info_frame, values=["Africa", "Antartica", "Asia", "Europe", "North America", "South America", "Oceania"])
nationality_label.grid(row=2, column=1)
nationality_combobox.grid(row=3, column=1)

#This is for-loop for all widgets in the frame to increase the spacing between the input lines
for widget in user_info_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

#Lower Frame
#This creates the courses frame, at the very bottom on of the screen
#The command sticky="news" ("news" = "north east south and west") makes frame expand in all directions
courses_frame = tkinter.LabelFrame(frame)
courses_frame.grid(row=1, column=0, sticky="news", padx=20, pady=10)

#Registration Status Checkbox
registered_label = tkinter.Label(courses_frame, text="Registration Status")
reg_status_var = tkinter.StringVar(value="Not Registered")
registered_check = tkinter.Checkbutton(courses_frame, text="Currently Registered", 
                                       variable=reg_status_var, onvalue="Registered", offvalue="Not Registered")
registered_label.grid(row=0, column=0)
registered_check.grid(row=1, column=0)

#This creates a spinbox on the bottom for number of completed courses
numcourses_label = tkinter.Label(courses_frame, text="# Completed Courses")
numcourses_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
numcourses_label.grid(row=0, column=1)
numcourses_spinbox.grid(row=1, column=1)
#This creates space between the "already registered" widget and numbered courses completed widgets
for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)
    
#This gives the # semesters attended    
numsemesters_label = tkinter.Label(courses_frame, text="# Semesters")
numsemesters_spinbox = tkinter.Spinbox(courses_frame, from_=0, to="infinity")
numsemesters_label.grid(row=0, column=2)
numsemesters_spinbox.grid(row=1, column=2)
#gives more pad spacing for # semesters attended
for widget in courses_frame.winfo_children():
    widget.grid_configure(padx=10, pady=5)

# Accept terms
terms_frame = tkinter.LabelFrame(frame, text="Terms & Conditions")
terms_frame.grid(row=2, column=0, sticky="news", padx=20, pady=10)
#Accept terms checkbox

accept_var = tkinter.StringVar(value="Not Accepted")
terms_check = tkinter.Checkbutton(terms_frame, text="I accept the terms and conditions.", 
                                  variable=accept_var, onvalue="Accepted", offvalue="Not Accepted")
terms_check.grid(row=0, column=0)

#Button
button = tkinter.Button(frame, text="Enter data", command = enter_data)
button.grid(row=3, column=0, sticky="news", padx=20, pady=10)

window.mainloop()
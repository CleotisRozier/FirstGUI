#This is the SQLite version of the mainInputGUI. It will create or append to a single sql .db file
#I created this following a tutorial playlist.
#Link: https://youtube.com/playlist?list=PLs3IFJPw3G9LQfo8GYbnPRaD8DKHj9aOJ&si=ipJZ6f54LpjBSiJL


import tkinter
from tkinter import ttk
from tkinter import messagebox
import sqlite3


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
        #Proceed if name requirements met
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
            
            
            #Begin SQL code 
            
            #Builds DB and populates headlines
            conn = sqlite3.connect('data.db')
            table_create_query = ''' CREATE TABLE IF NOT EXISTS Student_Data 
                                     (firstname TEXT, lastname TEXT, title TEXT, age INT, nationality TEXT, 
                                     registration_status TEXT, num_courses INT, num_semesters INT)'''
            conn.execute(table_create_query)
            
            #Data begins getting saved        
            data_insert_query = '''INSERT INTO Student_Data (firstname, lastname, title, 
                                age, nationality, registration_status, num_courses, num_semesters) 
                                VALUES (?,?,?,?,?,?,?,?)'''                                          #The question marks are the place holdrs for the variables
            data_insert_tuple = (firstname, lastname, title, age, nationality,                       #These variables replace the question marks
                                 registration_status, numcourses, numsemesters)
            
            cursor = conn.cursor()                                                                   #The cursor is what runs all of the inquiries and decides what goes where in the database
            cursor.execute(data_insert_query, data_insert_tuple)
            conn.commit()           
            conn.close()
            
            
            
            
            
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
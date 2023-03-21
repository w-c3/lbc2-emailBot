import customtkinter
import tkinter as tk
from tkinter import filedialog
import win32com.client as win32
import re
import tkinter.font as font

file_path = ""

def checkSchoolEmail():
        email = entry.get()
        for i in range (len(email) -1):
            if "loyolablakefield.org" in email:
                 return True
        return False

def checkFile():
    global file_path
    if file_path == "":
        tk.messagebox.showwarning(title= "Ooops", message="You need to add a file... ")
        file_select()
        

def submit():
    global successText
    checkFile()
    checkEmail = checkSchoolEmail()
    senderEmail = entry.get()
    sender = senderEmail.split("@")[0]
    school = entry1.get()
    principalEmail = entry2.get()
    name = entry3.get()
    gradClass = re.split(":|0", sender, maxsplit=1)[-1].strip()
    principalName = entry4.get()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = str(principalEmail)
    mail.cc = "cmhale2023@loyolablakefield.org"
    mail.Subject = 'LBC2 Invite Reminder'
    mail.Body = "Greetings " + str(principalName) + "!\nThis is a reminder that the Loyola Blakefield Cyber Challenge or better known as LBC2.\nWe are less than one week away from LBC2 and are excited to invite you to the event on Saturday March 25th!  Our staff and industry mentors are looking forward to possibly seeing you on March 26th. Everything you need to bring can be found on the site lbc2.org\nA few quick reminders that will help you prepare for the day:\n-Doors and registration open at 8:00am\n-Morning refreshments and lunch will be provided by Chick-Fil-A\n -Bring a laptop and your charger. Power and wireless internet will be available at each table\n-There will be lots of help during registration and over the course of the day if you have any questions!\n-Bring a headset – there may be some audio you will need to listen to\n-First time at an event like this?  There will be industry professionals available at each table to provide guidance and support\n-UMGC, CCBC, and The Maryland Army National Guard will have a team on site for you to visit with\n-Raffle prizes throughout the day!\n-The competition will be in Knott Hall in the 4-court gym. Directions and campus map can be found here\nReminder, registration closes on March 23rd.\nPlease let us know if you have any questions and share this email with your teammates!\nWe hope to see you there\nBest,\n" + str(name) + "\nClass of '" + str(gradClass)

    attachment  = file_path
    mail.Attachments.Add(attachment)

    if checkEmail == True:
         mail.Send()
         print("Successful email sent to " + str(school))
         successText.pack_forget()
         successText = customtkinter.CTkLabel(master=frame, text="Successful email sent to " + principalName + " at " + school)
         successText.pack()
         successText.configure(font = pdf_fontTuple)
         entry1.delete(0, "end")
         entry2.delete(0, "end")
         entry4.delete(0, "end")
    elif checkEmail == False:
         tk.messagebox.showwarning(title= "Ooops", message="You must use your Loyola email... ")
         entry.delete(0, "end")

def file_select():
    global file_path
    file_path = filedialog.askopenfilename()
    file = ""
    for i in range (len(file_path) -1):
        if file_path[-1 - i] == "/":
            file = file_path[0 - i:]
            break
    label6 = customtkinter.CTkLabel(master=frame, text="Selected file is.. " + file)
    label6.pack()
    label6.configure(font = pdf_fontTuple)
    
customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("dark-blue")
root = customtkinter.CTk()
root.title("LBC2 Email Sender")
root.geometry("450x620")
root.resizable(False, False)
frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="LBC2 Email Sender")
label.pack(pady = 20)

label0 = customtkinter.CTkLabel(master=frame, text="Your School Email:")
label0.pack()

entry = customtkinter.CTkEntry(master=frame, placeholder_text="ex. me@gmail.com")
entry.pack()

label1 = customtkinter.CTkLabel(master=frame, text="School Sending To:")
label1.pack()

entry1 = customtkinter.CTkEntry(master=frame, placeholder_text="ex. Loyola")
entry1.pack()

label2 = customtkinter.CTkLabel(master=frame, text="Pricipal Email:")
label2.pack()

entry2 = customtkinter.CTkEntry(master=frame, placeholder_text="ex. smith@school.org")
entry2.pack()

label3 = customtkinter.CTkLabel(master=frame, text="Your Name:")
label3.pack()

entry3 = customtkinter.CTkEntry(master=frame, placeholder_text="ex. Dylan")
entry3.pack()

label4 = customtkinter.CTkLabel(master=frame, text="Name of Principal:")
label4.pack()

entry4 = customtkinter.CTkEntry(master=frame, placeholder_text="ex. Ms. Smith")
entry4.pack()

submit_button = customtkinter.CTkButton(master=frame, text="Submit", command=submit)
submit_button.pack(pady=24)

label5 = customtkinter.CTkLabel(master=frame, text="Select the PDF Flyer you would \nlike to attach to the email:")
label5.pack()

file_select_button = customtkinter.CTkButton(master=frame, text="Select File", command=file_select)
file_select_button.pack(pady=12)

successText = customtkinter.CTkLabel(master=frame, text="")

title_fontTuple = ("Comic Sans MS", 28, "bold")
fontTuple = ("Comic Sans MS", 20)
pdf_fontTuple = ("Comic Sans MS", 12)

label.configure(font = title_fontTuple)
label0.configure(font = fontTuple)
label1.configure(font = fontTuple)
label2.configure(font = fontTuple)
label3.configure(font = fontTuple)
label4.configure(font = fontTuple)
label5.configure(font = pdf_fontTuple)
submit_button.configure(font = pdf_fontTuple)
file_select_button.configure(font = pdf_fontTuple)

root.mainloop()

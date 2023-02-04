import customtkinter
#import tkinter as tk
import win32com.client as win32
import re

def submit():
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
    mail.Subject = 'LBC2 Invite'
    mail.Body = "Dear " + str(principalName) + ",\nMy name is " + str(name) + " and I'm a student at Loyola Blakefield High School. I am a leader of Loyola Blakefield’s cybersecurity program, and I would love to invite your students at " + str(school) + " to our student-created cybersecurity competition, LBC2.\nLBC2 is an incredible opportunity for students interested in computers and cybersecurity, and absolutely no prior experience is required. The best part? The competition is entirely free! We would love to have some of your students attend! There will be many adult and student mentors in both in-person and virtual environments to assist our competitors and make the day a real learning experience. We will offer various prizes to winners, as well as breakfast and lunch offered to those participating. I have attached the flyer with more information about the competition, which is on March 25th, 2023, from 9:00am to 3:00pm on Loyola Blakefield’s campus or virtually. Registration for in-person attendance ends March 23rd and for virtual ends at the start of the competition. Register here at our website: https://lbc2.org/\nIf you have any questions at all, please don’t hesitate to reach out. Loyola Blakefield Cyber is excited to extend our outreach and education to you this year! Thank you for your time.\n\nBest,\n" + str(name) + "\nClass of '" + str(gradClass)

    attachment  = "C:/Users/" + str(sender) + "/Downloads/LBC2_2023.pdf"
    mail.Attachments.Add(attachment)

    mail.Send()
    successText = customtkinter.CTkLabel(master=frame, text="")
    successText = customtkinter.CTkLabel(master=frame, text="Successful email sent to " + str(school))
    successText.pack()
    print("Successful email sent to " + str(school))

customtkinter.set_appearance_mode("System")
customtkinter.set_default_color_theme("green")
root = customtkinter.CTk()
root.title("LBC2 Email Sender")
root.geometry("450x400")
frame = customtkinter.CTkFrame(master=root)
frame.pack(pady=20, padx=60, fill="both", expand=True)

label = customtkinter.CTkLabel(master=frame, text="Your Email:")
label.pack()

entry = customtkinter.CTkEntry(master=frame, placeholder_text="ex. me@gmail.com")
entry.pack()

label1 = customtkinter.CTkLabel(master=frame, text="School:")
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

root.mainloop()
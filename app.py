import tkinter as tk
import win32com.client as win32

def submit():
    school = entry1.get()
    principalEmail = entry2.get()
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = str(principalEmail)
    mail.Subject = 'LBC2 Invite'
    mail.Body = "Dear " + str(school) + ",\nMy name is Dylan Keler and I am a Junior at Loyola Blakefield High School. I am a leader of Loyola Blakefield’s cybersecurity program, and I would love to invite your students at " + str(school) + " to our student-created cybersecurity competition, LBC2.\nLBC2 is an incredible opportunity for students interested in computers and cybersecurity, and absolutely no prior experience is required. The best part? The competition is entirely free! We would love to have some of your students attend! There will be many adult and student mentors in both in-person and virtual environments to assist our competitors and make the day a real learning experience. We will offer various prizes to winners, as well as breakfast and lunch offered to those participating. I have attached the flyer with more information about the competition, which is on March 25th, 2023, from 9:00am to 3:00pm on Loyola Blakefield’s campus or virtually. Registration for in-person attendance ends March 23rd and for virtual ends at the start of the competition. Register here at our website: https://lbc2.org/\nIf you have any questions at all, please don’t hesitate to reach out. Loyola Blakefield Cyber is excited to extend our outreach and education you this year! Thank you for your time.\n\nBest,\nDylan Keller\nClass of 24"

    # To attach a file to the email (optional):
    attachment  = "C:/Users/dekeller2024/Desktop/LBC2_2023 (1).pdf"
    mail.Attachments.Add(attachment)

    mail.Send()
    print("Complete")

root = tk.Tk()
root.title("Simple Input Application")
root.geometry("800x600")


label1 = tk.Label(root, text="School:")
label1.pack()

entry1 = tk.Entry(root)
entry1.pack()

label2 = tk.Label(root, text="Pricipal Email:")
label2.pack()

entry2 = tk.Entry(root)
entry2.pack()

submit_button = tk.Button(root, text="Submit", command=submit)
submit_button.pack()

root.mainloop()
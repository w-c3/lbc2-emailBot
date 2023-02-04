import tkinter as tk

def submit():
    input1 = entry1.get()
    input2 = entry2.get()
    print("Input 1:", input1)
    print("Input 2:", input2)

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
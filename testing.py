import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_path = filedialog.askopenfilename()
file = ""
for i in range (len(file_path) -1):
    if file_path[-1 - i] == "/":
        file = file_path[0 - i:]
        break

print("file path is... " + file_path)
print("file is... " + file)
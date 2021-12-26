# Importing tkinter module 
# and all functions 
from tkinter import *
import tkinter 
from tkinter.ttk import *
  
def focus_next_widget(event):
    event.widget.tk_focusNext().focus()
    return("break")

window = Tk()
window.title("What's Your Message?")
window.configure(bg="black")
tkinter.Label(window, text="Type Your Message:\n", bg="Black", fg="white", font="none 25 bold").pack(anchor=N)

e = tkinter.Text(window, width=75, height=10,bg='white')
e.pack()
e.bind("<Tab>", focus_next_widget)
a = tkinter.Button(text="轉成excel檔")
a.pack()
a.bind("<Tab>", focus_next_widget)

window.mainloop()
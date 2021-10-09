import tkinter as tk
from tkinter import ttk

# --- functions ---

def newselection(event, other):
    print('selected:', event.widget.get())
    print('other:', other)

# --- main ---

root = tk.Tk()

cb1 = ttk.Combobox(root, values=('a', 'c', 'g', 't'))
cb1.pack()
cb1.bind("<<ComboboxSelected>>", lambda event:newselection(event, "Hello"))

cb2 = ttk.Combobox(root, values=('X', 'Y', 'XX', 'XY'))
cb2.pack()
cb2.bind("<<ComboboxSelected>>", lambda event:newselection(event, "World"))

root.mainloop()
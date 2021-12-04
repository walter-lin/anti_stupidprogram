from tkinter import *
import tkinter.ttk as ttk


class Application:

    def __init__(self, parent):
        self.parent = parent
        self.combo()

    def combo(self):
        self.box_value = StringVar()
        self.box = ttk.Combobox(self.parent, textvariable=self.box_value,font=("Helvetica",20))
        
        self.box['values'] = ('A', 'B', 'C')
        self.box.current(0)
        self.box.grid(column=0, row=0)

if __name__ == '__main__':
    root = Tk()
    app = Application(root)
    root.mainloop()
bigfont = tkFont.Font(family="Helvetica",size=20)
root.option_add("*TCombobox*Listbox*Font", bigfont)
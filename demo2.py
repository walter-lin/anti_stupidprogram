import tkinter as tk
import tkinter.ttk as ttk
 
root = tk.Tk()
 
last_text = ''
def text_change(event):
    global last_text
    got = text.get('1.0', 'end')
    if got != last_text:
        last_text = got
        print('文本被修改了')
 
text = tk.Text(root)
text.bind('<Key>', text_change)
text.grid(column=0,row=0)
lab = tk.Text(root)
lab.grid(column=0,row=1)
 
root.mainloop()
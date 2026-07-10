from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import Listbox, END
import os
os.environ['TCLLIBPATH'] = r'D:/python/tcl/tkdnd2.9'

def on_drop(event):
    print("拖拽事件触发", event.data)

root = TkinterDnD.Tk()
lb = Listbox(root)
lb.pack(fill='both', expand=True)
lb.drop_target_register(DND_FILES)
lb.dnd_bind('<<Drop>>', on_drop)
root.drop_target_register(DND_FILES)
root.dnd_bind('<<Drop>>', on_drop)
root.mainloop()
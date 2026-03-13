import tkinter as tk
import win32gui

class WindowSelector:
    def __init__(self, parent):
        self.parent=parent
        self.selected=None

    def open(self):
        self.titles=[]
        def cb(hwnd,_):
            if win32gui.IsWindowVisible(hwnd):
                t=win32gui.GetWindowText(hwnd)
                if t: self.titles.append((hwnd,t))
        win32gui.EnumWindows(cb,None)

        w=tk.Toplevel(self.parent)
        lb=tk.Listbox(w,width=60)
        lb.pack(fill="both",expand=True)
        for h,t in self.titles: lb.insert("end",t)

        def sel(e):
            i=lb.curselection()
            if i:
                self.selected=self.titles[i[0]]
                w.destroy()
        lb.bind("<Double-1>",sel)
        self.parent.wait_window(w)
        return self.selected

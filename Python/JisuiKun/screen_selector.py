import tkinter as tk

class ScreenRectSelector:
    def __init__(self, alpha=0.3):
        self.root=tk.Toplevel()
        self.root.attributes("-fullscreen",True)
        self.root.attributes("-alpha",alpha)
        self.root.configure(bg="black")
        self.canvas=tk.Canvas(self.root,bg="black",highlightthickness=0)
        self.canvas.pack(fill="both",expand=True)

        self.start=None
        self.rect=None
        self.coords=None

        self.canvas.bind("<Button-1>",self.start_sel)
        self.canvas.bind("<B1-Motion>",self.drag)
        self.canvas.bind("<ButtonRelease-1>",self.end_sel)
        self.root.protocol("WM_DELETE_WINDOW",self.on_close)

    def start_sel(self,e):
        self.start=(e.x,e.y)

    def drag(self,e):
        if self.rect: self.canvas.delete(self.rect)
        self.rect=self.canvas.create_rectangle(self.start[0],self.start[1],e.x,e.y,outline="red",width=2)

    def end_sel(self,e):
        x1,y1=self.start
        x2,y2=e.x,e.y
        self.coords=[min(x1,x2),min(y1,y2),max(x1,x2),max(y1,y2)]
        self.root.destroy()

    def on_close(self):
        self.coords=None
        self.root.destroy()

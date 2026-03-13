import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os, threading, time, csv, sys, re
import pyautogui, win32gui

from screen_selector import ScreenRectSelector
from window_selector import WindowSelector
from pdf_builder import build_pdf
from PIL import Image

APP_TITLE = "自炊君"
WINDOW_SIZE = "420x270"
PDF_OUTPUT_DIR = "pdfOutput"
CONFIG_CSV = "config.csv"

DEFAULT_PAGE_COUNT = 1
DEFAULT_INTERVAL = 1.0
DEFAULT_KEY = "Left"
IME_WAIT = 1.0
MAX_LABEL_LEN = 30
PROGRESS_LENGTH = 370

CAP_FILE_PATTERN = r'^cap_\d{4}(_\d+)?\.(png|pdf)$'

class App:
    def __init__(self):
        self.coords = None
        self.target_hwnd = None
        self.running = False

        self.root = tk.Tk()
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_SIZE)
        self.root.columnconfigure(0, weight=1)
        self.root.resizable(False, False)

        if getattr(sys,'frozen',False):
            self.base_dir = os.path.dirname(sys.executable)
        else:
            self.base_dir = os.path.dirname(os.path.abspath(__file__))

        self.default_save = os.path.join(self.base_dir, PDF_OUTPUT_DIR)
        os.makedirs(self.default_save, exist_ok=True)
        self.config_path = os.path.join(self.base_dir, CONFIG_CSV)

        self.build_ui()
        self.load_config()
        self.root.mainloop()

    def build_ui(self):
        row = 0

        # 保存先
        f = tk.Frame(self.root)
        f.grid(row=row, column=0, columnspan=4, sticky="we", padx=5)
        tk.Label(f, text="保存先").pack(side="left")
        self.save_entry = tk.Entry(f)
        self.save_entry.pack(side="left", fill="x", expand=True, padx=(5,5))
        self.save_entry.insert(0, self.default_save)
        tk.Button(f, text="参照", command=self.browse_dir, width=6).pack(side="left")
        # 保存先変更時に自動保存
        self.save_entry.bind("<FocusOut>", lambda e: self.save_config())
        row += 1

        # ページ設定
        f = tk.Frame(self.root)
        f.grid(row=row, column=0, columnspan=4, sticky="w", padx=5)
        tk.Label(f,text="ページ送り").grid(row=0,column=0)
        self.key_var = tk.StringVar(value=DEFAULT_KEY)
        tk.Radiobutton(f,text="左",variable=self.key_var,value="Left").grid(row=0,column=1)
        tk.Radiobutton(f,text="右",variable=self.key_var,value="Right").grid(row=0,column=2)
        # ページ送り変更時に自動保存
        self.key_var.trace_add("write", lambda *args: self.save_config())

        tk.Label(f,text="ページ数").grid(row=0,column=3,padx=(10,0))
        self.count = tk.Entry(f,width=5)
        self.count.insert(0,str(DEFAULT_PAGE_COUNT))
        self.count.grid(row=0,column=4)
        self.count.bind("<FocusOut>", lambda e: self.save_config())  # 自動保存

        tk.Label(f,text="間隔").grid(row=0,column=5,padx=(10,0))
        self.interval = tk.Entry(f,width=5)
        self.interval.insert(0,str(DEFAULT_INTERVAL))
        self.interval.grid(row=0,column=6)
        self.interval.bind("<FocusOut>", lambda e: self.save_config())  # 自動保存
        row += 1

        # 分割
        f = tk.Frame(self.root)
        f.grid(row=row,column=0,columnspan=4,sticky="w",padx=5)
        self.split_var = tk.BooleanVar(value=False)
        tk.Checkbutton(f,text="画面分割",variable=self.split_var,command=lambda:[self.update_split_state(), self.save_config()]).grid(row=0,column=0)
        self.split_order = tk.StringVar(value="右→左")
        self.split_combo = ttk.Combobox(f,textvariable=self.split_order,values=["右→左","左→右"],state="disabled",width=7)
        self.split_combo.grid(row=0,column=1)
        # 分割順序変更時も自動保存
        self.split_order.trace_add("write", lambda *args: self.save_config())
        row += 1

        # 座標
        self.coord_btn = tk.Button(self.root, text="座標選択", command=lambda:[self.set_coords(), self.save_config()], width=12)
        self.coord_btn.grid(row=row, column=0, sticky="w", padx=5)
        self.coord_label = tk.Label(text="未設定", width=40, anchor="w")
        self.coord_label.grid(row=row, column=1, columnspan=3, sticky="w")
        row += 1

        # ウィンドウ（保存しない）
        self.window_btn = tk.Button(self.root, text="ウィンドウ選択", command=self.select_window, width=12)
        self.window_btn.grid(row=row, column=0, sticky="w", padx=5)
        self.window_label = tk.Label(text="未選択", width=40, anchor="w")
        self.window_label.grid(row=row, column=1, columnspan=3, sticky="w")
        row += 1

        # 表紙
        self.cover_var = tk.BooleanVar(value=False)
        tk.Checkbutton(self.root,text="表紙画像",variable=self.cover_var,command=self.update_cover_state).grid(row=row,column=0,sticky="w",padx=5)
        row += 1
        f = tk.Frame(self.root)
        f.grid(row=row,column=0,columnspan=4,sticky="we",padx=5)
        self.cover_entry = tk.Entry(f,state="disabled")
        self.cover_entry.pack(side="left", fill="x", expand=True, padx=(0,5))
        self.cover_btn = tk.Button(f,text="参照",command=self.browse_cover,width=6,state="disabled")
        self.cover_btn.pack(side="left")
        row += 1



        # 実行
        self.start_btn = tk.Button(self.root,text="開始",command=lambda:threading.Thread(target=self.run_process,daemon=True).start(),height=2,font=("TkDefaultFont",10,"bold"))
        self.start_btn.grid(row=row,column=0,columnspan=4,sticky="we",pady=8,padx=5)
        row += 1

        # プログレスバー
        self.progress = ttk.Progressbar(self.root,length=PROGRESS_LENGTH)
        self.progress.grid(row=row,column=0,columnspan=4,sticky="w",padx=5)
        self.progress['value'] = 0
        # 停止ボタン
        self.stop_btn = tk.Button(self.root,text="停止",command=self.stop_process,state="disabled")
        self.stop_btn.grid(row=row, column=0, columnspan=4, sticky="e", padx=(0,5))

    # ---------- UI制御 ----------
    def update_cover_state(self):
        state="normal" if self.cover_var.get() else "disabled"
        self.cover_entry.config(state=state)
        self.cover_btn.config(state=state)

    def update_split_state(self):
        self.split_combo.config(state="readonly" if self.split_var.get() else "disabled")

    # ---------- utils ----------
    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.save_entry.delete(0,"end")
            self.save_entry.insert(0,d)
            self.save_config()

    def browse_cover(self):
        f = filedialog.askopenfilename(filetypes=[("Image","*.png;*.jpg;*.jpeg")])
        if f:
            self.cover_entry.delete(0,"end")
            self.cover_entry.insert(0,f)
            self.save_config()

    def set_coords(self):
        # ボタンを無効化
        self.coord_btn.config(state="disabled")

        sel = ScreenRectSelector()
        self.root.wait_window(sel.root)  # 選択が終わるまで待機

        if sel.coords:
            self.coords = sel.coords
            self.coord_label.config(text=str(self.coords))
            self.save_config()  # 選択直後に保存

        # ボタンを元に戻す
        self.coord_btn.config(state="normal")

    def select_window(self):
        # ボタンを無効化
        self.window_btn.config(state="disabled")

        ws = WindowSelector(self.root)
        res = ws.open()
        if res:
            self.target_hwnd, title = res
            self.window_label.config(text=title[:MAX_LABEL_LEN])

        # ボタンを元に戻す
        self.window_btn.config(state="normal")


    def activate_window(self):
        if not self.target_hwnd: return False
        win32gui.SetForegroundWindow(self.target_hwnd)
        time.sleep(IME_WAIT)
        return True

    # ---------- CSV ----------
    def load_config(self):
        if not os.path.exists(self.config_path): return
        try:
            with open(self.config_path,newline="",encoding="utf-8") as f:
                row = next(csv.reader(f))
                self.save_entry.delete(0,"end")
                self.save_entry.insert(0,row[0])
                self.key_var.set(row[1])
                self.count.delete(0,"end")
                self.count.insert(0,row[2])
                self.interval.delete(0,"end")
                self.interval.insert(0,row[3])
                self.split_var.set(row[4]=="True")
                self.update_split_state()
                self.split_order.set(row[5])
                self.cover_var.set(row[6]=="True")
                self.update_cover_state()
                self.cover_entry.delete(0,"end")
                self.cover_entry.insert(0,row[7])
                # coords
                if len(row)>8 and row[8]:
                    try:
                        self.coords = tuple(map(int,row[8].split(",")))
                        self.coord_label.config(text=str(self.coords))
                    except:
                        self.coords=None
                        self.coord_label.config(text="未設定")
        except:
            pass

    def save_config(self):
        try:
            coords_str = ",".join(map(str,self.coords)) if self.coords else ""
            with open(self.config_path,"w",newline="",encoding="utf-8") as f:
                csv.writer(f).writerow([
                    self.save_entry.get(),
                    self.key_var.get(),
                    self.count.get(),
                    self.interval.get(),
                    str(self.split_var.get()),
                    self.split_order.get(),
                    str(self.cover_var.get()),
                    self.cover_entry.get(),
                    coords_str
                ])
        except:
            pass
            
    def prepare_cover(self, cover_path):
        from PIL import Image, ImageOps

        # スクショ幅・高さ取得（分割を考慮）
        x1, y1, x2, y2 = self.coords
        full_width = x2 - x1
        full_height = y2 - y1
        target_width = full_width // 2 if self.split_var.get() else full_width
        target_height = full_height

        # 表紙画像を開く
        img = Image.open(cover_path)
        w, h = img.size

        # 幅をtarget_widthに合わせる
        new_h = int(h * (target_width / w))
        img = img.resize((target_width, new_h), Image.LANCZOS)

        # 高さが不足している場合は上下に黒帯を追加
        if new_h < target_height:
            top_bottom = (target_height - new_h) // 2
            img = ImageOps.expand(img, border=(0, top_bottom), fill="black")

            # 端数分を下に追加（余りがある場合）
            extra = target_height - img.size[1]
            if extra > 0:
                img = ImageOps.expand(img, border=(0, 0, 0, extra), fill="black")

        # 保存先はPDF出力先フォルダに cap_0000.png として
        save_dir = self.save_entry.get()
        os.makedirs(save_dir, exist_ok=True)
        temp_path = os.path.join(save_dir, "cap_0000.png")
        img.save(temp_path)

        return temp_path


    # ---------- 停止 ----------
    def stop_process(self):
        self.running = False

    # ---------- main ----------
    def run_process(self):
        if not self.coords:
            messagebox.showerror("エラー","座標未設定")
            return
        if not self.activate_window():
            messagebox.showerror("エラー","ウィンドウ未選択")
            return

        try:
            count=int(self.count.get())
            interval=float(self.interval.get())
        except:
            messagebox.showerror("エラー","ページ数または間隔が不正です")
            return

        # 分割補正
        if self.split_var.get():
            count = (count + 1) // 2  # 分割で2倍になるので半分に

        # 保存先ディレクトリを作成（存在しなければ作る）
        save = self.save_entry.get()
        os.makedirs(save, exist_ok=True)

        # ---------- 既存キャプチャ画像・PDF削除 ----------
        pattern = re.compile(CAP_FILE_PATTERN, re.IGNORECASE)
        for f in os.listdir(save):
            if pattern.match(f):
                try:
                    os.remove(os.path.join(save, f))
                except Exception as e:
                    print(f"削除失敗: {f} - {e}")

        self.running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progress['value'] = 0

        files=[]
        x1,y1,x2,y2=self.coords

        for i in range(1,count+1):
            if not self.running: break

            img = pyautogui.screenshot(region=(x1,y1,x2-x1,y2-y1))

            if self.split_var.get():
                parts = self.split_image(img)
                for j,p in enumerate(parts):
                    n = f"cap_{i:04}_{j+1}.png"
                    p.save(os.path.join(save,n))
                    files.append(n)
            else:
                n = f"cap_{i:04}.png"
                img.save(os.path.join(save,n))
                files.append(n)

            pyautogui.press(self.key_var.get())
            self.progress['value'] = i / count * 100
            self.root.update_idletasks()
            time.sleep(interval)

        cover = None
        if self.cover_var.get() and os.path.exists(self.cover_entry.get()):
            cover = self.prepare_cover(self.cover_entry.get())

        if self.running:
            build_pdf(save, sorted(files), cover)
            messagebox.showinfo("完了","PDF作成完了")

        self.running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.progress['value'] = 0

    def split_image(self,img):
        w,h = img.size
        mid = w//2
        l = img.crop((0,0,mid,h))
        r = img.crop((mid,0,w,h))
        return [r,l] if self.split_order.get()=="右→左" else [l,r]


import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from ttkthemes import ThemedTk
from PIL import Image, ImageDraw, ImageFont
import pandas as pd

def crop_hukum(img): return img.crop((88, 180, *img.size))
def crop_tahanapi(img): return img.crop((88, 170, *img.size))

def add_text(img, text, pos, size=28):
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("arial.ttf", size)
    draw.text(pos, text, font=font, fill="black")
    return img

def add_tanggal_hukum(img, tanggal):
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype("arial.ttf", 28)
    draw.text((220, 640), tanggal.strftime("%Y%m%d") + "-GJ", fill="black", font=font)
    draw.text((475, 655), tanggal.strftime("%m"), fill="black", font=font)
    draw.text((570, 655), tanggal.strftime("%Y"), fill="black", font=font)
    return img

class LabelApp:
    def __init__(self, root):
        self.root = root
        root.title("Label Generator")
        root.geometry("640x420")
        root.resizable(False, False)

        self.excel_path = tk.StringVar()
        self.sumber_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        style = ttk.Style()
        style.configure("TButton", padding=6, font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))
        style.configure("TEntry", font=("Segoe UI", 10))
        style.configure("TProgressbar", thickness=8)

        def make_row(text, var, row, cmd):
            ttk.Label(self.root, text=text).grid(row=row, column=0, padx=10, pady=6, sticky="e")
            ttk.Entry(self.root, textvariable=var, width=50).grid(row=row, column=1, padx=5)
            ttk.Button(self.root, text="Browse", command=cmd).grid(row=row, column=2, padx=5)

        make_row("File Excel:", self.excel_path, 0, self.browse_excel)
        make_row("Folder SUMBER:", self.sumber_path, 1, self.browse_sumber)
        make_row("Folder OUTPUT:", self.output_path, 2, self.browse_output)

        ttk.Button(self.root, text="🚀 Mulai Proses", command=self.proses).grid(row=3, column=0, columnspan=3, pady=15)

        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.grid(row=4, column=0, columnspan=3, padx=10, pady=5, sticky="ew")

        self.log = tk.Text(self.root, height=10, font=("Consolas", 9), state="disabled", bg="#f7f7f7")
        self.log.grid(row=5, column=0, columnspan=3, padx=10, pady=5, sticky="nsew")

    def browse_excel(self): self.excel_path.set(filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")]))
    def browse_sumber(self): self.sumber_path.set(filedialog.askdirectory())
    def browse_output(self): self.output_path.set(filedialog.askdirectory())

    def log_message(self, msg):
        self.log.config(state="normal")
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state="disabled")

    def proses(self):
        try:
            df = pd.read_excel(self.excel_path.get())
        except Exception as e:
            messagebox.showerror("Error", f"Gagal baca Excel:\n{e}")
            return

        total = len(df)
        self.progress["maximum"] = total

        for i, row in df.iterrows():
            try:
                sku = str(row["SKU"])
                tanggal = pd.to_datetime(row["TANGGAL"])
                jadwal = str(row["JADWAL"]).strip()

                base = os.path.join(self.output_path.get(), jadwal)
                hukum_folder = os.path.join(base, "法律标HUKUM")
                tahan_folder = os.path.join(base, "防火标TAHAN_API")
                os.makedirs(hukum_folder, exist_ok=True)
                os.makedirs(tahan_folder, exist_ok=True)

                hukum_file = sku + "-法律标.jpg"
                tahan_file = sku + "-防火标.jpg"

                hukum_path = next((os.path.join(r, hukum_file) for r, _, f in os.walk(self.sumber_path.get()) if hukum_file in f), None)
                tahan_path = next((os.path.join(r, tahan_file) for r, _, f in os.walk(self.sumber_path.get()) if tahan_file in f), None)

                if hukum_path:
                    img = Image.open(hukum_path).convert("RGB")
                    img = crop_hukum(img)
                    img = add_tanggal_hukum(img, tanggal)
                    img.save(os.path.join(hukum_folder, f"{sku}_法律标HUKUM.pdf"), "PDF")
                    self.log_message(f"[OK] {sku} HUKUM")
                else:
                    self.log_message(f"[MISS] {sku} - HUKUM not found")

                if tahan_path:
                    img = Image.open(tahan_path).convert("RGB")
                    img = crop_tahanapi(img)
                    label = tanggal.strftime("%Y%m%d-GJ")
                    img = add_text(img, label, (130, 240))
                    img.save(os.path.join(tahan_folder, f"{sku}_防火标TAHANAPI.pdf"), "PDF")
                    self.log_message(f"[OK] {sku} TAHAN_API")
                else:
                    self.log_message(f"[MISS] {sku} - TAHAN_API not found")

            except Exception as e:
                self.log_message(f"[ERR] {sku}: {e}")

            self.progress["value"] = i + 1
            self.root.update_idletasks()

        messagebox.showinfo("Selesai", "Proses selesai!")

if __name__ == "__main__":
    app = ThemedTk(theme="arc")
    LabelApp(app)
    app.mainloop()

import os
from PIL import Image, ImageDraw, ImageFont
import pandas as pd
from datetime import datetime

# --- Pengaturan tetap ---
SHEET_PATH = "Data.xlsx"
SUMBER_PATH = r"Z:\2503 防火标&法律标设计图"
OUTPUT_PATH = r"C:\Users\pupu\Documents\HEHE"
FONT_PATH = "arial.ttf"

def crop_hukum(img): return img.crop((0, 0, *img.size))
def crop_tahanapi(img): return img.crop((0, 0, *img.size))

def add_text(img, text, pos, size=30):
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(FONT_PATH, size)
    draw.text(pos, text, font=font, fill="black")
    return img

def add_tanggal_hukum(img, tanggal):
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(FONT_PATH, 28)
    draw.text((320, 825), tanggal.strftime("%Y%m%d") + "-GJ", fill="black", font=font)
    draw.text((615, 845), tanggal.strftime("%m"), fill="black", font=font)
    draw.text((735, 845), tanggal.strftime("%Y"), fill="black", font=font)
    return img

def cari_file(folder_awal, nama_file):
    for root, _, files in os.walk(folder_awal):
        if nama_file in files:
            return os.path.join(root, nama_file)
    return None

def proses():
    try:
        df = pd.read_excel(SHEET_PATH)
    except Exception as e:
        print(f"[ERROR] Gagal baca Excel: {e}")
        return

    total = len(df)
    print(f"[INFO] Memproses {total} baris data...")

    for i, row in df.iterrows():
        try:
            sku = str(row["SKU"])
            tanggal = pd.to_datetime(row["TANGGAL"])
            jadwal = str(row["JADWAL"]).strip()

            base = os.path.join(OUTPUT_PATH, jadwal)
            hukum_folder = os.path.join(base, "法律标HUKUM")
            tahan_folder = os.path.join(base, "防火标TAHAN_API")
            os.makedirs(hukum_folder, exist_ok=True)
            os.makedirs(tahan_folder, exist_ok=True)

            hukum_file = sku + "-法律标.jpg"
            tahan_file = sku + "-防火标.jpg"

            hukum_path = cari_file(SUMBER_PATH, hukum_file)
            tahan_path = cari_file(SUMBER_PATH, tahan_file)

            if hukum_path:
                img = Image.open(hukum_path).convert("RGB")
                img = crop_hukum(img)
                img = add_tanggal_hukum(img, tanggal)
                img.save(os.path.join(hukum_folder, f"{sku}_法律标HUKUM.pdf"), "PDF")
                print(f"[OK] {sku} HUKUM")
            else:
                print(f"[MISS] {sku} - HUKUM not found")

            if tahan_path:
                img = Image.open(tahan_path).convert("RGB")
                img = crop_tahanapi(img)
                label = tanggal.strftime("%Y%m%d-GJ")
                img = add_text(img, label, (180, 310))
                img.save(os.path.join(tahan_folder, f"{sku}_防火标TAHANAPI.pdf"), "PDF")
                print(f"[OK] {sku} TAHAN_API")
            else:
                print(f"[MISS] {sku} - TAHAN_API not found")

        except Exception as e:
            print(f"[ERR] {sku}: {e}")

    print("[DONE] Semua proses selesai.")

if __name__ == "__main__":
    proses()

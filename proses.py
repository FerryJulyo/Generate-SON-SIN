import pandas as pd
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from tkinter import ttk
import threading
from pathlib import Path

class SongProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SON SIN Generator")
        self.root.geometry("800x600")
        
        # Variabel untuk mengontrol thread
        self.processing = False
        
        self.file_path = tk.StringVar()
        self.create_widgets()

    def create_widgets(self):
        frm = ttk.Frame(self.root, padding=10)
        frm.pack(fill='both', expand=True)

        # File selection
        ttk.Label(frm, text="File Database (.xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.file_path, width=70).grid(row=0, column=1, padx=5)
        ttk.Button(frm, text="Browse...", command=self.browse_file).grid(row=0, column=2)

        # Proses button
        self.process_btn = ttk.Button(frm, text="Proses", command=self.start_processing_thread)
        self.process_btn.grid(row=1, column=1, pady=10)

        # Log area
        self.log_text = scrolledtext.ScrolledText(frm, height=25)
        self.log_text.grid(row=2, column=0, columnspan=3, sticky="nsew", pady=(10,0))

        frm.rowconfigure(2, weight=1)
        frm.columnconfigure(1, weight=1)

    def log(self, msg):
        self.log_text.insert(tk.END, msg + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Pilih file Excel Lagu",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.file_path.set(filename)

    def start_processing_thread(self):
        """Memulai proses di thread terpisah"""
        if self.processing:
            return
            
        file_path = self.file_path.get()
        if not file_path:
            messagebox.showerror("Error", "Silakan pilih file terlebih dahulu.")
            return

        # Disable tombol proses selama processing
        self.process_btn.config(state=tk.DISABLED)
        self.processing = True
        
        # Bersihkan log
        self.log_text.delete(1.0, tk.END)
        
        # Jalankan di thread terpisah
        thread = threading.Thread(target=self.process_file, daemon=True)
        thread.start()
        
        # Cek status thread secara periodik
        self.check_thread_status(thread)

    def check_thread_status(self, thread):
        """Memeriksa status thread dan update UI"""
        if thread.is_alive():
            # Jika thread masih berjalan, cek lagi setelah 100ms
            self.root.after(100, lambda: self.check_thread_status(thread))
        else:
            # Thread selesai, update UI
            self.processing = False
            self.process_btn.config(state=tk.NORMAL)

    def process_file(self):
        """Fungsi utama untuk memproses file (dijalankan di thread terpisah)"""
        try:
            file_path = self.file_path.get()
            ref_path = os.path.join("Reference", "Reference.xlsx")
            
            if not os.path.exists(ref_path):
                self.root.after(0, lambda: messagebox.showerror("Error", f"File referensi tidak ditemukan di: {ref_path}"))
                return

            self.log(f"🔍️ Membaca file: {file_path}")
            LIMIT_ROWS = None  # Bisa diubah jika mau batasi baris

            # Load semua sheet
            song_df = pd.read_excel(file_path, sheet_name='Song')
            sing_df = pd.read_excel(file_path, sheet_name='Sing')
            ref_df = pd.read_excel(ref_path, sheet_name='Reference', header=None)
            lang_ref_df = pd.read_excel(ref_path, sheet_name='Ref2', header=None, names=["IHP_CODE", "STARNET_CODE"])
            refsing_df = pd.read_excel(ref_path, sheet_name='RefSing')


            if LIMIT_ROWS:
                song_df = song_df.head(LIMIT_ROWS)

            def normalize_id(x):
                try:
                    if pd.isna(x): return ''
                    return str(int(float(x))).strip()
                except: return ''

            for col in ['SingId1', 'SingId2', 'SingId3', 'SingId4']:
                song_df[col] = song_df[col].apply(normalize_id)
            sing_df['SingId'] = sing_df['SingId'].apply(normalize_id)

            volume_convert = ref_df.iloc[1:30, [0, 1]].dropna()
            volume_convert.columns = ["FFMPEG_VOLUME", "STARNET_VOLUME"]
            volume_convert["FFMPEG_VOLUME"] = volume_convert["FFMPEG_VOLUME"].astype(str).str.extract(r"(-?\d+)")
            volume_convert["STARNET_VOLUME"] = pd.to_numeric(volume_convert["STARNET_VOLUME"], errors='coerce')

            lang_ref_df["IHP_CODE"] = lang_ref_df["IHP_CODE"].astype(str).str.strip().str.upper()
            genre_col = ref_df.columns[ref_df.iloc[0] == "SongTypeID Convert"][0]
            genre_ref = ref_df.iloc[1:70, [genre_col, genre_col + 1]].dropna()
            genre_ref.columns = ["GENRE_NAME", "GENRE_ID"]
            genre_ref["GENRE_NAME"] = genre_ref["GENRE_NAME"].astype(str).str.strip().str.lower()

            def get_vol_ref(ffmpeg_value):
                try:
                    if pd.isna(ffmpeg_value): return ''
                    val = int(str(ffmpeg_value).split('.')[0])
                    match = volume_convert.loc[volume_convert["FFMPEG_VOLUME"] == str(val), "STARNET_VOLUME"]
                    return str(int(match.values[0])) if not match.empty else ''
                except: return ''

            def get_lang_ref(code):
                try:
                    if pd.isna(code): return ''
                    code = str(code).strip().upper()
                    match = lang_ref_df.loc[lang_ref_df["IHP_CODE"] == code, "STARNET_CODE"]
                    return str(match.values[0]) if not match.empty else ''
                except: return ''

            def get_genre_ref(name):
                try:
                    if pd.isna(name): return -1
                    name = str(name).strip().lower()
                    match = genre_ref.loc[genre_ref["GENRE_NAME"] == name, "GENRE_ID"]
                    return int(match.values[0]) if not match.empty else -1
                except: return -1

            def get_singer_name(sing_id, song_id):
                try:
                    if not sing_id: return ''
                    row = sing_df.loc[sing_df["SingId"] == sing_id]
                    if row.empty: return ''
                    song_id = str(song_id).zfill(8)
                    prefix = song_id[:2]
                    if prefix in ['01', '1', '02', '2', '06', '6', '07', '7', '91', '92']:
                        return row["Sing"].values[0]
                    elif prefix in ['03', '3', '93']:
                        return row["OriginalSing"].values[0]
                    elif prefix in ['04', '4', '05', '5', '08', '8']:
                        return row["RomanSing"].values[0]
                    return row["Sing"].values[0]
                except: return ''


            self.log("📦 Membuat SONGLIST.son...")
            son_lines = []

            for _, row in song_df.iterrows():
                    
                volref = get_vol_ref(row.get("FFMpeg"))
                songlangref = get_lang_ref(row.get("SongLan"))
                genre_ids = [get_genre_ref(row.get(g)) for g in ['Genre1', 'Genre2', 'Genre3', 'Genre4']]
                singer_names = [get_singer_name(row[s], row['SongId']) for s in ['SingId1', 'SingId2', 'SingId3', 'SingId4']]
                singer_string = ', '.join([s for s in singer_names if s])

                line = (
                    f"MUSIC||{volref}||{row['SongId']}||{row['SongId']}.{row['Format']}||{row['Song']}||{row['PYStr1']}||"
                    f"{row['SongLen']}||{row['SongType']}||{row['SongLan']}||{singer_string}||1||2||{songlangref}||-1||-1||-1||"
                    f"{row['SingId1'] or -1}||{row['SingId2'] or -1}||{row['SingId3'] or -1}||{row['SingId4'] or -1}||"
                    f"{genre_ids[0]}||{genre_ids[1]}||{genre_ids[2]}||{genre_ids[3]}||C||291308162||GOOD||GOOD||GOOD||GOOD||{row['SongId']}||1||1"
                )
                son_lines.append(line)

            out_dir = os.path.dirname(file_path)
            with open(os.path.join(out_dir, "SONGLIST.son"), "w", encoding="utf-8") as f:
                f.write("\n".join(son_lines))
            self.log("✅ SONGLIST.son selesai")


            self.log("📦 Membuat songinfo.txt...")
            composer_cols = [f"COMPOSER{i}" for i in range(1, 11)]
            lines = []
            for _, row in song_df.iterrows():
                    
                composers = [str(row.get(c)).strip() for c in composer_cols if pd.notna(row.get(c)) and str(row.get(c)).strip()]
                line = f"{row['SongId']}||{row.get('OriginalSong', '')}||||{', '.join(composers)}||{row.get('PYStr1', '')}"
                lines.append(line)
            with open(os.path.join(out_dir, "songinfo.txt"), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            self.log("✅ songinfo.txt selesai")


            self.log("📦 Membuat SINGERLIST.sin...")
            def get_sing_type_id(country, sex):
                match = refsing_df[(refsing_df["Src1"] == str(country).strip()) & (refsing_df["Src2"] == str(sex).strip())]
                return str(int(match["Data"].values[0])) if not match.empty else ''
            lines = []
            for _, row in sing_df.iterrows():
                    
                line = f"{row['SingId']}||{row['Sing']}||{row['PYStr']}||{get_sing_type_id(row.get('SingCountry', ''), row.get('SingSex', ''))}||1"
                lines.append(line)
            with open(os.path.join(out_dir, "SINGERLIST.sin"), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            self.log("✅ SINGERLIST.sin selesai")


            self.log("📦 Membuat singerinfo.txt...")
            lines = []
            for _, row in sing_df.iterrows():
                    
                line = f"{row['SingId']}||{row.get('OriginalSing', '')}||{row.get('RomanSing', '')}||{row.get('PYStr', '')}"
                lines.append(line)
            with open(os.path.join(out_dir, "singerinfo.txt"), "w", encoding="utf-8") as f:
                f.write("\n".join(lines))
            self.log("✅ singerinfo.txt selesai")


            self.log("📦 Memproses Delete Song...")
            try:
                delete_df = pd.read_excel(file_path, sheet_name="Delete Song", usecols=["SongId"])
                delete_df["CleanSongId"] = delete_df["SongId"].astype(str).apply(lambda x: re.sub(r"\s*\(.*?\)", "", x))
                delete_df["ValidSongId"] = delete_df["CleanSongId"].apply(lambda val: re.match(r"^\d{8}[A-Z]?$", str(val).strip()).group(0) if re.match(r"^\d{8}[A-Z]?$", str(val).strip()) else '')
                delete_df = delete_df[delete_df["ValidSongId"] != ""]

                existing_ids = set(song_df["SongId"].astype(str).str.strip())
                enable = sorted(set(delete_df[delete_df["ValidSongId"].isin(existing_ids)]["SongId"].astype(str)))
                disable = sorted(set(delete_df[~delete_df["ValidSongId"].isin(existing_ids)]["SongId"].astype(str)))

                with open(os.path.join(out_dir, "ENABLESONG.cbso"), "w", encoding="utf-8") as f:
                    f.write("\n".join(enable) + "\n")
                with open(os.path.join(out_dir, "DISABLESONG.bso"), "w", encoding="utf-8") as f:
                    f.write("\n".join(disable) + "\n")

                self.log("✅ ENABLESONG.cbso dan DISABLESONG.bso selesai")
            except Exception as e:
                self.log(f"❌ Sheet Delete Song error: {e}")

            self.log("✅ Semua proses selesai!")

        except Exception as e:
            self.log(f"❌ Terjadi kesalahan: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.process_btn.config(state=tk.NORMAL))
            self.processing = False

if __name__ == "__main__":
    root = tk.Tk()
    root.iconbitmap("icon.ico")
    app = SongProcessorApp(root)
    root.mainloop()
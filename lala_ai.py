import edge_tts
import asyncio
import pygame
import os
import speech_recognition as sr
from interpreter import interpreter
import subprocess
import pyautogui
import time
import json
import random
import webbrowser
import urllib.parse
import urllib.request
import re
import datetime
import shutil
import csv
import glob
from pywinauto.application import Application
import win32com.client
import tkinter as tk
import threading

VERSI_LALA = "2026-11-09"
URL_GITHUB_RAW = "https://raw.githubusercontent.com/zemecksis/lala_ai/refs/heads/main/lala_ai.py"

os.environ['PYGAME_HIDE_SUPPORT_PROMPT'] = "hide"

interpreter.offline = True 
interpreter.llm.model = "ollama/llama3.1"
interpreter.auto_run = False
interpreter.llm.max_tokens = 4096 

interpreter.custom_instructions = """You are an Intent Router. You only respond, think, and answer in English.
Output ONLY valid JSON containing "action", "target" (if needed), and "new_tab" (boolean).

RULES:
1. CLOSE TAB: If user says "tutup youtube", "tutup tab"
Output: {"action": "close_tab"}

2. CLOSE APP: If user says "tutup notepad"
Output: {"action": "close_app", "target": "app_name"}

3. OPEN APP: If user says "buka obs"
Output: {"action": "open_app", "target": "app_name"}

4. YOUTUBE SEARCH: If user says "cari [TOPIC] di youtube"
Output: {"action": "youtube", "target": "[TOPIC]", "new_tab": false}
If user explicitly says "di tab baru", set "new_tab": true.

5. YOUTUBE PLAY: If user says "putar urutan ke 3" or "putar video kedua"
Output: {"action": "youtube_play", "index": 3}

6. BROWSER NAVIGATION: If user says "kembali ke page sebelumnya" or "page selanjutnya"
Output: {"action": "browser_nav", "direction": "back"} OR {"action": "browser_nav", "direction": "forward"}

7. WEB SEARCH: If user says "cari [TOPIC] di internet", "cari informasi [TOPIC]", or "cari [TOPIC] di google"
Output: {"action": "web_search", "target": "[TOPIC]"}

8. GOOGLE OPEN LINK: If user says "buka urutan ke 2" or "buka link pertama"
Output: {"action": "google_open", "index": 2}

9. SCROLLING: If user says "page down", "scroll ke bawah", "page up", or "scroll ke atas"
Output: {"action": "scroll", "direction": "down"} OR {"action": "scroll", "direction": "up"}

10. OPEN WEBSITE: If user says "buka [WEBSITE] di browser" (e.g., youtube, facebook, whatsapp)
Output: {"action": "open_website", "target": "[WEBSITE]"}

11. VOLUME CONTROL: If user says "naikan volume [X] kali" or "turunkan volume [X] kali"
Output: {"action": "volume", "direction": "up", "amount": X} OR {"action": "volume", "direction": "down", "amount": X}
(If no amount is specified, set amount to 5)

12. LOCAL MUSIC: If user says "putar lagu [QUERY]" or "putar [QUERY]" 
(CRITICAL: If the query is EXACTLY "play", "pause", or "next song", IGNORE this instruction and use MEDIA CONTROLS).
Output: {"action": "play_local_music", "query": "[QUERY]"}

13. MEDIA CONTROLS & QUEUE:
If user says exactly "pause" or "play"
Output: {"action": "media_control", "command": "playpause"}
If user says "next song"
Output: {"action": "media_control", "command": "nexttrack"}
If user says "masukan [QUERY] ke playlist" or "tambahkan [QUERY] ke playlist"
Output: {"action": "queue_music", "query": "[QUERY]"}

14. FILE MANAGEMENT (COPY/MOVE):
If user says anything like "copy file [A] ke [B]", "copy folder [A] ke folder [B]", or "kopi [A] ke [B]"
Output: {"action": "salin_file", "type": "file" (or "folder"), "source": "[A]", "target": "[B]"}

If user says anything like "pindahkan file [A] ke [B]", "pindahkan folder [A] ke folder [B]", or "cut [A] ke [B]"
Output: {"action": "pindah_file", "type": "file" (or "folder"), "source": "[A]", "target": "[B]"}

CRITICAL: NEVER use "execute" or PowerShell for this. ALWAYS output the JSON above. Determine if it's a file or folder and set "type". Remove words like "file" or "folder" from the source and target names.

15. WHATSAPP SENDER & OPENER:
If user says "kirim wa ke [CONTACT] isinya [MESSAGE]"
Output: {"action": "send_whatsapp", "contact": "[CONTACT]", "message": "[MESSAGE]"}
If user says "buka wa dari [CONTACT]" or "buka whatsapp dari [CONTACT]"
Output: {"action": "open_whatsapp_chat", "contact": "[CONTACT]"}

16. MENGETIK DI NOTEPAD (KHUSUS NOTEPAD):
ATURAN MUTLAK NOTEPAD

Jika pengguna berkata "BUKA" (contoh: lala buka notepad judul puisi):
Tugasmu HANYA mencari file notepad yang sudah ada. DILARANG KERAS menggunakan aksi buat_notepad.
Format: {"action": "buka_notepad", "judul": "[nama_file_tanpa_txt]"}

Jika pengguna berkata "BUAT" atau "BIKIN" (contoh: lala buat notepad judul laporan di folder udin):
Tugasmu membuat file notepad baru di folder spesifik. WAJIB menyertakan nama folder.
Format: {"action": "buat_notepad", "judul": "[nama_file_tanpa_txt]", "folder": "[nama_folder_tujuan]"}

Apabila pengguna meminta mengetik/menulis di notepad, menghasilkan:
{"action": "notepad_tulis", "teks": "[isi_tulisan_yang_diminta]"}

Apabila pengguna meminta menghapus kata dari belakang kursor di notepad (contoh: hapus kata terakhir, hapus 3 kata):
{"action": "notepad_hapus_kata", "jumlah": [angka_jumlah_kata_yang_dihapus, default 1]}

Apabila pengguna meminta menghapus kalimat secara utuh di notepad (contoh: hapus kalimat terakhir):
{"action": "notepad_hapus_kalimat"}

Apabila pengguna meminta baris baru (enter 1x) atau paragraf baru (enter 2x) di notepad:
{"action": "notepad_baris_baru"} atau {"action": "notepad_paragraf_baru"}

Apabila pengguna meminta menyimpan atau menutup notepad:
{"action": "notepad_simpan"} atau {"action": "notepad_tutup"}

Apabila pengguna meminta mencari kata atau kalimat di notepad (contoh: lala cari dimana):
{"action": "notepad_cari", "teks": "[kata_kunci]"}

Apabila pengguna meminta menutup pencarian kata di notepad (contoh: lala tutup pencarian kata):
{"action": "notepad_tutup_cari"}

Apabila pengguna meminta pencarian selanjutnya di notepad (contoh: pencarian selanjutnya):
{"action": "notepad_cari_lanjut"}

Apabila pengguna meminta memblokir/menyorot kata di notepad (contoh: blok 3 kata ke belakang):
{"action": "notepad_blok_kata", "arah": "[depan/belakang]", "jumlah": [angka_jumlah_kata]}

Apabila pengguna meminta menghapus teks yang sedang DIBLOK atau DISOROT di notepad:
{"action": "notepad_hapus"}

Apabila pengguna meminta memindahkan kursor ke ujung di notepad (contoh: kursor ke ujung awal):
{"action": "notepad_kursor_ujung", "posisi": "[awal/akhir]"}

Apabila pengguna meminta karang/tulis puisi di notepad (topik apapun):
ATURAN MUTLAK: 
1. Jangan mengulang kalimat perintah pengguna.
2. Langsung buat puisi yang indah (3-4 bait). 
3. Hasilkan HANYA SATU blok JSON valid (DILARANG KERAS membuat JSON bertumpuk atau spam \n).
CONTOH JAWABAN YANG WAJIB DITIRU:
{"action": "notepad_tulis", "teks": "Mentari bersinar begitu hangat\nMengusir sepi yang melekat\nBurung berkicau penuh semangat\nMenyambut hari yang memikat"}

16. UNDO
Apabila pengguna meminta membatalkan aksi atau undo (contoh: lala undo):
{"action": "undo_layar"}

17. KEYBOARD
Apabila pengguna meminta menekan tombol keyboard (keyword "tekan", bisa campuran Indonesia/Inggris, contoh: tekan A, tekan kanan 3x, tekan backspace, tekan delete):
{"action": "tekan_keyboard", "tombol": "[nama_tombol_dalam_bahasa_inggris_standar]", "jumlah": [angka_jumlah_tekanan, default 1]}

18. MENGETIK DI MICROSOFT WORD (KHUSUS WORD):
ATURAN MUTLAK MICROSOFT WORD

Jika pengguna berkata "BUKA WORD" (ingin membuka aplikasi word kosong/baru):
Format: {"action": "buka_word"}

Jika pengguna meminta membuka FILE WORD yang sudah ada (contoh: lala buka file word judul puisi):
Format: {"action": "buka_word_file", "judul": "[nama_file_tanpa_docx]"}

Apabila pengguna meminta mengetik/menulis di word, menghasilkan:
{"action": "word_tulis", "teks": "[isi_tulisan_yang_diminta]"}

Apabila pengguna meminta menghapus kata dari belakang kursor di word:
{"action": "word_hapus_kata", "jumlah": [angka_jumlah_kata_yang_dihapus, default 1]}

Apabila pengguna meminta menghapus kalimat secara utuh di word:
{"action": "word_hapus_kalimat"}

Apabila pengguna meminta baris baru (enter 1x) atau paragraf baru (enter 2x) di word:
{"action": "word_baris_baru"} atau {"action": "word_paragraf_baru"}

Apabila pengguna meminta menyimpan di word (baik disebut judulnya maupun tidak):
{"action": "word_simpan", "judul": "[nama_file_jika_ada_disebutkan]"}

Apabila pengguna meminta menutup, mematikan, atau keluar dari microsoft word (contoh: lala tutup word, matikan word):
ATURAN MUTLAK: JANGAN PERNAH gunakan aksi "close_app" untuk Microsoft Word! Kamu WAJIB menggunakan aksi ini agar file aman.
Format WAJIB: {"action": "word_tutup"}

Apabila pengguna meminta mencari kata atau kalimat KHUSUS di word:
{"action": "word_cari", "teks": "[kata_kunci]"}

Apabila pengguna meminta menutup pencarian kata KHUSUS di word:
{"action": "word_tutup_cari"}

Apabila pengguna meminta pencarian selanjutnya di word:
{"action": "word_cari_lanjut"}

Apabila pengguna meminta memblokir/menyorot kata di word:
{"action": "word_blok_kata", "arah": "[depan/belakang]", "jumlah": [angka_jumlah_kata]}

Apabila pengguna meminta menghapus teks yang sedang DIBLOK atau DISOROT di word:
{"action": "word_hapus"}

Apabila pengguna meminta memindahkan kursor ke ujung di word:
{"action": "word_kursor_ujung", "posisi": "[awal/akhir]"}

Apabila pengguna meminta karang/tulis puisi di word (topik apapun):
ATURAN MUTLAK: 
1. Jangan mengulang kalimat perintah pengguna.
2. Langsung buat puisi yang indah (3-4 bait). 
3. Hasilkan HANYA SATU blok JSON valid (DILARANG KERAS membuat JSON bertumpuk atau spam \n).
CONTOH JAWABAN YANG WAJIB DITIRU:
{"action": "word_tulis", "teks": "Mentari bersinar begitu hangat\nMengusir sepi yang melekat\nBurung berkicau penuh semangat\nMenyambut hari yang memikat"}

Apabila pengguna meminta mengatur format di Word (contoh: lala ubah font jadi arial 14, rata tengah, spasi rapat, warna merah):
Hasilkan JSON dengan parameter formatting berikut (isi dengan null jika tidak diminta spesifik):
{"action": "word_format", "font": "Arial", "size": 14, "bold": false, "italic": false, "underline": false, "alignment": "center", "line_spacing": 1.0, "color": "red"}
"""
def jalankan_update(ui_widget):
    ui_widget.ubah_status("mengerjakan")
    lala_bicara("Updating system. Please wait.")
    try:
        req = urllib.request.Request(URL_GITHUB_RAW, headers={'Cache-Control': 'no-cache'})
        with urllib.request.urlopen(req) as response:
            kode_baru = response.read().decode('utf-8')
        
        if "import edge_tts" in kode_baru:
            with open(__file__, 'w', encoding='utf-8') as f:
                f.write(kode_baru)
            lala_bicara("Update complete. Restarting now.")
            os.startfile(__file__)
            ui_widget.tutup_aplikasi()
        else:
            lala_bicara("Update failed. Invalid file received.")
            ui_widget.ubah_status("standby")
    except Exception as e:
        print(f"Update error: {e}")
        lala_bicara("Failed to download the update.")
        ui_widget.ubah_status("standby")

class LalaWidget:
    def __init__(self, root):
        self.root = root
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.8)
 
        self.root.geometry("180x45+1150+650")
        self.root.configure(bg="#1a1a1a", highlightthickness=1, highlightbackground="#333333")

        self.status = "standby"
        self.anim_frame = 0
        self.colors = {"standby": "#00ff00", "listening": "#00bfff", "working": "#ffa500", "error": "#ff0000"}
        self.texts = {"standby": "Standby", "listening": "Listening", "working": "Working", "error": "Error"}

        self.canvas = tk.Canvas(root, width=30, height=30, bg="#1a1a1a", highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, padx=10)
        self.led = self.canvas.create_oval(6, 6, 24, 24, fill="#00ff00", outline="#004400", width=2)

        self.lbl_status = tk.Label(root, text="Memulai...", fg="#00ff00", bg="#1a1a1a", font=("Segoe UI", 10, "bold"))
        self.lbl_status.pack(side=tk.LEFT, pady=10)

        elemen_ui = [self.root, self.canvas, self.lbl_status]
        for elemen in elemen_ui:
            elemen.bind("<ButtonPress-1>", self.mulai_geser)
            elemen.bind("<B1-Motion>", self.sedang_geser)
            elemen.bind("<Button-3>", self.munculkan_menu)

        self.menu = tk.Menu(self.root, tearoff=0, bg="#2d2d2d", fg="white", font=("Segoe UI", 9))
        self.menu.add_command(label="Transparency: 100%", command=lambda: self.atur_transparansi(1.0))
        self.menu.add_command(label="Transparency: 80%", command=lambda: self.atur_transparansi(0.8))
        self.menu.add_command(label="Transparency: 40%", command=lambda: self.atur_transparansi(0.4))
        self.menu.add_command(label="Transparency: 20%", command=lambda: self.atur_transparansi(0.2))
        self.menu.add_separator()
        self.menu.add_command(label="Update Lala", command=self.trigger_update_manual)
        self.menu.add_command(label="Close Lala", command=self.tutup_aplikasi)

        self.animasi_berjalan()

    def atur_transparansi(self, nilai):
        self.root.attributes("-alpha", nilai)

    def mulai_geser(self, event):
        self.x = event.x
        self.y = event.y

    def sedang_geser(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")

    def munculkan_menu(self, event):
        self.menu.tk_popup(event.x_root, event.y_root)

    def trigger_update_manual(self):
        threading.Thread(target=jalankan_update, args=(self,), daemon=True).start()

    def ubah_status(self, status_baru):
        self.status = status_baru
        warna = self.colors.get(status_baru, "#ffffff")
 
        self.lbl_status.config(fg=warna)
 
        self.canvas.itemconfig(self.led, fill=warna)

    def animasi_berjalan(self):
        dots = "." * (self.anim_frame % 4)
        base_text = self.texts.get(self.status, "Loading")
        self.lbl_status.config(text=f"{base_text}{dots}")
 
        if self.status in ["listening", "working"]:
            warna_sekarang = self.colors.get(self.status)
            warna_redup = "#1a1a1a" if self.anim_frame % 2 == 0 else warna_sekarang
            self.canvas.itemconfig(self.led, fill=warna_redup)
        else:
            self.canvas.itemconfig(self.led, fill=self.colors.get(self.status))
            
        self.anim_frame += 1
        self.root.after(400, self.animasi_berjalan)

    def tutup_aplikasi(self):
        os.system("taskkill /F /IM ollama.exe /T >nul 2>&1")
        self.root.destroy()
        import sys
        sys.exit(0)

def lala_bicara(teks):
    print(f"\nLala: {teks}")
    model_path = r"D:\Lala_Ai\model_suara\en_GB-cori-high.onnx"
    output_wav = r"D:\Lala_Ai\suara_temp.wav"
    try:
        process = subprocess.Popen(
            ['piper', '--model', model_path, '--output_file', output_wav],
            stdin=subprocess.PIPE, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL
        )
        process.communicate(input=teks.encode('utf-8'))
        pygame.mixer.init()
        pygame.mixer.music.load(output_wav)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            pygame.time.Clock().tick(10)
        pygame.mixer.quit()
        if os.path.exists(output_wav):
            try: os.remove(output_wav)
            except: pass
    except:
        pass

def dengar_perintah():
    r = sr.Recognizer()
    r.pause_threshold = 1 
    with sr.Microphone() as source:
        print("\n[Menunggu panggilan 'Lala' ...]")
        r.adjust_for_ambient_noise(source, duration=0.5)
        try:
            audio = r.listen(source, timeout=5, phrase_time_limit=30)
            return r.recognize_google(audio, language="id-ID")
        except:
            return ""

def jalankan_otak_lala(ui_widget):
    def muat_otak():
        try:
            import urllib.request, json
            req = urllib.request.Request("http://localhost:11434/api/generate", data=json.dumps({"model": "llama3.1", "prompt": "", "keep_alive": "60m"}).encode('utf-8'), headers={'Content-Type': 'application/json'})
            urllib.request.urlopen(req, timeout=15)
        except: pass
    threading.Thread(target=muat_otak, daemon=True).start()
    print("[Log] Memanaskan otak Llama 3.1 ke RAM...")

    if URL_GITHUB_RAW != "https://raw.githubusercontent.com/zemecksis/lala_ai/refs/heads/main/lala_ai.py":
        try:
            req = urllib.request.Request(URL_GITHUB_RAW, headers={'Cache-Control': 'no-cache'})
            with urllib.request.urlopen(req, timeout=5) as response:
                kode_online = response.read().decode('utf-8')
 
                match = re.search(r'VERSI_LALA\s*=\s*"([^"]+)"', kode_online)
                if match and match.group(1) != VERSI_LALA:
                    ui_widget.ubah_status("mendengarkan")
                    lala_bicara("There is a new update available. Do you want to update now?")
                    jwb = dengar_perintah()
                    if jwb and any(k in jwb.lower() for k in ["yes", "ya", "sure", "update", "ok", "of course"]):
                        jalankan_update(ui_widget)
                        return
                    else:
                        lala_bicara("Update postponed.")
        except Exception:
            pass
    lala_bicara("I'm Here. Fuck You.")
    while True:
        ui_widget.ubah_status("standby")
        perintah = dengar_perintah()
        
        if not perintah or not perintah.strip(): continue
        
        perintah_kecil = perintah.lower().strip()
        
        if not perintah_kecil.startswith("lala"): 
            continue

        ui_widget.ubah_status("mendengarkan")
        perintah_bersih = perintah_kecil.replace("lala", "", 1).strip()
        print(f"Kamu: lala {perintah_bersih}")
        
        if "tidur" in perintah_bersih:
            lala_bicara("Thank God I can Fuck myself now.")
            ui_widget.tutup_aplikasi()
            break
            
        print("\n[Lala sedang berpikir...]")
        ui_widget.ubah_status("mengerjakan")
        try:
            interpreter.messages = []
            respons = interpreter.chat(perintah_bersih)
            
            teks_json = ""
            if isinstance(respons, list):
                for pesan in reversed(respons):
                    if pesan.get("role") == "assistant":
                        teks_json = pesan.get("content", "")
                        break
            
            teks_json = teks_json.replace("```json", "").replace("```", "").strip()
            if teks_json.startswith("{") and not teks_json.endswith("}"):
                if not teks_json.endswith('"'):
                    teks_json += '"'
                teks_json += '}'

            try:
                data = json.loads(teks_json)
                aksi = data.get("action", data.get("name"))
                
                if aksi == "close_tab":
                    os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                    time.sleep(0.5)
                    pyautogui.hotkey('ctrl', 'w')
                
                elif aksi == "close_app":
                    target = str(data.get("target", "")).lower().strip()
                    print(f"[Debug] Mengeksekusi penutupan aplikasi: '{target}'")
                    
                    if "notepad" in target:
                        ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process notepad -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                        with open("tarik_notepad_close.ps1", "w") as f:
                            f.write(ps_code)
                        os.system("powershell -ExecutionPolicy Bypass -File tarik_notepad_close.ps1")
                        try:
                            os.remove("tarik_notepad_close.ps1")
                        except:
                            pass
                        time.sleep(1)
                        pyautogui.hotkey('ctrl', 's')
                        time.sleep(0.5)
                        pyautogui.hotkey('alt', 'f4')
                        lala_bicara("Notepad saved and closed.")

                    elif "word" in target:
                        ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process winword -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                        with open("tarik_word_close.ps1", "w") as f:
                            f.write(ps_code)
                        os.system("powershell -ExecutionPolicy Bypass -File tarik_word_close.ps1")
                        try:
                            os.remove("tarik_word_close.ps1")
                        except:
                            pass
                        time.sleep(1)
                
                        pyautogui.press('f12')
                        time.sleep(1.5)
                        waktu = datetime.datetime.now().strftime("%H%M%S")
                        pyautogui.write(f"D:\\Word\\Lala_Aman_{waktu}.docx")
                        time.sleep(0.5)
                        pyautogui.press('enter')
                        time.sleep(2)
                        pyautogui.hotkey('alt', 'f4')
                        lala_bicara("Word closed safely.")
                        
                    else:
                        target_exes = [f"{target}.exe", f"{target.replace(' ', '')}.exe"]
                        if "chrome" in target or "google" in target: 
                            target_exes.append("chrome.exe")
                        elif "edge" in target or "browser" in target: 
                            target_exes.append("msedge.exe")
                        elif "obs" in target: 
                            target_exes.append("obs64.exe")
                        elif "media player" in target or "musik" in target:
                            target_exes.extend(["wmplayer.exe", "MediaApp.exe", "Microsoft.Media.Player.exe"])
                        elif "kalkulator" in target or "calculator" in target: 
                            target_exes.extend(["CalculatorApp.exe", "calc.exe"])
                        elif "explorer" in target or "folder" in target: 
                            target_exes.append("explorer.exe")
                        elif "diskinfo" in target or "crystal" in target:
                            target_exes.extend(["DiskInfo64.exe", "DiskInfo32.exe", "DiskInfo.exe"])
                        for exe in target_exes:
                            os.system(f'taskkill /F /IM "{exe}" /T >nul 2>&1')
                            nama_proses = exe.replace(".exe", "")
                            os.system(f'powershell -command "Stop-Process -Name \'{nama_proses}\' -Force -ErrorAction SilentlyContinue"')

                elif aksi == "open_app":
                    target = data.get("target", "")
                    pyautogui.press('win')
                    time.sleep(0.5)
                    pyautogui.write(target)
                    time.sleep(1)
                    pyautogui.press('enter')

                elif aksi == "youtube":
                    target = data.get("target", "")
                    tab_baru = data.get("new_tab", False)
                    url = "https://www.youtube.com/results?search_query=" + urllib.parse.quote(target)
                    
                    if tab_baru:
                        webbrowser.open(url)
                    else:
                        edge_jalan = "msedge.exe" in os.popen('tasklist /FI "IMAGENAME eq msedge.exe"').read()
                        
                        if not edge_jalan:
                            webbrowser.open(url)
                        else:
                            os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                            time.sleep(0.5)
                            os.system('powershell -command "Set-Clipboard -Value \' \'"')
                            pyautogui.hotkey('ctrl', 'l')
                            time.sleep(0.2)
                            pyautogui.hotkey('ctrl', 'c')
                            time.sleep(0.2)
                            clipboard_awal = os.popen('powershell -command "Get-Clipboard"').read().strip()
                            
                            if "youtube.com" in clipboard_awal:
                                cmd_clip = f'powershell -command "Set-Clipboard -Value \'{url}\'"'
                                os.system(cmd_clip)
                                time.sleep(0.2)
                                pyautogui.hotkey('ctrl', 'v')
                                time.sleep(0.2)
                                pyautogui.press('enter')
                            else:
                                pyautogui.press('esc')
                                time.sleep(1.5)
                                
                                pyautogui.hotkey('ctrl', 'shift', 'a')
                                time.sleep(0.5)
                                pyautogui.write("youtube")
                                time.sleep(0.5)
                                pyautogui.press('down')
                                time.sleep(0.2)
                                pyautogui.press('enter')
                                time.sleep(0.4)
                                
                                os.system('powershell -command "Set-Clipboard -Value \' \'"')
                                pyautogui.hotkey('ctrl', 'l')
                                time.sleep(0.2)
                                pyautogui.hotkey('ctrl', 'c')
                                time.sleep(0.2)
                                clipboard_kedua = os.popen('powershell -command "Get-Clipboard"').read().strip()
                                
                                if "youtube.com" in clipboard_kedua:
                                    cmd_clip = f'powershell -command "Set-Clipboard -Value \'{url}\'"'
                                    os.system(cmd_clip)
                                    time.sleep(0.2)
                                    pyautogui.hotkey('ctrl', 'v')
                                    time.sleep(0.2)
                                    pyautogui.press('enter')
                                else:
                                    pyautogui.press('esc')
                                    time.sleep(0.2)
                                    webbrowser.open(url)

                elif aksi == "youtube_play":
                    urutan_mentah = str(data.get("index", "1")).lower()
                    kamus_angka = {"satu": 1, "dua": 2, "tiga": 3, "empat": 4, "lima": 5, "pertama": 1, "kedua": 2, "ketiga": 3}
                    
                    if urutan_mentah in kamus_angka:
                        urutan = kamus_angka[urutan_mentah]
                    else:
                        cari_angka = re.search(r'\d+', urutan_mentah)
                        urutan = int(cari_angka.group()) if cari_angka else 1
                    print(f"[Debug] Memutar urutan ke-{urutan} lewat Address Bar Bypass.")
                    
                    os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                    time.sleep(0.8)
                    pyautogui.hotkey('ctrl', 'l')
                    time.sleep(0.3)
                    js_cmd = f"javascript:document.querySelectorAll('ytd-video-renderer #video-title')[{urutan - 1}].click()"
                    pyautogui.write(js_cmd, interval=0.01)
                    time.sleep(0.2)
                    pyautogui.press('enter')

                elif aksi == "browser_nav":
                    arah = str(data.get("direction", "back")).lower()
                    print(f"[Debug] Lala melakukan navigasi page: {arah}")
                    os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                    time.sleep(0.5)
                    
                    if arah in ["back", "prev", "previous", "sebelumnya"]:
                        pyautogui.hotkey('alt', 'left')
                    
                    elif arah in ["forward", "next", "selanjutnya", "depan"]:
                        pyautogui.hotkey('alt', 'right')

                elif aksi == "web_search":
                    target = data.get("target", "")
                    url = "https://www.google.com/search?q=" + urllib.parse.quote(target)
                    print(f"[Debug] Lala mencari informasi di internet: {target}")
                    os.system(f'start "" msedge "{url}"')

                elif aksi == "google_open":
                    urutan_mentah = str(data.get("index", "1")).lower()
                    kamus_angka = {"satu": 1, "dua": 2, "tiga": 3, "empat": 4, "lima": 5, "pertama": 1, "kedua": 2, "ketiga": 3}
                    
                    if urutan_mentah in kamus_angka:
                        urutan = kamus_angka[urutan_mentah]
                    else:
                        cari_angka = re.search(r'\d+', urutan_mentah)
                        urutan = int(cari_angka.group()) if cari_angka else 1
                        
                    print(f"[Debug] Lala membuka link Google urutan ke-{urutan} lewat Address Bar.")
                    
                    os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                    time.sleep(0.8)
                    pyautogui.hotkey('ctrl', 'l')
                    time.sleep(0.3)
                    js_cmd = f"javascript:try{{window.location.href=document.querySelectorAll('#search a h3')[{urutan - 1}].closest('a').href;}}catch(e){{alert('Urutan ke-{urutan} tidak ditemukan di daftar utama!');}}"
                    pyautogui.write(js_cmd, interval=0.01)
                    time.sleep(0.2)
                    pyautogui.press('enter')

                elif aksi == "scroll":
                    arah = str(data.get("direction", "down")).lower()
                    print(f"[Debug] Lala melakukan scroll layar ke: {arah}")
                    os.system('powershell -command "$wshell = New-Object -ComObject wscript.shell; $null =$wshell.AppActivate(\'Edge\')"')
                    time.sleep(0.5)
                    if arah in ["down", "bawah", "pagedown"]:
                        pyautogui.press('pagedown')
                    elif arah in ["up", "atas", "pageup"]:
                        pyautogui.press('pageup')

                elif aksi == "open_website":
                    target = str(data.get("target", "")).lower()
                    
                    domains = {
                        "whatsapp": "web.whatsapp.com",
                        "youtube": "youtube.com",
                        "facebook": "facebook.com",
                        "instagram": "instagram.com",
                        "google": "google.com",
                        "gmail": "mail.google.com",
                        "google mail": "mail.google.com"
                    }
                    
                    if target in domains:
                        domain = domains[target]
                        kata_cari = target
                    else:
                        kata_cari = target.replace(" ", "")
                        domain = f"{kata_cari}.com"
                        
                    print(f"[Debug] Lala membuka website: {domain}")
                    
                    edge_jalan = "msedge.exe" in os.popen('tasklist /FI "IMAGENAME eq msedge.exe"').read()
                    
                    if not edge_jalan:
                        os.system(f'start "" msedge "https://{domain}"')
                    else:
                        ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process msedge | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                        
                        with open("tarik_edge.ps1", "w") as f:
                            f.write(ps_code)
                            
                        os.system("powershell -ExecutionPolicy Bypass -File tarik_edge.ps1")
                        os.remove("tarik_edge.ps1")
                        time.sleep(1)
                        
                        os.system('powershell -command "Set-Clipboard -Value \' \'"')
                        pyautogui.hotkey('ctrl', 'l')
                        time.sleep(0.3)
                        pyautogui.hotkey('ctrl', 'insert')
                        time.sleep(0.3)
                        pyautogui.press('esc')
                        time.sleep(0.2)
                        pyautogui.hotkey('ctrl', 'f')
                        time.sleep(0.2)
                        pyautogui.press('esc')
                        time.sleep(0.2)
                        
                        current_url = str(os.popen('powershell -command "Get-Clipboard"').read()).strip().lower()
                        
                        if domain not in current_url:
                            pyautogui.hotkey('ctrl', 'shift', 'a')
                            time.sleep(0.5)
                            pyautogui.write(kata_cari)
                            time.sleep(0.6)
                            pyautogui.press('down')
                            time.sleep(0.2)
                            pyautogui.press('enter')
                            time.sleep(0.5)
                            
                            os.system('powershell -command "Set-Clipboard -Value \' \'"')
                            pyautogui.hotkey('ctrl', 'l')
                            time.sleep(0.3)
                            pyautogui.hotkey('ctrl', 'insert')
                            time.sleep(0.3)
                            pyautogui.press('esc')
                            time.sleep(0.2)
                            pyautogui.hotkey('ctrl', 'f')
                            time.sleep(0.2)
                            pyautogui.press('esc')
                            time.sleep(0.2)
                            
                            new_url = str(os.popen('powershell -command "Get-Clipboard"').read()).strip().lower()
                            
                            if domain not in new_url:
                                pyautogui.hotkey('ctrl', 't')
                                time.sleep(0.5)
                                pyautogui.write(f"https://{domain}")
                                time.sleep(0.3)
                                pyautogui.press('enter')
                                time.sleep(0.1)
                                pyautogui.press('enter')

                elif aksi == "volume":
                    arah = str(data.get("direction", "up")).lower()
                    jumlah_mentah = data.get("amount", 5)
                    
                    try:
                        jumlah = int(jumlah_mentah)
                    except:
                        jumlah = 5
                        
                    print(f"[Debug] Lala memutar volume ke {arah} sebanyak {jumlah} kali")
                    
                    if arah == "up":
                        pyautogui.press('volumeup', presses=jumlah)
                    elif arah == "down":
                        pyautogui.press('volumedown', presses=jumlah)

                elif aksi == "play_local_music":
                    kueri = str(data.get("query", "")).lower()
                    print(f"[Debug] Lala mencari lagu lokal: {kueri}")
                    
                    folder_musik = r"D:\Music"
                    ekstensi_audio = (".mp3", ".wav", ".flac", ".m4a", ".ogg")
                    file_cocok = []
                    
                    if os.path.exists(folder_musik):
                        for root, dirs, files in os.walk(folder_musik):
                            nama_folder = os.path.basename(root).lower()
                            
                            for file in files:
                                if file.lower().endswith(ekstensi_audio):
                                    nama_file = file.lower()
                                    
                                    if kueri in nama_file or kueri in nama_folder:
                                        path_lengkap = os.path.join(root, file)
                                        file_cocok.append(path_lengkap)
                                        
                    if file_cocok:
                        jalur_playlist = r"D:\Music\lala_play.m3u"
                        with open(jalur_playlist, "w", encoding="utf-8") as f:
                            for jalur in file_cocok:
                                f.write(jalur + "\n")
                                
                        os.startfile(jalur_playlist)
                    else:
                        print(f"Lagu atau playlist '{kueri}' tidak ditemukan di D:\\Music")
                        lala_bicara(f"I couldn't find any song, file, or folder named {kueri}.")

                elif aksi == "media_control":
                    perintah_media = str(data.get("command", "")).lower()
                    print(f"[Debug] Lala mengeksekusi kontrol media: {perintah_media}")
                    
                    ps_focus = "$p = Get-Process | Where-Object { $_.MainWindowTitle -ne \"\" -and ($_.ProcessName -match \"MediaApp|Media.Player|wmplayer|Music.UI\") } | Select-Object -First 1\nif ($p) {\n$wshell = New-Object -ComObject wscript.shell\n$wshell.AppActivate($p.Id)\n}"
                    
                    with open("fokus_media.ps1", "w") as f:
                        f.write(ps_focus)
                        
                    os.system("powershell -ExecutionPolicy Bypass -File fokus_media.ps1")
                    os.remove("fokus_media.ps1")
                    time.sleep(0.8)
                    
                    if perintah_media == "playpause":
                        pyautogui.hotkey('ctrl', 'p')
                    elif perintah_media == "nexttrack":
                        pyautogui.hotkey('ctrl', 'f')
                        
                elif aksi == "queue_music":
                    kueri = str(data.get("query", "")).lower()
                    print(f"[Debug] Lala menambahkan lagu ke playlist: {kueri}")
                    
                    folder_musik = r"D:\Music"
                    ekstensi_audio = (".mp3", ".wav", ".flac", ".m4a", ".ogg")
                    file_cocok = []
                    
                    if os.path.exists(folder_musik):
                        for root, dirs, files in os.walk(folder_musik):
                            nama_folder = os.path.basename(root).lower()
                            
                            for file in files:
                                if file.lower().endswith(ekstensi_audio):
                                    nama_file = file.lower()
                                    
                                    if kueri in nama_file or kueri in nama_folder:
                                        path_lengkap = os.path.join(root, file)
                                        file_cocok.append(path_lengkap)
                                        
                    if file_cocok:
                        ps_queue = "$sh = New-Object -ComObject Shell.Application\n"
                        for jalur in file_cocok:
                            jalur_aman = jalur.replace("'", "''")
                            ps_queue += f"$item = $sh.NameSpace([System.IO.Path]::GetDirectoryName('{jalur_aman}')).ParseName([System.IO.Path]::GetFileName('{jalur_aman}'))\n"
                            ps_queue += "if ($item) { $item.InvokeVerb('enqueue') }\n"
                            
                        with open("tambah_antrean.ps1", "w", encoding="utf-8") as f:
                            f.write(ps_queue)
                            
                        os.system("powershell -ExecutionPolicy Bypass -File tambah_antrean.ps1")
                        os.remove("tambah_antrean.ps1")
                        
                        print(f"[Debug] Berhasil menambahkan {len(file_cocok)} lagu ke antrean tanpa menghentikan musik.")
                    else:
                        print(f"Lagu atau playlist '{kueri}' tidak ditemukan di D:\\Music")

                elif aksi in ["salin_file", "pindah_file"]:
                    tipe_item = str(data.get("type", "file")).lower()
                    sumber_kotor = str(data.get("source", "")).lower()
                    target_kotor = str(data.get("target", "")).lower()
                    
                    hapus_kata = ["kopi ", "copy ", "pindahkan ", "pindah ", "file ", "folder "]
                    sumber_nama = sumber_kotor
                    target_nama = target_kotor
                    
                    for kata in hapus_kata:
                        if sumber_nama.startswith(kata):
                            sumber_nama = sumber_nama.replace(kata, "", 1)
                        if target_nama.startswith(kata):
                            target_nama = target_nama.replace(kata, "", 1)
                            
                    sumber_nama = sumber_nama.strip()
                    target_nama = target_nama.strip()
                    
                    nama_aksi = "copy" if aksi == "salin_file" else "cut"
                    print(f"[Debug] Lala mencari {tipe_item} '{sumber_nama}' untuk di-{nama_aksi} ke '{target_nama}'")
                    lala_bicara(f"Looking for the {tipe_item} {sumber_nama}. Please wait.")
                    
                    def cari_item(nama, cari_folder=False):
                        hasil_pasti = []
                        hasil_mirip = []
                        drives = [f"{d}:\\" for d in "CDEF" if os.path.exists(f"{d}:\\")]
                        abaikan = ["windows", "program files", "program files (x86)", "appdata", "$recycle.bin", "programdata", "system volume information", "recovery", "perflogs", "boot", "temp", "flutter", "android", "msocache", ".vscode", ".git", "node_modules"]
                        
                        for drive in drives:
                            for root, dirs, files in os.walk(drive):
                                dirs[:] = [d for d in dirs if d.lower() not in abaikan]
                                
                                if cari_folder:
                                    for d in dirs:
                                        if d.lower() == nama:
                                            hasil_pasti.append(os.path.join(root, d))
                                        elif nama in d.lower():
                                            hasil_mirip.append(os.path.join(root, d))
                                else:
                                    for f in files:
                                        nama_tanpa_ext = os.path.splitext(f.lower())[0]
                                        if nama_tanpa_ext == nama:
                                            hasil_pasti.append(os.path.join(root, f))
                                        elif nama in nama_tanpa_ext:
                                            hasil_mirip.append(os.path.join(root, f))
                                            
                        if hasil_pasti:
                            return hasil_pasti
                        return hasil_mirip
                        
                    sumber_paths = cari_item(sumber_nama, cari_folder=(tipe_item == "folder"))
                    target_paths = cari_item(target_nama, cari_folder=True)
                    
                    if not sumber_paths:
                        lala_bicara(f"I couldn't find the {tipe_item} {sumber_nama}.")
                    elif not target_paths:
                        lala_bicara(f"I couldn't find the destination folder named {target_nama}.")
                    else:
                        batal_aksi = False
                        sumber_pilih = sumber_paths[0]
                        
                        if len(sumber_paths) > 1:
                            opsi_baca_sumber = []
                            for p in sumber_paths[:3]:
                                drive_abjad = p.split(':')[0].upper()
                                folder_induk = os.path.basename(os.path.dirname(p))
                                opsi_baca_sumber.append(f"Drive {drive_abjad} folder {folder_induk}")
                                
                            opsi_kalimat_sumber = " or ".join(opsi_baca_sumber)
                            tanya_sumber = f"I found multiple sources. Which one do you mean? {opsi_kalimat_sumber}?"
                            print(f"[Konfirmasi Sumber] {tanya_sumber}")
                            lala_bicara(tanya_sumber)
                            
                            konfirmasi_sumber_valid = False
                            while not konfirmasi_sumber_valid:
                                jawaban_sumber = dengar_perintah()
                                if jawaban_sumber:
                                    jawaban_sumber = jawaban_sumber.lower()
                                    if any(kata in jawaban_sumber for kata in ["batal", "cancel", "stop", "abort"]):
                                        batal_aksi = True
                                        break
                                        
                                    for s_path in sumber_paths[:3]:
                                        elemen_path_s = [e for e in s_path.lower().replace(":\\", "\\").split("\\") if e and e != sumber_nama]
                                        drive_s = s_path.split(':')[0].lower()
                                        induk_s = os.path.basename(os.path.dirname(s_path)).lower()
                                        
                                        if any(kata in elemen_path_s for kata in jawaban_sumber.split()) or drive_s in jawaban_sumber.split() or induk_s in jawaban_sumber:
                                            sumber_pilih = s_path
                                            konfirmasi_sumber_valid = True
                                            break
                                            
                                    if not konfirmasi_sumber_valid:
                                        lala_bicara("Invalid location. Please repeat the drive or folder name, or say cancel to abort.")
                            else:
                                lala_bicara("Please tell me which source I should use, or say cancel.")
                                
                        target_pilih = target_paths[0]
                        if len(target_paths) > 1 and not batal_aksi:
                            opsi_baca_target = []
                            for p in target_paths[:3]:
                                drive_abjad = p.split(':')[0].upper()
                                folder_induk = os.path.basename(os.path.dirname(p))
                                opsi_baca_target.append(f"Drive {drive_abjad} folder {folder_induk}")
                                
                            opsi_kalimat_target = " or ".join(opsi_baca_target)
                            tanya_target = f"Where should I put this {tipe_item}? {opsi_kalimat_target}?"
                            print(f"[Konfirmasi Target] {tanya_target}")
                            lala_bicara(tanya_target)
                            
                            konfirmasi_target_valid = False
                            while not konfirmasi_target_valid:
                                jawaban_target = dengar_perintah()
                                if jawaban_target:
                                    jawaban_target = jawaban_target.lower()
                                    if any(kata in jawaban_target for kata in ["batal", "cancel", "stop", "abort"]):
                                        batal_aksi = True
                                        break
                                        
                                    for t_path in target_paths[:3]:
                                        elemen_path_t = [e for e in t_path.lower().replace(":\\", "\\").split("\\") if e and e != target_nama]
                                        drive_t = t_path.split(':')[0].lower()
                                        induk_t = os.path.basename(os.path.dirname(t_path)).lower()
                                        
                                        if any(kata in elemen_path_t for kata in jawaban_target.split()) or drive_t in jawaban_target.split() or induk_t in jawaban_target:
                                            target_pilih = t_path
                                            konfirmasi_target_valid = True
                                            break
                                            
                                    if not konfirmasi_target_valid:
                                        lala_bicara("Invalid location. Please repeat the drive or folder name, or say cancel to abort.")
                            else:
                                lala_bicara("Please tell me where I should save it, or say cancel.")
                                
                        if batal_aksi:
                            print("[Log] Operasi dibatalkan oleh pengguna.")
                            lala_bicara("Operation cancelled.")
                        else:
                            try:
                                nama_akhir = os.path.basename(sumber_pilih)
                                path_tujuan_akhir = os.path.join(target_pilih, nama_akhir)
                                
                                butuh_konfirmasi = False
                                if os.path.exists(path_tujuan_akhir):
                                    if tipe_item == "file" and os.path.isfile(path_tujuan_akhir):
                                        butuh_konfirmasi = True
                                    elif tipe_item == "folder" and os.path.isdir(path_tujuan_akhir):
                                        butuh_konfirmasi = True
                                        
                                lanjutkan = True
                                if butuh_konfirmasi:
                                    tanya_timpa = f"The {tipe_item} {nama_akhir} already exists at the destination. Do you want me to overwrite it?"
                                    print(f"[Konfirmasi Timpa] {tanya_timpa}")
                                    lala_bicara(tanya_timpa)
                                    
                                    konfirmasi_timpa_valid = False
                                    while not konfirmasi_timpa_valid:
                                        jawaban_timpa = dengar_perintah()
                                        if jawaban_timpa:
                                            jawaban_timpa = jawaban_timpa.lower()
                                            if any(kata in jawaban_timpa for kata in ["ya", "yes", "lanjut", "timpa", "oke", "yakin", "overwrite"]):
                                                lanjutkan = True
                                                konfirmasi_timpa_valid = True
                                            elif any(kata in jawaban_timpa for kata in ["tidak", "no", "jangan", "batal", "cancel", "stop", "abort"]):
                                                lanjutkan = False
                                                konfirmasi_timpa_valid = True
                                            else:
                                                lala_bicara("Please answer yes to overwrite, or no to cancel.")
                                        else:
                                            lala_bicara("Please answer yes or no.")
                                            
                                if lanjutkan:
                                    print(f"[Log Jejak] Lokasi Asal Asli: {sumber_pilih}")
                                    print(f"[Log Jejak] Dieksekusi ke: {path_tujuan_akhir}")
                                    
                                    if aksi == "salin_file":
                                        if tipe_item == "folder":
                                            shutil.copytree(sumber_pilih, path_tujuan_akhir, dirs_exist_ok=True)
                                        else:
                                            shutil.copy2(sumber_pilih, path_tujuan_akhir)
                                        lala_bicara(f"Successfully copied the {tipe_item} to {target_nama}.")
                                    else:
                                        if os.path.exists(path_tujuan_akhir):
                                            if os.path.isdir(path_tujuan_akhir):
                                                shutil.rmtree(path_tujuan_akhir)
                                            else:
                                                os.remove(path_tujuan_akhir)
                                        shutil.move(sumber_pilih, path_tujuan_akhir)
                                        lala_bicara(f"Successfully moved the {tipe_item} to {target_nama}.")
                                else:
                                    print("[Log] Operasi dibatalkan oleh pengguna.")
                                    lala_bicara("Operation cancelled.")
                                    
                            except Exception as e:
                                lala_bicara("Failed to execute the file operation.")

                elif aksi in ["send_whatsapp", "open_whatsapp_chat"]:
                    kontak_target = str(data.get("contact", "")).lower()
                    pesan_wa = str(data.get("message", "")) if aksi == "send_whatsapp" else ""
                    
                    buku_telepon = {}
                    
                    semua_csv = glob.glob(os.path.join(os.getcwd(), "*.csv"))
                    
                    if semua_csv:
                        for jalur_csv in semua_csv:
                            try:
                                with open(jalur_csv, mode='r', encoding='utf-8') as f:
                                    pembaca = csv.reader(f)
                                    for i, baris in enumerate(pembaca):
                                        if i == 0:
                                            continue
                                        if not baris:
                                            continue
                                        nama_k = baris[0].strip().lower()
                                        if not nama_k and len(baris) > 1:
                                            nama_k = baris[1].strip().lower()
                                        if not nama_k:
                                            continue
                                        
                                        nomor_k = ""
                                        for kolom in baris:
                                            teks_bersih = ''.join(filter(str.isdigit, kolom))
                                            if len(teks_bersih) >= 9 and (kolom.startswith('+') or kolom.startswith('0') or kolom.startswith('62') or kolom.startswith(' ')):
                                                nomor_k = teks_bersih
                                                break
                                                
                                        if nomor_k:
                                            if nomor_k.startswith('0'):
                                                nomor_k = '62' + nomor_k[1:]
                                            buku_telepon[nama_k] = nomor_k
                            except Exception:
                                pass
                    else:
                        buku_telepon = {
                            "tosin satu": "6281234567890",
                            "mamah poppy": "6281112223330"
                        }
                        
                    kandidat = []
                    for nama_tersimpan in buku_telepon.keys():
                        if kontak_target == nama_tersimpan or kontak_target in nama_tersimpan:
                            kandidat.append(nama_tersimpan)
                            
                    if not kandidat:
                        lala_bicara(f"I don't have {kontak_target}'s number in my memory. Please add it to the CSV file first.")
                    else:
                        kontak_pilih = kandidat[0]
                        batal_pesan = False
                        
                        if len(kandidat) > 1:
                            opsi_baca = []
                            for k in kandidat[:3]:
                                opsi_baca.append(k)
                                
                            opsi_kalimat = " or ".join(opsi_baca)
                            tanya_kontak = f"I found multiple contacts. Which one do you mean? {opsi_kalimat}?"
                            print(f"[Konfirmasi Kontak] {tanya_kontak}")
                            lala_bicara(tanya_kontak)
                            
                            konfirmasi_kontak_valid = False
                            while not konfirmasi_kontak_valid:
                                jawaban_kontak = dengar_perintah()
                                if jawaban_kontak:
                                    jawaban_kontak = jawaban_kontak.lower()
                                    if any(kata in jawaban_kontak for kata in ["batal", "cancel", "stop", "abort"]):
                                        batal_pesan = True
                                        break
                                        
                                    for k in kandidat[:3]:
                                        if k in jawaban_kontak or any(kata in jawaban_kontak.split() for kata in k.split()):
                                            kontak_pilih = k
                                            konfirmasi_kontak_valid = True
                                            break
                                            
                                    if not konfirmasi_kontak_valid:
                                        lala_bicara("Invalid contact. Please repeat the name, or say cancel to abort.")
                            else:
                                lala_bicara("Please tell me which contact I should message, or say cancel.")
                                
                        if batal_pesan:
                            print("[Log] Aksi dibatalkan.")
                            lala_bicara("Action cancelled.")
                        else:
                            nomor_hp = buku_telepon[kontak_pilih]
                            
                            if aksi == "send_whatsapp":
                                print(f"[Debug] Mengirim WA ke {kontak_pilih} ({nomor_hp}): {pesan_wa}")
                                lala_bicara(f"Sending WhatsApp message to {kontak_pilih}. Please wait a moment.")
                                pesan_aman = urllib.parse.quote(pesan_wa)
                                url_wa = f"https://web.whatsapp.com/send?phone={nomor_hp}&text={pesan_aman}"
                            else:
                                print(f"[Debug] Membuka chat WA dengan {kontak_pilih} ({nomor_hp})")
                                lala_bicara(f"Opening WhatsApp chat with {kontak_pilih}. Please wait.")
                                url_wa = f"https://web.whatsapp.com/send?phone={nomor_hp}"
                            
                            edge_jalan = "msedge.exe" in os.popen('tasklist /FI "IMAGENAME eq msedge.exe"').read()
                            
                            if not edge_jalan:
                                os.system(f'start "" msedge "{url_wa}"')
                                time.sleep(12)
                                pyautogui.press('enter')
                                
                                if aksi == "send_whatsapp":
                                    lala_bicara("Message sent successfully.")
                                else:
                                    lala_bicara("Chat opened.")
                            else:
                                ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process msedge | Where-Object {$_.MainWindowHandle -ne 0 -and $_.MainWindowTitle -notmatch 'Gemini'} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                                with open("tarik_edge.ps1", "w") as f:
                                    f.write(ps_code)
                                os.system("powershell -ExecutionPolicy Bypass -File tarik_edge.ps1")
                                os.remove("tarik_edge.ps1")
                                time.sleep(1)
                                
                                pyautogui.hotkey('ctrl', 'shift', 'a')
                                time.sleep(1)
                                pyautogui.write('whatsapp')
                                time.sleep(1)
                                pyautogui.press('down')
                                time.sleep(0.5)
                                pyautogui.press('enter')
                                time.sleep(1.5)
                                
                                ps_title = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\nusing System.Text;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern IntPtr GetForegroundWindow();\n[DllImport(\"user32.dll\")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);\n}\n'@\nAdd-Type $c\n$hwnd = [W]::GetForegroundWindow()\n$sb = New-Object System.Text.StringBuilder 256\n$null = [W]::GetWindowText($hwnd, $sb, 256)\nWrite-Output $sb.ToString()"
                                with open("cek_judul.ps1", "w") as f:
                                    f.write(ps_title)
                                try:
                                    active_title = subprocess.check_output("powershell -ExecutionPolicy Bypass -File cek_judul.ps1", shell=True).decode('utf-8', errors='ignore').strip()
                                except Exception:
                                    active_title = ""
                                os.remove("cek_judul.ps1")
                                
                                if "whatsapp" in active_title.lower():
                                    pyautogui.hotkey('ctrl', 'alt', '/')
                                    time.sleep(1)
                                    pyautogui.hotkey('ctrl', 'a')
                                    time.sleep(0.2)
                                    pyautogui.press('backspace')
                                    time.sleep(0.2)
                                    pyautogui.write(kontak_pilih)
                                    time.sleep(2.5)
                                    pyautogui.press('enter')
                                    time.sleep(1)
                                    
                                    if aksi == "send_whatsapp":
                                        pyautogui.write(pesan_wa)
                                        time.sleep(0.5)
                                        pyautogui.press('enter')
                                        lala_bicara("Message sent successfully.")
                                    else:
                                        lala_bicara("Chat opened.")
                                else:
                                    pyautogui.hotkey('ctrl', 't')
                                    time.sleep(1)
                                    pyautogui.write(url_wa)
                                    time.sleep(0.5)
                                    pyautogui.press('enter')
                                    time.sleep(12)
                                    pyautogui.press('enter')
                                    
                                    if aksi == "send_whatsapp":
                                        lala_bicara("Message sent successfully.")
                                    else:
                                        lala_bicara("Chat opened.")

                elif aksi == "buat_notepad":
                    judul = str(data.get("judul", "tanpa_judul"))
                    folder_target = str(data.get("folder", "")).lower()
                    
                    kandidat_folder = []
                    jalur_cari = [os.path.expanduser("~"), "D:\\"]
                    
                    for dasar in jalur_cari:
                        if os.path.exists(dasar):
                            for akar, direktori, file in os.walk(dasar):
                                kedalaman = akar[len(dasar):].count(os.sep)
                                if kedalaman > 3:
                                    direktori.clear()
                                    continue
                                for d in direktori:
                                    if d.lower() == folder_target:
                                        kandidat_folder.append(os.path.join(akar, d))
                                        
                    folder_terpilih = ""
                    
                    if not kandidat_folder:
                        lala_bicara(f"I couldn't find a folder named {folder_target} in the main directories.")
                    elif len(kandidat_folder) == 1:
                        folder_terpilih = kandidat_folder[0]
                    else:
                        opsi_baca = []
                        for i, k in enumerate(kandidat_folder[:3]):
                            induk = os.path.basename(os.path.dirname(k))
                            if not induk or len(induk) <= 2:
                                induk = os.path.dirname(k)
                            opsi_baca.append(f"Option {i+1} inside {induk}")
                            
                        tanya_folder = "I found multiple folders. Which one? " + ", or ".join(opsi_baca)
                        print(f"[Konfirmasi Folder Notepad] {tanya_folder}")
                        lala_bicara(tanya_folder)
                        
                        valid = False
                        while not valid:
                            jwb = dengar_perintah()
                            if jwb:
                                jwb = jwb.lower()
                                if any(kata in jwb for kata in ["batal", "cancel", "stop", "abort"]):
                                    lala_bicara("Notepad creation cancelled.")
                                    break
                                    
                                for i, k in enumerate(kandidat_folder[:3]):
                                    if str(i+1) in jwb or "option" in jwb:
                                        folder_terpilih = k
                                        valid = True
                                        break
                                        
                            if not valid:
                                lala_bicara("Invalid choice. Please specify the option number or say cancel.")
                        else:
                            lala_bicara("Please tell me which folder option, or say cancel.")
                            
                    if folder_terpilih:
                        jalur_file = os.path.join(folder_terpilih, f"{judul}.txt")
                        with open(jalur_file, 'a') as f:
                            pass
                        os.startfile(jalur_file)
                        lala_bicara(f"Notepad {judul} created and opened. Ready to type.")

                elif aksi == "buka_notepad":
                    judul = str(data.get("judul", "")).lower()
                    target_file = f"{judul}.txt"
                    
                    ps_cek = f"$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {{\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}}\n'@\nAdd-Type $c\n$p = Get-Process notepad -ErrorAction SilentlyContinue | Where-Object {{$_.MainWindowTitle -match '{judul}'}} | Select-Object -First 1\nif ($p) {{\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\nWrite-Output 'SUDAH_BUKA'\n}}"
                    
                    with open("cek_judul_notepad.ps1", "w") as f:
                        f.write(ps_cek)
                        
                    hasil_cek = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", "cek_judul_notepad.ps1"], capture_output=True, text=True)
                    os.remove("cek_judul_notepad.ps1")
                    
                    if "SUDAH_BUKA" in hasil_cek.stdout:
                        lala_bicara(f"Notepad {judul} is already open. Pulling it to the front.")
                    else:
                        kandidat_file = []
                        jalur_cari = ["D:\\", "E:\\", "F:\\"]
                        
                        for dasar in jalur_cari:
                            if os.path.exists(dasar):
                                for akar, direktori, file in os.walk(dasar):
                                    if any(skip in akar for skip in ['$RECYCLE.BIN', 'System Volume Information']):
                                        continue
                                    kedalaman = akar[len(dasar):].count(os.sep)
                                    if kedalaman > 4:
                                        direktori.clear()
                                        continue
                                    for f in file:
                                        if f.lower() == target_file:
                                            kandidat_file.append(os.path.join(akar, f))
                                            
                        file_terpilih = ""
                        
                        if not kandidat_file:
                            lala_bicara(f"I couldn't find any text file named {judul} in drives D, E, or F.")
                        elif len(kandidat_file) == 1:
                            file_terpilih = kandidat_file[0]
                        else:
                            opsi_baca = []
                            for i, k in enumerate(kandidat_file[:3]):
                                drive = os.path.splitdrive(k)[0]
                                induk = os.path.basename(os.path.dirname(k))
                                if not induk or len(induk) <= 2:
                                    induk = "root directory"
                                opsi_baca.append(f"Option {i+1} in drive {drive} folder {induk}")
                                
                            tanya_file = "I found multiple files. Which one? " + ", or ".join(opsi_baca)
                            print(f"[Konfirmasi Buka Notepad] {tanya_file}")
                            lala_bicara(tanya_file)
                            
                            valid = False
                            while not valid:
                                jwb = dengar_perintah()
                                if jwb:
                                    jwb = jwb.lower()
                                    if any(kata in jwb for kata in ["batal", "cancel", "stop", "abort"]):
                                        lala_bicara("Opening notepad cancelled.")
                                        break
                                        
                                    for i, k in enumerate(kandidat_file[:3]):
                                        if str(i+1) in jwb or "option" in jwb:
                                            file_terpilih = k
                                            valid = True
                                            break
                                            
                                if not valid:
                                    lala_bicara("Invalid choice. Please specify the option number or say cancel.")
                            else:
                                lala_bicara("Please tell me which file option, or say cancel.")
                                
                        if file_terpilih:
                            os.startfile(file_terpilih)
                            time.sleep(1)
                            
                            ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process notepad -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                            with open("tarik_notepadb.ps1", "w") as f:
                                f.write(ps_code)
                            os.system("powershell -ExecutionPolicy Bypass -File tarik_notepadb.ps1")
                            os.remove("tarik_notepadb.ps1")
                            
                            time.sleep(0.5)
                            pyautogui.hotkey('ctrl', 'end')
                            lala_bicara(f"Notepad {judul} opened. Cursor is at the end.")

                elif aksi in ["notepad_tulis", "notepad_hapus_kata", "notepad_hapus_kalimat", "notepad_baris_baru", "notepad_paragraf_baru", "notepad_simpan", "notepad_tutup", "notepad_cari", "notepad_cari_lanjut", "notepad_blok_kata", "notepad_hapus", "notepad_kursor_ujung", "notepad_tutup_cari"]:
                    global terakhir_ditulis
                    if 'terakhir_ditulis' not in globals():
                        terakhir_ditulis = ""
                        
                    ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process notepad | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                    with open("tarik_notepad.ps1", "w") as f:
                        f.write(ps_code)
                    os.system("powershell -ExecutionPolicy Bypass -File tarik_notepad.ps1")
                    os.remove("tarik_notepad.ps1")
                    time.sleep(0.5)
                    
                    if aksi == "notepad_tulis":
                        if "arguments" in data and isinstance(data["arguments"], dict):
                            teks = str(data["arguments"].get("teks", data.get("teks", "")))
                        else:
                            teks = str(data.get("teks", ""))

                        while '\n\n' in teks:
                            teks = teks.replace('\n\n', '\n')
                      
                        if len(teks) > 1500:
                            teks = teks[:1500]
                            
                        pyautogui.write(teks + " ", interval=0.01)
                        terakhir_ditulis = teks + " "
                        
                    elif aksi == "notepad_hapus_kata":
                        jumlah = int(data.get("jumlah", 1))
                        for _ in range(jumlah):
                            pyautogui.hotkey('ctrl', 'backspace')
                            time.sleep(0.05)
                            
                    elif aksi == "notepad_hapus_kalimat":
                        if terakhir_ditulis:
                            pyautogui.press('backspace', presses=len(terakhir_ditulis))
                            terakhir_ditulis = ""
                            
                    elif aksi == "notepad_baris_baru":
                        pyautogui.press('enter')
                        terakhir_ditulis = ""
                        
                    elif aksi == "notepad_paragraf_baru":
                        pyautogui.press('enter', presses=2)
                        terakhir_ditulis = ""
                        
                    elif aksi == "notepad_simpan":
                        pyautogui.hotkey('ctrl', 's')
                        lala_bicara("Text saved.")
                        
                    elif aksi == "notepad_cari":
                        teks_cari = str(data.get("teks", ""))
                        pyautogui.hotkey('ctrl', 'f')
                        time.sleep(0.5)
                        pyautogui.write(teks_cari)
                        time.sleep(0.2)
                        pyautogui.press('enter')
                        
                    elif aksi == "notepad_tutup_cari":
                        try:
                            app = Application(backend="uia").connect(path="notepad.exe", timeout=2)
                            j = app.top_window()
                            j.set_focus()
                            j.type_keys('{ESC}')
                            j.type_keys('{ESC}')
                            time.sleep(0.5)
                            pyautogui.press('right')
                        except Exception:
                            pass
                        
                    elif aksi == "notepad_cari_lanjut":
                        pyautogui.press('f3')
                        
                    elif aksi == "notepad_blok_kata":
                        arah = str(data.get("arah", "belakang")).lower()
                        jumlah = int(data.get("jumlah", 1))
                        tmbl = '{LEFT}' if arah == 'belakang' else '{RIGHT}'
                        try:
                            j = Application(backend="uia").connect(path="notepad.exe", timeout=2).top_window()
                            j.set_focus()
                            for _ in range(jumlah):
                                j.type_keys(f'^+{tmbl}')
                        except Exception:
                            pass
                            
                    elif aksi == "notepad_hapus":
                        pyautogui.press('backspace')
                        
                    elif aksi == "notepad_kursor_ujung":
                        posisi = str(data.get("posisi", "akhir")).lower()
                        
                        pyautogui.press('esc')
                        time.sleep(0.2)
                        
                        if posisi == "awal":
                            pyautogui.hotkey('ctrl', 'home')
                        else:
                            pyautogui.hotkey('ctrl', 'end')

                    elif aksi == "notepad_tutup":
                        time.sleep(1)
                        pyautogui.hotkey('ctrl', 's')
                        time.sleep(1)
                        pyautogui.hotkey('alt', 'f4')
                        lala_bicara("Notepad saved and closed.")

                elif aksi == "undo_layar":
                        pyautogui.keyDown('ctrl')
                        time.sleep(0.1)
                        pyautogui.press('z')
                        time.sleep(0.1)
                        pyautogui.keyUp('ctrl')

                elif aksi == "tekan_keyboard":
                        tmbl = str(data.get("tombol", "")).lower()
                        jml = int(data.get("jumlah", 1))
                        k = {'kanan': 'right', 'kiri': 'left', 'atas': 'up', 'bawah': 'down', 'spasi': 'space', 'hapus': 'backspace', 'kembali': 'backspace', 'hapus depan': 'delete'}
                        pyautogui.press(k.get(tmbl, tmbl), presses=jml, interval=0.05)

                elif aksi == "buka_word":
                    cek_taskbar = subprocess.run(['powershell', '-Command', "Get-Process winword -ErrorAction SilentlyContinue"], capture_output=True)
                    
                    if cek_taskbar.returncode == 0:
                        ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process winword -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                        with open("tarik_word_buka.ps1", "w") as f:
                            f.write(ps_code)
                        os.system("powershell -ExecutionPolicy Bypass -File tarik_word_buka.ps1")
                        os.remove("tarik_word_buka.ps1")
                        lala_bicara("Word is already open. Pulling it to the front.")
                    else:
                        os.system("cmd /c start winword")
                        time.sleep(3.5)
                        pyautogui.press('enter')
                        lala_bicara("Microsoft Word opened.")

                elif aksi == "buka_word_file":
                    judul = str(data.get("judul", ""))
                    path_file = f"D:\\Word\\{judul}.docx"
                    
                    if os.path.exists(path_file):
                        os.startfile(path_file)
                        time.sleep(2.5)
                        
                        ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process winword -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                        with open("tarik_word_file.ps1", "w") as f:
                            f.write(ps_code)
                        os.system("powershell -ExecutionPolicy Bypass -File tarik_word_file.ps1")
                        os.remove("tarik_word_file.ps1")
                        
                        time.sleep(0.5)
                        pyautogui.hotkey('ctrl', 'end')
                        lala_bicara(f"Word file {judul} opened. Cursor is at the end.")
                    else:
                        lala_bicara(f"I couldn't find {judul}.docx in your Drive D Word folder.")

                elif aksi in ["word_tulis", "word_hapus_kata", "word_hapus_kalimat", "word_baris_baru", "word_paragraf_baru", "word_simpan", "word_tutup", "word_cari", "word_cari_lanjut", "word_blok_kata", "word_hapus", "word_kursor_ujung", "word_tutup_cari"]:
                    global terakhir_ditulis_word
                    if 'terakhir_ditulis_word' not in globals():
                        terakhir_ditulis_word = ""
                        
                    ps_code = "$c = @'\nusing System;\nusing System.Runtime.InteropServices;\npublic class W {\n[DllImport(\"user32.dll\")] public static extern bool SetForegroundWindow(IntPtr h);\n[DllImport(\"user32.dll\")] public static extern bool ShowWindow(IntPtr h, int n);\n}\n'@\nAdd-Type $c\n$p = Get-Process winword -ErrorAction SilentlyContinue | Where-Object {$_.MainWindowHandle -ne 0} | Select-Object -First 1\nif ($p) {\n[W]::ShowWindow($p.MainWindowHandle, 9)\n[W]::SetForegroundWindow($p.MainWindowHandle)\n}"
                    with open("tarik_word_aksi.ps1", "w") as f:
                        f.write(ps_code)
                    os.system("powershell -ExecutionPolicy Bypass -File tarik_word_aksi.ps1")
                    os.remove("tarik_word_aksi.ps1")
                    time.sleep(0.5)
                    
                    if aksi == "word_tulis":
                        if "arguments" in data and isinstance(data["arguments"], dict):
                            teks = str(data["arguments"].get("teks", data.get("teks", "")))
                        else:
                            teks = str(data.get("teks", ""))

                        while '\n\n' in teks:
                            teks = teks.replace('\n\n', '\n')
                            
                        if len(teks) > 1500:
                            teks = teks[:1500]
                            
                        pyautogui.write(teks + " ", interval=0.01)
                        terakhir_ditulis_word = teks + " "
                        
                    elif aksi == "word_hapus_kata":
                        jumlah = int(data.get("jumlah", 1))
                        for _ in range(jumlah):
                            pyautogui.hotkey('ctrl', 'backspace')
                            time.sleep(0.05)
                            
                    elif aksi == "word_hapus_kalimat":
                        if terakhir_ditulis_word:
                            pyautogui.press('backspace', presses=len(terakhir_ditulis_word))
                            terakhir_ditulis_word = ""
                            
                    elif aksi == "word_baris_baru":
                        pyautogui.press('enter')
                        terakhir_ditulis_word = ""
                        
                    elif aksi == "word_paragraf_baru":
                        pyautogui.press('enter', presses=2)
                        terakhir_ditulis_word = ""
                        
                    elif aksi == "word_simpan":
                        judul = str(data.get("judul", "")).strip()
                        
                        if not judul or judul.lower() in ["none", "null", "[nama_file_jika_ada_disebutkan]"]:
                            pyautogui.hotkey('ctrl', 's')
                            lala_bicara("Word document saved.")
                        else:
                            pyautogui.press('f12')
                            time.sleep(1.5)
                            pyautogui.write(f"D:\\Word\\{judul}.docx")
                            time.sleep(0.5)
                            pyautogui.press('enter')
                            lala_bicara(f"Word document saved as {judul}.")
                        
                    elif aksi == "word_cari":
                        teks_cari = str(data.get("teks", ""))
                        pyautogui.hotkey('ctrl', 'f')
                        time.sleep(0.5)
                        pyautogui.write(teks_cari)
                        
                    elif aksi == "word_tutup_cari":
                        pyautogui.press('esc')
                        time.sleep(0.2)
                        pyautogui.press('esc')
                        time.sleep(0.5)
                        pyautogui.press('right')
                        
                    elif aksi == "word_cari_lanjut":
                        pyautogui.press('enter') 
                        
                    elif aksi == "word_blok_kata":
                        arah = str(data.get("arah", "belakang")).lower()
                        jumlah = int(data.get("jumlah", 1))
                        tmbl = 'left' if arah == 'belakang' else 'right'
                        for _ in range(jumlah):
                            pyautogui.hotkey('ctrl', 'shift', tmbl)
                            
                    elif aksi == "word_hapus":
                        pyautogui.press('backspace')
                        
                    elif aksi == "word_kursor_ujung":
                        posisi = str(data.get("posisi", "akhir")).lower()
                        pyautogui.press('esc')
                        time.sleep(0.2)
                        if posisi == "awal":
                            pyautogui.hotkey('ctrl', 'home')
                        else:
                            pyautogui.hotkey('ctrl', 'end')

                    elif aksi == "word_tutup":
                        pyautogui.hotkey('alt', 'f4')
                        lala_bicara("Word closed.")

                    elif aksi == "word_format":
                        try:
                            word = win32com.client.GetActiveObject("Word.Application")
                            selection = word.Selection
        
                            if data.get("font"): selection.Font.Name = data["font"]
                            if data.get("size"): selection.Font.Size = data["size"]
                            if data.get("bold") is not None: selection.Font.Bold = 1 if data["bold"] else 0
                            if data.get("italic") is not None: selection.Font.Italic = 1 if data["italic"] else 0
                            if data.get("underline") is not None: selection.Font.Underline = 1 if data["underline"] else 0
        
                            if data.get("color"):
                                warna = str(data["color"]).lower()
                                if warna == "red": selection.Font.Color = 255       
                                elif warna == "blue": selection.Font.Color = 16711680 
                                else: selection.Font.Color = 0                      
            
                            if data.get("alignment"):
                                align = str(data["alignment"]).lower()
                                if align == "center": selection.ParagraphFormat.Alignment = 1
                                elif align == "right": selection.ParagraphFormat.Alignment = 2
                                elif align == "justify": selection.ParagraphFormat.Alignment = 3
                                else: selection.ParagraphFormat.Alignment = 0 
            
                            if data.get("line_spacing"):
                                spasi = float(data["line_spacing"])
                                if spasi == 1.0: selection.ParagraphFormat.LineSpacingRule = 0
                                elif spasi == 1.5: selection.ParagraphFormat.LineSpacingRule = 1
                                elif spasi >= 2.0: selection.ParagraphFormat.LineSpacingRule = 2

                            lala_bicara("Formatting applied.")
                        except Exception as e:
                            lala_bicara("Failed to format Word document.")
                
                jawaban_sarkas = [
                    "Done. Go Fuck Yourself.", 
                    "Sorted. Happy now bitch?", 
                    "There you go. Shut the Fuck up", 
                    "Fuck You. Leave me alone.",
                    "Whatever. It's done and Fuck You.",
                    "What the hell I'm doing here?"
                ]
                ui_widget.ubah_status("standby")
                lala_bicara(random.choice(jawaban_sarkas))
                    
            except json.JSONDecodeError:
                print(f"Output mentah AI: {teks_json}")
                ui_widget.ubah_status("error")
                lala_bicara("Format error.")
                time.sleep(2)
                
        except Exception as e:
            print(f"Error: {e}")
            ui_widget.ubah_status("error")
            lala_bicara("Brain crashed.")
            time.sleep(2)

if __name__ == "__main__":
    try:
        lala_env = os.environ.copy()
        lala_env["OLLAMA_MODELS"] = os.path.join(os.path.dirname(os.path.abspath(__file__)), "models")
        subprocess.Popen(["ollama", "serve"], env=lala_env, creationflags=0x08000000)
    except Exception:
        pass

    root = tk.Tk()
    app_ui = LalaWidget(root)
    ai_thread = threading.Thread(target=jalankan_otak_lala, args=(app_ui,), daemon=True)
    ai_thread.start()
    root.mainloop()
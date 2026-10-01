from datetime import datetime
import hashlib
import io
import json
import os
import sys
import threading
import time
import urllib.parse
import urllib.request
import uuid
import warnings
import ctypes
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# Tambahan Anti-DPI Awareness agar PyAutoGUI akurat di resolusi & skala layar baru
#pyinstaller --noconsole --onefile --hidden-import=pystray --hidden-import=pyperclip --hidden-import=PIL --hidden-import=PIL.Image --hidden-import=PIL.ImageTk --hidden-import=PIL.ImageDraw --hidden-import=pandas --hidden-import=openpyxl --hidden-import=pyautogui --hidden-import=keyboard --hidden-import=uuid --hidden-import=tkinter --hidden-import=tkinter.filedialog --hidden-import=tkinter.messagebox --hidden-import=tkinter.scrolledtext launcher.py
try:
  ctypes.windll.shcore.SetProcessDpiAwareness(2)
except:
  try:
    ctypes.windll.user32.SetProcessDPIAware()
  except:
    pass

try:
  import keyboard
  import openpyxl
  import pandas as pd
  import pyautogui
  import pyperclip
  import pystray
  from PIL import Image, ImageTk, ImageDraw
except ImportError as e:
  root_err = tk.Tk()
  root_err.withdraw()
  messagebox.showerror(
      "Error Modul",
      f"Modul belum lengkap: {e}\nSilakan instal: pip install pandas openpyxl"
      " pyperclip pyautogui keyboard pystray Pillow",
  )
  sys.exit()

warnings.simplefilter("ignore", category=UserWarning)

if getattr(sys, "frozen", False):
  base_dir = os.path.dirname(os.path.abspath(sys.executable))
else:
  base_dir = os.path.dirname(os.path.abspath(__file__))

active_threads = {}
buttons_dict = {}
label_stats_dict = {}
progress_bars_dict = {}
rentang_entry_dict = {}

VERSI_APLIKASI = "3.1.13"

LICENSE_FILE = os.path.join(base_dir, "license.key")
CONFIG_FILE = os.path.join(base_dir, "config_kib.json")

SHEET_TRIAL_UTAMA = "Data"
HWID_LOKAL = str(uuid.getnode())

is_paused = False
is_stopped = False
app_is_running = True
stop_flags = {}
skip_current_row_flags = {}

is_paused_pc = {}
is_stopped_pc = {}
stop_flags_pc = {}

sp_klik = None
sp_drop = None
sp_scroll = None
sp_simpan = None

# Variabel Global UI untuk Sistem Dropdown Sheet Tunggal
combo_sheet_pilih = None
panel_utama_aktif = None
panel_rentang_aktif = None
lbl_stat_aktif = None
canvas_prog_aktif = None
rect_prog_aktif = None
btn_jalankan_aktif = None

# ==========================================
# 📱 TOKEN TELEGRAM & DEFAULT SUPER ADMIN ID
# ==========================================


# ==========================================
# 📂 FUNGSI CONFIG JSON BIASA
# ==========================================
def simpan_config(data_config):
  try:
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
      json.dump(data_config, f, indent=4, ensure_ascii=False)
  except Exception as e:
    print(f"Gagal menyimpan config: {e}")


def ekspor_konfigurasi():
  if not os.path.exists(CONFIG_FILE):
    messagebox.showerror(
        "Error", "File konfigurasi (config_kib.json) belum ada!"
    )
    return
  file_tujuan = filedialog.asksaveasfilename(
      defaultextension=".json",
      filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
      initialfile="config_kib_backup.json",
      title="Ekspor Konfigurasi KIB",
  )
  if file_tujuan:
    try:
      with open(CONFIG_FILE, "r", encoding="utf-8") as f_src:
        konten = f_src.read()
      with open(file_tujuan, "w", encoding="utf-8") as f_dst:
        f_dst.write(konten)
      messagebox.showinfo(
          "Berhasil", f"Konfigurasi berhasil diekspor ke:\n{file_tujuan}"
      )
    except Exception as e:
      messagebox.showerror(
          "Gagal Ekspor", f"Terjadi kesalahan saat mengekspor: {e}"
      )


def impor_konfigurasi():
  file_sumber = filedialog.askopenfilename(
      filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")],
      title="Pilih File Konfigurasi KIB",
  )
  if file_sumber:
    if messagebox.askyesno(
        "Konfirmasi Impor",
        "Mengimpor konfigurasi baru akan menimpa pengaturan koordinat &"
        " tahap saat ini. Lanjutkan?",
    ):
      try:
        with open(file_sumber, "r", encoding="utf-8") as f_src:
          konten = f_src.read()
        with open(CONFIG_FILE, "w", encoding="utf-8") as f_dst:
          f_dst.write(konten)
        messagebox.showinfo(
            "Berhasil",
            "Konfigurasi berhasil diimpor! Silakan klik 'Refresh Aplikasi'.",
        )
        refresh_aplikasi()
      except Exception as e:
        messagebox.showerror(
            "Gagal Impor", f"Terjadi kesalahan saat mengimpor: {e}"
        )


def ambil_daftar_sheet_excel():
  excel_path = os.path.join(base_dir, "data_kib.xlsx")
  default_sheet = ["Data"]
  if os.path.exists(excel_path):
    try:
      wb = openpyxl.load_workbook(excel_path, read_only=True)
      sheets = wb.sheetnames
      wb.close()
      if sheets:
        return sheets
    except:
      pass
  return default_sheet


def muat_config():
  sheet_aktual = ambil_daftar_sheet_excel()

  setting_default_baru = {
      "DELAY_KLIK": "0.30",
      "DELAY_DROPDOWN": "0.40",
      "DELAY_SCROLL": "0.8",
      "DELAY_KOSONGKAN": "1.5",
      "DELAY_CARI": "1.5",
      "SCROLL_AWAL": "1000",
      "SCROLL_JARAK": "-380",
      "TELEGRAM_CHAT_ID": "",
  }

  if not os.path.exists(CONFIG_FILE):
    base_config = {
        "setting": {},
        "kordinat": {},
        "tahap": {},
        "status_online": {},
        "global_control": {"pause": False, "stop": False},
    }
    for s in sheet_aktual:
      base_config["setting"][s] = setting_default_baru.copy()
      base_config["kordinat"][s] = {}
      base_config["tahap"][s] = []
    simpan_config(base_config)
    return base_config

  try:
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
      data = json.load(f)

    if "setting" not in data:
      data["setting"] = {}
    if "kordinat" not in data:
      data["kordinat"] = {}
    if "tahap" not in data:
      data["tahap"] = {}
    if "status_online" not in data:
      data["status_online"] = {}
    if "global_control" not in data:
      data["global_control"] = {"pause": False, "stop": False}

    ada_perubahan = False
    sheet_lama_setting = list(data["setting"].keys())

    if len(sheet_lama_setting) == len(sheet_aktual):
      for lama, baru in zip(sheet_lama_setting, sheet_aktual):
        if lama != baru:
          if lama in data["setting"]:
            data["setting"][baru] = data["setting"].pop(lama)
          if lama in data["kordinat"]:
            data["kordinat"][baru] = data["kordinat"].pop(lama)
          if lama in data["tahap"]:
            data["tahap"][baru] = data["tahap"].pop(lama)
          ada_perubahan = True

    for sheet_tersimpan in list(data["setting"].keys()):
      if sheet_tersimpan not in sheet_aktual:
        data["setting"].pop(sheet_tersimpan, None)
        data["kordinat"].pop(sheet_tersimpan, None)
        data["tahap"].pop(sheet_tersimpan, None)
        ada_perubahan = True

    for s in sheet_aktual:
      if s not in data["setting"]:
        data["setting"][s] = setting_default_baru.copy()
        ada_perubahan = True
      else:
        for key_param, val_param in setting_default_baru.items():
          if key_param not in data["setting"][s]:
            data["setting"][s][key_param] = val_param
            ada_perubahan = True

      if s not in data["kordinat"]:
        data["kordinat"][s] = {}
        ada_perubahan = True
      if s not in data["tahap"]:
        data["tahap"][s] = []
        ada_perubahan = True

    if ada_perubahan:
      simpan_config(data)

    return data
  except:
    return {
        "setting": {},
        "kordinat": {},
        "tahap": {},
        "status_online": {},
        "global_control": {"pause": False, "stop": False},
    }


def set_status_global(status_key, nilai):
  config = muat_config()
  if "global_control" not in config:
    config["global_control"] = {"pause": False, "stop": False}
  config["global_control"][status_key] = nilai

  if status_key == "pause" and not nilai:
    config["global_control"]["stop"] = False

  simpan_config(config)


def cek_status_global(status_key):
  config = muat_config()
  return config.get("global_control", {}).get(status_key, False)


def catat_komputer_aktif():
  hwid = HWID_LOKAL
  config = muat_config()

  if "status_online" not in config:
    config["status_online"] = {}

  config["status_online"][hwid] = {
      "terakhir_aktif": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
      "nama_host": os.getenv("COMPUTERNAME", "PC-KIB"),
  }

  simpan_config(config)


def cek_koneksi_internet():
  try:
    urllib.request.urlopen("https://api.telegram.org", timeout=3)
    return True
  except:
    return False


def pastikan_koneksi_stabil(hwid, sheet_name):
  while not cek_koneksi_internet():
    if (
        is_stopped
        or stop_flags.get(sheet_name, False)
        or is_stopped_pc.get(hwid, False)
        or cek_status_global("stop")
    ):
      raise InterruptedError("Dihentikan")
    time.sleep(3.0)


def inisialisasi_file_excel_otomatis():
  excel_path = os.path.join(base_dir, "data_kib.xlsx")
  if not os.path.exists(excel_path):
    try:
      wb = openpyxl.Workbook()
      default_sheet = wb.active
      default_sheet.title = "Data"

      daftar_sheet = ["Data"]
      for i, nama_s in enumerate(daftar_sheet):
        if i == 0:
          ws = default_sheet
        else:
          ws = wb.create_sheet(title=nama_s)
        ws.cell(row=1, column=1, value="Proses")
        ws.cell(row=1, column=2, value="Status")

      wb.save(excel_path)
      wb.close()
    except Exception as e:
      print(f"Gagal membuat template Excel otomatis: {e}")


inisialisasi_file_excel_otomatis()


def keluar_aplikasi():
  global is_stopped, app_is_running
  app_is_running = False
  is_stopped = True
  is_stopped_pc[HWID_LOKAL] = True
  for k in stop_flags:
    stop_flags[k] = True
  try:
    root.destroy()
  except:
    pass
  sys.exit()


def inisialisasi_header_excel_dari_koordinat(excel_path, sheet_name, nama_field):
  try:
    if not os.path.exists(excel_path):
      wb = openpyxl.Workbook()
      ws = wb.active
      ws.title = sheet_name
    else:
      wb = openpyxl.load_workbook(excel_path)
      if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(title=sheet_name)
      else:
        ws = wb[sheet_name]

    if ws.cell(row=1, column=1).value is None:
      ws.cell(row=1, column=1, value="Proses")
    if ws.cell(row=1, column=2).value is None:
      ws.cell(row=1, column=2, value="Status")

    sudah_ada = False
    for c in range(3, ws.max_column + 1):
      val_h = str(ws.cell(row=1, column=c).value or "").strip().lower()
      if val_h == str(nama_field).strip().lower():
        sudah_ada = True
        break

    if not sudah_ada and nama_field.lower() != "btn_simpan":
      col_idx = ws.max_column + 1
      ws.cell(row=1, column=col_idx, value=nama_field)

    wb.save(excel_path)
    wb.close()
  except Exception as e:
    print(f"Gagal memperbarui header Excel untuk sheet {sheet_name}: {e}")


def cetak_log(pesan):
  print(pesan)


def ambil_daftar_chat_id():
  config = muat_config()
  daftar_chat_id = set()
  if DEFAULT_CHAT_ID:
    daftar_chat_id.add(str(DEFAULT_CHAT_ID).strip())

  for sheet in ambil_daftar_sheet_excel():
    c_id = (
        str(config.get("setting", {}).get(sheet, {}).get("TELEGRAM_CHAT_ID", ""))
        .strip()
    )
    if c_id and c_id.lower() not in ["none", "nan", "", "masukkan_chat_id"]:
      daftar_chat_id.add(c_id)

  return list(daftar_chat_id)


def kirim_notifikasi_telegram(pesan, target_chat_id=None):
  chat_tujuan = (
      str(target_chat_id).strip()
      if target_chat_id
      else str(DEFAULT_CHAT_ID).strip()
  )
  if not chat_tujuan or "MASUKKAN_" in chat_tujuan:
    return
  try:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps(
        {"chat_id": chat_tujuan, "text": pesan, "parse_mode": "HTML"}
    ).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=5)
  except:
    pass


def kirim_foto_screenshot_telegram(
    chat_id, caption="📸 <b>Tangkapan Layar Komputer (AMSFMX)</b>"
):
  chat_id_str = str(chat_id) if chat_id else str(DEFAULT_CHAT_ID)
  if not chat_id_str or "MASUKKAN_" in chat_id_str:
    return
  try:
    screenshot = pyautogui.screenshot()
    buffer = io.BytesIO()
    screenshot.save(buffer, format="PNG")
    photo_data = buffer.getvalue()
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendPhoto"
    boundary = "----PythonTelegramBotBoundary"
    body = (
        b"--"
        + boundary.encode("utf-8")
        + b"\r\n"
        + b'Content-Disposition: form-data; name="chat_id"\r\n\r\n'
        + chat_id_str.encode("utf-8")
        + b"\r\n"
        + b"--"
        + boundary.encode("utf-8")
        + b"\r\n"
        + b'Content-Disposition: form-data; name="caption"\r\n\r\n'
        + caption.encode("utf-8")
        + b"\r\n"
        + b"--"
        + boundary.encode("utf-8")
        + b"\r\n"
        + b'Content-Disposition: form-data; name="parse_mode"\r\n\r\nHTML\r\n'
        + b"--"
        + boundary.encode("utf-8")
        + b"\r\n"
        + b'Content-Disposition: form-data; name="photo";'
        + b' filename="screenshot.png"\r\n'
        + b"Content-Type: image/png\r\n\r\n"
        + photo_data
        + b"\r\n"
        + b"--"
        + boundary.encode("utf-8")
        + b"--\r\n"
    )
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": f"multipart/form-data; boundary={boundary}"
        },
        method="POST",
    )
    urllib.request.urlopen(req, timeout=15)
  except:
    pass


def buat_text_progress_bar(selesai, total, panjang=10):
  if total <= 0:
    persen, terisi = 0, 0
  else:
    persen = int((selesai / total) * 100)
    terisi = int((selesai / total) * panjang)
  if terisi > panjang:
    terisi = panjang
  return f"[{'█' * terisi}{'░' * (panjang - terisi)}] {persen}% ({selesai}/{total})"


def ambil_ringkasan_statistik_kib():
  excel_file = os.path.join(base_dir, "data_kib.xlsx")
  if not os.path.exists(excel_file):
    return "❌ File Excel 'data_kib.xlsx' tidak ditemukan."
  hasil = f"📊 <b>STATISTIK PROGRES SHEET (PC: {HWID_LOKAL[:8]})</b>\n\n"
  try:
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    for sheet_name in wb.sheetnames:
      ws = wb[sheet_name]
      c_proses, c_status = None, None
      for col in range(1, ws.max_column + 1):
        v1 = str(ws.cell(row=1, column=col).value).strip().lower()
        if "proses" in v1:
          c_proses = col
        if "status" in v1:
          c_status = col
      total_ya, selesai = 0, 0
      for r in range(2, ws.max_row + 1):
        pros = (
            str(ws.cell(row=r, column=c_proses).value if c_proses else "")
            .strip()
            .lower()
        )
        stat = (
            str(ws.cell(row=r, column=c_status).value if c_status else "")
            .strip()
            .lower()
        )
        if pros == "ya":
          total_ya += 1
          if stat == "selesai":
            selesai += 1
      hasil += (
          f"• <b>{sheet_name}</b>:\n{buat_text_progress_bar(selesai, total_ya)}\n\n"
      )
    wb.close()
  except Exception as e:
    hasil += f"Gagal membaca statistik: {e}"
  return hasil


def simpan_resolusi_layar_otomatis():
  current_w, current_h = pyautogui.size()
  config = muat_config()
  try:
    for sheet in ambil_daftar_sheet_excel():
      if sheet not in config["setting"]:
        config["setting"][sheet] = {}
      config["setting"][sheet]["SCREEN_REF_WIDTH"] = str(current_w)
      config["setting"][sheet]["SCREEN_REF_HEIGHT"] = str(current_h)
    simpan_config(config)
    return True, f"{current_w}x{current_h}"
  except Exception as e:
    return False, str(e)


def kirim_menu_utama_telegram(chat_id):
  try:
    sheets = ambil_daftar_sheet_excel()
    keyboard_rows = []
    row_temp = []
    for i, s in enumerate(sheets):
      row_temp.append({"text": f"📁 {s}", "callback_data": f"menu_{s}"})
      if len(row_temp) == 2 or i == len(sheets) - 1:
        keyboard_rows.append(row_temp)
        row_temp = []
    keyboard_rows.extend([
        [
            {
                "text": "🖥️ Pilih ID Komputer / Perangkat",
                "callback_data": "cmd_jalankan_id_menu",
            }
        ],
        [
            {
                "text": "⏸ Jeda Global (Semua PC)",
                "callback_data": "global_pause_true",
            },
            {
                "text": "▶ Lanjut Global (Semua PC)",
                "callback_data": "global_pause_false",
            },
        ],
        [
            {"text": "📊 Cek Status", "callback_data": f"statuspc_{HWID_LOKAL}"},
            {
                "text": "🛑 STOP TOTAL GLOBAL (Semua PC)",
                "callback_data": "global_stop",
            },
        ],
    ])
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps({
        "chat_id": chat_id,
        "text": (
            "🤖 <b>PUSAT KONTROL UTAMA (SUPER ADMIN GLOBAL)</b>\n\nSemua perintah"
            " di bawah berdampak pada seluruh instance aktif:"
        ),
        "parse_mode": "HTML",
        "reply_markup": {"inline_keyboard": keyboard_rows},
    }).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=5)
  except:
    pass


def kirim_menu_sheet_lokal_user(chat_id):
  try:
    sheets = ambil_daftar_sheet_excel()
    keyboard_rows = []
    row_temp = []
    for i, s in enumerate(sheets):
      row_temp.append(
          {"text": f"📁 {s}", "callback_data": f"start_{HWID_LOKAL}_{s}"}
      )
      if len(row_temp) == 2 or i == len(sheets) - 1:
        keyboard_rows.append(row_temp)
        row_temp = []

    keyboard_rows.extend([
        [
            {"text": "⏸ Jeda", "callback_data": f"pausepc_{HWID_LOKAL}"},
            {"text": "▶ Lanjut", "callback_data": f"resumepc_{HWID_LOKAL}"},
        ],
        [
            {
                "text": "📊 Cek Status",
                "callback_data": f"statuspc_{HWID_LOKAL}",
            },
            {"text": "⏹ Stop Total", "callback_data": f"stoppc_{HWID_LOKAL}"},
        ],
    ])

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps({
        "chat_id": chat_id,
        "text": (
            "🤖 <b>MENU SHEET ANDA</b>\n\nPilih sheet atau kontrol di bawah"
            f" untuk mulai (PC ID: <code>{HWID_LOKAL[:8]}</code>):"
        ),
        "parse_mode": "HTML",
        "reply_markup": {"inline_keyboard": keyboard_rows},
    }).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=5)
  except:
    pass


def kirim_pilihan_mulai_kib(chat_id, sheet_name):
  try:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps({
        "chat_id": chat_id,
        "text": f"Pilih aksi untuk sheet <b>{sheet_name}</b>:",
        "parse_mode": "HTML",
        "reply_markup": {
            "inline_keyboard": [
                [
                    {
                        "text": f"▶ Mulai {sheet_name}",
                        "callback_data": f"start_{HWID_LOKAL}_{sheet_name}",
                    }
                ],
                [
                    {
                        "text": f"🔄 Reset Status {sheet_name}",
                        "callback_data": f"reset_{sheet_name}",
                    }
                ],
                [{"text": "🔙 Menu Utama", "callback_data": "menu_utama"}],
            ]
        },
    }).encode("utf-8")
    urllib.request.urlopen(
        urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        ),
        timeout=5,
    )
  except:
    pass


def kirim_daftar_pilihan_id_komputer(chat_id):
  try:
    config_data = muat_config()
    daftar_pc = config_data.get("status_online", {})
    keyboard_rows = []
    sudah_ditambahkan = set()

    if daftar_pc:
      for hwid, info in daftar_pc.items():
        nama_pc = info.get("nama_host", "PC")
        if hwid not in sudah_ditambahkan:
          keyboard_rows.append(
              [{"text": f"💻 {nama_pc} ({hwid})", "callback_data": f"pilihpc_{hwid}"}]
          )
          sudah_ditambahkan.add(hwid)

    if not keyboard_rows:
      hwid_lokal = HWID_LOKAL
      keyboard_rows.append(
          [{"text": f"💻 PC Ini ({hwid_lokal})", "callback_data": f"pilihpc_{hwid_lokal}"}]
      )

    keyboard_rows.append([{"text": "🔙 Menu Utama", "callback_data": "menu_utama"}])

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps({
        "chat_id": chat_id,
        "text": (
            "🖥️ <b>PILIH ID KOMPUTER</b>\n\nSilakan klik ID komputer di bawah"
            " untuk mengontrol instance tersebut:"
        ),
        "parse_mode": "HTML",
        "reply_markup": {"inline_keyboard": keyboard_rows},
    }).encode("utf-8")
    urllib.request.urlopen(
        urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        ),
        timeout=5,
    )
  except:
    pass


def kirim_menu_sheet_berdasarkan_id(chat_id, id_komputer):
  try:
    sheets = ambil_daftar_sheet_excel()
    keyboard_rows = []
    row_temp = []
    for i, s in enumerate(sheets):
      row_temp.append(
          {"text": f"📁 {s}", "callback_data": f"startid_{id_komputer}_{s}"}
      )
      if len(row_temp) == 2 or i == len(sheets) - 1:
        keyboard_rows.append(row_temp)
        row_temp = []

    keyboard_rows.extend([
        [
            {"text": "⏸ Jeda", "callback_data": f"pausepc_{id_komputer}"},
            {"text": "▶ Lanjut", "callback_data": f"resumepc_{id_komputer}"},
        ],
        [
            {
                "text": "📊 Cek Status",
                "callback_data": f"statuspc_{id_komputer}",
            },
            {"text": "⏹ Stop Total", "callback_data": f"stoppc_{id_komputer}"},
        ],
        [
            {
                "text": "🔙 Kembali Pilih PC",
                "callback_data": "cmd_jalankan_id_menu",
            },
            {"text": "🔙 Menu Utama", "callback_data": "menu_utama"},
        ],
    ])

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    data = json.dumps({
        "chat_id": chat_id,
        "text": (
            "📁 <b>PILIH SHEET UNTUK KOMPUTER</b>\n\n• <b>ID Komputer:</b>"
            f" <code>{id_komputer}</code>\n\nPilih sheet di bawah:"
        ),
        "parse_mode": "HTML",
        "reply_markup": {"inline_keyboard": keyboard_rows},
    }).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    urllib.request.urlopen(req, timeout=5)
  except:
    pass


def tele_reset_status_sheet(sheet_name):
  excel_path = os.path.join(base_dir, "data_kib.xlsx")
  if not os.path.exists(excel_path):
    return False, 0
  try:
    wb = openpyxl.load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
      wb.close()
      return False, 0
    ws = wb[sheet_name]

    col_status = None
    for r in range(1, 4):
      for col in range(1, ws.max_column + 1):
        val = str(ws.cell(row=r, column=col).value).strip().lower()
        if val == "status":
          col_status = col
          break
      if col_status:
        break

    if col_status:
      hitung_reset = 0
      for r in range(2, ws.max_row + 1):
        cell_target = ws.cell(row=r, column=col_status)
        if cell_target.value is not None and str(cell_target.value).strip() != "":
          cell_target.value = None
          hitung_reset += 1
      wb.save(excel_path)
      wb.close()
      update_statistik_gui(sheet_name)
      return True, hitung_reset
    wb.close()
  except:
    pass
  return False, 0


def tele_proses_aktivasi(chat_id, teks_pesan):
  hwid_asli = HWID_LOKAL
  pecah = teks_pesan.strip().split()

  if len(pecah) < 2:
    kirim_notifikasi_telegram(
        f"⚠️ <b>Format Kurang Lengkap!</b>\nGunakan format:\n<code>/aktivasi"
        f" ID_KOMPUTER KODE_LISENSI</code>\n\nID Komputer Anda saat ini:"
        f" <code>{hwid_asli}</code>",
        target_chat_id=chat_id,
    )
    return

  id_komputer_input = pecah[0]
  kode_lisensi_input = pecah[1]

  if id_komputer_input != hwid_asli:
    kirim_notifikasi_telegram(
        f"❌ <b>Aktivasi Gagal!</b>\nID Komputer yang Anda masukkan"
        f" (<code>{id_komputer_input}</code>) tidak cocok dengan ID komputer ini"
        f" (<code>{hwid_asli}</code>).",
        target_chat_id=chat_id,
    )
    return

  if kode_lisensi_input == KODE_FULL_RAHASIA:
    try:
      bound_hash = hashlib.sha256(
          (KODE_FULL_RAHASIA + hwid_asli).encode("utf-8")
      ).hexdigest()
      with open(LICENSE_FILE, "w") as f:
        f.write(bound_hash)
      kirim_notifikasi_telegram(
          f"✅ <b>Aktivasi Berhasil via Telegram!</b>\n\nID Komputer:"
          f" <code>{hwid_asli}</code>\nAplikasi telah terbuka penuh (Full"
          " Version).",
          target_chat_id=chat_id,
      )
    except Exception as e:
      kirim_notifikasi_telegram(
          f"❌ <b>Gagal Menyimpan Lisensi:</b> {e}", target_chat_id=chat_id
      )
  else:
    kirim_notifikasi_telegram(
        "❌ <b>Aktivasi Gagal!</b> Kode lisensi salah.", target_chat_id=chat_id
    )


def listener_perintah_telegram():
  global TELEGRAM_TOKEN, DEFAULT_CHAT_ID, KODE_FULL_RAHASIA
  global is_paused, is_stopped, stop_flags, app_is_running
  last_update_id = 0
  while True:
    if not app_is_running:
      break
    if not TELEGRAM_TOKEN or "MASUKKAN_" in TELEGRAM_TOKEN:
      time.sleep(5)
      continue
    try:
      url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getUpdates?offset={last_update_id + 1}&timeout=10"
      response = urllib.request.urlopen(
          urllib.request.Request(url, method="GET"), timeout=12
      )
      result = json.loads(response.read().decode("utf-8"))
      if result.get("ok"):
        for update in result.get("result", []):
          last_update_id = update["update_id"]

          chat_id_masuk = None
          if "callback_query" in update:
            chat_id_masuk = str(update["callback_query"]["message"]["chat"]["id"])
          elif "message" in update:
            chat_id_masuk = str(update["message"]["chat"]["id"])

          daftar_izin = ambil_daftar_chat_id()
          if chat_id_masuk not in daftar_izin:
            continue

          is_super_admin = chat_id_masuk == str(DEFAULT_CHAT_ID)

          if "callback_query" in update:
            cq = update["callback_query"]
            data_cmd, chat_id, message_id, cq_id = (
                cq.get("data"),
                cq["message"]["chat"]["id"],
                cq["message"]["message_id"],
                cq["id"],
            )
            jawab_callback_telegram(cq_id)

            if data_cmd == "menu_utama":
              if is_super_admin:
                kirim_menu_utama_telegram(chat_id)
              else:
                kirim_menu_sheet_lokal_user(chat_id)
            elif data_cmd.startswith("menu_"):
              if is_super_admin:
                sheet_target = data_cmd.replace("menu_", "", 1)
                kirim_pilihan_mulai_kib(chat_id, sheet_target)
            elif data_cmd == "cmd_jalankan_id_menu":
              if is_super_admin:
                kirim_daftar_pilihan_id_komputer(chat_id)
            elif data_cmd.startswith("pilihpc_"):
              if is_super_admin:
                kirim_menu_sheet_berdasarkan_id(
                    chat_id, data_cmd.replace("pilihpc_", "", 1)
                )
            elif data_cmd == "global_pause_true":
              if is_super_admin:
                set_status_global("pause", True)
                kirim_notifikasi_telegram(
                    "⏸ <b>JEDA GLOBAL DIAKTIFKAN!</b>\nSeluruh komputer yang"
                    " menjalankan otomasi akan dijeda.",
                    target_chat_id=chat_id,
                )
            elif data_cmd == "global_pause_false":
              if is_super_admin:
                set_status_global("pause", False)
                set_status_global("stop", False)
                kirim_notifikasi_telegram(
                    "▶ <b>LANJUT GLOBAL DIAKTIFKAN!</b>\nSeluruh komputer"
                    " dilanjutkan kembali.",
                    target_chat_id=chat_id,
                )
            elif data_cmd == "global_stop":
              if is_super_admin:
                set_status_global("stop", True)
                emergency_stop()
                kirim_notifikasi_telegram(
                    "🛑 <b>STOP TOTAL GLOBAL DIAKTIFKAN!</b>\nSeluruh komputer"
                    " dihentikan secara total.",
                    target_chat_id=chat_id,
                )
            elif data_cmd.startswith("pausepc_"):
              target_hwid = data_cmd.replace("pausepc_", "", 1)
              if target_hwid == HWID_LOKAL:
                is_paused_pc[HWID_LOKAL] = True
                is_paused = True
                kirim_notifikasi_telegram(
                    f"⏸ Bot pada PC <code>{HWID_LOKAL[:8]}</code> dijeda.",
                    target_chat_id=chat_id,
                )
            elif data_cmd.startswith("resumepc_"):
              target_hwid = data_cmd.replace("resumepc_", "", 1)
              if target_hwid == HWID_LOKAL:
                is_paused_pc[HWID_LOKAL] = False
                is_paused = False
                kirim_notifikasi_telegram(
                    f"▶ Bot pada PC <code>{HWID_LOKAL[:8]}</code> dilanjutkan.",
                    target_chat_id=chat_id,
                )
            elif data_cmd.startswith("stoppc_"):
              target_hwid = data_cmd.replace("stoppc_", "", 1)
              if target_hwid == HWID_LOKAL:
                is_stopped_pc[HWID_LOKAL] = True
                emergency_stop()
                kirim_notifikasi_telegram(
                    f"🛑 Bot pada PC <code>{HWID_LOKAL[:8]}</code> dihentikan.",
                    target_chat_id=chat_id,
                )
            elif data_cmd.startswith("statuspc_"):
              target_hwid = data_cmd.replace("statuspc_", "", 1)
              if target_hwid == HWID_LOKAL:
                kirim_notifikasi_telegram(
                    ambil_ringkasan_statistik_kib(), target_chat_id=chat_id
                )
            elif data_cmd.startswith("start_"):
              parts = data_cmd.replace("start_", "", 1).split("_", 1)
              if len(parts) == 2 and parts[0] == HWID_LOKAL:
                threading.Thread(
                    target=jalankan_otomasi_langsung_telegram,
                    args=(parts[1], chat_id),
                    daemon=True,
                ).start()
            elif data_cmd.startswith("startid_"):
              parts = data_cmd.replace("startid_", "", 1).split("_", 1)
              if len(parts) == 2 and parts[0] == HWID_LOKAL:
                threading.Thread(
                    target=jalankan_otomasi_langsung_telegram,
                    args=(parts[1], chat_id),
                    daemon=True,
                ).start()
            elif data_cmd.startswith("reset_"):
              sh_t = data_cmd.replace("reset_", "", 1)
              sukses, jum = tele_reset_status_sheet(sh_t)
              if sukses:
                kirim_notifikasi_telegram(
                    f"🔄 Sheet <b>{sh_t}</b> direset ({jum} baris).",
                    target_chat_id=chat_id,
                )
          elif "message" in update:
            msg = update["message"]
            text_raw = msg.get("text", "").strip()
            text_lower = text_raw.lower()
            chat_id = msg["chat"]["id"]

            if not app_is_running:
              continue

            if text_lower == "/start":
              if is_super_admin:
                kirim_menu_utama_telegram(chat_id)
              else:
                kirim_menu_sheet_lokal_user(chat_id)

            if is_super_admin:
              if text_lower in ["/jeda_global", "/pause_global"]:
                set_status_global("pause", True)
                kirim_notifikasi_telegram(
                    "⏸ Jeda Global aktif untuk semua PC.", target_chat_id=chat_id
                )
              elif text_lower in ["/lanjut_global", "/resume_global"]:
                set_status_global("pause", False)
                set_status_global("stop", False)
                kirim_notifikasi_telegram(
                    "▶ Lanjut Global aktif untuk semua PC.", target_chat_id=chat_id
                )
              elif text_lower in ["/stop_global"]:
                set_status_global("stop", True)
                emergency_stop()
                kirim_notifikasi_telegram(
                    "🛑 Stop Global aktif untuk semua PC.", target_chat_id=chat_id
                )

            if text_lower in ["/capture", "/screenshot", "/img"]:
              kirim_foto_screenshot_telegram(chat_id)
            elif text_lower.startswith("/aktivasi "):
              parameter_aktivasi = text_raw.split(" ", 1)[1]
              tele_proses_aktivasi(chat_id, parameter_aktivasi)
    except:
      time.sleep(3)
    time.sleep(1)


def jawab_callback_telegram(cq_id):
  try:
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/answerCallbackQuery"
    data = json.dumps({"callback_query_id": cq_id}).encode("utf-8")
    urllib.request.urlopen(
        urllib.request.Request(
            url,
            data=data,
            headers={"Content-Type": "application/json"},
            method="POST",
        ),
        timeout=3,
    )
  except:
    pass


def jalankan_otomasi_langsung_telegram(sheet_name, chat_id=None):
  excel_path = os.path.join(base_dir, "data_kib.xlsx")
  max_row = 100
  if os.path.exists(excel_path):
    try:
      wb = openpyxl.load_workbook(excel_path, data_only=True)
      if sheet_name in wb.sheetnames:
        max_row = wb[sheet_name].max_row
      wb.close()
    except:
      pass
  worker_otomasi_kib(
      sheet_name, "Online", 2, max_row, bypass_enter=True, chat_id=chat_id
  )


threading.Thread(target=listener_perintah_telegram, daemon=True).start()


def emergency_stop():
  global is_stopped, stop_flags
  is_stopped = True
  is_stopped_pc[HWID_LOKAL] = True
  for k in stop_flags:
    stop_flags[k] = True


def toggle_pause():
  global is_paused
  if not is_stopped:
    is_paused = not is_paused
    is_paused_pc[HWID_LOKAL] = is_paused


try:
  keyboard.add_hotkey("left", toggle_pause)
  keyboard.add_hotkey("right", emergency_stop)
except:
  pass


def cek_status_lisensi():
  if os.path.exists(LICENSE_FILE):
    try:
      with open(LICENSE_FILE, "r") as f:
        if (
            f.read().strip()
            == hashlib.sha256(
                (KODE_FULL_RAHASIA + HWID_LOKAL).encode("utf-8")
            ).hexdigest()
        ):
          return True
    except:
      pass
  return False


def cek_jeda_dan_berhenti_pc(hwid, sheet_name):
  if (
      is_stopped
      or stop_flags.get(sheet_name, False)
      or is_stopped_pc.get(hwid, False)
      or cek_status_global("stop")
  ):
    raise InterruptedError("Dihentikan")
  if skip_current_row_flags.get(sheet_name, False):
    raise StopIteration("SkipBaris")

  while (
      is_paused
      or is_paused_pc.get(hwid, False)
      or cek_status_global("pause")
  ):
    if (
        is_stopped
        or stop_flags.get(sheet_name, False)
        or is_stopped_pc.get(hwid, False)
        or cek_status_global("stop")
    ):
      raise InterruptedError("Dihentikan")
    time.sleep(0.2)


def skip_baris_kib(sheet_name):
  global skip_current_row_flags
  skip_current_row_flags[sheet_name] = True


def parse_tanggal(nilai):
  if pd.isna(nilai):
    return None
  val_str = str(nilai).strip()
  if val_str in ["", "nan", "NaT", "None", "-"]:
    return None
  if isinstance(nilai, (pd.Timestamp, datetime)):
    return nilai.strftime("%d/%m/%Y")
  try:
    if isinstance(nilai, (int, float)) or val_str.replace(".", "", 1).isdigit():
      v_int = int(float(val_str))
      if v_int > 10000000:
        return f"{(v_int // 1000000):02d}/{((v_int // 10000) % 100):02d}/{(v_int % 10000):04d}"
  except:
    pass
  for fmt in ["%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"]:
    try:
      return datetime.strptime(val_str, fmt).strftime("%d/%m/%Y")
    except ValueError:
      continue
  return None


def update_statistik_gui(sheet_name):
  excel_file = os.path.join(base_dir, "data_kib.xlsx")
  if not os.path.exists(excel_file):
    return
  try:
    wb = openpyxl.load_workbook(excel_file, data_only=True)
    if sheet_name not in wb.sheetnames:
      return
    ws = wb[sheet_name]
    c_p, c_s = None, None
    for col in range(1, ws.max_column + 1):
      h = str(ws.cell(row=1, column=col).value).strip().lower()
      if "proses" in h:
        c_p = col
      if "status" in h:
        c_s = col
    total_ya, selesai = 0, 0
    for r in range(2, ws.max_row + 1):
      if (
          str(ws.cell(row=r, column=c_p).value if c_p else "").strip().lower()
          == "ya"
      ):
        total_ya += 1
        if (
            str(ws.cell(row=r, column=c_s).value if c_s else "").strip().lower()
            == "selesai"
        ):
          selesai += 1
    sisa = total_ya - selesai

    if lbl_stat_aktif:
      lbl_stat_aktif.config(
          text=f"Total: {total_ya} | Selesai: {selesai} | Sisa: {sisa} | ETA: 0 dtk"
      )

    if canvas_prog_aktif and rect_prog_aktif:
      canvas_prog_aktif.update_idletasks()
      w = canvas_prog_aktif.winfo_width()
      if w <= 1:
        w = canvas_prog_aktif.winfo_reqwidth()
      if w <= 1:
        w = 230

      persen = min(1.0, (selesai / total_ya) if total_ya > 0 else 0.0)
      canvas_prog_aktif.coords(rect_prog_aktif, 0, 0, int(w * persen), 10)
      canvas_prog_aktif.update()
    wb.close()
  except Exception as e:
    print(f"Error update statistik: {e}")


def refresh_aplikasi():
  try:
    muat_config()
    daftar_sheets = ambil_daftar_sheet_excel()
    if combo_sheet_pilih:
      combo_sheet_pilih["values"] = daftar_sheets
      current_sel = combo_sheet_pilih.get().strip()
      if current_sel not in daftar_sheets and daftar_sheets:
        combo_sheet_pilih.set(daftar_sheets[0])
        ganti_sheet_aktif(None)

    sheet_aktif = (
        combo_sheet_pilih.get().strip()
        if combo_sheet_pilih
        else SHEET_TRIAL_UTAMA
    )
    update_statistik_gui(sheet_aktif)

    status_lisensi = (
        "Teraktivasi ✅" if cek_status_lisensi() else "Trial Mode ⚠️"
    )
    messagebox.showinfo(
        "Refresh Berhasil",
        f"🔄 Aplikasi berhasil di-refresh!\nStatus Lisensi:"
        f" {status_lisensi}\nData Excel & konfigurasi telah diperbarui.",
    )
  except Exception as e:
    messagebox.showerror("Error Refresh", f"Gagal melakukan refresh: {e}")


def worker_otomasi_kib(
    sheet_name,
    mode_koneksi,
    start_r=None,
    end_r=None,
    bypass_enter=False,
    chat_id=None,
):
  global stop_flags, is_stopped, skip_current_row_flags, is_stopped_pc

  hwid = HWID_LOKAL
  is_stopped_pc[hwid] = False
  stop_flags[sheet_name] = False
  skip_current_row_flags[sheet_name] = False

  is_full = cek_status_lisensi()

  if not is_full and sheet_name != SHEET_TRIAL_UTAMA:
    pesan_peringatan = (
        f"⚠️ <b>PERINGATAN AKSES TERKUNCI!</b>\n\nSeseorang mencoba menjalankan"
        f" sheet <b>'{sheet_name}'</b>, namun aplikasi masih dalam mode Trial"
        f" (Belum Aktivasi).\n\nHanya sheet <b>'{SHEET_TRIAL_UTAMA}'</b> yang"
        " diizinkan."
    )
    kirim_notifikasi_telegram(pesan_peringatan, target_chat_id=DEFAULT_CHAT_ID)

    messagebox.showerror(
        "Aktivasi Diperlukan",
        f"Sheet '{sheet_name}' terkunci!\n\nSelama masa trial (belum"
        f" aktivasi), Anda hanya dapat menggunakan sheet '{SHEET_TRIAL_UTAMA}'."
        " Silakan masukkan kode lisensi untuk membuka semua sheet.",
    )
    return

  excel_file = os.path.join(base_dir, "data_kib.xlsx")
  if not os.path.exists(excel_file):
    pesan_err = f"File '{excel_file}' tidak ditemukan!"
    messagebox.showerror("Error", pesan_err)
    return

  if mode_koneksi == "Online":
    if not cek_koneksi_internet():
      messagebox.showerror(
          "Koneksi Error",
          "Tidak ada koneksi internet! Mode Online memerlukan koneksi aktif.",
      )
      return

  if btn_jalankan_aktif:
    try:
      btn_jalankan_aktif.config(bg="#c0392b", text="⏳ Berjalan...")
    except:
      pass

  config = muat_config()

  try:
    setting_dict = {
        "DELAY_KLIK": "0.30",
        "DELAY_DROPDOWN": "0.40",
        "DELAY_SCROLL": "0.8",
        "DELAY_KOSONGKAN": "1.5",
        "DELAY_CARI": "1.5",
        "SCROLL_AWAL": "1000",
        "SCROLL_JARAK": "-380",
    }
    setting_dict.update(config.get("setting", {}).get(sheet_name, {}))

    try:
      if sp_klik:
        setting_dict["DELAY_KLIK"] = sp_klik.get()
      if sp_drop:
        setting_dict["DELAY_DROPDOWN"] = sp_drop.get()
      if sp_scroll:
        setting_dict["DELAY_SCROLL"] = sp_scroll.get()
      if sp_simpan:
        setting_dict["DELAY_KOSONGKAN"] = sp_simpan.get()
    except:
      pass

    coord_dict = config.get("kordinat", {}).get(sheet_name, {})
    df_tahap = pd.DataFrame(config.get("tahap", {}).get(sheet_name, []))
  except:
    return

  if not df_tahap.empty:
    for _, t in df_tahap.iterrows():
      fname = str(t.get("fname"))
      c_val = coord_dict.get(fname, {})
      if c_val.get("x", 0) == 0 and c_val.get("y", 0) == 0:
        if btn_jalankan_aktif:
          try:
            btn_jalankan_aktif.config(
                bg="#27ae60", text=f"🚀 Jalankan {sheet_name}"
            )
          except:
            pass
        messagebox.showerror(
            "Error Koordinat",
            f"Koordinat untuk field/tombol '{fname}' pada sheet '{sheet_name}'"
            " belum diatur (masih 0,0)!\n\nSilakan atur melalui menu '⚙️ Input"
            " Koordinat' terlebih dahulu.",
        )
        return

  is_simulasi = mode_koneksi == "Simulasi"
  mode_teks = "SIMULASI" if is_simulasi else mode_koneksi
  if not bypass_enter:
    cetak_log(f"siap ({mode_teks}).Tekan ENTER di keyboard untuk mulai...")
    messagebox.showinfo(
        "Informasi",
        "Arahkan kursor ke form AMS, lalu tekan ENTER di keyboard untuk mulai.",
    )
    while (
        not is_stopped
        and not is_stopped_pc.get(hwid, False)
        and not cek_status_global("stop")
    ):
      try:
        if keyboard.is_pressed("enter"):
          time.sleep(1.0)
          break
      except:
        pass
      time.sleep(0.1)

  def get_xy(fname):
    c = coord_dict.get(fname, {})
    rx, ry = c.get("x", 0), c.get("y", 0)
    return int(rx), int(ry)

  try:
    wb = openpyxl.load_workbook(excel_file)
    ws = wb[sheet_name]
    c_p, c_s = None, None
    for col in range(1, ws.max_column + 1):
      h = str(ws.cell(row=1, column=col).value).strip().lower()
      if "proses" in h:
        c_p = col
      if "status" in h:
        c_s = col

    sr, er = (start_r or 2), (end_r or ws.max_row)
    counter = 0

    for r_idx in range(sr, er + 1):
      if not is_full and counter >= 5:
        pesan_batas_trial = (
            f"⚠️ <b>BATAS TRIAL TERCAPAI!</b>\n\nSheet <b>'{sheet_name}'</b> telah"
            " mencapai batas maksimal 5 baris pada mode Trial.\n\nSilakan"
            " lakukan aktivasi master untuk memproses baris selanjutnya."
        )
        kirim_notifikasi_telegram(
            pesan_batas_trial, target_chat_id=DEFAULT_CHAT_ID
        )
        messagebox.showwarning(
            "Batas Trial",
            "Batas maksimal 5 baris untuk mode trial telah tercapai. Masukkan"
            " kode aktivasi untuk melanjutkan.",
        )
        break

      cek_jeda_dan_berhenti_pc(hwid, sheet_name)

      if mode_koneksi == "Online":
        pastikan_koneksi_stabil(hwid, sheet_name)

      if (
          c_s
          and str(ws.cell(row=r_idx, column=c_s).value).strip().lower()
          == "selesai"
          and not is_simulasi
      ):
        continue
      if (
          str(ws.cell(row=r_idx, column=c_p).value).strip().lower() != "ya"
      ):
        continue

      row_dict = {
          str(ws.cell(row=1, column=c).value): ws.cell(row=r_idx, column=c).value
          for c in range(1, ws.max_column + 1)
      }
      pyautogui.moveTo(500, 400)
      pyautogui.scroll(int(setting_dict.get("SCROLL_AWAL", 1000)))
      time.sleep(0.5)

      skip_this_row = False

      if not df_tahap.empty:
        for _, t in df_tahap.iterrows():
          cek_jeda_dan_berhenti_pc(hwid, sheet_name)
          aksi, fname, colex = (
              str(t.get("aksi")).lower(),
              str(t.get("fname")),
              str(t.get("colex")),
          )
          if aksi == "scroll":
            pyautogui.scroll(int(setting_dict.get("SCROLL_JARAK", -380)))
            time.sleep(float(setting_dict.get("DELAY_SCROLL", 0.8)))
            continue
          val = next(
              (
                  v
                  for k, v in row_dict.items()
                  if colex.lower() in str(k).lower()
              ),
              "",
          )
          if pd.isna(val) or str(val).strip() in ["", "-"]:
            continue
          x, y = get_xy(fname)
          if x == 0 and y == 0:
            continue

          if aksi == "tanggal":
            tgl = parse_tanggal(val)
            if tgl:
              pyautogui.click(x, y)
              time.sleep(0.3)
              pyautogui.hotkey("ctrl", "a")
              pyautogui.press("backspace")
              pyperclip.copy(tgl)
              pyautogui.hotkey("ctrl", "v")
              time.sleep(0.4)
          elif aksi in ["field"]:
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_KLIK", 0.30)))
            time.sleep(0.3)
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            pyperclip.copy(str(val).strip())
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.4)
          elif aksi in ["field_aman"]:
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_KLIK", 0.30)))
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.2)
            pyautogui.press("backspace")
            time.sleep(0.2)
            teks_dengan_spasi = str(val).strip() + " "
            pyperclip.copy(teks_dengan_spasi)
            time.sleep(0.2)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.5)
          elif aksi in ["dropdown"]:
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_DROPDOWN", 0.40)))
            pyautogui.hotkey("ctrl", "a")
            pyautogui.press("backspace")
            pyperclip.copy(str(val).strip())
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.5)
            pyautogui.press("enter")
            time.sleep(0.3)
          elif aksi in ["dropdown_klik"]:
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_KLIK", 0.40)))
            pyautogui.hotkey("ctrl", "a")
            time.sleep(0.15)
            pyautogui.press("backspace")
            time.sleep(0.2)
            pyperclip.copy(str(val).strip())
            time.sleep(0.15)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.8)
            pyautogui.press("down")
            time.sleep(0.3)
            pyautogui.press("enter")
            time.sleep(0.4)
          elif aksi == "tombol":
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_KOSONGKAN", 1.5)))
          elif aksi == "pencarian":
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_CARI", 1.5)))
            
            waktu_mulai_tunggu = time.time()
            max_waktu_tunggu = 15.0
            
            while True:
              cek_jeda_dan_berhenti_pc(hwid, sheet_name)
              try:
                r_px, g_px, b_px = pyautogui.pixel(x, y)
                if r_px > 230 and g_px > 230 and b_px > 230:
                  break
              except Exception:
                pass
              if time.time() - waktu_mulai_tunggu > max_waktu_tunggu:
                break
              time.sleep(0.3)
          elif aksi == "baris_hasil":
            try:
              r_px, g_px, b_px = pyautogui.pixel(x, y)
              if b_px > 150 and r_px < 100:
                pyautogui.click(x, y)
                time.sleep(float(setting_dict.get("DELAY_KOSONGKAN", 1.5)))
              else:
                skip_this_row = True
                break
            except Exception:
              skip_this_row = True
              break
          else:
            pyautogui.click(x, y)
            time.sleep(float(setting_dict.get("DELAY_KOSONGKAN", 1.5)))

      if skip_this_row:
        continue

      if is_simulasi:
        counter += 1
        time.sleep(0.5)
        continue

      xs, ys = get_xy("Btn_Simpan")
      if xs and ys:
        pyautogui.click(xs, ys)
        time.sleep(float(setting_dict.get("DELAY_KOSONGKAN", 1.5)))

      if c_s:
        ws.cell(row=r_idx, column=c_s, value="Selesai")
        wb.save(excel_file)
        update_statistik_gui(sheet_name)
      counter += 1
      time.sleep(1.0)

    if btn_jalankan_aktif:
      try:
        btn_jalankan_aktif.config(
            bg="#27ae60", text=f"🚀 Jalankan {sheet_name}"
        )
      except:
        pass

    pesan_selesai = (
        f"✅ <b>Otomasi Selesai!</b>\nPC: <b>{HWID_LOKAL[:8]}</b>\nSheet:"
        f" <b>{sheet_name}</b>\nMode: <b>{mode_koneksi}</b>"
    )

    if chat_id:
      kirim_notifikasi_telegram(pesan_selesai, target_chat_id=chat_id)
    else:
      kirim_notifikasi_telegram(
          pesan_selesai, target_chat_id=DEFAULT_CHAT_ID
      )

    messagebox.showinfo(
        "Selesai", f"Otomasi sheet {sheet_name} ({mode_koneksi}) Selesai!"
    )

  except InterruptedError:
    if btn_jalankan_aktif:
      try:
        btn_jalankan_aktif.config(
            bg="#27ae60", text=f"🚀 Jalankan {sheet_name}"
        )
      except:
        pass
  except Exception as e:
    if btn_jalankan_aktif:
      try:
        btn_jalankan_aktif.config(
            bg="#27ae60", text=f"🚀 Jalankan {sheet_name}"
        )
      except:
        pass
    if (
        not is_stopped
        and not is_stopped_pc.get(hwid, False)
        and not cek_status_global("stop")
    ):
      messagebox.showerror("Error", f"{e}")


def cek_koordinat_belum_diatur():
  config = muat_config()
  daftar_sheets = ambil_daftar_sheet_excel()

  laporan = "📋 LAPORAN KOORDINAT KOSONG (0,0)\n\n"
  ada_yang_kosong = False

  for sheet in daftar_sheets:
    coord_dict = config.get("kordinat", {}).get(sheet, {})
    tahap_list = config.get("tahap", {}).get(sheet, [])

    field_dibutuhkan = set()
    for t in tahap_list:
      fn = t.get("fname")
      if fn:
        field_dibutuhkan.add(fn)

    field_dibutuhkan.add("Btn_Simpan")

    kosong_di_sheet = []
    for field in sorted(list(field_dibutuhkan)):
      c_val = coord_dict.get(field, {})
      x = c_val.get("x", 0)
      y = c_val.get("y", 0)
      if x == 0 and y == 0:
        kosong_di_sheet.append(field)
        ada_yang_kosong = True

    if kosong_di_sheet:
      laporan += f"📁 Sheet: {sheet}\n"
      for f_kosong in kosong_di_sheet:
        laporan += f"    • ❌ {f_kosong}\n"
      laporan += "\n"

  if not ada_yang_kosong:
    messagebox.showinfo(
        "Status Koordinat",
        "✅ Sempurna! Seluruh koordinat field dan tombol pada semua sheet sudah"
        " diatur (> 0).",
    )
  else:
    win_lap = tk.Toplevel(root)
    win_lap.title("Daftar Koordinat Kosong")
    win_lap.geometry("380x320")
    win_lap.config(bg="#1e1e2f")
    win_lap.attributes("-topmost", True)

    tk.Label(
        win_lap,
        text="⚠️ Field Berikut Masih Koordinat (0,0):",
        font=("Segoe UI", 9, "bold"),
        bg="#1e1e2f",
        fg="#e67e22",
    ).pack(pady=10)

    txt_area = scrolledtext.ScrolledText(
        win_lap,
        width=42,
        height=12,
        font=("Segoe UI", 9),
        bg="#2a2a3e",
        fg="white",
        insertbackground="white",
    )
    txt_area.pack(pady=5, padx=15)
    txt_area.insert(tk.END, laporan)
    txt_area.config(state="disabled")

    tk.Button(
        win_lap,
        text="Tutup",
        bg="#e74c3c",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=win_lap.destroy,
    ).pack(pady=8)


def buka_menu_pengaturan_koordinat():
  win_set = tk.Toplevel(root)
  win_set.title("Pengaturan Koordinat")
  win_set.geometry("320x330")
  win_set.config(bg="#1e1e2f")
  win_set.resizable(False, False)
  win_set.attributes("-topmost", True)

  tk.Label(
      win_set,
      text="⚙️ Pengaturan Koordinat",
      font=("Segoe UI", 10, "bold"),
      bg="#1e1e2f",
      fg="#3498db",
  ).pack(pady=10)

  f_form = tk.Frame(win_set, bg="#1e1e2f")
  f_form.pack(pady=5, padx=20, fill="x")

  tk.Label(
      f_form,
      text="Pilih Sheet:",
      font=("Segoe UI", 8, "bold"),
      bg="#1e1e2f",
      fg="white",
  ).pack(anchor="w")
  daftar_sheets = ambil_daftar_sheet_excel()
  combo_kib = ttk.Combobox(
      f_form, values=daftar_sheets, state="readonly", font=("Segoe UI", 9)
  )
  if daftar_sheets:
    combo_kib.set(daftar_sheets[0])
  combo_kib.pack(fill="x", pady=(2, 6))

  tk.Label(
      f_form,
      text="Nama Field / Tombol (Ketik/Pilih):",
      font=("Segoe UI", 8, "bold"),
      bg="#1e1e2f",
      fg="white",
  ).pack(anchor="w")
  combo_field = ttk.Combobox(f_form, font=("Segoe UI", 9))
  combo_field.pack(fill="x", pady=(2, 6))

  excel_path = os.path.join(base_dir, "data_kib.xlsx")

  f_xy = tk.Frame(f_form, bg="#1e1e2f")
  f_xy.pack(fill="x", pady=2)
  f_x = tk.Frame(f_xy, bg="#1e1e2f")
  f_x.pack(side="left", expand=True, fill="x", padx=(0, 5))
  tk.Label(
      f_x, text="X:", font=("Segoe UI", 8, "bold"), bg="#1e1e2f", fg="white"
  ).pack(anchor="w")
  entry_x = tk.Entry(f_x, font=("Segoe UI", 9))
  entry_x.pack(fill="x")

  f_y = tk.Frame(f_xy, bg="#1e1e2f")
  f_y.pack(side="left", expand=True, fill="x", padx=(5, 0))
  tk.Label(
      f_y, text="Y:", font=("Segoe UI", 8, "bold"), bg="#1e1e2f", fg="white"
  ).pack(anchor="w")
  entry_y = tk.Entry(f_y, font=("Segoe UI", 9))
  entry_y.pack(fill="x")

  def ambil_posisi_mouse_spasi():
    win_set.withdraw()

    def listener():
      while True:
        try:
          if keyboard.is_pressed("space"):
            x, y = pyautogui.position()
            win_set.deiconify()
            win_set.attributes("-topmost", True)
            entry_x.delete(0, tk.END)
            entry_x.insert(0, str(x))
            entry_y.delete(0, tk.END)
            entry_y.insert(0, str(y))
            break
        except:
          pass
        time.sleep(0.05)

    threading.Thread(target=listener, daemon=True).start()

  tk.Button(
      win_set,
      text="🎯 Arahkan Mouse & Tekan SPASI",
      font=("Segoe UI", 8, "bold"),
      bg="#2980b9",
      fg="white",
      command=ambil_posisi_mouse_spasi,
  ).pack(pady=4, padx=20, fill="x")

  def simpan_koordinat_baru():
    sheet_pilihan = combo_kib.get().strip()
    nama_field = combo_field.get().strip()
    if not nama_field:
      messagebox.showerror("Error", "Nama field tidak boleh kosong!")
      return
    try:
      x_int, y_int = int(entry_x.get().strip()), int(entry_y.get().strip())
    except:
      messagebox.showerror("Error", "Koordinat harus angka!")
      return

    config = muat_config()
    if sheet_pilihan not in config["kordinat"]:
      config["kordinat"][sheet_pilihan] = {}
    config["kordinat"][sheet_pilihan][nama_field] = {"x": x_int, "y": y_int}

    if sheet_pilihan not in config["setting"]:
      config["setting"][sheet_pilihan] = {}
    config["setting"][sheet_pilihan][nama_field] = "Ya"
    simpan_config(config)

    inisialisasi_header_excel_dari_koordinat(excel_path, sheet_pilihan, nama_field)
    messagebox.showinfo(
        "Berhasil",
        f"Koordinat '{nama_field}' & setting 'Ya' tersimpan, Header Excel"
        " dibuat!",
    )
    muat_pilihan_field()

  tk.Button(
      win_set,
      text="💾 Simpan / Update ke Config",
      font=("Segoe UI", 9, "bold"),
      bg="#27ae60",
      fg="white",
      command=simpan_koordinat_baru,
  ).pack(pady=4, padx=20, fill="x")

  def muat_pilihan_field(event=None):
    sheet_pilihan = combo_kib.get().strip()
    config = muat_config()
    field_set = set(config.get("kordinat", {}).get(sheet_pilihan, {}).keys())
    for t in config.get("tahap", {}).get(sheet_pilihan, []):
      if t.get("fname"):
        field_set.add(t.get("fname"))
    if os.path.exists(excel_path):
      try:
        wb_t = openpyxl.load_workbook(excel_path, data_only=True)
        if sheet_pilihan in wb_t.sheetnames:
          ws_t = wb_t[sheet_pilihan]
          for col in range(3, ws_t.max_column + 1):
            h = str(ws_t.cell(row=1, column=col).value or "").strip()
            if h:
              field_set.add(h)
        wb_t.close()
      except:
        pass
    field_set.add("Btn_Simpan")
    combo_field["values"] = sorted(list(field_set))

  def muat_koordinat_terpilih(event=None):
    sheet_pilihan = combo_kib.get().strip()
    field_pilihan = combo_field.get().strip()
    if not field_pilihan:
      return
    config = muat_config()
    coord = (
        config.get("kordinat", {}).get(sheet_pilihan, {}).get(field_pilihan, {})
    )
    entry_x.delete(0, tk.END)
    entry_x.insert(0, str(coord.get("x", "0")))
    entry_y.delete(0, tk.END)
    entry_y.insert(0, str(coord.get("y", "0")))

  combo_kib.bind("<<ComboboxSelected>>", muat_pilihan_field)
  combo_field.bind("<<ComboboxSelected>>", muat_koordinat_terpilih)
  muat_pilihan_field()

  tk.Button(
      win_set,
      text="🔍 Cek Koordinat Belum Diatur",
      font=("Segoe UI", 9, "bold"),
      bg="#e67e22",
      fg="white",
      command=cek_koordinat_belum_diatur,
  ).pack(pady=4, padx=20, fill="x")


def buka_menu_setting_dan_tahap():
  win_st = tk.Toplevel(root)
  win_st.title("Kelola Setting & Tahap JSON")
  win_st.geometry("320x510")
  win_st.config(bg="#1e1e2f")
  win_st.resizable(False, False)
  win_st.attributes("-topmost", True)

  notebook = ttk.Notebook(win_st)
  notebook.pack(pady=5, padx=15, fill="both", expand=True)

  tab_setting = tk.Frame(notebook, bg="#2a2a3e")
  notebook.add(tab_setting, text=" ⚙️ Kelola Setting ")
  tk.Label(
      tab_setting,
      text="Pilih Sheet:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10, pady=(6, 2))
  combo_set_sh = ttk.Combobox(
      tab_setting,
      values=ambil_daftar_sheet_excel(),
      state="readonly",
      font=("Segoe UI", 9),
  )
  if ambil_daftar_sheet_excel():
    combo_set_sh.set(ambil_daftar_sheet_excel()[0])
  combo_set_sh.pack(fill="x", padx=10, pady=(0, 4))

  tk.Label(
      tab_setting,
      text="Pilih Parameter:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  combo_param = ttk.Combobox(tab_setting, font=("Segoe UI", 9))
  combo_param.pack(fill="x", padx=10, pady=(2, 4))
  tk.Label(
      tab_setting,
      text="Nilai / Status:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  entry_nilai = tk.Entry(tab_setting, font=("Segoe UI", 9))
  entry_nilai.pack(fill="x", padx=10, pady=(2, 6))

  def muat_setting(e=None):
    sh = combo_set_sh.get().strip()
    config = muat_config()
    params = list(config.get("setting", {}).get(sh, {}).keys())
    combo_param["values"] = params
    if params:
      combo_param.set(params[0])
      upd_val(None)

  def upd_val(e=None):
    sh, p = combo_set_sh.get().strip(), combo_param.get().strip()
    val = muat_config().get("setting", {}).get(sh, {}).get(p, "")
    entry_nilai.delete(0, tk.END)
    entry_nilai.insert(0, str(val))

  combo_set_sh.bind("<<ComboboxSelected>>", muat_setting)
  combo_param.bind("<<ComboboxSelected>>", upd_val)
  muat_setting()

  def simpan_set():
    sh, p, v = (
        combo_set_sh.get().strip(),
        combo_param.get().strip(),
        entry_nilai.get().strip(),
    )
    if not p:
      return
    config = muat_config()
    if sh not in config["setting"]:
      config["setting"][sh] = {}
    config["setting"][sh][p] = v
    simpan_config(config)
    messagebox.showinfo("Berhasil", "Setting diperbarui!")

  tk.Button(
      tab_setting,
      text="💾 Simpan Setting",
      font=("Segoe UI", 8, "bold"),
      bg="#27ae60",
      fg="white",
      command=simpan_set,
  ).pack(pady=8, padx=10, fill="x")

  tab_tahap = tk.Frame(notebook, bg="#2a2a3e")
  notebook.add(tab_tahap, text=" 📋 Kelola Tahap ")
  tk.Label(
      tab_tahap,
      text="Pilih Sheet:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10, pady=(4, 2))
  combo_th_sh = ttk.Combobox(
      tab_tahap,
      values=ambil_daftar_sheet_excel(),
      state="readonly",
      font=("Segoe UI", 9),
  )
  if ambil_daftar_sheet_excel():
    combo_th_sh.set(ambil_daftar_sheet_excel()[0])
  combo_th_sh.pack(fill="x", padx=10, pady=(0, 4))

  tk.Label(
      tab_tahap,
      text="Pilih Tahap (Acuan Sisip/Hapus/Update):",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  combo_th_ada = ttk.Combobox(tab_tahap, state="readonly", font=("Segoe UI", 9))
  combo_th_ada.pack(fill="x", padx=10, pady=(2, 4))

  tk.Label(
      tab_tahap,
      text="Aksi:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  combo_aksi = ttk.Combobox(
      tab_tahap,
      values=[
          "dropdown",
          "dropdown_klik",
          "field",
          "field_aman",
          "tanggal",
          "scroll",
          "tombol",
          "pencarian",
          "baris_hasil",
      ],
      font=("Segoe UI", 9),
  )
  combo_aksi.set("field")
  combo_aksi.pack(fill="x", padx=10, pady=(2, 4))

  tk.Label(
      tab_tahap,
      text="Field Name:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  entry_fn = tk.Entry(tab_tahap, font=("Segoe UI", 9))
  entry_fn.pack(fill="x", padx=10, pady=(2, 4))

  tk.Label(
      tab_tahap,
      text="Kolom Excel:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  combo_colex = ttk.Combobox(tab_tahap, font=("Segoe UI", 9))
  combo_colex.pack(fill="x", padx=10, pady=(2, 4))

  def on_kolom_excel_selected(event=None):
    selected_col = combo_colex.get().strip()
    if selected_col:
      entry_fn.delete(0, tk.END)
      entry_fn.insert(0, selected_col)
      entry_lbl.delete(0, tk.END)
      entry_lbl.insert(0, selected_col)

  combo_colex.bind("<<ComboboxSelected>>", on_kolom_excel_selected)

  tk.Label(
      tab_tahap,
      text="Label:",
      font=("Segoe UI", 8, "bold"),
      bg="#2a2a3e",
      fg="white",
  ).pack(anchor="w", padx=10)
  entry_lbl = tk.Entry(tab_tahap, font=("Segoe UI", 9))
  entry_lbl.pack(fill="x", padx=10, pady=(2, 6))

  list_tahap = []

  def muat_th(e=None):
    nonlocal list_tahap
    sh = combo_th_sh.get().strip()
    config = muat_config()
    list_tahap = config.get("tahap", {}).get(sh, [])
    combo_th_ada["values"] = [
        f"{i+1}. [{t.get('aksi')}] {t.get('label', t.get('fname'))}"
        for i, t in enumerate(list_tahap)
    ]
    if list_tahap:
      combo_th_ada.set(combo_th_ada["values"][0])
      upd_th_det(None)
    else:
      combo_th_ada.set("")

    cols = list(config.get("kordinat", {}).get(sh, {}).keys())
    combo_colex["values"] = [c for c in cols if c.lower() != "btn_simpan"]

  def upd_th_det(e=None):
    nonlocal list_tahap
    try:
      val_combo = combo_th_ada.get()
      if not val_combo:
        return
      idx = int(val_combo.split(".")[0]) - 1
      t = list_tahap[idx]
      combo_aksi.set(t.get("aksi", "field"))
      entry_fn.delete(0, tk.END)
      entry_fn.insert(0, t.get("fname", ""))
      combo_colex.set(t.get("colex", ""))
      entry_lbl.delete(0, tk.END)
      entry_lbl.insert(0, t.get("label", ""))
    except:
      pass

  combo_th_sh.bind("<<ComboboxSelected>>", muat_th)
  combo_th_ada.bind("<<ComboboxSelected>>", upd_th_det)
  muat_th()

  def simpan_th():
    sh = combo_th_sh.get().strip()
    config = muat_config()
    if sh not in config["tahap"]:
      config["tahap"][sh] = []
    config["tahap"][sh].append({
        "aksi": combo_aksi.get(),
        "fname": entry_fn.get(),
        "colex": combo_colex.get(),
        "label": entry_lbl.get(),
    })
    simpan_config(config)
    entry_fn.delete(0, tk.END)
    entry_lbl.delete(0, tk.END)
    combo_colex.set("")
    messagebox.showinfo("Berhasil", "Tahap ditambahkan di akhir!")
    muat_th()

  def sisip_th():
    sh = combo_th_sh.get().strip()
    try:
      val_combo = combo_th_ada.get()
      if not val_combo:
        simpan_th()
        return
      idx = int(val_combo.split(".")[0]) - 1
    except:
      idx = 0

    config = muat_config()
    if sh not in config["tahap"]:
      config["tahap"][sh] = []

    tahap_baru = {
        "aksi": combo_aksi.get(),
        "fname": entry_fn.get(),
        "colex": combo_colex.get(),
        "label": entry_lbl.get(),
    }

    config["tahap"][sh].insert(idx, tahap_baru)
    simpan_config(config)
    messagebox.showinfo(
        "Berhasil", f"Tahap berhasil disisipkan pada urutan ke-{idx + 1}!"
    )
    muat_th()

  def edit_th():
    sh = combo_th_sh.get().strip()
    try:
      idx = int(combo_th_ada.get().split(".")[0]) - 1
    except:
      return
    config = muat_config()
    if sh in config["tahap"] and 0 <= idx < len(config["tahap"][sh]):
      config["tahap"][sh][idx] = {
          "aksi": combo_aksi.get(),
          "fname": entry_fn.get(),
          "colex": combo_colex.get(),
          "label": entry_lbl.get(),
      }
      simpan_config(config)
      messagebox.showinfo("Berhasil", "Tahap diperbarui!")
      muat_th()

  def hapus_th():
    sh = combo_th_sh.get().strip()
    try:
      val_combo = combo_th_ada.get()
      if not val_combo:
        return
      idx = int(val_combo.split(".")[0]) - 1
    except:
      return

    config = muat_config()
    if sh in config["tahap"] and 0 <= idx < len(config["tahap"][sh]):
      tahap_dihapus = config["tahap"][sh].pop(idx)
      simpan_config(config)
      messagebox.showinfo(
          "Berhasil",
          f"Tahap '{tahap_dihapus.get('label', tahap_dihapus.get('fname'))}'"
          " berhasil dihapus!",
      )
      muat_th()

  f_bth1 = tk.Frame(tab_tahap, bg="#2a2a3e")
  f_bth1.pack(pady=(4, 2), padx=10, fill="x")
  tk.Button(
      f_bth1,
      text="💾 Tambah",
      font=("Segoe UI", 8, "bold"),
      bg="#27ae60",
      fg="white",
      command=simpan_th,
  ).pack(side="left", expand=True, fill="x", padx=(0, 2))
  tk.Button(
      f_bth1,
      text="➕ Sisip",
      font=("Segoe UI", 8, "bold"),
      bg="#e67e22",
      fg="white",
      command=sisip_th,
  ).pack(side="left", expand=True, fill="x", padx=(2, 0))

  f_bth2 = tk.Frame(tab_tahap, bg="#2a2a3e")
  f_bth2.pack(pady=(2, 4), padx=10, fill="x")
  tk.Button(
      f_bth2,
      text="✏️ Update",
      font=("Segoe UI", 8, "bold"),
      bg="#2980b9",
      fg="white",
      command=edit_th,
  ).pack(side="left", expand=True, fill="x", padx=(0, 2))
  tk.Button(
      f_bth2,
      text="🗑️ Hapus",
      font=("Segoe UI", 8, "bold"),
      bg="#c0392b",
      fg="white",
      command=hapus_th,
  ).pack(side="left", expand=True, fill="x", padx=(2, 0))


# Fungsi Toggle Tampilan Panel Kontrol Utama vs Panel Rentang
def tampilkan_panel_rentang():
  if panel_utama_aktif and panel_rentang_aktif:
    panel_utama_aktif.pack_forget()
    panel_rentang_aktif.pack(pady=4, fill="x", padx=8)


def sembunyikan_panel_rentang():
  if panel_utama_aktif and panel_rentang_aktif:
    panel_rentang_aktif.pack_forget()
    panel_utama_aktif.pack(pady=4, fill="x", padx=8)


def jalankan_otomasi_dari_panel(mode):
  global is_stopped, is_stopped_pc
  is_stopped = False
  is_stopped_pc[HWID_LOKAL] = False

  sheet_name = combo_sheet_pilih.get().strip() if combo_sheet_pilih else "Data"

  e_dr, e_sp = rentang_entry_dict.get(sheet_name, (None, None))
  try:
    sr = int(e_dr.get().strip()) if e_dr else 2
    er = int(e_sp.get().strip()) if e_sp else 100
  except:
    sr, er = 2, 100

  sembunyikan_panel_rentang()

  root.iconify()
  threading.Thread(
      target=worker_otomasi_kib,
      args=(sheet_name, mode, sr, er, False, DEFAULT_CHAT_ID),
      daemon=True,
  ).start()


def reset_status_sheet():
  sheet_name = combo_sheet_pilih.get().strip() if combo_sheet_pilih else "Data"
  excel_path = os.path.join(base_dir, "data_kib.xlsx")

  if not os.path.exists(excel_path):
    messagebox.showerror("Error", f"File tidak ditemukan di:\n{excel_path}")
    return

  if messagebox.askyesno(
      "Konfirmasi Reset",
      f"Apakah Anda yakin ingin mereset semua status 'Selesai' di {sheet_name}"
      " menjadi kosong?",
  ):
    try:
      wb = openpyxl.load_workbook(excel_path)
      if sheet_name not in wb.sheetnames:
        messagebox.showerror("Error", f"Sheet '{sheet_name}' tidak ditemukan!")
        return
      ws = wb[sheet_name]

      col_status = None
      for r in range(1, 4):
        for col in range(1, ws.max_column + 1):
          val = str(ws.cell(row=r, column=col).value).strip().lower()
          if val == "status":
            col_status = col
            break
        if col_status:
          break

      if col_status:
        hitung_reset = 0
        for r in range(2, ws.max_row + 1):
          cell_target = ws.cell(row=r, column=col_status)
          if (
              cell_target.value is not None
              and str(cell_target.value).strip() != ""
          ):
            cell_target.value = None
            hitung_reset += 1

        wb.save(excel_path)
        wb.close()
        update_statistik_gui(sheet_name)
        messagebox.showinfo(
            "Berhasil",
            f"Berhasil mereset {hitung_reset} baris status di {sheet_name}"
            " menjadi kosong!",
        )
      else:
        messagebox.showerror("Error", "Kolom 'Status' tidak ditemukan!")
        wb.close()
    except PermissionError:
      messagebox.showerror(
          "Gagal Reset",
          "File Excel sedang dibuka! Tutup file terlebih dahulu.",
      )
    except Exception as e:
      messagebox.showerror("Error", f"Terjadi kesalahan: {e}")


def undo_status_sheet():
  sheet_name = combo_sheet_pilih.get().strip() if combo_sheet_pilih else "Data"
  excel_path = os.path.join(base_dir, "data_kib.xlsx")

  if not os.path.exists(excel_path):
    messagebox.showerror("Error", f"File tidak ditemukan di:\n{excel_path}")
    return

  try:
    wb = openpyxl.load_workbook(excel_path)
    if sheet_name not in wb.sheetnames:
      messagebox.showerror("Error", f"Sheet '{sheet_name}' tidak ditemukan!")
      return
    ws = wb[sheet_name]

    col_status = None
    for r in range(1, 5):
      for col in range(1, ws.max_column + 1):
        val = ws.cell(row=r, column=col).value
        if val and str(val).strip().lower() == "status":
          col_status = col
          break
      if col_status:
        break

    if not col_status:
      messagebox.showerror(
          "Error", "Kolom 'Status' tidak ditemukan di baris 1-4!"
      )
      wb.close()
      return

    target_baris = None
    for r in range(ws.max_row, 1, -1):
      cell_val = ws.cell(row=r, column=col_status).value
      if cell_val is not None and str(cell_val).strip().lower() == "selesai":
        target_baris = r
        break

    if target_baris:
      cell_target = ws.cell(row=target_baris, column=col_status)
      cell_target.value = None
      wb.save(excel_path)
      wb.close()
      update_statistik_gui(sheet_name)
      messagebox.showinfo(
          "Undo Berhasil",
          f"Status 'Selesai' pada baris ke-{target_baris} di {sheet_name}"
          " berhasil dihapus!",
      )
    else:
      messagebox.showwarning(
          "Informasi",
          f"Tidak ada baris dengan status 'Selesai' yang bisa di-undo di"
          f" {sheet_name}.",
      )
      wb.close()
  except PermissionError:
    messagebox.showerror(
        "Gagal Undo",
        f"File '{excel_path}' sedang dibuka di Microsoft Excel!",
    )
  except Exception as e:
    messagebox.showerror("Error", f"Terjadi kesalahan saat undo: {e}")


def stop_sheet():
  sheet_name = combo_sheet_pilih.get().strip() if combo_sheet_pilih else "Data"
  stop_flags[sheet_name] = True
  is_stopped_pc[HWID_LOKAL] = True
  sembunyikan_panel_rentang()


def buka_excel():
  path = os.path.join(base_dir, "data_kib.xlsx")
  if os.path.exists(path):
    os.startfile(path)


def buka_aplikasi_ams():
  ams_path = r"C:\Program Files\AMSFMX\AMSFMX.exe"
  if os.path.exists(ams_path):
    try:
      os.startfile(ams_path)
    except Exception as e:
      messagebox.showerror("Error", f"{e}")
  else:
    try:
      os.startfile("AMS.lnk")
    except:
      try:
        import webbrowser

        webbrowser.open("https://aset.bpkpd.wajokab.my.id/dasbor/")
      except:
        messagebox.showerror(
            "Error", "Aplikasi AMS atau shortcut AMS.lnk tidak ditemukan."
        )


def menu_aktivasi_lisensi():
  win_act = tk.Toplevel(root)
  win_act.title("Aktivasi Full Version")
  win_act.geometry("360x260")
  win_act.config(bg="#1e1e2f")
  win_act.resizable(False, False)
  win_act.attributes("-topmost", True)

  hwid = HWID_LOKAL

  status_teks = (
      "Status: Teraktivasi ✅"
      if cek_status_lisensi()
      else (
          f"Status: Trial Mode (Maks. 5 baris & Sheet '{SHEET_TRIAL_UTAMA}') ⚠️"
      )
  )
  tk.Label(
      win_act,
      text=status_teks,
      font=("Segoe UI", 8, "bold"),
      bg="#1e1e2f",
      fg="#3498db" if cek_status_lisensi() else "#e67e22",
      wraplength=340,
  ).pack(pady=(10, 5))

  tk.Label(
      win_act,
      text="ID Komputer Anda:",
      font=("Segoe UI", 8, "bold"),
      bg="#1e1e2f",
      fg="white",
  ).pack()
  entry_hwid = tk.Entry(
      win_act,
      width=32,
      font=("Segoe UI", 9),
      justify="center",
      bg="#2a2a3e",
      fg="#2ecc71",
      insertbackground="white",
  )
  entry_hwid.pack(pady=2)
  entry_hwid.insert(0, hwid)
  entry_hwid.config(state="readonly")

  tk.Label(
      win_act,
      text="Masukkan Kode Lisensi Master:",
      font=("Segoe UI", 8, "bold"),
      bg="#1e1e2f",
      fg="white",
  ).pack(pady=(6, 2))
  entry_kode = tk.Entry(
      win_act, width=32, font=("Segoe UI", 9), show="*", justify="center"
  )
  entry_kode.pack(pady=2)

  def proses_aktivasi():
    if entry_kode.get().strip() == KODE_FULL_RAHASIA:
      try:
        bound_hash = hashlib.sha256(
            (KODE_FULL_RAHASIA + hwid).encode("utf-8")
        ).hexdigest()
        with open(LICENSE_FILE, "w") as f:
          f.write(bound_hash)
        messagebox.showinfo(
            "Berhasil",
            "Aktivasi Berhasil! Silakan klik tombol Refresh Aplikasi.",
        )
        win_act.destroy()
      except Exception as e:
        messagebox.showerror("Error", f"Gagal menyimpan: {e}")
    else:
      messagebox.showerror("Gagal", "Kode lisensi salah!")

  tk.Button(
      win_act,
      text="🔑 Aktifkan",
      font=("Segoe UI", 8, "bold"),
      bg="#27ae60",
      fg="white",
      width=30,
      command=proses_aktivasi,
  ).pack(pady=8)


if __name__ == "__main__":
  try:
    catat_komputer_aktif()
    simpan_resolusi_layar_otomatis()

    root = tk.Tk()
    root.title("Automation Inputan - Center Control")
    root.geometry("530x550")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.config(bg="#1e1e2f")

    def toggle_hide_show():
      try:
        if root.winfo_viewable():
          root.withdraw()
        else:
          root.deiconify()
          root.attributes("-topmost", True)
          root.attributes("-topmost", False)
      except:
        pass

    try:
      keyboard.add_hotkey("end", toggle_hide_show)
    except:
      pass

    try:
      logo_img = Image.open(os.path.join(base_dir, "logo.png"))
      icon_photo = ImageTk.PhotoImage(logo_img)
      root.iconphoto(True, icon_photo)
    except:
      pass

    root.protocol("WM_DELETE_WINDOW", keluar_aplikasi)

    tk.Label(
        root,
        text="AUTOMATION INPUTAN AMS",
        font=("Segoe UI", 14, "bold"),
        bg="#1e1e2f",
        fg="#ffffff",
    ).pack(pady=(8, 0))
    tk.Label(
        root,
        text="Pusat Kontrol (Tekan END untuk Sembunyikan/Tampilkan)",
        font=("Segoe UI", 8),
        bg="#1e1e2f",
        fg="#a0a0a5",
    ).pack(pady=2)

    f_delay = tk.LabelFrame(
        root,
        text=" ⚙️ Pengaturan Delay Cepat/Lambat ",
        font=("Segoe UI", 8, "bold"),
        bg="#1e1e2f",
        fg="#3498db",
        bd=2,
        relief="solid",
    )
    f_delay.pack(pady=4, padx=20, fill="x", ipady=2)

    config_awal = muat_config()
    first_sheet = (
        ambil_daftar_sheet_excel()[0] if ambil_daftar_sheet_excel() else ""
    )
    set_vals = config_awal.get("setting", {}).get(first_sheet, {})

    f_d_in = tk.Frame(f_delay, bg="#1e1e2f")
    f_d_in.pack(fill="x", padx=5, pady=2)

    def buat_spinbox_delay(parent, label_text, key_name, default_val):
      tk.Label(
          parent, text=label_text, font=("Segoe UI", 8), bg="#1e1e2f", fg="white"
      ).pack(side="left", padx=(2, 0))
      sp = tk.Spinbox(
          parent, from_=0.1, to=10.0, increment=0.1, width=4, font=("Segoe UI", 8)
      )
      val_exist = set_vals.get(key_name, default_val)
      sp.delete(0, tk.END)
      sp.insert(0, val_exist)
      sp.pack(side="left", padx=(0, 6))
      return sp

    sp_klik = buat_spinbox_delay(f_d_in, "D.Klik:", "DELAY_KLIK", "0.30")
    sp_drop = buat_spinbox_delay(f_d_in, "D.DropDwon:", "DELAY_DROPDOWN", "0.40")
    sp_scroll = buat_spinbox_delay(f_d_in, "D.Scroll:", "DELAY_SCROLL", "0.8")
    sp_simpan = buat_spinbox_delay(f_d_in, "D.Btn_Simpan:", "DELAY_KOSONGKAN", "1.5")

    # --- Frame Utama Pemilihan Sheet menggunakan Dropdown (Combobox) ---
    f_pilih_sheet = tk.Frame(root, bg="#1e1e2f")
    f_pilih_sheet.pack(pady=4, padx=20, fill="x")

    tk.Label(
        f_pilih_sheet,
        text="📁 Pilih Sheet Excel:",
        font=("Segoe UI", 9, "bold"),
        bg="#1e1e2f",
        fg="#3498db",
    ).pack(anchor="w", pady=(2, 2))

    daftar_sheets_excel = ambil_daftar_sheet_excel()
    combo_sheet_pilih = ttk.Combobox(
        f_pilih_sheet,
        values=daftar_sheets_excel,
        state="readonly",
        font=("Segoe UI", 10),
    )
    if daftar_sheets_excel:
      combo_sheet_pilih.set(daftar_sheets_excel[0])
    combo_sheet_pilih.pack(fill="x", pady=2)

    # --- Panel Kontrol Utama & Rentang (Sistem Tunggal Interaktif) ---
    f_box_container = tk.Frame(root, bg="#2a2a3e", bd=2, relief="solid")
    f_box_container.pack(pady=4, fill="x", padx=20, ipady=4)

    lbl_judul_panel = tk.Label(
        f_box_container,
        text="📌 Panel Kontrol Utama",
        font=("Segoe UI", 9, "bold"),
        bg="#2a2a3e",
        fg="#e67e22",
    )
    lbl_judul_panel.pack(anchor="w", padx=10, pady=(4, 0))

    lbl_stat_aktif = tk.Label(
        f_box_container,
        text="Total: 0 | Selesai: 0 | Sisa: 0 | ETA: 0 dtk",
        font=("Segoe UI", 8),
        bg="#2a2a3e",
        fg="#3498db",
    )
    lbl_stat_aktif.pack(anchor="w", padx=10, pady=(2, 2))

    canvas_prog_aktif = tk.Canvas(
        f_box_container, bg="#1e1e2f", height=10, highlightthickness=0
    )
    canvas_prog_aktif.pack(fill="x", padx=10, pady=(2, 6))
    rect_prog_aktif = canvas_prog_aktif.create_rectangle(
        0, 0, 0, 10, fill="#27ae60", width=0
    )

    # 1. Panel Tombol Utama (Jalankan, Undo, Rst, Stop)
    panel_utama_aktif = tk.Frame(f_box_container, bg="#2a2a3e")
    panel_utama_aktif.pack(pady=4, fill="x", padx=8)

    btn_jalankan_aktif = tk.Button(
        panel_utama_aktif,
        text="🚀 Jalankan",
        bg="#27ae60",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=tampilkan_panel_rentang,
    )
    btn_jalankan_aktif.pack(side="left", expand=True, fill="x", padx=1)

    tk.Button(
        panel_utama_aktif,
        text="↩️ Undo",
        bg="#8e44ad",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=undo_status_sheet,
    ).pack(side="left", expand=True, fill="x", padx=1)
    tk.Button(
        panel_utama_aktif,
        text="🔄 Rst",
        bg="#f39c12",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=reset_status_sheet,
    ).pack(side="left", expand=True, fill="x", padx=1)
    tk.Button(
        panel_utama_aktif,
        text="⏹️ Stop",
        bg="#e74c3c",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=stop_sheet,
    ).pack(side="left", expand=True, fill="x", padx=1)

    # 2. Panel Rentang Baris & Tombol Mode (Disembunyikan Awalnya)
    panel_rentang_aktif = tk.Frame(f_box_container, bg="#2a2a3e")

    f_input_baris = tk.Frame(panel_rentang_aktif, bg="#2a2a3e")
    f_input_baris.pack(fill="x", padx=2, pady=(2, 4))

    # Input Dari
    f_dari = tk.Frame(f_input_baris, bg="#2a2a3e")
    f_dari.pack(side="left", expand=True, fill="x", padx=(0, 4))
    tk.Label(
        f_dari,
        text="Dari Baris:",
        font=("Segoe UI", 8, "bold"),
        bg="#2a2a3e",
        fg="white",
    ).pack(side="left", padx=(0, 4))
    e_dr = tk.Entry(f_dari, width=6, font=("Segoe UI", 9))
    e_dr.insert(0, "2")
    e_dr.pack(side="left", fill="x", expand=True)

    # Input Sampai
    f_sampai = tk.Frame(f_input_baris, bg="#2a2a3e")
    f_sampai.pack(side="left", expand=True, fill="x", padx=(4, 0))
    tk.Label(
        f_sampai,
        text="Sampai Baris:",
        font=("Segoe UI", 8, "bold"),
        bg="#2a2a3e",
        fg="white",
    ).pack(side="left", padx=(0, 4))

    e_sp = tk.Entry(f_sampai, width=6, font=("Segoe UI", 9))
    e_sp.pack(side="left", fill="x", expand=True)

    # Dictionary untuk menampung entry dari & sampai per sheet secara dinamis
    for sh in daftar_sheets_excel:
      max_r_def = 100
      excel_path = os.path.join(base_dir, "data_kib.xlsx")
      if os.path.exists(excel_path):
        try:
          wb_tmp = openpyxl.load_workbook(excel_path, data_only=True)
          if sh in wb_tmp.sheetnames:
            max_r_def = wb_tmp[sh].max_row
          wb_tmp.close()
        except:
          pass
      e_d_temp = tk.Entry(f_dari, width=6, font=("Segoe UI", 9))
      e_d_temp.insert(0, "2")
      e_s_temp = tk.Entry(f_sampai, width=6, font=("Segoe UI", 9))
      e_s_temp.insert(0, str(max_r_def))
      rentang_entry_dict[sh] = (e_d_temp, e_s_temp)

    # Tombol Mode ONLINE & SIMULASI
    f_btn_mode = tk.Frame(panel_rentang_aktif, bg="#2a2a3e")
    f_btn_mode.pack(fill="x", padx=2, pady=2)

    tk.Button(
        f_btn_mode,
        text="🚀 PROSES",
        bg="#2980b9",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=lambda: jalankan_otomasi_dari_panel("Online"),
    ).pack(side="left", expand=True, fill="x", padx=(0, 2))

    tk.Button(
        f_btn_mode,
        text="🧪 SIMULASI",
        bg="#e67e22",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=lambda: jalankan_otomasi_dari_panel("Simulasi"),
    ).pack(side="left", expand=True, fill="x", padx=(2, 0))


    def ganti_sheet_aktif(event=None):
      sh_terpilih = combo_sheet_pilih.get().strip()
      lbl_judul_panel.config(text=f"📌 Panel Kontrol Utama - {sh_terpilih}")
      if btn_jalankan_aktif:
        btn_jalankan_aktif.config(text=f"🚀 Jalankan {sh_terpilih}")

      # Sembunyikan panel rentang kembali ke menu utama tombol
      sembunyikan_panel_rentang()

      # Atur nilai entry rentang sesuai sheet yang dipilih
      if sh_terpilih in rentang_entry_dict:
        ed, es = rentang_entry_dict[sh_terpilih]
        e_dr.delete(0, tk.END)
        e_dr.insert(0, ed.get())
        e_sp.delete(0, tk.END)
        e_sp.insert(0, es.get())

      update_statistik_gui(sh_terpilih)


    combo_sheet_pilih.bind("<<ComboboxSelected>>", ganti_sheet_aktif)
    if daftar_sheets_excel:
      ganti_sheet_aktif(None)

    # --- Grid Menu Tombol Bawah ---
    f_grid_tombol = tk.Frame(root, bg="#1e1e2f")
    f_grid_tombol.pack(pady=4, padx=20, fill="x")

    f_grid_tombol.grid_columnconfigure(0, weight=1)
    f_grid_tombol.grid_columnconfigure(1, weight=1)

    tk.Button(
        f_grid_tombol,
        text="⚙️ Koordinat",
        bg="#16a085",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=buka_menu_pengaturan_koordinat,
    ).grid(row=0, column=0, sticky="nsew", padx=(0, 2), pady=2)
    tk.Button(
        f_grid_tombol,
        text="📋 Setting & Tahap",
        bg="#d35400",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=buka_menu_setting_dan_tahap,
    ).grid(row=0, column=1, sticky="nsew", padx=(2, 0), pady=2)

    tk.Button(
        f_grid_tombol,
        text="📁 Buka AMS",
        bg="#16a085",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=buka_aplikasi_ams,
    ).grid(row=1, column=0, sticky="nsew", padx=(0, 2), pady=2)
    tk.Button(
        f_grid_tombol,
        text="📊 Buka Excel",
        bg="#d35400",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=buka_excel,
    ).grid(row=1, column=1, sticky="nsew", padx=(2, 0), pady=2)

    tk.Button(
        f_grid_tombol,
        text="📤 Ekspor Config",
        bg="#16a085",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=ekspor_konfigurasi,
    ).grid(row=2, column=0, sticky="nsew", padx=(0, 2), pady=2)
    tk.Button(
        f_grid_tombol,
        text="📥 Impor Config",
        bg="#d35400",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=impor_konfigurasi,
    ).grid(row=2, column=1, sticky="nsew", padx=(2, 0), pady=2)

    tk.Button(
        f_grid_tombol,
        text="🔄 Refresh Aplikasi",
        bg="#2980b9",
        fg="white",
        font=("Segoe UI", 8, "bold"),
        command=refresh_aplikasi,
    ).grid(row=3, column=0, columnspan=2, sticky="nsew", pady=2)

    if not cek_status_lisensi():
      tk.Button(
          f_grid_tombol,
          text="🔑 Aktivasi",
          bg="#8e44ad",
          fg="white",
          font=("Segoe UI", 8, "bold"),
          command=menu_aktivasi_lisensi,
      ).grid(row=4, column=0, sticky="nsew", padx=(0, 2), pady=2)
      tk.Button(
          f_grid_tombol,
          text="❌ Keluar",
          bg="#7f8c8d",
          fg="white",
          font=("Segoe UI", 8, "bold"),
          command=keluar_aplikasi,
      ).grid(row=4, column=1, sticky="nsew", padx=(2, 0), pady=2)
    else:
      tk.Button(
          f_grid_tombol,
          text="❌ Keluar Aplikasi",
          bg="#7f8c8d",
          fg="white",
          font=("Segoe UI", 8, "bold"),
          command=keluar_aplikasi,
      ).grid(row=4, column=0, columnspan=2, sticky="nsew", pady=2)

    tk.Label(
        root,
        text="Zainal Abidin S.Kep Ns",
        font=("Segoe UI", 8),
        bg="#1e1e2f",
        fg="#60606f",
    ).pack(pady=1)

    root.mainloop()
  except Exception as e:
    messagebox.showerror("Fatal Error", f"{e}")
    messagebox.showerror("Fatal Error", f"{e}")

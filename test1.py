import tkinter as tk
import subprocess
from datetime import datetime
import signal
import time
from pathlib import Path
from rtsp_config import rtsp_url
from openpyxl import Workbook, load_workbook

recording = False
ffmpeg_process = None
current_video_path = None

pending_barcode = None
current_barcode = None

BASE_DIR = Path(__file__).resolve().parent
VIDEO_ROOT = BASE_DIR / "video"
EXCEL_ROOT = BASE_DIR / "excel"


def build_daily_dir(root_dir, dt):
    return root_dir / dt.strftime("%Y") / dt.strftime("%m") / dt.strftime("%d")


def get_excel_path(dt):
    daily_dir = build_daily_dir(EXCEL_ROOT, dt)
    daily_dir.mkdir(parents=True, exist_ok=True)
    return daily_dir / f"{dt.strftime('%Y%m%d')}.xlsx"


def ensure_excel_file(excel_path):
    if excel_path.exists():
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "log"
    ws.append(["Thoi gian dong", "Ma vach", "Duong dan video"])
    wb.save(excel_path)


def log_closed_code(barcode, video_path):
    if not barcode or not video_path:
        return

    now = datetime.now()
    excel_path = get_excel_path(now)
    ensure_excel_file(excel_path)

    wb = load_workbook(excel_path)
    ws = wb.active
    ws.append([now.strftime("%Y-%m-%d %H:%M:%S"), barcode, str(video_path)])
    wb.save(excel_path)


def new_filename():
    global current_barcode

    now = datetime.now()
    ts = now.strftime("%Y%m%d_%H%M%S")
    video_dir = build_daily_dir(VIDEO_ROOT, now)
    video_dir.mkdir(parents=True, exist_ok=True)

    if current_barcode:
        name = f"{current_barcode}_{ts}.mp4"
        current_barcode = None
        return video_dir / name

    return video_dir / f"record_{ts}.mp4"

def start_ffmpeg():
    global ffmpeg_process, recording, current_video_path

    if recording:
        return

    filename = new_filename()

    cmd = [
        "ffmpeg", "-y",
        "-rtsp_transport", "tcp",
        "-i", rtsp_url,
        "-c", "copy",
        "-movflags", "+faststart",
        str(filename)
    ]

    ffmpeg_process = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    recording = True
    current_video_path = filename

    barcode_entry.config(state="normal")   # ✅ CHO BẮN
    barcode_entry.focus_set()

    status.config(text=f"🔴 Đang ghi: {filename.name}")

def start_record():
    global recording

    if recording:
        status.config(text="⚠️ Đang ghi rồi")
        return

    recording = True
    start_ffmpeg()

def cut_record():
    global ffmpeg_process, recording, current_video_path
    global pending_barcode, current_barcode

    if not recording or not ffmpeg_process:
        return

    closed_barcode = pending_barcode
    closed_video_path = current_video_path

    current_barcode = pending_barcode
    pending_barcode = None

    try:
        ffmpeg_process.stdin.write(b"q")
        ffmpeg_process.stdin.flush()
    except:
        pass

    ffmpeg_process.wait()

    ffmpeg_process = None
    recording = False
    current_video_path = None

    log_closed_code(closed_barcode, closed_video_path)

    start_ffmpeg()   # ghi tiếp



def stop_record():
    global ffmpeg_process, recording, current_video_path

    if not recording:
        return

    try:
        ffmpeg_process.stdin.write(b"q")
        ffmpeg_process.stdin.flush()
    except:
        pass

    ffmpeg_process.wait()

    ffmpeg_process = None
    recording = False
    current_video_path = None

    barcode_entry.config(state="disabled")  # ❌ KHÔNG CHO BẮN

    status.config(text="⏹️ Đã dừng ghi")

def on_barcode_enter(event):
    global pending_barcode

    if not recording:
        status.config(text="⚠️ Chưa bấm Bắt đầu ghi")
        barcode_entry.delete(0, tk.END)
        return

    code = barcode_entry.get().strip()
    if not code:
        return

    pending_barcode = code
    barcode_entry.delete(0, tk.END)

    status.config(text=f"📦 Đã bắn: {pending_barcode}")

    cut_record()


# ===== GUI =====
root = tk.Tk()
root.title("RTSP Recorder")

# ▶ Start
btn_start = tk.Button(root, text="▶ Bắt đầu ghi", command=start_ffmpeg)
btn_start.pack(pady=5)

# 📦 Input bắn mã vạch
barcode_entry = tk.Entry(root, font=("Arial", 14), state="disabled")
barcode_entry.pack(pady=8)
barcode_entry.bind("<Return>", on_barcode_enter)

# ⏹ Stop
btn_stop = tk.Button(root, text="⏹ Dừng hẳn", command=stop_record)
btn_stop.pack(pady=5)

status = tk.Label(root, text="⏸️ Chưa ghi")
status.pack(pady=10)

root.mainloop()

import tkinter as tk
import subprocess
from datetime import datetime
import signal
import time

rtsp_url = "rtsp://admin:GJTMIL@192.168.1.91:554/ch1/main"

recording = False
ffmpeg_process = None

pending_barcode = None
current_barcode = None

def new_filename():
    global current_barcode

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    if current_barcode:
        name = f"{current_barcode}_{ts}.mp4"
        current_barcode = None
        return name

    return f"record_{ts}.mp4"

def start_ffmpeg():
    global ffmpeg_process, recording

    if recording:
        return

    filename = new_filename()

    cmd = [
        "ffmpeg", "-y",
        "-rtsp_transport", "tcp",
        "-i", rtsp_url,
        "-c", "copy",
        "-movflags", "+faststart",
        filename
    ]

    ffmpeg_process = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )

    recording = True

    barcode_entry.config(state="normal")   # ✅ CHO BẮN
    barcode_entry.focus_set()

    status.config(text=f"🔴 Đang ghi: {filename}")

def start_record():
    global recording

    if recording:
        status.config(text="⚠️ Đang ghi rồi")
        return

    recording = True
    start_ffmpeg()

def cut_record():
    global ffmpeg_process, recording
    global pending_barcode, current_barcode

    if not recording or not ffmpeg_process:
        return

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

    start_ffmpeg()   # ghi tiếp



def stop_record():
    global ffmpeg_process, recording

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

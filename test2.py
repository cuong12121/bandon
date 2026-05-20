import argparse
import os
import subprocess
import sys
import threading
import queue
import socket
import signal
import time
from datetime import datetime

# Default RTSP URL (if you have a fixed RTSP, set it here)
# Example taken from your `test1.py` attachment — change as needed.
DEFAULT_RTSP = "rtsp://admin:GJTMIL@192.168.1.91:554/ch1/main"


def new_filename(prefix=None, out_dir="."):
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    name = f"{prefix}_{ts}.mp4" if prefix else f"record_{ts}.mp4"
    return os.path.join(out_dir, name)


def start_ffmpeg_process(rtsp, out_file, overwrite):
    cmd = [
        "ffmpeg",
        "-y" if overwrite else "-n",
        "-rtsp_transport", "tcp",
        "-i", rtsp,
        "-c", "copy",
        "-movflags", "+faststart",
        out_file,
    ]
    print("Starting ffmpeg:", " ".join(cmd))
    try:
        return subprocess.Popen(cmd, stdin=subprocess.PIPE)
    except FileNotFoundError:
        print("Error: ffmpeg not found in PATH.")
        return None


def stop_ffmpeg_process(p, timeout=5):
    if not p:
        return
    try:
        p.stdin.write(b"q")
        p.stdin.flush()
    except Exception:
        pass
    try:
        p.wait(timeout=timeout)
    except subprocess.TimeoutExpired:
        try:
            p.terminate()
        except Exception:
            pass
        p.wait()


def control_server(port, cmd_queue, host="127.0.0.1"):
    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    s.bind((host, port))
    s.listen(5)
    print(f"Control server listening on {host}:{port}")
    while True:
        try:
            conn, addr = s.accept()
            data = conn.recv(1024).decode(errors="ignore").strip()
            conn.close()
            if not data:
                continue
            # expected: CUT[:barcode]
            if data.upper().startswith("CUT"):
                parts = data.split(":", 1)
                barcode = parts[1].strip() if len(parts) > 1 else None
                cmd_queue.put(("cut", barcode))
                print(f"Received CUT command (barcode={barcode}) from {addr}")
            elif data.upper().startswith("STOP"):
                cmd_queue.put(("stop", None))
                print(f"Received STOP command from {addr}")
        except Exception:
            time.sleep(0.1)


def main():
    parser = argparse.ArgumentParser(description="Record RTSP and accept external cut commands over TCP.")
    parser.add_argument("--rtsp", required=False, default=DEFAULT_RTSP, help=f"RTSP URL (default: {DEFAULT_RTSP})")
    parser.add_argument("--out-dir", default=".", help="Output directory")
    parser.add_argument("--barcode", help="Optional initial barcode prefix for filenames")
    parser.add_argument("--duration", type=int, help="Record for this many seconds then stop (single file)")
    parser.add_argument("--segment-time", type=int, help="Split into segments of N seconds (ffmpeg segment)")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite existing files")
    parser.add_argument("--control-port", type=int, default=9999, help="TCP port for control commands (CUT:barcode)")

    args = parser.parse_args()

    os.makedirs(args.out_dir, exist_ok=True)

    # segment mode: leave original ffmpeg segment behavior (no manual cuts)
    if args.segment_time:
        prefix = args.barcode if args.barcode else "record"
        out_pattern = os.path.join(args.out_dir, f"{prefix}_%Y%m%d_%H%M%S.mp4")
        cmd = [
            "ffmpeg",
            "-y" if args.overwrite else "-n",
            "-rtsp_transport", "tcp",
            "-i", args.rtsp,
            "-c", "copy",
            "-f", "segment",
            "-segment_time", str(args.segment_time),
            "-reset_timestamps", "1",
            "-strftime", "1",
            out_pattern,
        ]
        print("Segment mode: starting ffmpeg (no manual CUT support).")
        p = subprocess.Popen(cmd)
        try:
            p.wait()
        except KeyboardInterrupt:
            try:
                p.stdin.write(b"q")
                p.stdin.flush()
            except Exception:
                pass
        return

    # continuous/manual-cut mode
    cmd_queue = queue.Queue()
    server_thread = threading.Thread(target=control_server, args=(args.control_port, cmd_queue), daemon=True)
    server_thread.start()

    ffmpeg_proc = None
    running = True

    def shutdown(signum, frame):
        nonlocal running, ffmpeg_proc
        print("Shutting down...")
        running = False
        stop_ffmpeg_process(ffmpeg_proc)

    signal.signal(signal.SIGINT, shutdown)
    signal.signal(signal.SIGTERM, shutdown)

    current_barcode = None

    # start initial recording
    filename = new_filename(args.barcode if args.barcode else None, args.out_dir)
    ffmpeg_proc = start_ffmpeg_process(args.rtsp, filename, args.overwrite)
    if not ffmpeg_proc:
        return

    print(f"Recording to {filename}")

    while running:
        try:
            cmd, payload = cmd_queue.get(timeout=0.5)
        except queue.Empty:
            continue

        if cmd == "cut":
            barcode = payload
            print(f"Cut requested (barcode={barcode})")
            stop_ffmpeg_process(ffmpeg_proc)
            # after stopping, start new file with barcode prefix if provided
            new_prefix = barcode if barcode else None
            filename = new_filename(new_prefix, args.out_dir)
            ffmpeg_proc = start_ffmpeg_process(args.rtsp, filename, args.overwrite)
            if not ffmpeg_proc:
                print("Failed to restart ffmpeg after cut.")
                running = False
                break
            print(f"Recording to {filename}")
        elif cmd == "stop":
            print("Stop command received")
            running = False
            stop_ffmpeg_process(ffmpeg_proc)


if __name__ == "__main__":
    main()
import sys

def on_barcode_enter(event):
    # TODO: thay đổi xử lý theo nhu cầu của bạn
    print(f"on_barcode_enter received: {event}")


if __name__ == '__main__':
    arg = sys.argv[1] if len(sys.argv) > 1 else ''
    on_barcode_enter(arg)

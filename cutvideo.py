#!/usr/bin/env python3
"""
cutvideo.py

Cắt một đoạn từ file MP4 bằng 4 tham số dòng lệnh:
    - input file mp4
    - start second (float)
    - end second (float)
    - output file

Sử dụng ffmpeg nếu có, ngược lại dùng moviepy nếu được cài.
"""

import argparse
import os
import shutil
import subprocess
import sys
import codecs

# Ensure UTF-8 output on Windows consoles to avoid UnicodeEncodeError
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    try:
        sys.stdout = codecs.getwriter("utf-8")(sys.stdout.buffer)
        sys.stderr = codecs.getwriter("utf-8")(sys.stderr.buffer)
    except Exception:
        pass


def has_ffmpeg():
    return shutil.which("ffmpeg") is not None


def run_ffmpeg(input_file, start, duration, output_file):
    cmd = [
        "ffmpeg",
        "-y",
        "-ss",
        str(start),
        "-i",
        input_file,
        "-t",
        str(duration),
        "-c",
        "copy",
        output_file,
    ]
    try:
        subprocess.run(cmd, check=True)
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"ffmpeg failed: {e}")


def run_moviepy(input_file, start, end, output_file):
    try:
        from moviepy.editor import VideoFileClip
    except Exception as e:
        raise RuntimeError("moviepy is not available: " + str(e))

    clip = VideoFileClip(input_file).subclip(start, end)
    # write_videofile will choose sensible defaults; use mp4 H.264 + AAC
    clip.write_videofile(output_file, codec="libx264", audio_codec="aac")
    clip.close()


def parse_args():
    p = argparse.ArgumentParser(description="Cắt video mp4: input start end output")
    p.add_argument("input", help="Input mp4 file")
    p.add_argument("start", help="Start time in seconds (float allowed)")
    p.add_argument("end", help="End time in seconds (float allowed)")
    p.add_argument("output", help="Output file path (mp4 recommended)")
    return p.parse_args()


def main():
    args = parse_args()
    input_file = args.input
    output_file = args.output

    if not os.path.isfile(input_file):
        print(f"Input file not found: {input_file}")
        sys.exit(2)

    try:
        start = float(args.start)
        end = float(args.end)
    except ValueError:
        print("Start and end must be numeric (seconds).")
        sys.exit(2)

    if end <= start:
        print("End time must be greater than start time.")
        sys.exit(2)

    duration = end - start

    print(f"Cắt {input_file} từ {start}s đến {end}s → {output_file}")

    if has_ffmpeg():
        try:
            run_ffmpeg(input_file, start, duration, output_file)
            print("Hoàn tất (ffmpeg).")
            return
        except Exception as e:
            print(f"ffmpeg thất bại: {e}")
            print("Thử dùng moviepy làm phương án dự phòng...")

    # fallback to moviepy
    try:
        run_moviepy(input_file, start, end, output_file)
        print("Hoàn tất (moviepy).")
    except Exception as e:
        print("Không thể cắt video: " + str(e))
        print("Cài ffmpeg (khuyến nghị) hoặc moviepy (pip install moviepy)")
        sys.exit(1)


if __name__ == "__main__":
    main()

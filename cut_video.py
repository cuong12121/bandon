import argparse
import os
import subprocess
import sys


def parse_time(t: str) -> float:
    """Parse a time string into seconds.

    Accepts seconds as float or h:m:s / m:s formats.
    """
    if isinstance(t, (int, float)):
        return float(t)
    t = str(t).strip()
    if not t:
        raise ValueError("Empty time string")
    if ":" in t:
        parts = t.split(":")
        parts = [float(p) for p in parts]
        parts = list(reversed(parts))
        secs = 0.0
        mul = 1.0
        for p in parts:
            secs += p * mul
            mul *= 60.0
        return secs
    return float(t)


def format_time_for_ffmpeg(seconds: float) -> str:
    # ffmpeg accepts plain seconds or hh:mm:ss[.ms]
    return str(seconds)


def try_ffmpeg_cut(input_path: str, start: float, duration: float, output_path: str) -> int:
    cmd = [
        "ffmpeg",
        "-y",
        "-ss",
        format_time_for_ffmpeg(start),
        "-i",
        input_path,
        "-t",
        format_time_for_ffmpeg(duration),
        "-c",
        "copy",
        output_path,
    ]
    proc = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
    return proc.returncode


def ffmpeg_cut_reencode(input_path: str, start: float, duration: float, output_path: str) -> int:
    cmd = [
        "ffmpeg",
        "-y",
        "-ss",
        format_time_for_ffmpeg(start),
        "-i",
        input_path,
        "-t",
        format_time_for_ffmpeg(duration),
        "-c:v",
        "libx264",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        output_path,
    ]
    proc = subprocess.run(cmd)
    return proc.returncode


def ffmpeg_cut_with_overlay(input_path: str, start: float, duration: float, output_path: str, text: str, fontsize: int = 24, fontcolor: str = "white", x: int = 10, y: int = 10) -> int:
    # Build drawtext filter. Use a semi-opaque box for readability.
    # Note: escaping may vary by platform; this works in most environments.
    draw = f"drawtext=text='{text}':fontcolor={fontcolor}:fontsize={fontsize}:x={x}:y={y}:box=1:boxcolor=black@0.5"
    cmd = [
        "ffmpeg",
        "-y",
        "-ss",
        format_time_for_ffmpeg(start),
        "-i",
        input_path,
        "-t",
        format_time_for_ffmpeg(duration),
        "-vf",
        draw,
        "-c:v",
        "libx264",
        "-c:a",
        "aac",
        "-b:a",
        "192k",
        output_path,
    ]
    proc = subprocess.run(cmd)
    return proc.returncode


def build_default_output(input_path: str, start: float, end: float) -> str:
    base, ext = os.path.splitext(os.path.basename(input_path))
    s = int(start) if float(start).is_integer() else start
    e = int(end) if float(end).is_integer() else end
    return f"{base}_cut_{s}_{e}{ext}"


def main():
    parser = argparse.ArgumentParser(description="Cut a video between two times using ffmpeg.")
    parser.add_argument("input", help="Input video file path")
    parser.add_argument("start", help="Start time (seconds or HH:MM:SS or MM:SS)")
    parser.add_argument("end", help="End time (seconds or HH:MM:SS or MM:SS)")
    parser.add_argument("-o", "--output", help="Output file path (optional)")
    parser.add_argument("--overlay", action="store_true", help="Overlay default text 'video cut' on the output")
    parser.add_argument("--text", help="Custom overlay text (implies overlay if provided)")
    parser.add_argument("--fontsize", type=int, default=24, help="Overlay font size")
    parser.add_argument("--fontcolor", default="white", help="Overlay font color")
    parser.add_argument("--x", type=int, default=10, help="Overlay X position")
    parser.add_argument("--y", type=int, default=10, help="Overlay Y position")
    args = parser.parse_args()

    try:
        start = parse_time(args.start)
        end = parse_time(args.end)
    except Exception as ex:
        print(f"Error parsing times: {ex}")
        sys.exit(2)

    if end <= start:
        print("End time must be greater than start time")
        sys.exit(2)

    duration = end - start

    input_path = args.input
    if not os.path.isfile(input_path):
        print(f"Input file not found: {input_path}")
        sys.exit(2)

    output_path = args.output or build_default_output(input_path, start, end)

    overlay_text = None
    if args.text:
        overlay_text = args.text
    elif args.overlay:
        overlay_text = "video cut"

    print(f"Cutting {input_path} from {start}s to {end}s (duration {duration}s) -> {output_path}")

    if overlay_text:
        print(f"Applying overlay text: {overlay_text}")
        code = ffmpeg_cut_with_overlay(
            input_path,
            start,
            duration,
            output_path,
            overlay_text,
            fontsize=args.fontsize,
            fontcolor=args.fontcolor,
            x=args.x,
            y=args.y,
        )
        if code == 0:
            print("Finished (re-encoded with overlay)")
            sys.exit(0)
        print("ffmpeg failed to cut/overlay the video")
        sys.exit(1)

    # No overlay requested: try stream copy first, then re-encode fallback
    code = try_ffmpeg_cut(input_path, start, duration, output_path)
    if code == 0:
        print("Finished (stream copy)")
        sys.exit(0)

    print("Stream copy failed or produced non-zero exit. Falling back to re-encode...")
    code2 = ffmpeg_cut_reencode(input_path, start, duration, output_path)
    if code2 == 0:
        print("Finished (re-encoded)")
        sys.exit(0)

    print("ffmpeg failed to cut the video")
    sys.exit(1)


if __name__ == "__main__":
    main()

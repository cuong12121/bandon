import subprocess
import shutil
import sys

cmd = [
    "ffmpeg",
    "-i", "1.mp4",
    "-vf",
    "drawtext=text='VIDEO':x=w-tw-20:y=20:fontsize=40:fontcolor=white",
    "-codec:a", "copy",
    "11_cv.mp4"
]

ffmpeg_path = shutil.which("ffmpeg")
if not ffmpeg_path:
    print("Error: 'ffmpeg' not found in PATH.")
    print("Please install ffmpeg and ensure it's on your system PATH.")
    print("Download: https://ffmpeg.org/download.html")
    sys.exit(1)

cmd[0] = ffmpeg_path

try:
    subprocess.run(cmd, check=True)
except subprocess.CalledProcessError as e:
    print(f"ffmpeg failed with exit code {e.returncode}")
    sys.exit(e.returncode)
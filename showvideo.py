import argparse
import threading
import re
import mimetypes
from pathlib import Path

from flask import Flask, request, Response, send_file
from werkzeug.serving import make_server

try:
    import webview
except Exception:
    webview = None


def send_file_partial(file_path):
    file_path = Path(file_path)
    if not file_path.exists():
        return Response('File not found', status=404)

    file_size = file_path.stat().st_size
    range_header = request.headers.get('Range', None)
    if not range_header:
        return send_file(str(file_path))

    m = re.match(r'bytes=(\d+)-(\d*)', range_header)
    if m:
        start = int(m.group(1))
        end = m.group(2)
        end = int(end) if end else file_size - 1
    else:
        start = 0
        end = file_size - 1

    if start >= file_size:
        return Response(status=416)

    length = end - start + 1

    def generate():
        with open(file_path, 'rb') as f:
            f.seek(start)
            remaining = length
            chunk_size = 64 * 1024
            while remaining > 0:
                read_size = min(chunk_size, remaining)
                chunk = f.read(read_size)
                if not chunk:
                    break
                remaining -= len(chunk)
                yield chunk

    mime_type, _ = mimetypes.guess_type(str(file_path))
    mime_type = mime_type or 'application/octet-stream'

    rv = Response(generate(), 206, mimetype=mime_type, direct_passthrough=True)
    rv.headers.add('Content-Range', f'bytes {start}-{end}/{file_size}')
    rv.headers.add('Accept-Ranges', 'bytes')
    rv.headers.add('Content-Length', str(length))
    return rv


def make_app(video_path: Path):
    app = Flask(__name__)

    @app.route('/')
    def index():
                video_uri = '/video'
                html = """<!doctype html>
<html>
    <head><meta charset="utf-8"><title>Show Video</title></head>
    <body style="margin:0; background:#000; display:flex; align-items:center; justify-content:center; height:100vh;">
        <video id="player" controls autoplay style="max-width:100%; max-height:100%;" src="{video_uri}"></video>
        <script>
            const p = document.getElementById('player');
            p.addEventListener('error', e => console.error('Video error', e));
            p.addEventListener('play', () => { if (p.muted) p.muted = false; });
        </script>
    </body>
</html>""".replace('{video_uri}', video_uri)
                return html

    @app.route('/video')
    def video():
        return send_file_partial(video_path)

    return app


def run(video: str):
    path = Path(video).expanduser().resolve()
    if not path.exists():
        print('Video file not found:', path)
        return

    app = make_app(path)
    server = make_server('127.0.0.1', 0, app)
    port = server.socket.getsockname()[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    url = f'http://127.0.0.1:{port}/'
    print('Serving', path)
    print('Open', url)

    try:
        if webview:
            webview.create_window('Show Video', url=url, width=900, height=600)
            webview.start()
        else:
            import webbrowser
            webbrowser.open(url)
            input('Press Enter to stop...')
    finally:
        server.shutdown()


def main():
    parser = argparse.ArgumentParser(description='Show a single video in a webview with ranged streaming')
    parser.add_argument('video', help='Path to video file')
    args = parser.parse_args()
    run(args.video)


if __name__ == '__main__':
    main()

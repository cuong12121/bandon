import json
from datetime import datetime
from pathlib import Path

import webview
from openpyxl import load_workbook, Workbook
import threading
import urllib.parse
import re
import mimetypes
import os
import json
from datetime import datetime
from pathlib import Path

import webview
from openpyxl import load_workbook, Workbook
import threading
import urllib.parse
import re
import mimetypes
import os
import subprocess
import webbrowser
from flask import Flask, request, Response, send_file, jsonify
from werkzeug.serving import make_server
from define_config import VALID_DEFINES

BASE_DIR = Path(__file__).resolve().parent
EXCEL_ROOT = BASE_DIR / "excel"
PAGE_SIZE = 20
VIDEO_ROOT = BASE_DIR / "video"
rtsp_url = "rtsp://admin:GJTMIL@192.168.1.91:554/ch1/main"
ffmpeg_process = None
recording = False
current_video_path = None


def build_daily_dir(root_dir, dt):
    return root_dir / dt.strftime("%Y") / dt.strftime("%m") / dt.strftime("%d")


def new_filename():
    global current_video_path
    now = datetime.now()
    video_dir = build_daily_dir(VIDEO_ROOT, now)
    video_dir.mkdir(parents=True, exist_ok=True)
    ts = now.strftime("%Y%m%d_%H%M%S")
    filename = video_dir / f"record_{ts}.mp4"
    current_video_path = filename
    return filename


def get_today_excel_path():
    now = datetime.now()
    daily_dir = EXCEL_ROOT / now.strftime("%Y") / now.strftime("%m") / now.strftime("%d")
    return daily_dir / f"{now.strftime('%Y%m%d')}.xlsx"


def resolve_video_path(raw_path):
    video_path = Path(str(raw_path).strip())
    if not video_path.is_absolute():
        video_path = (BASE_DIR / video_path).resolve()
    return video_path


def parse_close_time(value):
    if isinstance(value, datetime):
        return value

    text = str(value).strip()
    if not text:
        return datetime.min

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y/%m/%d %H:%M:%S"):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(text)
    except ValueError:
        return datetime.min


def load_today_orders():
    excel_path = get_today_excel_path()
    if not excel_path.exists():
        return [], excel_path

    workbook = load_workbook(excel_path, data_only=True)
    worksheet = workbook.active

    rows = []
    for row in worksheet.iter_rows(min_row=2, max_col=5, values_only=True):
        if not any(row):
            continue

        close_time = row[0] or ""
        barcode = str(row[1] or "")
        # user is defined as the last character of the barcode if it's in VALID_DEFINES
        user_char = ''
        if barcode:
            last = barcode.strip()[-1:]
            if last in VALID_DEFINES:
                user_char = last

        video_raw = row[3] or ""
        elapsed = row[4] if len(row) > 4 and row[4] is not None else ""
        sort_key = parse_close_time(close_time)

        resolved_video = resolve_video_path(video_raw)
        exists = resolved_video.exists()
        rows.append({
            "close_time": str(close_time),
            "barcode": str(barcode),
            "user": user_char,
            "video_path": str(resolved_video),
            "elapsed": str(elapsed),
            "exists": exists,
            "sort_key": sort_key.isoformat(),
        })

    rows.sort(key=lambda item: item["sort_key"], reverse=True)
    for item in rows:
        item.pop("sort_key", None)

    return rows, excel_path


def append_order(barcode_value, elapsed_seconds=None):
    excel_path = get_today_excel_path()
    excel_path.parent.mkdir(parents=True, exist_ok=True)

    if not excel_path.exists():
        wb = Workbook()
        ws = wb.active
        ws.append(["close_time", "barcode", "user", "video_path", "elapsed_seconds"])
        wb.save(excel_path)

    wb = load_workbook(excel_path)
    ws = wb.active
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    el = '' if elapsed_seconds is None else float(elapsed_seconds)
    b = str(barcode_value or "")
    user_char = ''
    if b:
        last = b.strip()[-1:]
        if last in VALID_DEFINES:
            user_char = last

    ws.append([now, b, user_char, "", el])
    wb.save(excel_path)


def build_html(rows, excel_path, base_url):
    # Use a template with placeholders to avoid f-string brace interpolation issues
    data_json = json.dumps(rows)
    template = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Danh sach don hom nay</title>
  <style>
    :root {--bg: #f4efe4;--surface: #fff8ee;--ink: #1f2328;--muted: #5f676f;--accent: #005f73;--accent-2: #ee9b00;--line: #d9cbb0;}
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", Tahoma, sans-serif; color: var(--ink); background: radial-gradient(circle at 15% 10%, #fff5dd 0%, var(--bg) 45%); min-height: 100vh; }
    .wrap { max-width: 1100px; margin: 24px auto; padding: 0 16px 24px; }
    .head { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 14px 16px; margin-bottom: 14px; }
    .title { margin: 0; font-size: 24px; color: var(--accent); }
    .meta { margin-top: 6px; color: var(--muted); font-size: 13px; word-break: break-all; }
    .grid { display: grid; grid-template-columns: 1fr; gap: 14px; }
    @media (min-width: 980px) { .grid { grid-template-columns: 1.2fr .8fr; } }
    .card { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 12px; box-shadow: 0 8px 20px rgba(31, 35, 40, 0.06); }
    table { width: 100%; border-collapse: collapse; font-size: 14px; }
    th, td { border-bottom: 1px solid var(--line); text-align: left; padding: 10px 8px; vertical-align: top; }
    th { color: var(--accent); font-weight: 700; background: #fff2dc; position: sticky; top: 0; }
    .table-wrap { max-height: 66vh; overflow: auto; border: 1px solid var(--line); border-radius: 10px; background: #fff; }
    .btn { border: 0; border-radius: 8px; padding: 7px 10px; background: var(--accent); color: #fff; cursor: pointer; font-weight: 600; }
    .btn:disabled { background: #8a8f95; cursor: not-allowed; }
    .pager { margin-top: 10px; display: flex; align-items: center; gap: 8px; }
    .pager-info { color: var(--muted); font-size: 13px; margin-left: 4px; }
    .video-box { border: 1px dashed var(--line); border-radius: 10px; padding: 10px; min-height: 240px; background: #fff; }
    .video-title { margin: 0 0 8px; color: var(--accent); font-size: 15px; }
    #videoInfo { font-size: 12px; color: var(--muted); word-break: break-all; margin: 0 0 8px; }
    video { width: 100%; max-height: 420px; background: #000; border-radius: 8px; }
    .empty { color: var(--muted); font-style: italic; }
    .scan-row { display:flex; gap:8px; margin-bottom:8px; }
    .scan-input { flex:1; padding:8px; border-radius:8px; border:1px solid var(--line); font-size:14px; }
  </style>
</head>
<body>
  <div class="wrap">
    <div class="head">
      <h1 class="title">Danh sach don hom nay</h1>
      <div class="meta">Excel: %%EXCEL_PATH%%</div>
      <div class="meta" id="count"></div>
    </div>

    <div class="grid">
      <section class="card">
        <div class="scan-row">
          <input id="barcodeInput" class="scan-input" placeholder="Bắn mã vạch ở đây rồi Enter..." />
          <button class="btn" id="startBtn">Bắt đầu</button>
          <button class="btn" id="stopBtn" disabled> Dừng </button>
          <span id="timerDisplay" style="align-self:center;margin-left:8px;color:var(--muted)">0.0s</span>
          <button class="btn" id="scanBtn">Bắn</button>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
                            <tr>
                                <th>Thoi gian dong</th>
                                <th>Ma vach</th>
                                <th>Người dùng</th>
                                <th>Thời gian đếm</th>
                                <th>Video</th>
                            </tr>
            </thead>
            <tbody id="rows"></tbody>
          </table>
        </div>
        <div class="pager">
          <button class="btn" id="prevBtn" onclick="prevPage()">Trang truoc</button>
          <button class="btn" id="nextBtn" onclick="nextPage()">Trang sau</button>
          <span class="pager-info" id="pageInfo">Trang 1/1</span>
        </div>
      </section>

      <section class="card">
        <div class="video-box">
          <h3 class="video-title">Xem video don</h3>
          <p id="videoInfo">Chon mot don de xem video.</p>
          <video id="player" controls></video>
        </div>
      </section>
    </div>
  </div>

  <script>
    let data = %%DATA_JSON%%;
    const pageSize = %%PAGE_SIZE%%;
    let currentPage = 1;
    const rowsEl = document.getElementById('rows');
    const countEl = document.getElementById('count');
    const player = document.getElementById('player');
    const info = document.getElementById('videoInfo');
    const pageInfoEl = document.getElementById('pageInfo');
    const prevBtn = document.getElementById('prevBtn');
    const nextBtn = document.getElementById('nextBtn');
    const scanInput = document.getElementById('barcodeInput');
    const scanBtn = document.getElementById('scanBtn');
    const startBtn = document.getElementById('startBtn');
    const stopBtn = document.getElementById('stopBtn');
    const timerDisplay = document.getElementById('timerDisplay');

    let timerStart = null;
    let timerInterval = null;

    function formatSeconds(s) {
      if (s === null || s === undefined || isNaN(s)) return '';
      return (Math.round(s * 1000) / 1000) + 's';
    }

    function startTimer() {
      if (timerInterval) clearInterval(timerInterval);
      timerStart = Date.now();
      startBtn.textContent = 'Đang...';
      startBtn.disabled = true;
      stopBtn.disabled = false;
      timerInterval = setInterval(() => {
        const elapsed = (Date.now() - timerStart) / 1000;
        timerDisplay.textContent = formatSeconds(elapsed);
      }, 200);
    }

    function stopTimer() {
      if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
      // freeze display; allow starting a fresh timer
      startBtn.textContent = 'Bắt đầu';
      startBtn.disabled = false;
      stopBtn.disabled = true;
      timerStart = null;
    }

    function playVideo(item) {
      var url = item.video_uri || item.video_path || '';
      if (!url) {
        info.textContent = 'Khong tim thay file video.';
        return;
      }
      if (window.pywebview && window.pywebview.api && window.pywebview.api.open_in_browser) {
        window.pywebview.api.open_in_browser(url);
      } else {
        window.open(url, '_blank');
      }
      info.textContent = item.video_path;
    }

    function renderTable() {
            if (data.length === 0) {
                rowsEl.innerHTML = '<tr><td colspan="5" class="empty">Khong co du lieu don cho ngay hom nay.</td></tr>';
        countEl.textContent = 'Tong don: 0';
        pageInfoEl.textContent = 'Trang 0/0';
        prevBtn.disabled = true;
        nextBtn.disabled = true;
        return;
      }
      const totalPages = Math.ceil(data.length / pageSize);
      if (currentPage > totalPages) currentPage = totalPages;
      const start = (currentPage - 1) * pageSize;
      const end = start + pageSize;
      const pageRows = data.slice(start, end);
      countEl.textContent = 'Tong don: ' + data.length;
      pageInfoEl.textContent = 'Trang ' + currentPage + '/' + totalPages;
      prevBtn.disabled = currentPage <= 1;
      nextBtn.disabled = currentPage >= totalPages;
            rowsEl.innerHTML = pageRows.map((item, index) => {
        const actualIndex = start + index;
        const disabled = item.exists ? '' : 'disabled';
        const btn = '<button class="btn" ' + disabled + ' onclick="playVideo(data[' + actualIndex + '])">Xem</button>';
                return '<tr>' + '<td>' + item.close_time + '</td>' + '<td>' + item.barcode + '</td>' + '<td>' + (item.user || '') + '</td>' + '<td>' + (item.elapsed || '') + '</td>' + '<td>' + btn + '</td>' + '</tr>';
      }).join('');
    }

    function nextPage() { const totalPages = Math.ceil(data.length / pageSize); if (currentPage < totalPages) { currentPage += 1; renderTable(); } }
    function prevPage() { if (currentPage > 1) { currentPage -= 1; renderTable(); } }

    async function refreshData() {
      try { const resp = await fetch('/data'); if (!resp.ok) return; const json = await resp.json(); data = json; currentPage = 1; renderTable(); } catch (e) { console.error('Refresh error', e); }
    }

    async function sendBarcode(barcode) {
      try {
        let elapsed = null;
        if (timerStart) elapsed = (Date.now() - timerStart) / 1000;
        const resp = await fetch('/add', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ barcode, elapsed_seconds: elapsed }) });
        if (resp.ok) {
          scanInput.value = ''; scanInput.focus();
          // do NOT stop or reset the timer here; timer keeps running until user presses Stop
          await refreshData();
        } else { alert('Lỗi khi gửi mã vạch'); }
      } catch (e) { alert('Lỗi kết nối: ' + e.message); }
    }

    scanBtn.addEventListener('click', () => { const v = scanInput.value.trim(); if (!v) return; sendBarcode(v); });
    scanInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { const v = scanInput.value.trim(); if (!v) return; sendBarcode(v); } });
    startBtn.addEventListener('click', async () => {
      startTimer();
      try {
        await fetch('/rec_start', { method: 'POST' });
      } catch (e) {
        console.error('Start recording proxy failed', e);
      }
    });

    stopBtn.addEventListener('click', async () => {
      stopTimer();
      try {
                const resp = await fetch('/rec_stop', { method: 'POST' });
                if (resp.ok) {
                    await refreshData();
                    const j = await resp.json();
                    if (j && j.video) console.info('Ghi video xong: ' + j.video);
                }
      } catch (e) {
        console.error('Stop recording proxy failed', e);
      }
    });
    renderTable();
  </script>
</body>
</html>
"""

    html = template.replace('%%DATA_JSON%%', data_json).replace('%%PAGE_SIZE%%', str(PAGE_SIZE)).replace('%%EXCEL_PATH%%', str(excel_path))
    return html


def main():
    VIDEO_ROOT.mkdir(parents=True, exist_ok=True)

    app = Flask(__name__)

    class Api:
        def open_in_browser(self, url):
            try:
                if not url:
                    return False
                chrome_paths = [
                    os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'Google\\Chrome\\Application\\chrome.exe'),
                    os.path.join(os.environ.get('PROGRAMFILES(X86)', 'C:\\Program Files (x86)'), 'Google\\Chrome\\Application\\chrome.exe'),
                    os.path.join(os.environ.get('LOCALAPPDATA', os.path.expanduser('~\\AppData\\Local')), 'Google\\Chrome\\Application\\chrome.exe'),
                ]
                for ch in chrome_paths:
                    if os.path.exists(ch):
                        try:
                            subprocess.Popen([ch, f'--app={url}', '--window-size=900,600'], shell=False)
                            return True
                        except Exception:
                            try:
                                subprocess.Popen([ch, '--new-window', url, '--window-size=900,600'], shell=False)
                                return True
                            except Exception:
                                continue

                webbrowser.open(url)
                return True
            except Exception:
                try:
                    webbrowser.open(url)
                    return True
                except Exception:
                    return False


    @app.route('/')
    def index():
        # load rows fresh each request so newly appended orders appear after reload
        rows, excel_path = load_today_orders()

        # convert video paths to flask URLs served by /video/<path:filename>
        for item in rows:
            if item.get('exists'):
                try:
                    p = Path(item['video_path'])
                    rel = p.relative_to(VIDEO_ROOT)
                    quoted = "/".join(urllib.parse.quote(part) for part in rel.parts)
                    item['video_uri'] = base_url + f"/video/{quoted}"
                except Exception:
                    try:
                        item['video_uri'] = Path(item['video_path']).as_uri()
                    except Exception:
                        item['video_uri'] = ""
            else:
                item['video_uri'] = ""

        html = build_html(rows, excel_path, base_url)
        return html


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

    def start_record():
        global recording

        if recording:
            return

        start_ffmpeg()

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
    


    @app.route('/add', methods=['POST'])
    def add():
        try:
            data = request.get_json(force=True)
            barcode = data.get('barcode') if isinstance(data, dict) else None
            elapsed = data.get('elapsed_seconds') if isinstance(data, dict) else None
            if not barcode:
                return jsonify({'ok': False, 'error': 'missing barcode'}), 400
            append_order(barcode, elapsed)
            return jsonify({'ok': True})
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)}), 500


    def find_latest_video_file():
        now = datetime.now()
        video_dir = VIDEO_ROOT / now.strftime("%Y") / now.strftime("%m") / now.strftime("%d")
        if not video_dir.exists():
            return None
        files = list(video_dir.glob('**/*.mp4'))
        if not files:
            return None
        files.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        return files[0]


    def attach_latest_video_to_last_row(video_path: Path):
        excel_path = get_today_excel_path()
        if not excel_path.exists():
            return False
        wb = load_workbook(excel_path)
        ws = wb.active
        # find last row (append style)
        last = ws.max_row
        # video_path is column 4 now (1-based): close_time, barcode, user, video_path, elapsed_seconds
        ws.cell(row=last, column=4, value=str(video_path))
        wb.save(excel_path)
        return True


    @app.route('/rec_start', methods=['POST'])
    def rec_start():
        try:
            start_record()
            return jsonify({'ok': True})
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)}), 500


    @app.route('/rec_stop', methods=['POST'])
    def rec_stop():
        try:
            stop_record()
        except Exception as e:
            return jsonify({'ok': False, 'error': str(e)}), 500

        # after stopping, find latest video and attach
        latest = find_latest_video_file()
        if latest:
            attached = attach_latest_video_to_last_row(latest)
            return jsonify({'ok': True, 'video': str(latest), 'attached': attached})
        return jsonify({'ok': True, 'video': None, 'attached': False})


    @app.route('/data')
    def data_api():
        rows, _ = load_today_orders()
        for item in rows:
            if item.get('exists'):
                try:
                    p = Path(item['video_path'])
                    rel = p.relative_to(VIDEO_ROOT)
                    quoted = "/".join(urllib.parse.quote(part) for part in rel.parts)
                    item['video_uri'] = base_url + f"/video/{quoted}"
                except Exception:
                    try:
                        item['video_uri'] = Path(item['video_path']).as_uri()
                    except Exception:
                        item['video_uri'] = ""
            else:
                item['video_uri'] = ""
        return jsonify(rows)


    def send_file_partial(path):
        file_path = Path(path)
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
        with open(file_path, 'rb') as f:
            f.seek(start)
            data = f.read(length)

        mime_type, _ = mimetypes.guess_type(str(file_path))
        mime_type = mime_type or 'application/octet-stream'

        rv = Response(data, 206, mimetype=mime_type, direct_passthrough=True)
        rv.headers.add('Content-Range', f'bytes {start}-{end}/{file_size}')
        rv.headers.add('Accept-Ranges', 'bytes')
        rv.headers.add('Content-Length', str(length))
        return rv


    @app.route('/video/<path:filename>')
    def video(filename):
        safe_path = (VIDEO_ROOT / filename).resolve()
        try:
            safe_path.relative_to(VIDEO_ROOT.resolve())
        except Exception:
            return Response('Forbidden', status=403)
        return send_file_partial(safe_path)

    server = make_server('127.0.0.1', 0, app)
    port = server.socket.getsockname()[1]
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()

    global base_url
    base_url = f"http://127.0.0.1:{port}"

    try:
        api = Api()
        webview.create_window(base_url, url=base_url + '/', width=1200, height=760, js_api=api)

        use_cef = False
        try:
            import cefpython3  # type: ignore
            use_cef = True
        except Exception:
            use_cef = False

        if use_cef:
            webview.start(gui='cef')
        else:
            webview.start()
    finally:
        server.shutdown()


if __name__ == "__main__":
    main()

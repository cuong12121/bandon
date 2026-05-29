import json
from datetime import datetime
from pathlib import Path
import webbrowser
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
EXCEL_ROOT = BASE_DIR / "excel"
PAGE_SIZE = 20
VIDEO_ROOT = BASE_DIR / "video"


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


def get_today_excel_path():
    now = datetime.now()
    daily_dir = EXCEL_ROOT / now.strftime("%Y") / now.strftime("%m") / now.strftime("%d")
    return daily_dir / f"{now.strftime('%Y%m%d')}.xlsx"


def load_today_orders():
    excel_path = get_today_excel_path()
    if not excel_path.exists():
        return [], excel_path

    workbook = load_workbook(excel_path, data_only=True)
    worksheet = workbook.active

    rows = []
    for row in worksheet.iter_rows(min_row=2, max_col=3, values_only=True):
        if not any(row):
            continue

        close_time = row[0] or ""
        barcode = row[1] or ""
        video_raw = row[2] or ""
        sort_key = parse_close_time(close_time)

        resolved_video = resolve_video_path(video_raw)
        exists = resolved_video.exists()
        rows.append({
            "close_time": str(close_time),
            "barcode": str(barcode),
            "video_path": str(resolved_video),
            "exists": exists,
            "sort_key": sort_key.isoformat(),
        })

    rows.sort(key=lambda item: item["sort_key"], reverse=True)
    for item in rows:
        item.pop("sort_key", None)

    # add playback URI and cleaned display link
    for item in rows:
        if item.get('exists'):
            try:
                p = Path(item['video_path'])
                item['video_uri'] = p.as_uri()
                try:
                    rel = p.relative_to(BASE_DIR)
                    item['video_link'] = str(rel).replace('\\', '/')
                except Exception:
                    item['video_link'] = p.as_posix()
            except Exception:
                item['video_uri'] = ""
                item['video_link'] = ""
        else:
            item['video_uri'] = ""
            item['video_link'] = ""

    return rows, excel_path


def build_html(rows, excel_path):
    data_json = json.dumps(rows)
    return f"""<!doctype html>
<html>
<head>
  <meta charset=\"utf-8\" />
  <title>Danh sach don hom nay</title>
  <style>body{{font-family:Segoe UI, Tahoma, sans-serif;}}</style>
</head>
<body>
<h1>Danh sach don hom nay</h1>
<div>Excel: {excel_path}</div>
<table border="1" cellpadding="6">
<thead><tr><th>Thoi gian dong</th><th>Ma vach</th><th>Video</th><th>Link Video</th></tr></thead>
<tbody id="rows"></tbody>
</table>
<video id="player" controls style="width:100%;max-height:420px;margin-top:12px;background:#000"></video>

<script>
const data = {data_json};
const rowsEl = document.getElementById('rows');
const player = document.getElementById('player');
function playVideo(item){
  if(!item.video_uri){ player.removeAttribute('src'); player.load(); return; }
  player.src = item.video_uri; player.load();
}
rowsEl.innerHTML = data.map(item => '<tr>' +
  '<td>' + item.close_time + '</td>' +
  '<td>' + item.barcode + '</td>' +
  '<td>' + (item.exists ? '<button onclick="playVideo('+JSON.stringify(item).replace(/</g,'\\u003c') +')">Xem</button>' : 'N/A') + '</td>' +
  '<td>' + item.video_link + '</td>' +
  '</tr>').join('');
</script>
</body>
</html>"""


def main():
    rows, excel_path = load_today_orders()
    html = build_html(rows, excel_path)
    out = BASE_DIR / 'list_don.html'
    out.write_text(html, encoding='utf-8')
    webbrowser.open(out.as_uri())


if __name__ == '__main__':
    main()

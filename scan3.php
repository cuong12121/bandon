<?php
// scan3.php - simple web UI to control test2.py recorder
// Usage: place on a PHP-enabled web server on the same machine.

$PYTHON = 'python';
$TEST2 = 'c:\\Users\\cuong\\Desktop\\New folder (11)\\test2.py';
$OUT_DIR = 'c:\\Users\\cuong\\Desktop\\New folder (11)\\out';
$CONTROL_PORT = 9999;
$STATE_FILE = __DIR__ . DIRECTORY_SEPARATOR . 'recorder_state.json';
$SCANS_FILE = __DIR__ . DIRECTORY_SEPARATOR . 'scans.json';

if (!file_exists($OUT_DIR)) {
    @mkdir($OUT_DIR, 0777, true);
}

function load_state() {
    global $STATE_FILE;
    if (!file_exists($STATE_FILE)) return ['running'=>false];
    $j = @file_get_contents($STATE_FILE);
    $a = @json_decode($j, true);
    return is_array($a) ? $a : ['running'=>false];
}

function save_state($state) {
    global $STATE_FILE;
    file_put_contents($STATE_FILE, json_encode($state));
}

function load_scans() {
    global $SCANS_FILE;
    if (!file_exists($SCANS_FILE)) return [];
    $j = @file_get_contents($SCANS_FILE);
    $a = @json_decode($j, true);
    return is_array($a) ? $a : [];
}

function save_scan($entry) {
    global $SCANS_FILE;
    $scans = load_scans();
    $scans[] = $entry;
    file_put_contents($SCANS_FILE, json_encode($scans));
}

function send_tcp($host, $port, $msg) {
    $fp = @fsockopen($host, $port, $errno, $errstr, 2);
    if (!$fp) return false;
    fwrite($fp, $msg);
    fclose($fp);
    return true;
}

// Actions: start, cut, stop, status
$action = $_REQUEST['action'] ?? '';
if ($action === 'start') {
    $state = load_state();
    if (!empty($state['running'])) {
        echo json_encode(['ok'=>false,'msg'=>'Already running']);
        exit;
    }
    $out = escapeshellarg($OUT_DIR);
    $cmd = "start /B \"\" $PYTHON " . escapeshellarg($TEST2) . " --out-dir $out --control-port $CONTROL_PORT --overwrite";
    // Windows: use pclose(popen("start /B ...","r"));
    pclose(popen($cmd, 'r'));
    $state = ['running'=>true,'start_time'=>time(),'out_dir'=>$OUT_DIR];
    save_state($state);
    echo json_encode(['ok'=>true]);
    exit;
}

if ($action === 'cut') {
    $state = load_state();
    if (empty($state['running'])) {
        echo json_encode(['ok'=>false,'msg'=>'Recorder not running']);
        exit;
    }
    $barcode = trim($_REQUEST['barcode'] ?? '');
    // send CUT:barcode
    $msg = 'CUT' . ($barcode !== '' ? ':' . $barcode : '');
    $sent = send_tcp('127.0.0.1', $CONTROL_PORT, $msg);
    if (!$sent) {
        echo json_encode(['ok'=>false,'msg'=>'Failed to send to control port']);
        exit;
    }
    // guess filename that will be created: prefix + current timestamp
    $prefix = $barcode !== '' ? $barcode : 'record';
    $ts = date('Ymd_His');
    $filename = $prefix . '_' . $ts . '.mp4';
    $entry = ['barcode'=>$barcode,'time'=>time(),'file'=>$filename];
    save_scan($entry);
    echo json_encode(['ok'=>true,'entry'=>$entry]);
    exit;
}

if ($action === 'stop') {
    $state = load_state();
    if (empty($state['running'])) {
        echo json_encode(['ok'=>false,'msg'=>'Not running']);
        exit;
    }
    $sent = send_tcp('127.0.0.1', $CONTROL_PORT, 'STOP');
    if (!$sent) {
        echo json_encode(['ok'=>false,'msg'=>'Failed to send STOP']);
        exit;
    }
    $state['running'] = false;
    $state['stop_time'] = time();
    save_state($state);
    echo json_encode(['ok'=>true]);
    exit;
}

if ($action === 'status') {
    $state = load_state();
    $scans = load_scans();
    echo json_encode(['state'=>$state,'scans'=>$scans]);
    exit;
}

// HTML UI
$state = load_state();
$scans = load_scans();
?>
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>Scan3 Recorder Control</title>
  <style> body{font-family:Arial,Helvetica,sans-serif;margin:20px} button{margin-right:8px} table{border-collapse:collapse;width:100%;margin-top:12px} th,td{border:1px solid #ddd;padding:8px} th{background:#f2f2f2}</style>
  <script>
    async function api(action, data={}){
      const params = new URLSearchParams({action, ...data});
      const res = await fetch('scan3.php', {method:'POST', body:params});
      return res.json();
    }
    function updateUI(status){
      document.getElementById('startBtn').disabled = status.running;
      document.getElementById('stopBtn').disabled = !status.running;
      document.getElementById('barcode').disabled = !status.running;
      if (!status.running) document.getElementById('timer').textContent = '00:00:00';
    }
    async function refresh(){
      const r = await api('status');
      const state = r.state || {running:false};
      updateUI(state);
      const scans = r.scans || [];
      const tbody = document.getElementById('scans');
      tbody.innerHTML = '';
      scans.reverse().forEach(s => {
        const tr = document.createElement('tr');
        const t1 = document.createElement('td'); t1.textContent = s.barcode;
        const t2 = document.createElement('td'); t2.textContent = new Date(s.time*1000).toLocaleString();
        const t3 = document.createElement('td'); t3.textContent = s.file;
        tr.appendChild(t1); tr.appendChild(t2); tr.appendChild(t3);
        tbody.appendChild(tr);
      });
    }
    async function startRec(){
      const r = await api('start');
      if (r.ok) { refresh(); startTimer(); }
      else alert(r.msg||'Failed to start');
    }
    async function stopRec(){
      const r = await api('stop');
      if (r.ok) { refresh(); stopTimer(); }
      else alert(r.msg||'Failed to stop');
    }
    async function doCut(){
      const code = document.getElementById('barcode').value.trim();
      if (code===''){
        if (!confirm('Gửi CUT không kèm barcode?')) return;
      }
      const r = await api('cut', {barcode: code});
      if (r.ok) { document.getElementById('barcode').value=''; refresh(); }
      else alert(r.msg||'Failed to cut');
    }

    let timerInterval=null;
    function startTimer(){
      if (timerInterval) return;
      timerInterval = setInterval(async ()=>{
        const r = await api('status');
        const st = r.state||{};
        if (!st.running){ document.getElementById('timer').textContent='00:00:00'; return; }
        const start = st.start_time || Math.floor(Date.now()/1000);
        const diff = Math.floor(Date.now()/1000) - start;
        const h = String(Math.floor(diff/3600)).padStart(2,'0');
        const m = String(Math.floor(diff%3600/60)).padStart(2,'0');
        const s = String(diff%60).padStart(2,'0');
        document.getElementById('timer').textContent = h+':'+m+':'+s;
      }, 800);
    }
    function stopTimer(){ if (timerInterval){ clearInterval(timerInterval); timerInterval=null; document.getElementById('timer').textContent='00:00:00'; }}

    window.addEventListener('load', ()=>{ refresh(); if (<?php echo json_encode(!empty($state['running'])); ?>) startTimer(); document.getElementById('barcode').addEventListener('keydown', e=>{ if (e.key==='Enter') doCut(); }); setInterval(refresh, 5000); });
  </script>
</head>
<body>
  <h2>Scan3 Recorder Control</h2>
  <div>
    <button id="startBtn" onclick="startRec()" <?php echo !empty($state['running']) ? 'disabled' : ''; ?>>▶ Bắt đầu ghi</button>
    <button id="stopBtn" onclick="stopRec()" <?php echo empty($state['running']) ? 'disabled' : ''; ?>>⏹ Dừng hẳn</button>
    <span style="margin-left:20px">Thời gian ghi: <b id="timer">00:00:00</b></span>
  </div>

  <div style="margin-top:12px">
    <label>Mã vạch:</label>
    <input id="barcode" style="font-size:16px;padding:6px;width:260px" <?php echo empty($state['running']) ? 'disabled' : ''; ?> />
    <button onclick="doCut()">Gửi CUT</button>
  </div>

  <table>
    <thead><tr><th>Mã vạch</th><th>Thời gian quét</th><th>Video (dự kiến)</th></tr></thead>
    <tbody id="scans">
      <?php foreach(array_reverse($scans) as $s): ?>
      <tr>
        <td><?php echo htmlspecialchars($s['barcode']); ?></td>
        <td><?php echo date('Y-m-d H:i:s', $s['time']); ?></td>
        <td><?php echo htmlspecialchars($s['file']); ?></td>
      </tr>
      <?php endforeach; ?>
    </tbody>
  </table>

</body>
</html>

?>

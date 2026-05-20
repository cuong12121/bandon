<?php
// scan.php: Form nhập mã QR, gọi Python qua exec để xử lý ngay và ghi log
$msg = '';
$logFile = __DIR__ . DIRECTORY_SEPARATOR . 'scan.log';
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $barcode = trim($_POST['barcode'] ?? '');
    if ($barcode) {
        try {
            $python = 'python';
            $script = __DIR__ . DIRECTORY_SEPARATOR . 'test1.py';
            $cmd = escapeshellcmd($python) . ' ' . escapeshellarg($script) . ' ' . escapeshellarg($barcode);

            // Ghi lệnh vào log
            file_put_contents($logFile, "[" . date('c') . "] Command: $cmd\n", FILE_APPEND);

            // Thực thi và thu thập output
            exec($cmd . ' 2>&1', $output, $returnVar);

            // Ghi output và return code
            file_put_contents($logFile, "[" . date('c') . "] Output: " . implode("\n", $output) . "\nReturn: $returnVar\n", FILE_APPEND);

            if ($returnVar === 0) {
                // Thành công -> chuyển về template để xem video
                header('Location: template.php');
                exit;
            } else {
                $msg = 'Lỗi khi gọi Python. Xem scan.log.';
            }
        } catch (Exception $e) {
            $msg = 'Đã xảy ra lỗi: ' . $e->getMessage();
            file_put_contents($logFile, "[" . date('c') . "] Exception: " . $e->getMessage() . "\n", FILE_APPEND);
        }
    } else {
        $msg = 'Vui lòng nhập mã.';
    }
}
?>
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <title>Bắn mã QR</title>
</head>
<body>
    <h2>Nhập mã QR (bắn mã vạch)</h2>
    <form method="post">
        <input type="text" name="barcode" autofocus required style="font-size:20px;">
        <button type="submit">Gửi mã</button>
    </form>
    <p style="color:green; font-weight:bold;">
        <?php echo $msg; ?>
    </p>
    <p style="font-size:12px;color:#666;">Logs: <a href="scan.log">scan.log</a></p>
</body>
</html>

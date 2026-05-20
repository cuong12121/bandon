<?php
// scan_test2.php
// POST param: 'code' -> calls test2.py via shell and returns JSON

$code = 111;
// Simple escaping for double-quotes on Windows paths
$code_escaped = str_replace('"', '\\"', $code);

// Adjust Python path if needed
$python = 'C:\\laragon\\bin\\python\\python-3.10\\python.exe';
$script = 'C:\\Users\\cuong\\Desktop\\New folder (11)\\test2.py';

$cmd = "\"$python\" \"$script\" \"$code_escaped\" 2>&1";
$output = shell_exec($cmd);

header('Content-Type: application/json');
echo json_encode(['output' => trim($output)]);

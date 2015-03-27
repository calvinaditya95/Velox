<?php
[COM_DOT_NET]

$runCommand = "D:\\ScanBoyConsole\\ScanBoy_Console.exe COM1 9600 8 1 0 1"; 
$WshShell = new COM("WScript.Shell");
$output = $WshShell->Exec($runCommand)->StdOut->ReadAll;
json_decode($output);

?>
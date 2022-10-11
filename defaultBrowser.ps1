


#SendKeys
$wshell = New-Object -ComObject wscript.shell 

Start-Process -FilePath "control"

Start-Sleep 1
for ($i = 0; $i -lt 12; $i++) {
    $wshell.SendKeys("+{TAB}")
    Start-sleep 1
}
$wshell.SendKeys("{ENTER}")

Start-Sleep 1
for ($i = 0; $i -lt 6; $i++) {
    $wshell.SendKeys("{TAB}")
    Start-sleep 1
}
$wshell.SendKeys("{ENTER}")

Start-Sleep 1
for ($i = 0; $i -lt 3; $i++) {
    $wshell.SendKeys("{TAB}")
    Start-sleep 1
}
$wshell.SendKeys("{ENTER}")

Start-Sleep 1
for ($i = 0; $i -lt 5; $i++) {
    $wshell.SendKeys("{TAB}")
    Start-sleep 1
}
$wshell.SendKeys("{ENTER}")

Start-Sleep 1
for ($i = 0; $i -lt 1; $i++) {
    $wshell.SendKeys("{TAB}")
    Start-sleep 0.5
}
$wshell.SendKeys("{ENTER}")

Start-Sleep 3
$wshell.SendKeys("%{F4}")
Start-Sleep 1
$wshell.SendKeys("%{F4}")







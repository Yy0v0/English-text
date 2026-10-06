$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$env:USERPROFILE\Desktop\英语闯关游戏.lnk")
$Shortcut.TargetPath = "C:\Users\屈子悦\WorkBuddy\20260425161018\yyxx\英语闯关游戏.html"
$Shortcut.Description = "英语趣味闯关学习游戏"
$Shortcut.Save()
Write-Host "Shortcut updated"

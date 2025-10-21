$wshell = New-Object -ComObject wscript.shell

$keyToPress = "{SCROLLLOCK}"

Write-Host "Running..."

while ($true) {
    try {
        $wshell.SendKeys($keyToPress)
        Start-Sleep -Seconds 30
    }
    catch {
        Write-Host "Bye..."
        break
    }
}

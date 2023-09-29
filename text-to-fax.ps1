function stringToFax {
    $inputString = Read-Host "my fax"
    $filePath = '.\output.txt'
    $inputString | Out-File -FilePath $filePath -Force -Encoding utf8
    Start-Process -FilePath $filePath
    Start-Sleep -Seconds 2

    $wshell = New-Object -ComObject wscript.shell
    $wshell.AppActivate('Untitled - Notepad')
    $wshell.SendKeys('^(p)')  

    Start-Sleep -Seconds 2
    $wshell.SendKeys("{ENTER}")
    $faxWindow = $null
    while ($faxWindow -eq $null) {
        $faxWindow = $wshell.AppActivate("Neues Fax")
        Start-Sleep -Milliseconds 250
    }
   
    $wshell.SendKeys("**2")
    $wshell.SendKeys("%{s}")
}

while ($true) {
    stringToFax
    $runAgain = Read-Host "Do you want to run the process again? (yes/no)"
    if ($runAgain.ToLower() -ne "yes") {
        break  
    }
}

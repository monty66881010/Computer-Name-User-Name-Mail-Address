$PCinfo1 = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -Property UserName)
$PCinfo2 = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -Property Name)
$PCinfo3 = (Get-CimInstance -ClassName Win32_OperatingSystem | Select-Object -Property version)
$PCinfo4 = (Get-ChildItem -Path C:\Users\$env:username\AppData\Local\Microsoft\Outlook | Where-Object {$_.Name -like "*ost"} |Select-Object -Property name)
$PCinfo5 = (Get-ChildItem -Path C:\Users\$env:username\AppData\Local\Microsoft\Outlook | Where-Object {$_.Name -like "*ost"} |Select-Object -Property LastWriteTime)
$PCinfo6 = (Get-ChildItem -Path C:\Users\$env:username\AppData\Local\Microsoft\Outlook | Where-Object {$_.Name -like "*ost"} |Select-Object -Property Length)
$PC = @(
 [pscustomobject]@{
  UserName = $PCinfo1
   ComputerName = $PCinfo2
    OSversion = $PCinfo3
     OstName = $PCinfo4
      OstLastWriteTime = $PCinfo5
        OstSize = $PCinfo6
}
)
$PC | Export-Csv \\share file\Shares\temp\APPlist\PCinfo1.csv -Append -Force -NoTypeInformation

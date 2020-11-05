$ServerArray=Get-Content c:\users\barry.bradford.a\desktop\servers.csv
$DefineSaveLocation=""
if ($DefineSaveLocation -eq "")
    {$DefineSaveLocation="C:\users\barry.bradford.a\desktop\services.csv"}
$SaveLocaPath=Test-Path $DefineSaveLocation
if ($SaveLocaPath -eq $False)
    {New-Item -ItemType directory -Path $DefineSaveLocation}
cd $DefineSaveLocation
Foreach ($Server in $ ServerArray )
 {
  Write-Host "Retrieving Servers for $Server "    
  Get-WmiObject win32_service -ComputerName $Server  | select Name,
  @{N="Startup Type";E={$_.StartMode}},
  @{N="Service Account";E={$_.StartName}},
  @{N="System Name";E={$_.Systemname}} | Sort-Object "Name" > ".\$Server -Services.txt"
 }
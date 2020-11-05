Get-ADGroup -filter "GroupCategory -eq 'Security'" -properties Member |
 Select Name,@{Name="Members";
 Expression={($_.member | Measure-Object).count}},
 GroupCategory,GroupScope,Distinguishedname |
 Out-GridView -Title "Select one or more groups to export" -OutputMode Multiple |
 foreach {
   Write-Host "Exporting $($_.name)" -ForegroundColor cyan
   #replace spaces in name with a dash
   $name = $_.name -replace " ","-"
   $file = Join-Path -path "C:\temp" -ChildPath "$name.csv"
   Get-ADGroupMember -identity $_.distinguishedname -Recursive |
   Get-ADUser -Properties Title,Department |
   Select Name,Title,Department,SamAccountName,DistinguishedName |
   Export-CSV -Path $file -NoTypeInformation
 Get-Item -Path $file
 }
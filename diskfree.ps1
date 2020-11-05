
clear 
$cre = get-Credential # This user account should be Adminstrator  
$file = get-Content c:\temp\computers.txt  # Replace it with your TXT file which contain Name of Computers  
 
foreach ( $args in $file) { 
get-WmiObject win32_logicaldisk -Credential $cre -ComputerName $args -Filter "Drivetype=3"  |  
ft SystemName,DeviceID,VolumeName,@{Label="Total SIze";Expression={$_.Size / 1gb -as [int] }},@{Label="Free Size";Expression={$_.freespace / 1gb -as [int] }} -autosize | export-csv C:\temp\diskfree.csv
} 
 
## END OF SCRIPT ### 
#This script searches for .encrypted files within a specified list of computers (computers.txt) and exports the list of files to CSV
# INITIAL RELEASE: 15/07/15
#AUTHOR: BARRY BRADFORD
#$StrComputer = "c:\temp\computers.txt"
Get-WmiObject Win32_LogicalDisk -filter "DriveType = 3" | -include *.avi -recurse | Export-Csv "C:\temp\EncryptedFiles.csv"
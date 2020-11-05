$InputFile="C:\users\barry.bradford.a\desktop\serv.csv"
get-ADComputer $InputFile | Get-WmiObject win32_service | format-table Name, StartName | Out-File c:\users\barry.bradford.a\desktop\serv2.csv

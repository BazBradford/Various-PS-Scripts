#Created By Barry Bradford - C2020''''
##E:barry.bradford@westnet.com.au'''''
#'''''''''''''''''''''''''''''''''''''

#Create Instances
$objSnmpObject = new-object -comobject AxNetwork.SnmpObject
#Create Network specific GUIDS for reference
Write-Host "Network Component Version " $objSnmpObject.Version
Write-Host "  Build  " $objSnmpObject.Build 
Write-Host "  Module " $objSnmpObject.Module
Write-Host "License Status: " $objSnmpObject.LicenseStatus
# ********************************************************************
#  Function PrintSnmpTrapData
# ********************************************************************
Function PrintSnmpTrap ($objSnmpTrap)
{
   Write-Host " "
   $res =   "Trap from      : " + $objSnmpTrap.Host
   Write-Host $res
   Write-Host " "
   Write-Host "Variables         :"
   Write-Host " "

   $objSnmpObject = $objSnmpTrap.GetFirstObject()
 
   While($objSnmpTrap.LastError -eq 0)
   {
      $result = GetTypeString $objSnmpObject.Type
      Write-Host "OID        : "  $objSnmpObject.OID
      Write-Host "Value      : "  $objSnmpObject.Value
      Write-Host "Type       : "  $result
      Write-Host "Request ID : "  $objSnmpObject.RequestID
      Write-Host " "
     $objSnmpObject = $objSnmpTrap.GetNextObject()
   }

}

# ********************************************************************
#  Function GetTypeString()
# ********************************************************************
Function GetTypeString ($lType)
{   
   $objConstants	 = new-object -comobject AxNetwork.NwConstants
   switch($lType)
   {
     $objConstants.nwSNMP_TYPE_INTEGER32{"nwSNMP_TYPE_INTEGER32"}
     $objConstants.nwSNMP_TYPE_BITS{"nwSNMP_TYPE_BITS"}
     $objConstants.nwSNMP_TYPE_OCTETSTRING{"nwSNMP_TYPE_OCTETSTRING"}
     $objConstants.nwSNMP_TYPE_NULL{"nwSNMP_TYPE_NULL"}
     $objConstants.nwSNMP_TYPE_OBJECTIDENTIFIER{"nwSNMP_TYPE_OBJECTIDENTIFIER"}
     $objConstants.nwSNMP_TYPE_SEQUENCE{"nwSNMP_TYPE_SEQUENCE"}
     $objConstants.nwSNMP_TYPE_IPADDRESS{"nwSNMP_TYPE_IPADDRESS"}
     $objConstants.nwSNMP_TYPE_COUNTER32 {"nwSNMP_TYPE_COUNTER32"}
     $objConstants.nwSNMP_TYPE_GAUGE32{"nwSNMP_TYPE_GAUGE32"}
     $objConstants.nwSNMP_TYPE_TIMETICKS{"nwSNMP_TYPE_TIMETICKS"}
     $objConstants.nwSNMP_TYPE_OPAQUE{"nwSNMP_TYPE_OPAQUE"}
     $objConstants.nwSNMP_TYPE_COUNTER64{"nwSNMP_TYPE_COUNTER64"}
     $objConstants.nwSNMP_TYPE_UNSIGNED32{"nwSNMP_TYPE_UNSIGNED32"}
     default{"UNKNOWN"}
  }
}

# ***************************************************************************
# Function ReadInput
# ***************************************************************************
Function ReadInput($strPrompt, $strDefaultValue, $bAllowEmpty)
{ 
  $strReturn = ""  
  If ($strDefaultValue -ne "")
  {
     $strPrompt += " (leave empty for " + $strDefaultValue + "): "
  }
  Do 
  {       
    Write-Host $strPrompt
    $strReturn = read-host
    
    If ($strReturn -eq "" -and $strDefaultValue -ne "")
    {
      $strReturn = $strDefaultValue
      Write-Host $strReturn
    }
    elseif ($strReturn -eq "" -and $bAllowEmpty -eq $True)
    {
      break
    }   
  } While ($strReturn -eq "") 
  Write-Host ""
  return $strReturn
}


# ***************************************************************************
# MAIN SCRIPT
# ***************************************************************************

cls

# Create a SnmpTrapOut instance
$objSnmpTrapManager = new-object -comobject AxNetwork.SnmpTrapManager
$objConstants	 = new-object -comobject AxNetwork.NwConstants

# Display ActiveXperts Network Component Version
Write-Host "ActiveXperts Network Component " $objSnmpTrapManager.Version "`nBuild: " $objSnmpTrapManager.Build "`nModule: "  $objSnmpTrapManager.Module "`nLicense Status: " $objSnmpTrapManager.LicenseStatus "`nLicense Key: " $objSnmpTrapManager.LicenseKey "`n`n";

# Logfile
$objSnmpTrapManager.Logfile = $env:temp + "\objSnmpTrapManagerRecv.log"
Write-Host "Log file used: " $objSnmpTrapManager.Logfile "`n"

# Get Host, community name and optionally a MIB file
$strCommunity   = ReadInput "Enter community" "public" $True
$strMibFile     = ReadInput "Enter location of MIB file (optional, only required when using alpha-numeric OID's)" "" $True

# Initialize SNMP
$objSnmpTrapManager.Initialize()
$res = "Initialize, result: " + $objSnmpTrapManager.LastError + " (" + $objSnmpTrapManager.GetErrorDescription( $objSnmpTrapManager.LastError ) + ")"
Write-Host $res
If($objSnmpTrapManager.LastError -ne 0 )
{
  exit
}

if($strMibFile -ne "")
{
  $objSnmpTrapManager.LoadMibFile($strMibFile)
}

# Start listening for incoming SNMP traps
$objSnmpTrapManager.StartListening($strCommunity)
$res =  "StartListening, result: " + $objSnmpTrapManager.LastError + " (" + $objSnmpTrapManager.GetErrorDescription( $objSnmpTrapManager.LastError ) + ")"
Write-Host $res

# Connection established; receive incoming traps
Write-Host "Waiting for incoming traps ..."
While($true)
{
  $objSnmpTrap = $objSnmpTrapManager.GetFirstTrap()

  While ($objSnmpTrapManager.LastError -eq 0)
  { 
    PrintSnmpTrap ($objSnmpTrap)
    trap [Exception]
    {
      $objSnmpTrap = $objSnmpTrapManager.GetNextTrap()
    }
    break
  }
    
  start-Sleep -m 1000
  if($objSnmpTrapManager.LastError -eq 0)
  {
    $strEnter     = ReadInput "type y to stop" "" $true
    if($strEnter -eq "y")
    {
      break
    }
  }
}

# Stop listening
$objSnmpTrapManager.StopListening()
$res = "StopListening, result: " + $objSnmpTrapManager.LastError + " (" + $objSnmpTrapManager.GetErrorDescription( $objSnmpTrapManager.LastError ) + ")"
Write-Host $res

# Shutdown SNMP
$objSnmpTrapManager.Shutdown()
$res =  "Shutdown, result: " + $objSnmpTrapManager.LastError + " (" + $objSnmpTrapManager.GetErrorDescription( $objSnmpTrapManager.LastError ) + ")"
Write-Host $res

# Finished
Write-Host "Finished."

if (!(Get-Module ActiveDirectory))
{
	Import-Module ActiveDirectory -ErrorAction Stop
}

function Remove-LastBackSlash
{
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[String] $Str
	)
	
	$Index = $Str.LastIndexOf("\")
	$Length = $Str.Length -1
	
	If ($Index -eq $Length)
	{
		$Str = $Str.SubString(0,$Index)
	}
	
	return $Str
}

function Ask-ForChoice
{
	param
	(
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("CT")]
		[String] $ChoiceTle,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("CM")]
		[String] $ChoiceMsg,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("YM")]
		[String] $YesMsg,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("NM")]
		[String] $NoMsg
	)
	
	$yes = New-Object System.Management.Automation.Host.ChoiceDescription '&Y',$YesMsg
	$no = New-Object System.Management.Automation.Host.ChoiceDescription '&N',$NoMsg
	
	$ChoiceOpt = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
	
	$Choice = $host.ui.PromptForChoice($ChoiceTle,$ChoiceMsg,$ChoiceOpt,0)
	
	return $Choice
}

function Search-ADObjQuicker
{
	param
	(
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet('Base','OneLevel','Subtree')]
        [Alias("SC")]
		[String] $Scope = 'Subtree'
	)
    
	$Index = $ADObjDN.IndexOf('DC=')
	$DomainDN = $ADObjDN.Substring($Index)
	$LdapPath = 'LDAP://' + $DomainDN
	$Filter = '(&distinguishedName=' + $ADObjDN + ')'
	
	$ADSearch = New-Object DirectoryServices.DirectorySearcher 
    $ADSearch.SearchRoot = $LdapPath 
    $ADSearch.SearchScope = $Scope 
    $ADSearch.Filter = $Filter
    $ADResult = $ADSearch.FindAll() 
     
	return $ADResult
}

function List-ADGCOnePerDomainReachable
{
	param
	(	
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("FQDN")]
		[String[]] $ADDomainFQDNCol,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("3268","3269")]
		[Alias("Port")]
		[String] $ADGCPort = "3268",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateScript({Test-Path $_ -PathType 'Container'})]
        [Alias("Path")]
		[String] $ADForestGCListRoot = "$env:USERPROFILE\Desktop",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("ASCII","UTF8")]
		[Alias("CS")]
		[String] $Charset = "UTF8",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
        [Boolean]$Clear = $true
	)
	
	if ($ADForestGCListRoot -ne "$env:USERPROFILE\Desktop")
	{
		$ADForestGCListRoot = Remove-LastBackSlash $ADForestGCListRoot
	}
	
	$ADForestGCListPath = $ADForestGCListRoot + "\ADForestGCList.csv"
	
	$ADForestGCListExists = Test-Path -Path $ADForestGCListPath -PathType Leaf
	
	if ($Clear -and $ADForestGCListExists)
	{
		$Choice_Tle = 'Flush Global Catalog list'
		$Choice_Msg = 'Do you really want to flush and rediscover Global Catalog list rather than reusing "' + $ADForestGCListPath + '" ?'
		$Choice_YesMsg = 'GC list flush confirmed.'
		$Choice_NoMsg = 'GC list flush canceled.'
		
		$Choice = Ask-ForChoice $Choice_Tle $Choice_Msg $Choice_YesMsg $Choice_NoMsg
	}
	elseif (!$ADForestGCListExists)
	{
		$Choice = 0
	}
	else
	{
		$Choice = 1
	}
	
	if ($Choice -eq 0)
	{
		$ADDomainPSHostFQDN = (Get-ADDomain).DNSRoot
		#$ADDomainPSHostSite = (Get-ADDomainController).Site #based on ADWS
		$ADDomainPSHostSite = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).Name
	
		$ADForestGCPsoCol = @()

		foreach ($ADDomainFQDN in $ADDomainFQDNCol)
		{
			$Msg = "`n" + 'Discovering one Global Catalog avalaible in Domain "' + $ADDomainFQDN + '":'
			Write-Host $Msg -ForegroundColor White
		
			if ($ADDomainFQDN -eq $ADDomainPSHostFQDN)
			{
				$ADDomainGCColInt = Get-ADDomainController -Server $ADDomainFQDN -Filter { IsGlobalCatalog -eq "True" -and Site -eq $ADDomainPSHostSite }
				$ADDomainGCColExt = Get-ADDomainController -Server $ADDomainFQDN -Filter { IsGlobalCatalog -eq "True" -and Site -ne $ADDomainPSHostSite }
				$ADDomainGCCol = @($ADDomainGCColInt) + @($ADDomainGCColExt)
			}
			
			else
			{
				$ADDomainGCCol = Get-ADDomainController -Server $ADDomainFQDN -Filter { IsGlobalCatalog -eq "True" }
			}
		
			$ADDomainGCCtr = $null	

			:nextdom while (!$ADDomainGCCtr)
			{
				foreach ($ADDomainGC in $ADDomainGCCol)
				{
					$ADDomainGCFQDN = $ADDomainGC.HostName
					$ADDomainGCConxStr = $ADDomainGCFQDN + ":" + $ADGCPort
					
					$ADDomainGCConx = $null
					
					try
					{
						$ADDomainGCConx = Get-ADObject $ADDomainGC.ComputerObjectDN -Server $ADDomainGCConxStr -ErrorAction Stop | Select Name
					}
					catch [Microsoft.ActiveDirectory.Management.ADServerDownException]
					{
						$Msg = 'Global Catalog "' + $ADDomainGCFQDN + '" is not responding on port ' + $ADGCPort + '. Seeking for another Domain Controller...'
						Write-Host $Msg
					}

					if ($ADDomainGCConx)
					{
						$Msg = "`n" + 'Global Catalog "' + $ADDomainGCFQDN + '" is responding on port "' + $ADGCPort + '". Now it is the target Domain Controller for LDAP search on "'+ $ADDomainFQDN + '".' + "`n"
						Write-Host $Msg
					
						$ADDomainGCPso = New-Object -TypeName PSObject
						$ADDomainGCPso | Add-Member -MemberType NoteProperty -Name HostName -Value $ADDomainGC.HostName
						$ADDomainGCPso | Add-Member -MemberType NoteProperty -Name Port -Value $ADGCPort
						$ADDomainGCPso | Add-Member -MemberType NoteProperty -Name DomainFQDN -Value $ADDomainGC.Domain
						$ADDomainGCPso | Add-Member -MemberType NoteProperty -Name DomainDN -Value $ADDomainGC.DefaultPartition

						$ADForestGCPsoCol = $ADForestGCPsoCol + $ADDomainGCPso
						
						break nextdom
					}
				}

				$ADDomainGCCtr++
			}
		}

		$Msg = "`n" + 'Backup Global Catalog targeted servers in "' + $ADForestGCListPath + '".'
		Write-Host $Msg -ForegroundColor White
		
		$ADForestGCPsoCol | Export-Csv -Path $ADForestGCListPath -NoTypeInformation -Encoding $Charset
	}
	elseif ($Choice -eq 1)
	{
		$Msg = "`n" + 'Importing existing Global Catalog list from "' + $ADForestGCListPath + '".'
		Write-Host $Msg -ForegroundColor White
		
		$Exists = Test-Path -Path $ADForestGCListPath -PathType Leaf
		
		$ADForestGCPsoCol = @(Import-Csv -Path $ADForestGCListPath)
	}
	
	if ($ADForestGCPsoCol.count -ge $ADDomainFQDNCol.count)
	{
		return $ADForestGCPsoCol
	}
}

function Get-ADGCOneForADObj
{
	param
	(
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
				
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("GC")]
	    [PSObject[]] $ADForestGCPsoCol
	)

	$Index = $ADObjDN.IndexOf("DC=")
	$ADDomainDN = $ADObjDN.Substring($Index)
	$ADDomainGCPso = $ADForestGCPsoCol |
	? {
		$_.DomainDN -eq  $ADDomainDN
	} |
	Select-Object -First 1
	
	return $ADDomainGCPso
}

function Get-ADSecurityGroupDup
{
	param
	(	
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[PSObject] $ADSGroupPsoDupCol,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("For")]
		[String] $ADObjStr,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("By")]
		[String] $Property
	)
	
	$Msg = @()
		
	foreach ($ADSGroupPsoDup in $ADSGroupPsoDupCol)
	{	
		if ($ADSGroupPsoDup.Duplicated -eq $true)
		{	
			if ($ADSGroupPsoDup.CanonicalName -eq $ADObjStr)
			{
				$DupStr = '- ' + ' loop on ' + '"' + $ADSGroupPsoDup.CanonicalName + '"' + "`n"
			}	
			else
			{
				$DupStr = '- ' + '"' + $ADSGroupPsoDup.CanonicalName + '"' + ' appears more than one time' + "`n"
			}
			
			$Msg = $Msg + @($DupStr)
		}
	}
	
	if ($Msg.Count)
	{
		if ($Msg.Count -gt 1)
		{
			$Msg = $Msg[0..($Msg.Count-2)] + @($Msg[$Msg.Count-1].SubString(0,$Msg[$Msg.Count-1].LastIndexof("`n")))
		}
		else
		{
			$Msg = $Msg[$Msg.Count-1].SubString(0,$Msg[$Msg.Count-1].LastIndexof("`n"))
		}
		
		switch ($Property)
		{
			"MemberOf"
			{
				$DupTle = 'MemberOf nesting chain for "' + $ADObjStr + '" seems not optimal on some points:' + "`n"
			}
			
			"Member"
			{
				$DupTle = 'Member nesting chain for "' + $ADObjStr + '" seems not optimal on some points:' + "`n"
			}
		}
		
		$Msg = @($DupTle) + $Msg
		Write-Host $Msg -BackgroundColor DarkRed
	}
}

function Dig-ADSecurityGroupMemberOf
{
	param
	(  
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[ValidateScript({$_.GroupCategory -eq "Security"})]
        [Microsoft.ActiveDirectory.Management.ADGroup] $ADSGroup,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("GC")]
        [PSObject[]] $ADForestGCPsoCol
    )
 
	foreach ($ADSGroupPsoDup in $ADSGroupPsoDupCol)
	{ 
		$i = [Array]::IndexOf($ADSGroupPsoDupCol,$ADSGroupPsoDup)
		
		if ($ADSGroupPsoDup.CanonicalName -eq $ADSGroup.CanonicalName)
		{
			[Array]::Clear($ADSGroupPsoDupCol,$i,1)
			
			$ADSGrpPsoDup = New-Object PSObject -Property @{
				CanonicalName = $ADSGroup.CanonicalName
				Duplicated	= $true
			}
	
			$global:ADSGroupPsoDupCol = $ADSGroupPsoDupCol + @($ADSGrpPsoDup)
		
			return
		}
    }
	
	$ADSGroupMbrOfAll = @()
	$ADSGroupMbrOf =  @()
	
	$Msg = 'Analyzing Security Group "' + $ADSGroup.CanonicalName + '"'
	Write-Host $Msg
	
	if ($ADSGroup.GroupScope -eq "DomainLocal")
	{	
		$DomainGCPso = Get-ADGCOneForADObj -DN $ADSGroup.DistinguishedName -GC $ADForestGCPsoCol
		$GC = $DomainGCPso.HostName + ":" + $DomainGCPso.Port
		
		$Msg = 'from "' + $GC + '"'
		Write-Host $Msg
		
		$ADSGroupMbrOf = (Get-ADGroup $ADSGroup.DistinguishedName -Properties MemberOf -Server $GC).MemberOf | 
		? -FilterScript {
			(Get-ADGroup $_ -Server $GC).GroupCategory -eq "Security"
		}
	}
	else
	{   
	    foreach ($ADForestGCPso in $ADForestGCPsoCol)
		{
			$GC = $ADForestGCPso.HostName + ":" + $ADForestGCPso.Port
			
			$Msg = 'from "' + $GC + '"'
			Write-Host $Msg

			$ADSGroupMbrOfAll = (Get-ADGroup $ADSGroup.DistinguishedName -Properties MemberOf -Server $GC).MemberOf
			
			if ($ADSGroupMbrOfAll)
			{
				$ADSGroupMbrOfSec = $ADSGroupMbrOfAll |
				? -FilterScript {
					(Get-ADGroup $_ -Server $GC).GroupCategory -eq "Security"
				}
			
				$ADSGroupMbrOf = $ADSGroupMbrOf + @($ADSGroupMbrOfSec)
			}
		}
		
		$ADSGroupMbrOf = $ADSGroupMbrOf | Sort-Object -Unique
	}
    
    $ADSGroupPso = New-Object PSObject -Property @{
        Name				= $ADSGroup.Name
		CanonicalName		= $ADSGroup.CanonicalName
        DistinguishedName	= $ADSGroup.DistinguishedName
        GroupScope			= $ADSGroup.GroupScope
        MemberOf			= $ADSGroupMbrOf
    }
	
    $ADSGroupPso
	
	$ADSGroupPsoDup = New-Object PSObject -Property @{
		CanonicalName = $ADSGroup.CanonicalName
		Duplicated	= $false
	}
	
	$global:ADSGroupPsoDupCol = $ADSGroupPsoDupCol + @($ADSGroupPsoDup)
	
    if ($ADSGroupMbrOf)
    {
        foreach ($ADSGroupDN in $ADSGroupMbrOf)
		{
        	Get-ADGroup $ADSGroupDN -Properties CanonicalName -Server $GC |
			Dig-ADSecurityGroupMemberOf -GC $ADForestGCPsoCol
    	}
	}
}

function Dig-ADSecurityGroupMember
{
	param
	(  
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[ValidateScript({$_.GroupCategory -eq "Security"})] 
        [Microsoft.ActiveDirectory.Management.ADGroup] $ADSGroup,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
        [String] $DC,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
        [Alias("DGC")]
		[String] $DomainGC,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("GC")]
        [PSObject[]] $ADForestGCPsoCol
    )
	
	foreach ($ADSGroupPsoDup in $ADSGroupPsoDupCol)
	{ 
		$i = [Array]::IndexOf($ADSGroupPsoDupCol,$ADSGroupPsoDup)
		
		if ($ADSGroupPsoDup.CanonicalName -eq $ADSGroup.CanonicalName)
		{
			[Array]::Clear($ADSGroupPsoDupCol,$i,1)
			
			$ADSGrpPsoDup = New-Object PSObject -Property @{
				CanonicalName = $ADSGroup.CanonicalName
				Duplicated	= $true
			}
	
			$global:ADSGroupPsoDupCol = $ADSGroupPsoDupCol + @($ADSGrpPsoDup)
		
			return
		}
    }
 
	$ADSGroupMbrAll = @()
	$ADSGroupMbr =  @()
	
	if ($ADForestGCPsoCol.count -gt 1)
	{
		$DC = (Get-ADGCOneForADObj -DN $ADSGroup.DistinguishedName -GC $ADForestGCPsoCol).HostName
	}
	
	$Msg = 'Analyzing Security Group "' + $ADSGroup.CanonicalName + '" from "' + $DC + '".'
	Write-Host $Msg
	
	try
	{
		$ADSGroupMbrAll = Get-ADGroupMember -Identity $ADSGroup.DistinguishedName -Server $DC
	}
	catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException]
	{
		$Msg = "`n" + '"' + $ADSGroup.CanonicalName + '" is out of domain scope, his members are ignored' + "`n"
		Write-Warning $Msg
	}
	catch [Microsoft.ActiveDirectory.Management.ADException]
	{
		if ($_.Exception.ServerErrorMessage -eq "Exceeded groups or group members limit.")
		{
			$Msg = "`n" + 'Retrieving "' + $ADSGroup.CanonicalName + '" members failed because of ADWS limitations, his members are ignored.'
			$Msg += "`n" + 'Please refer to "ADWS configuration" to adapt ADWS "MaxGroupOrMemberEntries" value:'
			$Msg += "`n" + "http://technet.microsoft.com/en-us/library/dd391908%28WS.10%29.aspx" + "`n"
			Write-Warning $Msg
		}
	}
	
	if ($ADSGroupMbrAll)
	{
		$ADSGroupMbr = $ADSGroupMbrAll | ? -FilterScript { ($_.ObjectClass -eq "group")} |
		? -FilterScript {
			(Get-ADGroup -Identity $_.DistinguishedName -Server $DomainGC).GroupCategory -eq "Security"
		}
	}

	$ADSGroupPso = New-Object PSObject -Property @{
		Name				= $ADSGroup.Name
		CanonicalName		= $ADSGroup.CanonicalName		
		DistinguishedName	= $ADSGroup.DistinguishedName
    	GroupScope			= $ADSGroup.GroupScope
    	Member				= $ADSGroupMbr | Select-Object distinguishedName
	}	
	
	$ADSGroupPso
	
	$ADSGroupPsoDup = New-Object PSObject -Property @{
		CanonicalName = $ADSGroup.CanonicalName
		Duplicated	= $false
	}
	
	$global:ADSGroupPsoDupCol = $ADSGroupPsoDupCol + @($ADSGroupPsoDup)
		
	if ($ADSGroupMbr)
	{
	    foreach ($ADSGrpMbr in $ADSGroupMbr)
		{	
			Get-ADGroup -Identity $ADSGrpMbr.DistinguishedName -Properties CanonicalName -Server $DomainGC |
			Dig-ADSecurityGroupMember -DC $DC -DGC $DomainGC -GC $ADForestGCPsoCol
		}
	}
}

function Measure-ADTokenSize
{
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[PSObject[]] $ADSGroupPsoCol,
		
		[Parameter(ValueFromPipeline = $true, Mandatory = $false)]
		[Alias("DN")]
		[String] $ADObjDN
	)
	
	foreach ($ADSGroupPso in $ADSGroupPsoCol)
	{
		switch ($ADSGroupPso.GroupScope)
		{
			"DomainLocal"
			{
				$D++
			}
			
			"Global"
			{
				$S++
			}
			
			"Universal"
			{
				if ( $ADSGroupPso.DistinguishedName.Substring($ADSGroupPso.DistinguishedName.IndexOf("DC=")) -eq $ADObjDN.Substring($ADObjDN.IndexOf("DC=")) )
				{
					$S++
				}
				else
				{
					$D++
				}
			}
		}
	}
	
	$ADTokenSize = 1200 + 40*$D + 8*$S
	
	return $ADTokenSize
}
 
function List-ADSecurityGroupMemberOf
{  
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
        [Alias("SC")]
		[Microsoft.ActiveDirectory.Management.ADSearchScope] $Scope,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("GC")]
        [PSObject[]] $ADForestGCPsoCol
    )
	
	$DomainGCPso = Get-ADGCOneForADObj -DN $ADObjDN -GC $ADForestGCPsoCol
	$DomainGC = $DomainGCPso.HostName + ":" + $DomainGCPso.Port
	
	$ADObj = Get-ADObject $ADObjDN -Properties CanonicalName,GroupType -Server $DomainGC
	
	$global:ADSGroupPsoDupCol = @()
	
	switch ($ADObj)
	{
		{($_.ObjectClass -eq "domainDNS") -or ($_.ObjectClass -eq "organizationalUnit") -or ($_.ObjectClass -eq "container") -or ($_.ObjectClass -eq "builtinDomain")}
		{		
			$ADSGroupPsoCol = Get-ADGroup -Filter {GroupCategory -eq "Security"} -SearchBase $ADObjDN -SearchScope $Scope -Server $DomainGC |
			% {	
				List-ADSecurityGroupMemberOf -DN $_.DistinguishedName -SC $Scope -GC $ADForestGCPsoCol
	    	}
	    	
			$ADSGroupPsoCol = $ADSGroupPsoCol |
			Sort-Object -Unique -Property DistinguishedName 
	    
			return $ADSGroupPsoCol
			
			break
		}
		
		{($_.ObjectClass -eq "group") -and ($_.GroupType -like "-2*")}
		{
			$ADSGroupPsoCol =  Get-ADGroup $ADObjDN -Properties CanonicalName -Server $DomainGC |
			Dig-ADSecurityGroupMemberOf -GC $ADForestGCPsoCol
			
			Get-ADSecurityGroupDup $ADSGroupPsoDupCol -For $ADObj.CanonicalName -By "MemberOf"		
			
			return $ADSGroupPsoCol
			
			break
		}
		
		{($_.ObjectClass -eq "computer")}
		{
			$ADComputer = Get-ADComputer $ADObjDN -Properties CanonicalName,PrimaryGroup -Server $DomainGC
			
			$Msg = 'Analyzing Computer "' + $ADComputer.CanonicalName + '"'
			Write-Host $Msg
			
			$ADComputerMbrOf = @()
			
    		foreach ($ADForestGCPso in $ADForestGCPsoCol)
			{
				$ForestGC = $ADForestGCPso.HostName + ":" + $ADForestGCPso.Port
				
				$Msg = 'from "' + $ForestGC + '"'
				Write-Host $Msg
				
				$ADComputerMbrOfAll = (Get-ADComputer $ADComputer.DistinguishedName -Properties MemberOf -Server $ForestGC).MemberOf
				
				if ($ADComputerMbrOfAll)
				{
					$ADComputerMbrOfSec = $ADComputerMbrOfAll |
					? -FilterScript {
						(Get-ADGroup $_ -Server $ForestGC).GroupCategory -eq "Security" 
					}
					
					$ADComputerMbrOf = $ADComputerMbrOf + @($ADComputerMbrOfSec)
				}
			}

			$ADComputerMbrOf = @($ADComputerMbrOf | Sort-Object -Unique)
			
			$ADComputerMbrOf = $ADComputerMbrOf + @($ADComputer.PrimaryGroup)
			
			$ADSGroupPsoCol = foreach ($ADSGroupDN in $ADComputerMbrOf)
			{
				Get-ADGroup $ADSGroupDN -Properties CanonicalName -Server $DomainGC |
				Dig-ADSecurityGroupMemberOf -GC $ADForestGCPsoCol
	    	}
			
			Get-ADSecurityGroupDup $ADSGroupPsoDupCol -For $ADComputer.CanonicalName -By "MemberOf"
			
			$ADPso = New-Object PSObject -Property @{
				Name				= $ADComputer.Name
				CanonicalName		= $ADComputer.CanonicalName
				DistinguishedName	= $ADComputer.DistinguishedName
    			GroupScope			= "None"
    			MemberOf			= $ADComputerMbrOf
			}
			
			$ADSGroupPsoCol = @($ADSGroupPsoCol) + $ADPso
			
			return $ADSGroupPsoCol
			
			break
		}
		
		{($_.ObjectClass -eq "user") -or ($_.ObjectClass -eq "inetOrgPerson")}
		{
			$ADUser = Get-ADUser $ADObjDN -Properties CanonicalName,PrimaryGroup -Server $DomainGC
			
			$Msg = 'Analyzing User "' + $ADUser.CanonicalName + '"'
			Write-Host $Msg
			
			$ADUserMbrOf = @()
			
			foreach ($ADForestGCPso in $ADForestGCPsoCol)
    		{
				$ForestGC = $ADForestGCPso.HostName + ":" + $ADForestGCPso.Port
				
				$Msg = 'from "' + $ForestGC + '"'
				Write-Host $Msg
				
				$ADUserMbrOfAll = (Get-ADUser $ADUser.DistinguishedName -Properties MemberOf -Server $ForestGC).MemberOf
				
				if ($ADUserMbrOfAll)
				{
					$ADUserMbrOfSec = $ADUserMbrOfAll |
					? -FilterScript {
						(Get-ADGroup $_ -Server $ForestGC).GroupCategory -eq "Security" 
					}
					
					$ADUserMbrOf = $ADUserMbrOf + @($ADUserMbrOfSec)
				}
			}
	
			$ADUserMbrOf = @($ADUserMbrOf | Sort-Object -Unique)
			
			$ADUserMbrOf = $ADUserMbrOf + @($ADUser.PrimaryGroup)
			
			$ADSGroupPsoCol = foreach ($ADSGroupDN in $ADUserMbrOf)
			{
				Get-ADGroup $ADSGroupDN -Properties CanonicalName -Server $DomainGC |
				Dig-ADSecurityGroupMemberOf -GC $ADForestGCPsoCol
	    	}

			Get-ADSecurityGroupDup $ADSGroupPsoDupCol -For $ADUser.CanonicalName -By "MemberOf"
			
			$ADSGroupCount = $ADSGroupPsoCol.Count
			$ADTokenSize = Measure-ADTokenSize $ADSGroupPsoCol -DN $ADUser.DistinguishedName
			
			$ADSGroupCountAndTokenSize = 'Group Count: '+ $ADSGroupCount + ' | Estimated Token Size: ' + $ADTokenSize + ' bytes' 
			
			$ADPso = New-Object PSObject -Property @{
				Name				= $ADUser.Name
				CanonicalName		= $ADUser.CanonicalName
				DistinguishedName	= $ADSGroupCountAndTokenSize
    			GroupScope			= "None"
    			MemberOf			= $ADUserMbrOf
			}
			
			$ADSGroupPsoCol = @($ADSGroupPsoCol) + $ADPso
			
			return $ADSGroupPsoCol
			
			break
		}
		
		default
		{
			$Msg = 'Script aborted.'
			$Msg += "`n" + 'Please select a Domain, an OU, a Container, a Security Group, a User or a Computer'
			Write-Warning $Msg
		}
	}
}

function List-ADSecurityGroupMember
{  
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
        [Alias("SC")]
		[Microsoft.ActiveDirectory.Management.ADSearchScope] $Scope,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("GC")]
        [PSObject[]] $ADForestGCPsoCol
    )
	
	$DomainGCPso = Get-ADGCOneForADObj -DN $ADObjDN -GC $ADForestGCPsoCol
	$DomainGC = $DomainGCPso.HostName + ":" + $DomainGCPso.Port
	$DC = $DomainGCPso.HostName
	
	$ADObj = Get-ADObject $ADObjDN -Properties CanonicalName,GroupType -Server $DomainGC
	
	$global:ADSGroupPsoDupCol = @()
	
	switch ($ADObj)
	{
		{($_.ObjectClass -eq "domainDNS") -or ($_.ObjectClass -eq "organizationalUnit") -or ($_.ObjectClass -eq "container") -or ($_.ObjectClass -eq "builtinDomain")}
		{ 
			$ADSGroupPsoCol = Get-ADGroup -Filter {GroupCategory -eq "Security"} -SearchBase $ADObjDN -SearchScope $Scope -Server $DomainGC | 
			% {
				List-ADSecurityGroupMember -DN $_.DistinguishedName -SC $Scope -GC $ADForestGCPsoCol
	    	} 
	    
			$ADSGroupPsoCol = $ADSGroupPsoCol |
			Sort-Object -Unique -Property DistinguishedName
	    
			return $ADSGroupPsoCol
			
			break
		}
	
		{($_.ObjectClass -eq "group") -and ($_.GroupType -like "-2*")}
		{
			$ADSGroupPsoCol =  Get-ADGroup $ADObjDN -Properties CanonicalName -Server $DomainGC |
			Dig-ADSecurityGroupMember -DC $DC -DGC $DomainGC  -GC $ADForestGCPsoCol
			
			Get-ADSecurityGroupDup $ADSGroupPsoDupCol -For $ADObj.CanonicalName -By "Member"
			
			return $ADSGroupPsoCol
			
			break
		}
	
		default
		{
			$Msg = 'Script aborted.'
			$Msg += "`n" + 'Please select a Domain, an OU, a Container or a Security Group' 
			Write-Warning $Msg
		}
	}
}

function Graph-ADSecurityGroupMemberOf
{   
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[PSObject[]] $ADSGroupPsoCol,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("SC")]
		[String] $Scope,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[String] $Mode,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("CS")]
		[String] $Charset,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("DGC")]
		[String] $DomainLocalColor = "Red",
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("GGC")]
        [String] $GlobalColor = "Green",
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("UGC")]
        [String] $UniversalColor = "Cyan",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
        [Alias("UC")]
		[String] $UserColor = "Black"
    )

    $GroupColor = {
        
		param
		(	
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[String] $GroupScope
		)

        switch ($GroupScope)
		{
            "DomainLocal" { return $DomainLocalColor }
            "Global" { return $GlobalColor }
            "Universal" { return $UniversalColor }
			"None" { return $UserColor }
        }
    }

    $GraphNode = {
	
		param
		(
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[PSObject] $ADSGroupPso
		)
		
		'ADObj_{0} [' -f [Array]::indexOf($ADSGroupPsoCol,$ADSGroupPso) | Write-Output

		if ($ADSGroupPso.DistinguishedName.IndexOf("DC=") -ne -1)
		{
			$StrMark = $ADSGroupPso.DistinguishedName.IndexOf("DC=")
			$DomainName = $ADSGroupPso.DistinguishedName.Substring($StrMark)

        	'label="{0}|{1}",' -f $ADSGroupPso.Name,$DomainName | Write-Output
		}
		
		else
		{
			'label="{0}|{1}",' -f $ADSGroupPso.Name,$ADSGroupPso.DistinguishedName | Write-Output
		}
		
        'color={0}' -f (&$GroupColor $ADSGroupPso.GroupScope) | Write-Output
        '];' | Write-Output
    }

    'digraph G {' | Write-Output
	
	if ($Charset -eq "UTF8")
	{
		'charset="utf-8";' | Write-Output
	}
	
	'fontsize=9;' | Write-Output
	'fontname=serif;' | Write-Output
	'label="~*~' | Write-Output
	' ' | Write-Output
	'Security Groups and nesting found with {0} search starting from {1} and using MemberOf back-link attribute over {2}' -f $Scope,$ADObjDN,$Mode | Write-Output
	'{0}' -f (Get-Date).ToShortDateString() | Write-Output
	' ' | Write-Output
	'~*~' | Write-Output
	' ' | Write-Output
	'Green: Global Group' | Write-Output
	'Red: Domain Local Group' | Write-Output
	'Cyan: Universal Group"' | Write-Output
    'graph [overlap=false,rankdir=LR];' | Write-Output
	'node [fontsize=8,fontname=serif,shape=record,style=rounded];' | Write-Output
	'edge [dir=forward,arrowsize=0.5,arrowhead=empty];' | Write-Output
		
	foreach ($ADSGroupPso in $ADSGroupPsoCol)
	{
		$Msg = 'Graphing "' + $ADSGroupPso.CanonicalName + '".'
		Write-Host $Msg
		
		$ADSGroupChildPso = $ADSGroupPso
        
		&$GraphNode $ADSGroupChildPso

		foreach ($ADSGroupDN in $ADSGroupChildPso.MemberOf)
		{
			foreach ($ADSGrpPso in $ADSGroupPsoCol)
			{	
				if ($ADSGrpPso.DistinguishedName -eq $ADSGroupDN)
				{
					$ADSGroupParentPso = $ADSGrpPso
					
					'ADObj_{0} -> ADObj_{1};' -f [Array]::indexOf($ADSGroupPsoCol,$ADSGroupChildPso), [Array]::indexOf($ADSGroupPsoCol,$ADSGroupParentPso) |
					Write-Output
					
					break
				}
		    }
		}
	} 
	
	'}'	| Write-Output
}

function Graph-ADSecurityGroupMember
{
	param
	(
		[Parameter(ValueFromPipeline = $true, Mandatory = $true)]
		[PSObject[]] $ADSGroupPsoCol,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("DN")]
		[String] $ADObjDN,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("SC")]
		[String] $Scope,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[String] $Mode,
		[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
		[Alias("CS")]
		[String] $Charset,
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("DGC")]
		[String] $DomainLocalColor = "Red",
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("GGC")]
        [String] $GlobalColor = "Green",
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("UGC")]
        [String] $UniversalColor = "Cyan"
    )

    $GroupColor = {
        
		param
		(	
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[String] $GroupScope
		)

        switch ($GroupScope)
		{
            "DomainLocal" { return $DomainLocalColor }
            "Global"      { return $GlobalColor }
            "Universal"   { return $UniversalColor }
        }
    }

    $GraphNode = {
	
		param
		(
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[PSObject] $ADSGroupPso
		)
	
		$StrMark = $ADSGroupPso.DistinguishedName.IndexOf("DC=")
		$DomainName = $ADSGroupPso.DistinguishedName.Substring($StrMark)

        'ADObj_{0} [' -f [Array]::indexOf($ADSGroupPsoCol,$ADSGroupPso) | Write-Output
        'label="{0}|{1}",' -f $ADSGroupPso.Name, $DomainName | Write-Output
        'color={0}' -f (&$GroupColor $ADSGroupPso.GroupScope) | Write-Output
        '];' | Write-Output
    }

    'digraph G {' | Write-Output
	
	if ($Charset -eq "UTF8")
	{
		'charset="utf-8";' | Write-Output
	}
	
	'fontsize=9;' | Write-Output
	'fontname=serif;' | Write-Output
	'label="~*~' | Write-Output
	' ' | Write-Output
	'Security Groups and nesting found with {0} search starting from {1} and using Member attribute over {2}' -f $Scope,$ADObjDN,$Mode | Write-Output
	'{0}' -f (Get-Date).ToShortDateString() | Write-Output
	' ' | Write-Output
	'~*~' | Write-Output
	' ' | Write-Output
	'Green: Global Group' | Write-Output
	'Red: Domain Local Group' | Write-Output
	'Cyan: Universal Group"' | Write-Output		
    'graph [overlap=false,rankdir=LR];' | Write-Output
	'node [fontsize=8,fontname=serif,shape=record,style=rounded];' | Write-Output
	'edge [dir=back,arrowsize=0.5,arrowtail=empty];' | Write-Output

	foreach ($ADSGroupPso in $ADSGroupPsoCol)
	{
		$Msg = 'Graphing "' + $ADSGroupPso.CanonicalName + '".'
		Write-Host $Msg
	
		$ADSGroupParentPso = $ADSGroupPso

        &$GraphNode $ADSGroupParentPso

    	foreach ($ADSGroupMbr in $ADSGroupParentPso.Member)
		{
        	foreach ($ADSGrpPso in $ADSGroupPsoCol)
			{
				if ($ADSGrpPso.DistinguishedName -eq $ADSGroupMbr.DistinguishedName)
				{
					$ADSGroupChildPso = $ADSGrpPso
					
					'ADObj_{0} -> ADObj_{1};' -f [Array]::indexOf($ADSGroupPsoCol,$ADSGroupParentPso), [Array]::indexOf($ADSGroupPsoCol,$ADSGroupChildPso) |
					Write-Output
					
					break
				}
			}
		}
	}

	'}'	| Write-Output
}

function Draw-ADSecurityGroupNesting
{
	[CmdletBinding()]
	
	param
	(
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[Alias("DN")]
		[String[]] $ADObjDN_List = @((Get-ADDomain).UsersContainer),
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
        [Alias("SC")]
		[Microsoft.ActiveDirectory.Management.ADSearchScope] $Scope = "Subtree",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("3268","3269")]
		[Alias("Port")]
		[String] $ADGCPort = "3268",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("MemberOf","Member")]
		[Alias("By")]
		[String] $Property = "MemberOf",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("Domain","Forest")]
		[String] $Mode = "Domain",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateSet("ASCII","UTF8")]
		[Alias("CS")]
		[String] $Charset = "UTF8",
		
		[Parameter(ValueFromPipeline = $false, Mandatory = $false)]
		[ValidateScript({Test-Path $_ -PathType 'Container'})]
		[Alias("Dir")]
		[String] $Directory = "$env:USERPROFILE\Desktop"
    )
	
	if ($Directory -ne "$env:USERPROFILE\Desktop")
	{
		$Directory = Remove-LastBackSlash $Directory
	}
	
	$global:VizPath = $null
	
	$Ending = {
	
		param
		(
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[String] $ADObjDN,
			
			[Parameter(ValueFromPipeline = $false, Mandatory = $true)]
			[String] $Directory
		)

		$Msg = "`n" + 'Ending'
		Write-Host $Msg -Foregroundcolor Green -Backgroundcolor DarkBlue

		$global:ADSGroupPsoDupCol = $null

		$ADObj = Get-ADObject $ADObjDN -Properties name,canonicalName -Server $DomainGC
		$Index = ($ADObj.CanonicalName).IndexOf("/")
		$ADObjDomainFQDN = ($ADObj.CanonicalName).Substring(0,$Index)

		$FileName = $ADObj.Name + "_" + $ADObjDomainFQDN + "_" + $Property + "_" + $Mode

		$global:VizPath = $Directory + '\' + $FileName + '.viz'

		$Output | Out-File -FilePath $VizPath -Encoding $Charset

		$Msg = "`n" + 'GraphViz file to use with dot.exe or gvedit.exe is "' + $VizPath + '".'
		Write-Host $Msg -Foregroundcolor White
	}
	
	$Msg = "`n" + 'Initialization'
	Write-Host $Msg -Foregroundcolor Green -Backgroundcolor DarkBlue
	
	:nextDN foreach ($ADObjDN in $ADObjDN_List)
	{
		try
		{	
			Search-ADObjQuicker $ADObjDN | Out-Null
		}
		catch
		{
			$Msg = 'Non terminating error on "' + $ADObjDN + '".'
			$Msg += "`n" + 'Object does not exist or DistinguishedName is mistyped, it will be ignored in processing step.'
			Write-Warning $Msg
			
			continue nextDN
		}
		
		$Index = $ADObjDN.IndexOf("DC=")
		$ADDomainFQDN = ($ADObjDN.Substring($Index)).Replace("DC=","").Replace(",",".")
		
		$ADDomainFQDN_List = $ADDomainFQDN_List + @($ADDomainFQDN)
		
		$ADObjDN_List_Checked = $ADObjDN_List_Checked + @($ADObjDN)
	}
	
	if ($ADObjDN_List_Checked)
	{
		$ADDomainFQDN_List = $ADDomainFQDN_List | Select -Unique
	
		if ($Mode -eq "Domain")
		{
			$ADGCPsoCol = List-ADGCOnePerDomainReachable -FQDN $ADDomainFQDN_List -Port $ADGCPort -CS $Charset
		}
		else
		{
			$ADDomainFQDN_List = (Get-ADForest).Domains
				
			$ADGCPsoCol = List-ADGCOnePerDomainReachable -FQDN $ADDomainFQDN_List -Port $ADGCPort -CS $Charset
		}
		
		if (!$ADGCPsoCol)
		{
			$Msg = 'Script aborted.'
			$Msg += "`n" + 'Be sure that at least one Global Catalog in each Domain responds and that Global Catalog list is well formed.'
			$Msg += "`n" + 'You should flush and rediscover Global Catalog list.'
			Write-Warning $Msg
		
			return
		}
		
		:nextObj foreach ($ADObjDN in $ADObjDN_List_Checked)
		{
			$global:VizPath = $null
		
			$DomainGCPso = Get-ADGCOneForADObj -DN $ADObjDN -GC $ADGCPsoCol
			$DomainGC = $DomainGCPso.HostName + ":" + $DomainGCPso.Port
		
			if ($Mode -eq "Domain")
			{
				$ADForestGCPsoCol = @($DomainGCPso)
			}
			else
			{
				$ADForestGCPsoCol = $ADGCPsoCol
			}
			

			$Msg = "`n" + 'Processing'
			Write-Host $Msg -Foregroundcolor Green -Backgroundcolor DarkBlue
			$Msg = "`n" + 'Exploring security group nesting for "'+ $ADObjDN + '" using property "' + $Property + '":' + "`n"
			Write-Host $Msg -Foregroundcolor White
	
			switch ($Property)
			{
				"MemberOf"
				{
					$ADSGroupPsoCol = List-ADSecurityGroupMemberOf -DN $ADObjDN -SC $Scope -GC $ADForestGCPsoCol
				
					if ($ADSGroupPsoCol)
					{
						$Output = Graph-ADSecurityGroupMemberOf $ADSGroupPsoCol -DN $ADObjDN -SC $Scope -Mode $Mode -CS $Charset
						
						&$Ending $ADObjDN $Directory
					}
				}
		
				"Member"
				{
					$Msg = 'Request may time out or exceed size limit if one or more groups contain more than 5 000 members.'
					$Msg += "`n" + 'If so, prefer drawing using memberOf Attribute (-By MemberOf).' + "`n"
					Write-Host $Msg -ForegroundColor Magenta
				
					$ADSGroupPsoCol = List-ADSecurityGroupMember -DN $ADObjDN -SC $Scope -GC $ADForestGCPsoCol
				
					if ($ADSGroupPsoCol)
					{
						$Output = Graph-ADSecurityGroupMember $ADSGroupPsoCol -DN $ADObjDN -SC $Scope -Mode $Mode -CS $Charset
						
						&$Ending $ADObjDN $Directory
					}
				}
			}
			
			if (!$VizPath)
			{
				$Msg = 'Found no valid object for ' + $Property + ' property exploration.'
				Write-Warning $Msg
			}
		}
	}
	else
	{
		$Msg = 'Script aborted.'
		$Msg += "`n" + 'Found no valid DistinguishedName.'
		Write-Warning $Msg
		
		return		
	}
}
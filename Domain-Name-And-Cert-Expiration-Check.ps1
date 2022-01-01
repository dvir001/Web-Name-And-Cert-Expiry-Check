<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.191
	 Created on:   	29/12/2021 13:44
	 Created by:   	Dvir
	 Organization: 	
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		This script will get whois and cert expiry data, and save the data into HTML\XLSX.
#>

$domainlist = @(
	@{
		site = "google.com"
		url  = "https://google.com/"
	}
	@{
		site = "microsoft.com"
		url  = "https://microsoft.com/"
	}
	@{
		site = "office.com"
		url  = "https://office.com/"
	}
	@{
		site = "amazon.co.uk"
		url  = "https://amazon.co.uk/"
	}
	@{
		site = "w3.org"
		url  = "https://www.w3.org/"
	}
	@{
		site = "walla.co.il"
		url  = "https://walla.co.il/"
	}
)

$formatTime = "yyyy-MM-dd" <# Time format for report #>
$workDir = "$env:windir\Temp" <# Whois install folder #>
$htmlDir = "$env:SystemDrive\Temp" <# Report file folder #>
$htmlfile = "DomainExpiration.html" <# Report file name #>

#$xlsxDir = "$env:SystemDrive\Temp"
#$xlsxfile = "DomainExpiration.xlsx"

function Invoke-DomainCheck
{
	if (!(Test-Path -Path $htmlDir)) { New-Item -Path $htmlDir -ItemType "directory" -Force -Verbose:$false *>$null } <# Lookup if the HTML folder is there and create #>
	if (Test-Path -Path "$htmlDir\$htmlfile") { Remove-Item -Path "$htmlDir\$htmlfile" -Force } <# Clean before running #>
	#if (!(Test-Path -Path $xlsxDir)) { New-Item -Path "$xlsxDir" -ItemType "directory" -Force -Verbose:$false *>$null } <# Lookup if the XLSX folder is there and create #>
	#if (Test-Path -Path "$xlsxDir\$xlsxfile") { Remove-Item -Path "$xlsxDir\$xlsxfile" -Force } <# Clean before running #>
	
	# Grab today date in the formats required
	try { $todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	catch { [datetime]$todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	
	Write-Output "Today: $todayTime"
	$resultArray = @()
	
	foreach ($domain in $domainlist)
	{
		$domainSite = "$($domain.site)" <# Domain site short #>
		$url = "$($domain.url)" <# Domain url short #>
		
		# Run the qurry to get website information
		$domainNameExpirationCommand = Invoke-Process -FilePath "$workDir\whois.exe" -ArgumentList $domainSite
		
		$domainNameExpirationInfo = $null
		$domainNameExpirationDate = $null
		
		$infoPatterns = "Expiration Date:", "validity:", "Expiry date:", "Expiry"
		foreach ($infoPattern in $infoPatterns)
		{
			$domainNameExpirationInfo = $domainNameExpirationCommand -split "`n" | Select-String -Pattern $infoPattern -AllMatches | Select-Object -Unique <# Try to find a pattern #>
			$domainNameExpirationInfo = $domainNameExpirationInfo -replace 'Registrar Registration Expiration Date:', '' -replace 'validity:', '' -replace 'Expiry date:', '' -replace ' ', '' <# Remove text #>
			if ($domainNameExpirationInfo -ne $null) { <#Write-Output "$infoPattern" ;#> break } <# Stop the loop when found a line with exipry info #>
		} <# Clean the whois to get only the expiry date #>
		
		#Write-Output "$domainSite domainNameExpirationInfo: $domainNameExpirationInfo"
		if ($domainNameExpirationInfo -like "*T*")
		{
			$domainNameExpirationInfo = $domainNameExpirationInfo.split() -ne ""
			$domainNameExpirationInfo = $domainNameExpirationInfo.split('T')[0]
			$tryTimeFormats = $formatTime, "yyyy-MM-dd", "dd-MM-yyyy", "dd-MMM-yyyy"
			foreach ($tryTimeFormat in $tryTimeFormats)
			{
				#Write-Output "tryTimeFormat: $tryTimeFormat"
				$error.clear()
				try { $date2 = [datetime]::ParseExact($domainNameExpirationInfo, $tryTimeFormat, $null) }
				catch { }
				if (!$error) { break }
			}
			#Write-Output "$domainSite domainNameExpirationInfo: $domainNameExpirationInfo"
			$domainNameExpirationDate = Get-Date $($date2) -Format $formatTime
		} <# Remove hours and keep only date #>
		else
		{
			#Write-Output "domainNameExpirationInfo: $domainNameExpirationInfo"
			try { $domainNameExpirationDate = Get-Date $($domainNameExpirationInfo) -Format $formatTime }
			catch { $domainNameExpirationDate = $todayTime }
			<# In case there is no date found, reset to today #>
		}
		
		$domainNameExpirationDateLeft = New-Timespan -Start $todayTime -End $domainNameExpirationDate
		$domainNameExpirationDateLeft = $domainNameExpirationDateLeft -replace ".00:00:00", "" -replace "00:00:00", "0" <# Remove hours and keep only date #>
		
		#Write-Output "domainNameExpirationInfo: $domainNameExpirationInfo"
				
		# Run the qurry to get cert information
		$certFail = $false
		try
		{
			[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
			$req = [Net.HttpWebRequest]::Create($url)
			$req.GetResponse() | Out-Null
			$domainCertExpirationDate = Get-Date $($req.ServicePoint.Certificate.GetExpirationDateString()) -Format $formatTime
		}
		catch { $certFail = $true }
		
		if ($certFail)
		{
			try
			{
				$servicePoint = [System.Net.ServicePointManager]::FindServicePoint("$url")
				$domainCertExpirationDate = Get-Date $($servicePoint.Certificate.GetExpirationDateString()) -Format $formatTime
			}
			catch { $domainCertExpirationDate = $todayTime ; $certFail = $true }
		}
		
		# Fix the domain cert time format
		if (![string]::IsNullOrEmpty($domainCertExpirationDate)) <# In case no cert found #>
		{
			#Write-Output "Test domainCertExpirationDate: $domainCertExpirationDate"
			try { $domainCertExpirationDateLeft = New-Timespan -Start $todayTime -End $domainCertExpirationDate <# Get the domainCertExpirationdateLeft (EU Format) #> }
			catch { $domainCertExpirationDateLeft = New-Timespan -Start ([datetime]$todayTime) -End ([datetime]$domainCertExpirationDate) <# Get the domainCertExpirationdateLeft (US Format) #> }
			$domainCertExpirationDateLeft = $domainCertExpirationDateLeft -replace ".00:00:00", "" -replace "00:00:00", "0" <# Remove hours and keep only date #>
			#Write-Output "Test domainCertExpirationDateLeft: $domainCertExpirationDateLeft"
		}
		else { $domainCertExpirationDateLeft = "0" }
		
		#Write-Output "$domainSite, Name- $domainNameExpirationDateLeft($domainNameExpirationDate), Cert- $domainCertExpirationDateLeft($domainCertExpirationDate)"
		
		# Export result to HTML
		$ExportResult = @{ 'Domain' = $domainSite; 'URL' = $url; 'Name Expiration Date' = $domainNameExpirationDate; 'Name Expiration Left' = $domainNameExpirationDateLeft; 'Cert Expiration Date' = $domainCertExpirationDate; 'Cert Expiration Left' = $domainCertExpirationDateLeft }
		$ExportResults = New-Object PSObject -Property $ExportResult
		$ExportResults | Select-Object 'Domain', 'URL', 'Name Expiration Date', 'Name Expiration Left', 'Cert Expiration Date', 'Cert Expiration Left'
		#$ExportResults | Select-Object 'Domain', 'URL', 'Name Expiration Date', 'Name Expiration Left', 'Cert Expiration Date', 'Cert Expiration Left' | Export-XLSX -Path "$xlsxDir\$xlsxfile" -Force -Append
		$resultArray += $ExportResults | Select-Object 'Domain', 'URL', 'Name Expiration Date', 'Name Expiration Left', 'Cert Expiration Date', 'Cert Expiration Left'
	}
	$resultArray | Out-HtmlView -FilePath "$htmlDir\$htmlfile" -DisablePaging -PreventShowHTML <# Export to HTML #>
}

function Invoke-Program-Install
{
	$installSource = "https://download.sysinternals.com/files/WhoIs.zip"
	$installZip = "WhoIs.zip"
	
	if ((Test-Path -Path "$workDir\$installZip") -ne "True") <# Lookup if the exe is there #>
	{
		if (Get-Command 'Invoke-Webrequest') { Invoke-WebRequest $installSource -OutFile "$workDir\$installZip" | Wait-Job }
		else
		{
			$WebClient = New-Object System.Net.WebClient
			$webclient.DownloadFile($installSource, "$workDir\$installZip")
		}
	}
	
	Expand-Archive -Path "$workDir\$installZip" -DestinationPath "$workDir" -Force -ErrorAction SilentlyContinue
	#if (Test-Path -Path "$workDir\$installZip") { Remove-Item -Path "$workDir\$installZip" -Force }
}

function Invoke-Process
{
	[CmdletBinding(SupportsShouldProcess)]
	param
	(
		[Parameter(Mandatory)]
		[ValidateNotNullOrEmpty()]
		[string]$FilePath,
		[Parameter()]
		[ValidateNotNullOrEmpty()]
		[string]$ArgumentList
	)
	
	$ErrorActionPreference = 'Stop'
	
	try
	{
		$pinfo = New-Object System.Diagnostics.ProcessStartInfo
		$pinfo.FileName = $FilePath
		$pinfo.RedirectStandardError = $true
		$pinfo.RedirectStandardOutput = $true
		$pinfo.UseShellExecute = $false
		$pinfo.WindowStyle = 'Hidden'
		$pinfo.CreateNoWindow = $true
		$pinfo.Arguments = $ArgumentList
		$p = New-Object System.Diagnostics.Process
		$p.StartInfo = $pinfo
		$p.Start() | Out-Null
		$result = [pscustomobject]@{
			Title = ($MyInvocation.MyCommand).Name
			Command = $FilePath
			Arguments = $ArgumentList
			StdOut = $p.StandardOutput.ReadToEnd()
			StdErr = $p.StandardError.ReadToEnd()
			ExitCode = $p.ExitCode
		}
		$p.WaitForExit()
		
		return $result.StdOut; break
	}
	catch { exit }
}

function Invoke-Module-Install
{
	param (
		[Parameter(Mandatory = $true)]
		[array]$modules
	)
	foreach ($module in $modules)
	{
		try
		{
			#Write-Output "Importing module '$module'"
			Import-Module $module -ErrorAction Stop
		}
		catch
		{
			Write-Output "Could not find '$module' module, installing..."
			Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Verbose:$false *>$null
			Install-Module -Name $module -AllowClobber -Scope AllUsers -force
			Import-Module $module -ErrorAction Stop
			#Write-Output "Importing module '$module'"
		}
	}
}

Invoke-Module-Install -modules "PSWriteHTML"#, "PSExcel"
Invoke-Program-Install
Invoke-DomainCheck

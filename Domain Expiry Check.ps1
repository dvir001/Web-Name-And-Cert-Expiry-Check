<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2021 v5.8.191
	 Created on:   	04/07/2022 16:10
	 Created by:   	Dvir F
	 Organization: 	ITMS
	 Filename:     	
	===========================================================================
	.DESCRIPTION
		A description of the file.
#>

# If not using a json file for SMTP please provide the config here:
param (
	[Parameter(Mandatory = $false)]
	[String]$smtpUser = "apikey",
	[Parameter(Mandatory = $false)]
	[String]$smtpPass = "PASS",
	[Parameter(Mandatory = $false)]
	[String]$smtpServer = "smtp.sendgrid.net",
	[Parameter(Mandatory = $false)]
	[String]$smtpPort = "587",
	[Parameter(Mandatory = $false)]
	[Bool]$smtpUseSSL = $true,
	[Parameter(Mandatory = $false)]
	[String]$smtpFrom = "ExpiryReport@corp.com",
	[Parameter(Mandatory = $false)]
	[String]$smtpTo = "helpdesk+ExpiryReport@corp.com",
	[Parameter(Mandatory = $false)]
	[String]$smtpSubject = "Cert And Domain Expiry Report"
)

# If you want, you can use a json file instead for the smtp, leave it empty if using the params:
#$smtpFile = "$PSScriptRoot\smtp.json" # Link or file location for SMTP json

#===========================================================================

$inputFile = "https://XXXXXXXXX.sharepoint.com/:x:/s/Automation/EQx1SzOpO5hPt7YseUVjTC8Bi-xxxxxxxxx_ysTC7Q?e=rbkymn&download=1" # Link or file location for Expiry list
$inputFile = "C:\Temp\Domains.xlsx"

$formatTime = "yyyy-MM-dd" <# Time format for report #>

$warningDays = "30"
$outputFilesFolder = "C:\Temp"
$domainFileHTML = "DomainExpiry.html" <# Domain HTML Report file name #>
$certFileHTML = "CertExpiry.html" <# Cert HTML Report file name #>
$domainFileXLSX = "DomainExpiry.xlsx" <# Domain XLSX Report file name #>
$certFileXLSX = "CertExpiry.xlsx" <# Cert XLSX Report file name #>
$domainExpiryWarningJSON = "DomainExpiryWarning.json" <# Domain JSON Report file name #>
$certExpiryWarningJSON = "CertExpiryWarning.json" <# Cert JSON Report file name #>

function Get-DomainExpiry
{
	if (Test-Path -Path "$outputFilesFolder\$domainFileHTML") { Remove-Item -Path "$outputFilesFolder\$domainFileHTML" -Force } <# Clean before running #>
	if (Test-Path -Path "$outputFilesFolder\$domainFileXLSX") { Remove-Item -Path "$outputFilesFolder\$domainFileXLSX" -Force } <# Clean before running #>
	
	# Grab today date in the formats required
	try { $todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	catch { [datetime]$todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	
	Write-Output "Today: $todayTime"
	$resultArray = @()
	$warningArray = @()
	
	if (!([string]::IsNullOrEmpty($inputFile))) # if the domain names file isnt missing, run domain names testing
	{
		if ($inputFile -like "*://*") # if link
		{
			Write-Output "Downloading to `"$env:windir\Temp\Domains.xlsx`""
			try { Invoke-WebRequest $inputFile -OutFile "$env:windir\Temp\Domains.xlsx" -Verbose -ErrorAction Ignore | Wait-Job }
			catch [System.Net.WebException] { Write-Output "Link Broken / No network."; exit }
			
			$inputFile = "$env:windir\Temp\Domains.xlsx"
		}
		elseif ($inputFile -like "*.xlsx*") { Write-Output "Using `"$inputFile`"" }
		else { Write-Verbose "Cant find a link for or an .xlsx or a `".xlsx`" file."; exit }
		
		$importXLSX = Import-XLSX -Path $inputFile -Header "Domain", "Cert"
		$inputFile = $([Array]$importXLSX.Domain)
	}
	
	foreach ($item in $inputFile)
	{
		if (!([string]::IsNullOrEmpty($item)))
		{
			$domainExpiryCommand = Invoke-Process -FilePath "$env:windir\Temp\whois.exe" -ArgumentList $item
			
			$domainExpiryInfo = $null
			$domainExpiryDate = $null
			
			$infoPatterns = "Expiration Date:", "validity:", "Expiry date:", "Expiry"
			foreach ($infoPattern in $infoPatterns)
			{
				$domainExpiryInfo = $domainExpiryCommand -split "`n" | Select-String -Pattern $infoPattern -AllMatches | Select-Object -Unique <# Try to find a pattern #>
				$domainExpiryInfo = $domainExpiryInfo -replace 'Registrar Registration Expiration Date:', '' -replace 'validity:', '' -replace 'Expiry date:', '' -replace ' ', '' <# Remove text #>
				if ($domainExpiryInfo -ne $null) { <#Write-Output "$infoPattern" ;#> break } <# Stop the loop when found a line with exipry info #>
			} <# Clean the whois to get only the expiry date #>
			
			#Write-Output "$item domainExpiryInfo: $domainExpiryInfo"
			if ($domainExpiryInfo -like "*T*")
			{
				$domainExpiryInfo = $domainExpiryInfo.split() -ne ""
				$domainExpiryInfo = $domainExpiryInfo.split('T')[0]
				$tryTimeFormats = $formatTime, "yyyy-MM-dd", "dd-MM-yyyy", "dd-MMM-yyyy"
				foreach ($tryTimeFormat in $tryTimeFormats)
				{
					#Write-Output "tryTimeFormat: $tryTimeFormat"
					$error.clear()
					try { $date2 = [datetime]::ParseExact($domainExpiryInfo, $tryTimeFormat, $null) }
					catch { }
					if (!$error) { break }
				}
				#Write-Output "$item domainExpiryInfo: $domainExpiryInfo"
				$domainExpiryDate = Get-Date $($date2) -Format $formatTime
			} <# Remove hours and keep only date #>
			else
			{
				#Write-Output "domainExpiryInfo: $domainExpiryInfo"
				try { $domainExpiryDate = Get-Date $($domainExpiryInfo) -Format $formatTime }
				catch { $domainExpiryDate = $todayTime }
			<# In case there is no date found, reset to today #>
			}
			
			$domainExpiryLeft = New-Timespan -Start $todayTime -End $domainExpiryDate
			$domainExpiryLeft = $domainExpiryLeft -replace ".00:00:00", "" -replace "00:00:00", "0" <# Remove hours and keep only date #>
			
			#Write-Output "domainExpiryInfo: $domainExpiryInfo"
			
			#Write-Output "$item, Name- $domainExpiryLeft($domainExpiryDate), Cert- $certExpiryLeft($certExpiryDate)"
			
			if ([int]$domainExpiryLeft -le [int]$warningDays) <# feed domain into the warning list if its under or equal to $warningDays #>
			{
				$warningArray = @(
					@{
						Domain		   = $item
						DomainExpiryLeft = $domainExpiryLeft
					} <# warning into json format #>
				)
				[Array]$domainWarningList += $warningArray <# feed all warnings into array for json file #>
			}
			
			# Export results
			$exportResult = @{ 'Domain' = $item; 'Expiry Date' = $domainExpiryDate; 'Expiry Left' = $domainExpiryLeft }
			$exportResults = New-Object PSObject -Property $exportResult
			$exportResults | Select-Object 'Domain', 'Expiry Date', 'Expiry Left' <# Show in output #>
			$exportResults | Select-Object 'Domain', 'Expiry Date', 'Expiry Left' | Export-XLSX -Path "$outputFilesFolder\$domainFileXLSX" -Force -Append <# Export to XLSX #>
			$resultArray += $exportResults | Select-Object 'Domain', 'Expiry Date', 'Expiry Left' <# Feed the data into array for the HTML file #>
		}
	}
	
	$domainWarningList | ConvertTo-Json | Out-File "$outputFilesFolder\$domainExpiryWarningJSON" -Force
	$resultArray | Out-HtmlView -FilePath "$outputFilesFolder\$domainFileHTML" -DisablePaging -PreventShowHTML <# Export to HTML #>
}

function Get-CertExpiry
{
	if (Test-Path -Path "$outputFilesFolder\$certFileHTML") { Remove-Item -Path "$outputFilesFolder\$certFileHTML" -Force } <# Clean before running #>
	if (Test-Path -Path "$outputFilesFolder\$certFileXLSX") { Remove-Item -Path "$outputFilesFolder\$certFileXLSX" -Force } <# Clean before running #>
	
	# Grab today date in the formats required
	try { $todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	catch { [datetime]$todayTime = Get-Date -format $formatTime } <# Try to handle time issues on US or EU formats #>
	
	Write-Output "Today: $todayTime"
	$resultArray = @()
	$warningArray = @()
	
	if (!([string]::IsNullOrEmpty($inputFile))) # if the domain names file isnt missing, run domain names testing
	{
		if ($inputFile -like "*://*") # if link
		{
			Write-Output "Downloading to `"$env:windir\Temp\Domains.xlsx`""
			try { Invoke-WebRequest $inputFile -OutFile "$env:windir\Temp\Domains.xlsx" -Verbose -ErrorAction Ignore | Wait-Job }
			catch [System.Net.WebException] { Write-Output "Link Broken / No network."; exit }
			
			$inputFile = "$env:windir\Temp\Domains.xlsx"
		}
		elseif ($inputFile -like "*.xlsx*") { Write-Output "Using `"$inputFile`"" }
		else { Write-Verbose "Cant find a link for or an .xlsx or a `".xlsx`" file."; exit }
		
		$importXLSX = Import-XLSX -Path $inputFile -Header "Domain", "Cert"
		$inputFile = $([Array]$importXLSX.Cert)
	}
	
	foreach ($item in $inputFile)
	{
		if (!([string]::IsNullOrEmpty($item)))
		{
			# Run the qurry to get cert information
			$certFail = $false
			
			try
			{
				[Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
				$req = [Net.HttpWebRequest]::Create($item)
				$req.GetResponse() | Out-Null
				$certExpiryDate = Get-Date $($req.ServicePoint.Certificate.GetExpirationDateString()) -Format $formatTime
			}
			catch { $certFail = $true }
			
			if ($certFail)
			{
				try
				{
					$servicePoint = [System.Net.ServicePointManager]::FindServicePoint("$item")
					$certExpiryDate = Get-Date $($servicePoint.Certificate.GetExpirationDateString()) -Format $formatTime
				}
				catch { $certExpiryDate = $todayTime; $certFail = $true }
			}
			
			# Fix the domain cert time format
			if (![string]::IsNullOrEmpty($certExpiryDate)) <# In case no cert found #>
			{
				#Write-Output "Test domainCertExpiryDate: $certExpiryDate"
				try { $certExpiryLeft = New-Timespan -Start $todayTime -End $certExpiryDate <# Get the domainCertExpirydateLeft (EU Format) #> }
				catch { $certExpiryLeft = New-Timespan -Start ([datetime]$todayTime) -End ([datetime]$certExpiryDate) <# Get the domainCertExpirydateLeft (US Format) #> }
				$certExpiryLeft = $certExpiryLeft -replace ".00:00:00", "" -replace "00:00:00", "0" <# Remove hours and keep only date #>
				#Write-Output "Test domainCertExpiryDateLeft: $certExpiryLeft"
			}
			else { $certExpiryLeft = "0" }
			
			if ([int]$certExpiryLeft -le [int]$warningDays) <# feed domain into the warning list if its under or equal to $warningDays #>
			{
				$warningArray = @(
					@{
						Cert		   = $item
						CertExpiryLeft = $certExpiryLeft
					} <# warning into json format #>
				)
				[Array]$certWarningList += $warningArray <# feed all warnings into array for json file #>
			}
			
			# Export results
			$exportResult = @{ 'Cert' = $item; 'Expiry Date' = $certExpiryDate; 'Expiry Left' = $certExpiryLeft }
			$exportResults = New-Object PSObject -Property $exportResult
			$exportResults | Select-Object 'Cert', 'Expiry Date', 'Expiry Left' <# Show in output #>
			$exportResults | Select-Object 'Cert', 'Expiry Date', 'Expiry Left' | Export-XLSX -Path "$outputFilesFolder\$certFileXLSX" -Force -Append <# Export to XLSX #>
			$resultArray += $exportResults | Select-Object 'Cert', 'Expiry Date', 'Expiry Left' <# Feed the data into array for the HTML file #>
		}
	}
	
	$certWarningList | ConvertTo-Json | Out-File "$outputFilesFolder\$CertExpiryWarningJSON" -Force
	$resultArray | Out-HtmlView -FilePath "$outputFilesFolder\$certFileHTML" -DisablePaging -PreventShowHTML <# Export to HTML #>
}

function Send-CustomMailMessage
{
	if (!([string]::IsNullOrEmpty($smtpFile))) # if the domain names file isnt missing, run domain names testing
	{
		if ($smtpFile -like "*://*") # if link
		{
			Write-Output "Downloading to `"$env:windir\Temp\smtp.json`""
			try { Invoke-WebRequest $smtpFile -OutFile "$env:windir\Temp\smtp.json" -Verbose -ErrorAction Ignore | Wait-Job }
			catch [System.Net.WebException] { Write-Output "Link Broken / No network."; exit }
			
			$smtpFile = "$env:windir\Temp\smtp.json"
		}
		elseif ($smtpFile -like "*.json*") { Write-Output "Using `"$smtpFile`"" }
		
		$smtpConfig = (Get-Content "$smtpFile" -Raw) | ConvertFrom-Json <# Grab config from JSON file #>
		$sstr = ConvertTo-SecureString -string $($smtpConfig.Login_Pass) -AsPlainText -Force; $smtpCredential = New-Object System.Management.Automation.PSCredential -argumentlist $($smtpConfig.Login_User), $sstr # Convert Pass	
		
		# SSL true or false
		If ($smtpConfig.SSL -like "*true*") { $smtpUseSSL = $true }
		If ($smtpConfig.SSL -like "*false*") { $smtpUseSSL = $false }
		
		$smtpServer = $smtpConfig.Server
		$smtpPort = $smtpConfig.Port
		$smtpFrom = $smtpConfig.From
		$smtpTo = $smtpConfig.To
		$smtpSubject = $smtpConfig.Subject
	}
	else
	{
		Write-Output "Using provided SMTP Settings"
		$sstr = ConvertTo-SecureString -string $smtpPass -AsPlainText -Force; $smtpCredential = New-Object System.Management.Automation.PSCredential -argumentlist $smtpUser, $sstr # Convert Pass
	}
	
	# Add files if it can find the files
	$attachmentsFiles = "$outputFilesFolder\$domainFileHTML", "$outputFilesFolder\$domainFileXLSX", "$outputFilesFolder\$certFileHTML", "$outputFilesFolder\$certFileXLSX"
	foreach ($attachmentsFile in $attachmentsFiles) { if (Test-Path $attachmentsFile) { [Array]$smtpAttachments += $attachmentsFile } }
	
	# Add the warnings from JSON files.
	$domainWarningList = (Get-Content "$outputFilesFolder\$domainExpiryWarningJSON" -Raw) | ConvertFrom-Json
	foreach ($domainWarning in $domainWarningList)
	{
		$domainWarningOutput += @"

`'$($domainWarning.Domain)`' - $($domainWarning.DomainExpiryLeft) Days left

"@
	}
	
	$certWarningList = (Get-Content "$outputFilesFolder\$CertExpiryWarningJSON" -Raw) | ConvertFrom-Json
	foreach ($certWarning in $certWarningList)
	{
		$certWarningNotHyper = $($certWarning.Cert).replace("://", " ")
		$certWarningOutput += @"

`'$certWarningNotHyper`' - $($certWarning.CertExpiryLeft) Days left

"@
	}
	
	$smtpBody = @"
Domain certs with less then $warningDays days:
$certWarningOutput

Domain names with less then $warningDays days:
$domainWarningOutput
"@
	
	Send-MailMessage -SmtpServer $smtpServer -Port $smtpPort -UseSsl:$smtpUseSSL -Credential $smtpCredential -From $smtpFrom -To $smtpTo -Subject $smtpSubject -Attachments $smtpAttachments -Body $smtpBody
}

function Invoke-Program-Install
{
	$installSource = "https://download.sysinternals.com/files/WhoIs.zip"
	$installZip = "WhoIs.zip"
	
	if (!(Test-Path -Path "$env:windir\Temp\$installZip")) <# Lookup if the exe is there #>
	{ Invoke-WebRequest $installSource -OutFile "$env:windir\Temp\$installZip" | Wait-Job }
	
	Expand-Archive -Path "$env:windir\Temp\$installZip" -DestinationPath "$env:windir\Temp" -Force -ErrorAction SilentlyContinue
	#if (Test-Path -Path "$env:windir\Temp\$installZip") { Remove-Item -Path "$env:windir\Temp\$installZip" -Force }
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

function Install-CustomModule
{
	param (
		[Parameter(Mandatory = $true)]
		[Array]$modules
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
			Install-Module -Name $module -Scope AllUsers -AllowClobber -Force
			Import-Module $module -ErrorAction Stop
			#Write-Output "Importing module '$module'"
		}
	}
}

Invoke-Program-Install
Install-CustomModule -modules "PSWriteHTML", "PSExcel"
#Get-DomainExpiry
#Get-CertExpiry
Send-CustomMailMessage

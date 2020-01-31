<#
.DESCRIPTION
Connects to Office 365 and Exchange Online then creates a online user and sets license
#>
# 
#
#
# Add -WhatIf to test

[CmdletBinding(ConfirmImpact="Medium",SupportsShouldProcess)]
Param(
    #[PSCredential]
    #$Credential=(Get-Credential),
    [Parameter(Mandatory)]
    [ValidateScript({
    Test-Path -LiteralPath $_
    })]
    [String]
    $CSVFilePath = 'D:\PowershellScripts\o365NewUsers\UserList.csv'
)

#region Parameters
#
$Domain = "@CMPY.com"


###########################################################
#Variables created for file locations
$path = Split-Path -parent "D:\PowershellScripts\o365NewUsers\*.*"
$logfile = $path + "\logfile.txt"
$i        = 0
$date     = Get-Date

#endregion Paramters

#region Functions
# 
# Functions
#

#Function set to generate random password which meets complexity requirements

function New-Password {

    $Upper = [char[]]([int][char]'A'..[int][char]'Z') | Get-Random -Count 2
    $Lower = [char[]]([int][char]'a'..[int][char]'z') | Get-Random -Count 2
    $Special = "!@#$%^&*-./,<>" -split "" | Where-Object {$_} | Get-Random -Count 2
    $Number = 0..9 | Get-Random -Count 2

    $Password = ($Upper, $Lower, $Special, $Number | Get-Random -count 9999) -join ''


    Write-Output (ConvertTo-SecureString -AsPlainText $Password -force)
}

#Function tto be created for validation
#Function Test-MSOlUser {
#Add code to check if user exists
#}

#Function to connect to Office 365
	function Connect-Office365
{
<#
    .DESCRIPTION
        Connect to different Office 365 Services using PowerShell function. Supports MFA.
        
    .EXAMPLE
		Description: Connect to Exchange Online and Azure AD V2 using Multi-Factor Authentication
        C:\PS> Connect-Office365 -Service Exchange, MSOnline -MFA

#>
	
	[OutputType()]
	[CmdletBinding(DefaultParameterSetName)]
	Param (
		[Parameter(Mandatory = $True, Position = 1)]
		[ValidateSet('AzureAD', 'Exchange', 'MSOnline')]
		[string[]]$Service,
		[Parameter(Mandatory = $False, Position = 2)]
		[Alias('SPOrgName')]
		[string]$SharePointOrganizationName,
		[Parameter(Mandatory = $False, Position = 3, ParameterSetName = 'Credential')]
		[PSCredential]$Credential,
		[Parameter(Mandatory = $False, Position = 3, ParameterSetName = 'MFA')]
		[Switch]$MFA
	)
	
	$getModuleSplat = @{
		ListAvailable = $True
		Verbose	      = $False
	}
	
	If ($MFA -ne $True)
	{
		Write-Verbose "Gathering PSCredentials object for non MFA sign on"
		$Credential = Get-Credential -Message "Please enter your Office 365 credentials"
	}
	
	ForEach ($Item in $PSBoundParameters.Service)
	{
		Write-Verbose "Attempting connection to $Item"
		Switch ($Item)
		{
			AzureAD {
				If ($null -eq (Get-Module @getModuleSplat -Name "AzureAD"))
				{
					Write-Error "AzureAD Module is not present!"
					continue
				}
				Else
				{
					If ($MFA -eq $True)
					{
						$Connect = Connect-AzureAD
						If ($null -ne $Connect)
						{
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
							{
								$host.ui.RawUI.WindowTitle += " - Connected To: AzureAD"
							}
							Else
							{
								$host.ui.RawUI.WindowTitle += " - AzureAD"
							}
						}
						
					}
					Else
					{
						$Connect = Connect-AzureAD -Credential $Credential
						If ($Null -ne $Connect)
						{
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
							{
								$host.ui.RawUI.WindowTitle += " - Connected To: AzureAD"
							}
							Else
							{
								$host.ui.RawUI.WindowTitle += " - AzureAD"
							}
						}
					}
				}
				continue
			}
			
			Exchange {
				If ($MFA -eq $True)
				{
					$getChildItemSplat = @{
						Path = "$Env:LOCALAPPDATA\Apps\2.0\*\CreateExoPSSession.ps1"
						Recurse = $true
						ErrorAction = 'SilentlyContinue'
						Verbose = $false
					}
					$MFAExchangeModule = ((Get-ChildItem @getChildItemSplat | Select-Object -ExpandProperty Target -First 1).Replace("CreateExoPSSession.ps1", ""))
					
					If ($null -eq $MFAExchangeModule)
					{
						Write-Error "The Exchange Online MFA Module was not found!
        https://docs.microsoft.com/en-us/powershell/exchange/exchange-online/connect-to-exchange-online-powershell/mfa-connect-to-exchange-online-powershell?view=exchange-ps"
						continue
					}
					Else
					{
						Write-Verbose "Importing Exchange MFA Module"
						. "$MFAExchangeModule\CreateExoPSSession.ps1"
						
						Write-Verbose "Connecting to Exchange Online"
						Connect-EXOPSSession
						If ($Null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" }))
						{
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
							{
								$host.ui.RawUI.WindowTitle += " - Connected To: Exchange"
							}
							Else
							{
								$host.ui.RawUI.WindowTitle += " - Exchange"
							}
						}
					}
				}
				Else
				{
					$newPSSessionSplat = @{
						ConfigurationName = 'Microsoft.Exchange'
						ConnectionUri	  = "https://ps.outlook.com/powershell/"
						Authentication    = 'Basic'
						Credential	      = $Credential
						AllowRedirection  = $true
					}
					$Session = New-PSSession @newPSSessionSplat
					Write-Verbose "Connecting to Exchange Online"
					Import-PSSession $Session -AllowClobber
					If ($Null -ne (Get-PSSession | Where-Object { $_.ConfigurationName -like "*Exchange*" }))
					{
						If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
						{
							$host.ui.RawUI.WindowTitle += " - Connected To: Exchange"
						}
						Else
						{
							$host.ui.RawUI.WindowTitle += " - Exchange"
						}
					}
					
				}
				continue
			}
			
			MSOnline {
				If ($null -eq (Get-Module @getModuleSplat -Name "MSOnline"))
				{
					Write-Error "MSOnline Module is not present!"
					continue
				}
				Else
				{
					Write-Verbose "Connecting to MSOnline"
					If ($MFA -eq $True)
					{
						Connect-MsolService
						If ($Null -ne (Get-MsolCompanyInformation -ErrorAction SilentlyContinue))
						{
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
							{
								$host.ui.RawUI.WindowTitle += " - Connected To: MSOnline"
							}
							Else
							{
								$host.ui.RawUI.WindowTitle += " - MSOnline"
							}
						}
					}
					Else
					{
						Connect-MsolService -Credential $Credential
						If ($Null -ne (Get-MsolCompanyInformation -ErrorAction SilentlyContinue))
						{
							If (($host.ui.RawUI.WindowTitle) -notlike "*Connected To:*")
							{
								$host.ui.RawUI.WindowTitle += " - Connected To: MSOnline"
							}
							Else
							{
								$host.ui.RawUI.WindowTitle += " - MSOnline"
							}
						}
					}
				}
				continue
			}
			
			Default { }
		}
	}
}
#Commands run to connect to Exchange Online and Office 365
Connect-Office365 -Service Exchange, MSOnline -MFA

#Import the CSV File
$Users = Import-Csv -LiteralPath $CSVFilePath -ErrorAction Stop


#Main Block

ForEach ($User in $Users) { 

    "" | Out-File $logfile -append
    "AD user creation logs for( " + $date + "): " | Out-File $logfile -append
    "--------------------------------------------" | Out-File $logfile -append
#Define FirstName and LastName variables prior to hastable  
    $FirstName = $User.'First Name'
    $LastName = $User.'Last Name'
    
    $MSOlParams = @{
        ImmutableId = "{0}.{1}" -f $FirstName.Substring(0, 1).ToLower(), $LastName.Replace(' ', '').Replace('-', '').ToLower()
        FirstName = $User.'First Name'
        LastName = $User.'Last Name'
        DisplayName = "{0} {1}" -f $FirstName, $LastName
        Office = $User.Office
        UserPrincipalName = "{0}.{1}{2}" -f $FirstName.ToLower(), $LastName.Replace(' ', '').Replace('-', '').ToLower(), $Domain
        Department = $User.Department
        Password = $User.Password
        LicenseAssignment = $User.AccountSkuId
        UsageLocation = $User.UsageLocation
        ForceChangePassword = $true
        Title = $User.Title
        StreetAddress = $User.StreetAddress
        City = $User.City
        State = $User.State
        Country = $User.Country
        ErrorAction = 'Stop'
    }
    
    if ($PSCmdlet.ShouldProcess("New-MSOlUser $($MSOlParams.UserPrincipalName)")) {
        try {
          $NewMsolUsersCreated += @(New-MSOlUser @MSOlParams)
          
          #Update log file with users created successfully
          $MSOlParams.DisplayName + " Created successfully" | Out-File $logfile -append
    
          # Output confirmation
          #Write-Host "Creating AD account for $UPN" -ForegroundColor Yellow
          Write-Host "Creating AD account for $($MSOlParams.UserPrincipalName)" -ForegroundColor Yellow
          Start-Sleep -Seconds 2
        } catch {
            $_ 
            # 
        }
        Write-Host "Office 365 account has been created for $($MSOlParams.UserPrincipalName)" -ForegroundColor Green
		 
    }
    #code to implement CustomAttribute10 based on users Office value - $User.Office ||||  Set-Mailbox User@domain.com -CustomAttribute10 <the new value>
    while (-not ($Mailbox = Get-Mailbox -Identity $MsolParams.UserPrincipalName -ErrorAction SilentlyContinue)) {
        Write-Host "Sleeping 30 seconds until Mailbox is available..." -ForegroundColor DarkCyan
        Start-Sleep -Seconds 30
    }
    
    Write-Host "Mailbox has been found for $($Mailbox.DisplayName)" -ForegroundColor Green
    Start-Sleep -Seconds 3
    try {
        Write-Host "Setting user based properties for $($MSOlParams.UserPrincipalName)" -ForegroundColor Yellow
        Start-Sleep -Seconds 5
        Set-Mailbox $MsolParams.UserPrincipalName -CustomAttribute10 $User.Office -ErrorAction Stop
        Set-Mailbox $MsolParams.UserPrincipalName -HiddenFromAddressListsEnabled $true -ErrorAction Stop
        Set-User -Identity $MsolParams.UserPrincipalName -Title $User.Title -ErrorAction Stop -WarningAction SilentlyContinue
        Set-User -Identity $MsolParams.UserPrincipalName -Company 'CMPY Group' -ErrorAction Stop -WarningAction SilentlyContinue
        Set-User -Identity $MsolParams.UserPrincipalName -StreetAddress $User.StreetAddress -ErrorAction Stop -WarningAction SilentlyContinue
        Set-User -Identity $MsolParams.UserPrincipalName -City $User.City -ErrorAction Stop -WarningAction SilentlyContinue
        Set-User -Identity $MsolParams.UserPrincipalName -State $User.State -ErrorAction Stop -WarningAction SilentlyContinue
        Set-User -Identity $MsolParams.UserPrincipalName -Country $User.Country -ErrorAction Stop -WarningAction SilentlyContinue
    }
    catch {
        Write-Warning "Failed to set CustomAttribute10 for $($MSOlParams.UserPrincipalName)"
    }
        Start-Sleep -Seconds 1
        Write-Host "User based properties are now set for: $($Mailbox.DisplayName)" -ForegroundColor Green
        Start-Sleep -Seconds 2

}
        Write-Host "New User confirmation" -ForegroundColor Cyan
        Start-Sleep -Seconds 2
        $NewMsolUsersCreated
<#
.SYNOPSIS
    Sync O365 licenses 
.VERSION 
    1.0
.AUTHOR 
    Bart Tacken - Avensus
.PREREQUISITES
    AzureDirectory Module
#>
#---------------------------------------------------------[Initialisations]--------------------------------------------------------
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string    
[string]$Today = Get-Date -Format "dd/MM/yyyy"  
Start-Transcript ('c:\windows\temp\' + "$Datestr" + '_Sync-Office365_E3.log') # Start logging  
$ErrorActionPreference = "Stop" # Try/Catch
  
Write-Warning "Importing AD module.."
If (!(Get-Module ActiveDirectory)) { Import-Module ActiveDirectory}

# Specific for request
$CloudStoreSubscriptionId = ""
$CloudStoreResourceId = "" 

# AD
$Office365LicenseSecurityGroup = "Office365_E3"
# O365 Licenses
$O365LicenseName = "ENTERPRISEPACK" # E3
$O365TenantName = ""
$CredPath = ".\O365Cred.xml"
$KeyFilePath = ".\O365.key"
$Key = Get-Content $KeyFilePath
$credXML = Import-Clixml $CredPath # Import encrypted credential file into XML format
$secureStringPWD = ConvertTo-SecureString -String $credXML.Password -Key $key
$Credentials = New-Object System.Management.Automation.PsCredential($credXML.UserName, $secureStringPWD) # Create PScredential Object

Connect-MsolService -Credential $Credentials

#-----------------------------------------------------------[Functions]------------------------------------------------------------


#-----------------------------------------------------------[Execution]------------------------------------------------------------
# Get members of Office security group
$Users = (Get-ADGroupMember -Identity $Office365LicenseSecurityGroup)
$UserCount = $Users.count

# Region Cloud Store config
$CloudStoreHeaders = New-Object "System.Collections.Generic.Dictionary[[String],[String]]"
$CloudStoreHeaders.Add("")
$CloudStoreHeaders.Add("Content-Type", "application/json")

$O365LicenseName = $O365LicenseName.trim() # fixes formatting issue
$FullLicenseName = $O365TenantName + ":" + $O365LicenseName
$SKUid = Get-MsolAccountSku | where {$_.accountSKUid -like "$FullLicenseName"}
$TotalLicenses = ($($SKUid.activeunits)) # - $($SKUid.ConsumedUnits))

$TotalLicenses = $TotalLicenses.ToString()
$licensesOverview = Get-MsolAccountSku
$licensesOverview = $licensesOverview | Out-String
$licensesOverview # Log to transcript

Write-Warning "There are [$TotalLicenses] licenses activated for license [$FullLicenseName]"
Write-Warning "and [$($SKUid.ConsumedUnits))] licenses IN USE."
Write-Warning "The security group [$Office365LicenseSecurityGroup] has [$UserCount] members"

# Check if the user account is not deviate more than 5% from the current active licenses

# 5% of Total activated licenses
$FivePercentCurrentActiveLic = (($TotalLicenses / 100) * 5)
$FivePercentCurrentActiveLic = [math]::Round($FivePercentCurrentActiveLic)

# Difference between membercount and current total O365 licenses
$Difference = $UserCount - $TotalLicenses

# To prevent a problem where there are (accidentally) added a lot of members to the security group 
# there is a security feature build in that will stop the script when this amount is bigger than 5% of total licenses
If (($Difference -gt $FivePercentCurrentActiveLic) -or ($Difference -lt (-$FivePercentCurrentActiveLic))) {
    Write-Warning "The difference between current total O365 licenses for [$FullLicenseName]"
    Write-Warning "and members of security group [$Office365LicenseSecurityGroup] is [$Difference] and"
    Write-Warning "bigger or less as 5% of current total O365 licenses ([$FivePercentCurrentActiveLic])."
    Write-Warning "This is unusual behaviour, script will abort now!"
    exit # Exit script
}

Write-Warning "The difference between current total O365 licenses for [$FullLicenseName]"
Write-Warning "and members of security group [$Office365LicenseSecurityGroup] is [$Difference] and"
Write-Warning "therefore less than 5% of current total O365 licenses ([$FivePercentCurrentActiveLic])."
Write-Warning "This is normal behaviour, script will continue to change licenses"

# Added security to prevent change of licenses when the usercount is zero or $null
Try {
    If (($UserCount -ne 0) -or ($UserCount -ne $null)) {

        # Composing body for updating license amount
        $body = "{
        `n	`"type`": `"CHANGE`",
        `n	`"subscriptionId`": `"$CloudStoreSubscriptionId`",
        `n	`"paymentMethodId`": `"0`",
        `n	`"resources`": [{
        `n		`"resourceId`": `"$CloudStoreResourceId`",
        `n		`"amount`": `"$UserCount`"
        `n	}]
        `n}"

        $response = Invoke-RestMethod '<URL>' -Method 'POST' -Headers $CloudStoreHeaders -Body $body -ErrorAction Stop
        #$response | ConvertTo-Json

        If ($Difference -gt 0) {Write-Warning "Web request for upgrading licenses with [$Difference] is processed succesfully!"}
        If ($Difference -lt 0) {Write-Warning "Web request for downgrading licenses with [$Difference] is processed succesfully!"}
                
        $OrderID = $response.orderId
        Write-Host "Order ID: [$OrderID]"
    } # End If
          
} catch {
    # Catch all errors and sent an email to monitor the script
    Write-Warning "Something went wrong! Processing code in catch block!"
    $ErrorMessage = $Error[0]
    $Error[0] # Log error in transcript
    # todo sent email
}
Stop-Transcript

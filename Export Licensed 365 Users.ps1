# Author: Chris Dunphy
# Notes: Run this in PowerShell 5.1, the one built-in with Windows. You need the MSOnline module installed in your PowerShell.

# First, you need to connect to the MSOnline service
Connect-MsolService

# Read the license name mapping from a CSV file
$licenseNameMapping = @{}
Import-Csv -Path ".\LicenseMaps.csv" | ForEach-Object {
    $licenseNameMapping[$_.String_Id] = $_.Product_Display_Name
}

# Get all the licensed users
$licensedUsers = Get-MsolUser -All | Where-Object { $_.IsLicensed -eq $true }

# Create an empty array to store the user information
$userInfoArray = @()

# Loop through each user and get their license information
foreach ($user in $licensedUsers) {
    # Get the user's display name and email address
    $displayName = $user.DisplayName
    $emailAddress = $user.UserPrincipalName

    # Get the user's licenses and use the part after the colon
    $licenses = $user.Licenses.AccountSkuId
    $licenseList = ""
    foreach ($license in $licenses) {
        # Split the license name into its parts and use the part after the colon
        # This is necessary because Microsoft likes to put colons before their licenses from a partner reseller
        # TODO: Does this work if the licenses do not come from a reseller still?
        $licenseParts = $license -split ":"
    
        # Check if there is a mapping for this license name
        if ($licenseNameMapping.ContainsKey($licenseParts[1])) {
            # Use the advertised name from the mapping
            $licenseList += $licenseNameMapping[$licenseParts[1]]
        }
        else {
            # Use the original license name if there is no mapping
            $licenseList += $licenseParts[1]
        }
    
        # Check if this is the last iteration of the loop
        if ($license -ne $licenses[-1]) {
            # Add a comma if this is not the last iteration
            $licenseList += ", "
        }
    }

    # Create a custom object to store the user's information
    $userInfo = New-Object -TypeName PSObject
    $userInfo | Add-Member -MemberType NoteProperty -Name "Display Name" -Value $displayName
    $userInfo | Add-Member -MemberType NoteProperty -Name "Email Address" -Value $emailAddress
    $userInfo | Add-Member -MemberType NoteProperty -Name "Licenses" -Value $licenseList

    # Add the custom object to the array
    $userInfoArray += $userInfo
}

# Get all the available license types in the organization
$AccountSku = Get-MsolAccountSku

# Create an empty array to store the results
$orgLicenseCounts = @()

# Loop through each SKU
# Create an empty array to store the license counts
$orgLicenseCounts = @()

foreach ($Sku in $AccountSku) {
    # Split the SKU name into its parts and use the part after the colon
    $skuParts = $Sku.AccountSkuId -split ":"

    # Check if there is a mapping for this SKU name
    if ($licenseNameMapping.ContainsKey($skuParts[1])) {
        # Use the advertised name from the mapping
        $skuName = $licenseNameMapping[$skuParts[1]]
    }
    else {
        # Use the original SKU name if there is no mapping
        $skuName = $skuParts[1]
    }

    # Create a custom object to store the SKU and license count
    $Result = New-Object -TypeName PSObject
    $Result | Add-Member -MemberType NoteProperty -Name "SKU" -Value $skuName
    $Result | Add-Member -MemberType NoteProperty -Name "Total Licenses" -Value $Sku.ActiveUnits
    $Result | Add-Member -MemberType NoteProperty -Name "Consumed Licenses" -Value $Sku.ConsumedUnits

    # Add the result to the array
    $orgLicenseCounts += $Result
}

# Export the user information and license totals to a CSV file in the current directory
$userInfoArray | Export-Csv -Path ".\User Licenses.csv" -NoTypeInformation
$orgLicenseCounts | Export-Csv -Path ".\Organization Licenses.csv" -NoTypeInformation
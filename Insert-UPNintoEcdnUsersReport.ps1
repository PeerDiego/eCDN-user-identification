<#
.SYNOPSIS
    This script reads a CSV file containing user Object IDs, retrieves their User Principal Names (UPNs) and display names from Azure AD, and exports the results to a new CSV file.

.DESCRIPTION
    The script takes an input CSV file with a column named 'Users' containing the Object IDs of users. It connects to Azure AD, retrieves the UPN and display name for each user, and adds these details to the CSV. The results are then exported to a specified output CSV file. Optionally, the results can also be output to the host.

.PARAMETER InputCsvPath
    The path to the input CSV file containing the user Object IDs.

.PARAMETER OutputCsvPath
    The path to the output CSV file where the results will be saved. Defaults to "output.csv".

.PARAMETER AlsoOutputToHost
    A switch parameter that, if specified, will also output the results to the host.

.EXAMPLE
    .\Insert-UPNintoEcdnUsersReport.ps1 -InputCsvPath "input.csv" -OutputCsvPath "output.csv" -AlsoOutputToHost

.NOTES
    The script requires the AzureAD module to be installed and the user to be connected to Azure AD.

#>
param (
    [Parameter(Mandatory = $true)]
    [string] $InputCsvPath,
    [string] $OutputCsvPath = "output.csv",
    [switch] $AlsoOutputToHost
)

# Define the GUID pattern
$guidPattern = '^[0-9a-fA-F]{8}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{4}\b-[0-9a-fA-F]{12}$'

# Read the input CSV
try {
    $users = Import-Csv -Path $InputCsvPath
}
catch {
    Write-Error "Failed to read the input CSV file. $_"
    exit 1
}

# Ensure the input CSV contains the required columns
if ($users[0].PSObject.Properties.Name -notcontains "Users") {
    Write-Error "The input CSV file must contain a 'Users' column containing the Object IDs"
    exit 1
}

# Import the AzureAD module
Import-Module AzureAD -Verbose:$false

# Connect to Azure AD
try {
    Connect-AzureAD | Out-Null
}
catch {
    Write-Error "Failed to connect to Azure AD. $_"
    if ($PSVersionTable.PSVersion.Major -ne 5) {
        Write-Host "Try running the script in PowerShell 5" -f Yellow
    }
    exit 1
}

# Create a hash table to cache user details
$hashTable = @{}

# Initialize progress bar variables
$totalUsers = $users.Count
$currentUser = 0

foreach ($user in $users) {
    $currentUser++
    $ObjectId = $user.Users

    # Update the progress bar
    Write-Progress -Activity "Processing users" -Status "Processing user $currentUser of $totalUsers" -PercentComplete (($currentUser / $totalUsers) * 100)

    try {
        # Get the user details
        $userDetails = if ($ObjectId -in $hashTable.Keys) {
            Write-Host "Using cached user details for $ObjectId" -f Yellow
            $hashTable[$ObjectId]
        }
        elseif ($ObjectId -match $guidPattern) {
            Get-AzureADUser -ObjectId $ObjectId -ErrorAction SilentlyContinue
        } else {
            Get-AzureADUser -Filter "UserPrincipalName eq '$ObjectId'" -ErrorAction SilentlyContinue
        }
        
        # Add the UPN and Name columns to the user object
        if ($userDetails) {
            $user | Add-Member -MemberType NoteProperty -Name UPN -Value $userDetails.UserPrincipalName
            $user | Add-Member -MemberType NoteProperty -Name Name -Value $userDetails.DisplayName
            if ($ObjectId -notin $hashTable.Keys) {
                Write-Verbose "Caching user details for $ObjectId"
                $hashTable[$ObjectId] = $userDetails
            }
        }
        else {
            Write-Host "Invalid Object ID: $ObjectId" -f DarkGray
        }
    } catch {
        Write-Host "Failed to get user details for $ObjectId" -f Red
    }
}

# Inserts the UPN and Name columns into the CSV
$Properties = @("Users", "UPN", "Name") + ($users[0].PSObject.Properties.Name | Where-Object { $_ -notin @("Users", "UPN", "Name") })

# Export the results to a CSV
# Retry until the export is successful or user cancels
while (1) {
    try {
        $users | Select-Object $Properties | Export-Csv -Path $OutputCsvPath -NoTypeInformation
        break
    }
    catch {
        Write-Error "Failed to export the results to $OutputCsvPath. $_"
        $null = Read-Host "Please enter to try again or press Ctrl + C to exit"
    }
}

Write-Host "Results have been exported to $OutputCsvPath"

if ($AlsoOutputToHost) {
    $users | Select-Object $Properties
}
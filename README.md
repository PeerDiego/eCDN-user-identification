# Insert UPN into eCDN Users Report

This PowerShell script reads a CSV file containing user Object IDs, retrieves their User Principal Names (UPNs) and display names from Azure AD, and exports the results to a new CSV file.

## Description

The script takes an input CSV file with a column named `Users` containing the Object IDs of users. It connects to Azure AD, retrieves the UPN and display name for each user, and adds these details to the CSV. The results are then exported to a specified output CSV file. Optionally, the results can also be output to the host.

## Parameters

- `InputCsvPath`: The path to the input CSV file containing the user Object IDs. (Mandatory)
- `OutputCsvPath`: The path to the output CSV file where the results will be saved. Defaults to `output.csv`.
- `AlsoOutputToHost`: A switch parameter that, if specified, will also output the results to the host.

## Example

```powershell
.\Insert-UPNintoEcdnUsersReport.ps1 -InputCsvPath "input.csv" -OutputCsvPath "output.csv" -AlsoOutputToHost
```

## Notes

The script requires the AzureAD module to be installed and the user to be connected to Azure AD.

## Usage

Ensure you have the AzureAD module installed.
Connect to Azure AD.
Run the script with the required parameters.

## Script

The script is located at `Insert-UPNintoEcdnUsersReport.ps1`.

## License

This project is licensed under the MIT License.

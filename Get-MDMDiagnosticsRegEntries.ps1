<#
.SYNOPSIS
    Retrieves and displays registry entries from the MdmDiagnostics area.

.DESCRIPTION
    This script iterates through the registry entries under the MdmDiagnostics area in the Windows registry,
    collects the values, and outputs them in a formatted table. It is useful for understanding which items are
    collected for diagnosing and troubleshooting MDM related issues.

.EXAMPLE
    .\Get-MDMDiagnosticsRegEntries.ps1
    This command runs the script and outputs the registry entries from the MdmDiagnostics area.

.NOTES
    Author: Ben Whitmore
    Date: 25/01/2025

.OUTPUTS
    System.Management.Automation.PSCustomObject
    The script outputs custom objects containing the registry path, area, item type, name, registry type, and data.
#>


# Define the base registry path
$basePath = "HKLM:\SOFTWARE\Microsoft\MdmDiagnostics\Area"

# Initialize an array to store results
$results = @()

# Iterate through each area under the base path
Get-ChildItem -Path $basePath | ForEach-Object {
    $area = $_.PSChildName

    # Iterate through each folder under the area
    Get-ChildItem -Path $_.PSPath | ForEach-Object {
        $folder = $_.PSChildName
        $keyPath = $_.PSPath

        # Retrieve all values from the folder
        $key = Get-Item -Path $keyPath
        $values = @()
        foreach ($name in $key.Property) {
        
            $type = $key.GetValueKind($name) 
            $data = $key.GetValue($name)

            $values += [PSCustomObject]@{
                Name     = $name
                Reg_Type = $type.ToString()
                Data     = $data
            }
        }

        # Add the result as a custom object
        $results += [PSCustomObject]@{
            BasePath  = $basePath
            Area      = $area
            Type      = $folder
            HashTable = $values
        }
    }
}

# Initialize an array to collect the output
$finalResults = @()

# Process the results and add to the array
foreach ($result in $results) {
    foreach ($entry in $result.HashTable) {
        $finalResults += [PSCustomObject]@{
            BasePath = $result.BasePath
            Area     = $result.Area
            ItemType = $result.Type
            Name     = $entry.Name
            Reg_Type = $entry.Reg_Type
            Data     = $entry.Data
        }
    }
}

# Output the collected results
$finalResults | Format-Table -AutoSize
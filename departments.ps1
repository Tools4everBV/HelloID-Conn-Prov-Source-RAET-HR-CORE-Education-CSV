$connectionSettings = ConvertFrom-Json $configuration
$importSourcePath = $($connectionSettings.path)
$delimiter = $($connectionSettings.delimiter)

#region Functions
function Get-SourceConnectorData { 
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]$SourceFile,
        [parameter(Mandatory = $true)][ref]$data
    )
    
    try {
        $importSourcePath = $importSourcePath -replace '[\\/]?[\\/]$'

        # the string '"t' causes the csv to be partially imported correctly. 
        $result = Get-Content -Path "$importSourcePath\$SourceFile"
        $result | ForEach-Object { $_-replace '"t ', "'t"} | Set-Content "$importSourcePath\$SourceFile"

        # CSV is exported weirdly, via excel propably? Only using the UTF7 encoding the special characters are imported correctly
        $dataset = Import-Csv -Path "$importSourcePath\$SourceFile" -Delimiter $delimiter -Encoding UTF7

        foreach ($record in $dataset) { 
            $null = $data.Value.add($record) 
        }
    }
    catch {
        $data.Value = $null
        Write-Verbose $_.Exception
    }
}
#endregion Functions

#region Script

# OU's
$organizationalUnits = New-Object System.Collections.ArrayList
Get-SourceConnectorData -SourceFile "Eenheid.csv" ([ref]$organizationalUnits)

# Manager references
$managerReferences = New-Object System.Collections.ArrayList
Get-SourceConnectorData -SourceFile "Manageridentificatie.csv" ([ref]$managerReferences)

$managerReferences | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$managerReferences | Add-Member -MemberType NoteProperty -Name "Volgnr_next" -Value $null -Force
$managerReferences | ForEach-Object {
    $_.ExternalId = $_.'[Dossiernummer]'
    $_.Volgnr_next = $_.'[Volgend Volgnummer]'
}
$managerReferences = $managerReferences | Where-Object Volgnr_next -eq '0' | Sort-Object ExternalId -unique | Group-Object ExternalId -AsHashTable

$organizationalUnits | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$organizationalUnits | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
$organizationalUnits | Add-Member -MemberType NoteProperty -Name "Code" -Value $null -Force
$organizationalUnits | Add-Member -MemberType NoteProperty -Name "ParentExternalId" -Value $null -Force
$organizationalUnits | Add-Member -MemberType NoteProperty -Name "ManagerExternalId" -Value $null -Force

$organizationalUnits | ForEach-Object {
    # join with manager
    $manager = $managerReferences[$_.'[Dossiernummer]']
    if ($null -ne $manager) {
        $_.ManagerExternalId = $manager.'[ID]'.trim()
    }
    $_.ParentExternalId = $_.'[Parent Eenheid]'
    $_.Code = $_.'[ID org-eenheid]'
    $_.ExternalId = $_.'[ID org-eenheid]'
    $_.DisplayName = $_.'[Lg naam org-eenhd]'
}

# only get the data needed for HelloID
$organizationalUnits = $organizationalUnits | Select-Object -Property ExternalId,DisplayName,Code,ManagerExternalId,ParentExternalId 

$organizationalUnits = $organizationalUnits | Sort-Object ExternalId -Unique

$json = $organizationalUnits | ConvertTo-Json -Depth 5

Write-Output $json
#endregion Script
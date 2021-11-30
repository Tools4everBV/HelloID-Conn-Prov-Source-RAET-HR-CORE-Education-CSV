$connectionSettings = ConvertFrom-Json $configuration
$importSourcePath = $($connectionSettings.path)
$delimiter = $($connectionSettings.delimiter)
$useCustomPrimaryPersonCalculation = $($connectionSettings.useCustomPrimaryPersonCalculation)

#region Functions
function Get-SourceConnectorData { 
    [CmdletBinding()]
    param (
        [parameter(Mandatory = $true)]$SourceFile,
        [parameter(Mandatory = $true)][ref]$data
    )

    try {
        # sanitize first
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
$persons = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Medewerker.csv" ([ref]$persons)

$partners = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Echtgenoot.csv" ([ref]$partners)

$metadata = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Telefoon.csv" ([ref]$metadata)

$employments = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Aanstellingen.csv" ([ref]$employments)

$professions = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Functie.csv" ([ref]$professions)

$departments = New-Object System.Collections.Generic.list[object]
Get-SourceConnectorData -SourceFile "Eenheid.csv" ([ref]$departments)

$partners | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$partners | Add-Member -MemberType NoteProperty -Name "Volgnr_next" -Value $null -Force

$partners | ForEach-Object {
    $_.ExternalId = $_.'[Dossiernummer]'
    $_.Volgnr_next = $_.'[Volgend Volgnummer]'
}
# fix for duplicate records!
$partners = $partners | Where-Object Volgnr_next -eq '0' | Sort-Object ExternalId -unique | Group-Object ExternalId -AsHashTable

# contact stuff
$metadata | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$metadata | Add-Member -MemberType NoteProperty -Name "Volgnr_next" -Value $null -Force
$metadata | ForEach-Object {
    $_.ExternalId = $_.'[Dossiernummer]'
    $_.Volgnr_next = $_.'[Volgend Volgnummer]'
}
$metadata = $metadata | Where-Object Volgnr_next -eq '0' | Group-Object ExternalId -AsHashTable

# function
$professions | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$professions | ForEach-Object {
    $_.ExternalId = $_.'[Werkidentificatie]'
}
$professions = $professions | Group-Object ExternalId -AsHashTable

# department
$departments | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$departments | ForEach-Object {
    $_.ExternalId = $_.'[ID org-eenheid]'.trim()
}
$departments = $departments | Group-Object ExternalId -AsHashTable

# contracts
$employments | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$employments | Add-Member -MemberType NoteProperty -Name "ExternalIdVault" -Value $null -Force
$employments | Add-Member -MemberType NoteProperty -Name "[Functieomschrijving]" -Value $null -Force
$employments | Add-Member -MemberType NoteProperty -Name "[Afdelingsomschrijving]" -Value $null -Force

# Join department and profession description
$employments | ForEach-Object {
    $_.ExternalId = $_.'[Dossiernummer]'
    $_.ExternalIdVault = $_.'[Dossiernummer]' + '_' + $_.'[Volgnummer]'

    $department = $departments[$_.'[Afdelingscode]']
    if ($null -ne $department) {
        $_.'[Afdelingsomschrijving]' = $department.'[Lg naam org-eenhd]'.trim()
    }    

    $profession = $professions[$_.'[Functiecode]']
    if ($null -ne $profession) {
        $_.'[Functieomschrijving]' = $profession.'[Functienaam]'.trim()
    }  

    # convert the dates
    $_.'[Begindatum]' = [datetime]::parseexact($_.'[Begindatum]', 'yyyyMMdd', $null)
    $_.'[Einddatum]' = [datetime]::parseexact($_.'[Einddatum]', 'yyyyMMdd', $null)
}

$employments = $employments | Select-Object ExternalId, 
                                            ExternalIdVault, 
                                            @{Name = 'Begindatum';Expression= {$_.'[BeginDatum]'}}, 
                                            @{Name = 'Einddatum';Expression= {$_.'[EindDatum]'}}, 
                                            @{Name = 'Afdelingscode';Expression= {$_.'[Afdelingscode]'}}, 
                                            @{Name = 'Afdelingsomschrijving';Expression= {$_.'[Afdelingsomschrijving]'}}, 
                                            @{Name = 'Functiecode';Expression= {$_.'[Functiecode]'}}, 
                                            @{Name = 'Functieomschrijving';Expression= {$_.'[Functieomschrijving]'}}, 
                                            @{Name = 'Kostendragercode';Expression= {$_.'[Kostendrager]'}}, 
                                            @{Name = 'Kostendrageromschrijving';Expression= {$_.'[Kostendrager korte omschrijving]'}}, 
                                            @{Name = 'Type';Expression= {$_.'[Code type aanstelling]'}}, 
                                            @{Name = 'Typeomschrijving';Expression= {$_.'[Code type aanstelling omschrijving]'}}, 
                                            @{Name = 'Kostenplaatscode';Expression= {$_.'[Kostenplaats]'}}, 
                                            @{Name = 'Kostenplaatsomschrijving';Expression= {$_.'[Kostenplaats omschrijving]'}}, 
                                            @{Name = 'Kostensoortcode';Expression= {$_.'[Kostensoort]'}},
                                            @{Name = 'Kostensoortomschrijving';Expression= {$_.'[Kostensoort omschrijving]'}},
                                            @{Name = 'IndicatieHoofdaanstelling';Expression= {$_.'[Indicatie hoofdaanstelling]'}},
                                            @{Name = 'Urenperweek';Expression= {$_.'[Uren per week]'}},
                                            @{Name = 'SoortContractCode';Expression= {$_.'[Soort contract]'}},
                                            @{Name = 'SoortContractOmschrijving';Expression= {$_.'[Soort contract omschrijving]'}},
                                            @{Name = 'UserId';Expression= {$_.'[User ID]'}}

# Group the employments
$employments = $employments | Group-Object ExternalId -AsHashTable

# Extend the persons with required and optional fields
$persons | Add-Member -MemberType NoteProperty -Name "Contracts" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "ExternalId" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "MobielWerk" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "VastWerk" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "EmailWerk" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "EmailPrive" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "Partnernaam" -Value $null -Force
$persons | Add-Member -MemberType NoteProperty -Name "Partnertussenvoegsel" -Value $null -Force

# Join info with Persons
$persons | ForEach-Object {
    $_.ExternalId = $_.'[Dossiernummer_2]'
    $_.DisplayName = $_.'[Volledige Naam]'
    $_.'[Geboortedatum]' = [datetime]::parseexact($_.'[Geboortedatum]', 'yyyyMMdd', $null)
    # Contact info
    $meta = $metadata[$_.'[Dossiernummer_2]']
    if ($null -ne $meta) {
        $zakelijk_mobiel_nummer = ''
        $zakelijk_vast_nummer = ''
        $zakelijk_email = ''
        $prive_email = ''
        # all the values are in 'Telefoonnummer' attribute, using the 'type nummmer' we can determine what the value stands for
        foreach ($item in $meta) {
            switch($item.'[Type nummer]') {
                'CPC' { $zakelijk_mobiel_nummer = $item.'[Telefoonnummer]';break }
                'WPN' { $zakelijk_vast_nummer = $item.'[Telefoonnummer]';break }
                'EML' { $zakelijk_email = $item.'[Telefoonnummer]';break }
                'EMP' { $prive_email = $item.'[Telefoonnummer]';break }
            }
        }
        $_.MobielWerk = $zakelijk_mobiel_nummer.Trim()
        $_.VastWerk = $zakelijk_vast_nummer.Trim()
        $_.EmailWerk = $zakelijk_email.Trim()
        $_.EmailPrive = $prive_email.Trim()
    }

    # Join partner stuff
    $partner = $partners[$_.ExternalId]
    if ($null -ne $partner) {
        $_.Partnernaam = $partner.'[Achternaam]'
        $_.Partnertussenvoegsel = $partner.'[Voorvoegsels]'
    }

    # Join contracts
    $contracts = $employments[$_.'[Dossiernummer_2]']
    if ($null -ne $contracts) {
        [array]$_.Contracts = $contracts
    }

    # Format name convention
    if ($_.'[Naam samenstelling]' -eq '1') {
        $_.'[Naam samenstelling]' = 'B'
    }
    if ($_.'[Naam samenstelling]' -eq '2') {
        $_.'[Naam samenstelling]' = 'BP'
    }
    if ($_.'[Naam samenstelling]' -eq '3') {
        $_.'[Naam samenstelling]' = 'PB'
    }
    if ($_.'[Naam samenstelling]' -eq '4') {
        $_.'[Naam samenstelling]' = 'P'
    }
    # if ($_.'[Naam samenstelling]' -eq '5') {
    #     $_.'[Naam samenstelling]' = 'B'
    # }
    # if ($_.'[Naam samenstelling]' -eq '6') {
    #     $_.'[Naam samenstelling]' = 'B'
    # }
    # if ($_.'[Naam samenstelling]' -eq '7') {
    #     $_.'[Naam samenstelling]' = 'B'
    # }

    # EmployeeNumber as ExternalId
    $_.ExternalId = $_.'[Personeelsnummer]'.trim()
}

# Below is the configuration part as used to identify the primary person, 
# all persons are grouped by the UniqueKey where only the person with the highest contract will be included for provisioning
if ($true -eq $useCustomPrimaryPersonCalculation) {
    #Extend the person model to automatically exclude persons by the source import
    $persons | Add-Member -MemberType NoteProperty -Name "ExcludedBySource" -Value "true" -Force

    #Define the logic used for ordering and grouping the person objects used to identify the primary person
    #Calculate identities base on your UniqueKey
    $persons = $persons | Select-Object *, @{name = 'UniqueKey'; expression = { "$($_.'[Voornaam]')$($_.'[Eigen naam]')$($_.'[Geslacht]')$($_.'[Geboortedatum]')" } }
    $identities = $persons | Select-Object -Property UniqueKey -Unique
    $personsGrouped = $persons | Group-Object -Property UniqueKey -AsHashTable -AsString

    #Define the property's used for sorting the persons priority based on one or more contract fields
    $prop1 = @{Expression = { $_.'[Volgnummer]' }; Descending = $True }
    # $prop2 = @{Expression = { $_.Aantal_FTE }; Descending = $True } # FTE niet beschikbaar?
    $prop3 = @{Expression = { if (($_.'[Einddatum]' -eq "") -or ($null -eq $_.'[Einddatum]') ) { (Get-Date -Year 2199 -Month 12 -Day 31) -as [datetime] } else { $_.'[Einddatum]' -as [datetime] } }; Descending = $true }
    $prop4 = @{Expression = { $_.'[Begindatum]' }; Ascending = $True }

    foreach ($identity in $identities ) {
        $latestContract = ($personsGrouped[$identity.UniqueKey] | Select-Object -ExpandProperty contracts | Sort-Object -Property $prop1, $prop3, $prop4 | Select-Object -First 1)
        $personToUpdate = ($persons | Where-Object { $_.ExternalId -eq $latestContract.ExternalId })

        if ($null -eq $personToUpdate ) {
            "No contracts found for $identity"
            continue
        }
        $personToUpdate.ExcludedBySource = "false"
    }
}

#Test selection to identify if the sorted results are correct
#$persons | Select-Object -Property Medewerker, ExcludedBySource | Format-Table

# Make sure persons are unique
$persons = $persons | Sort-Object ExternalId -Unique

$persons = $persons | Select-Object ExternalId,
                                    DisplayName,
                                    MobielWerk,
                                    VastWerk,
                                    EmailWerk,
                                    EmailPrive,
                                    @{Name = 'Roepnaam';Expression= {$_.'[Roepnaam]'.Trim()}}, 
                                    @{Name = 'Geboortenaam';Expression= {$_.'[Eigen naam]'.Trim()}}, 
                                    @{Name = 'Geboortetussenvoegsel';Expression= {$_.'[Tussenvoegsel]'.Trim()}}, 
                                    @{Name = 'Partnernaam';Expression= {$_.Partnernaam.Trim()}}, 
                                    @{Name = 'Partnernaamtussenvoegsel';Expression= {$_.Partnertussenvoegsel.Trim()}}, # not present, only in combined form
                                    @{Name = 'Convention';Expression= {$_.'[Naam samenstelling]'.Trim()}}, 
                                    @{Name = 'Initialen';Expression= {$_.'[Initialen]'.Trim()}}, 
                                    @{Name = 'Geslachtcode';Expression= {$_.'[Geslacht]'.Trim()}}, 
                                    @{Name = 'Geslachtomschrijving';Expression= {$_.'[Geslacht omschrijving]'.Trim()}}, 
                                    @{Name = 'Geboortedatum';Expression= {$_.'[Geboortedatum]'.Trim()}}, 
                                    @{Name = 'Status';Expression= {$_.'[Huidige Medewerker Status]'.Trim()}}, 
                                    @{Name = 'Roostercode';Expression= {$_.'[Roostercode]'.Trim()}}, 
                                    @{Name = 'ExterneMedewerker';Expression= {$_.'[Externe Medewerker]'.Trim()}}, 
                                    @{Name = 'BlokkerenOpnameFunctieMix';Expression= {$_.'[Blokkeren opname functiemix]'.Trim()}},
                                    Contracts
# test
# $persons = $persons | Where-Object contracts -ne $null

# Export and return the json
$json = $persons | ConvertTo-Json -Depth 10

Write-Output $json
#endregion Script
function Add-PSObjectToRegistry {
    <#
.SYNOPSIS
    Writes PSObject properties to Windows registry keys from the pipeline.

.DESCRIPTION
    Add-PSObjectToRegistry accepts one or more PSObjects or HashTables from the pipeline
    and serializes their properties as registry values under a specified key path. The
    target path is always rooted under HKLM:\SOFTWARE or HKCU:\SOFTWARE. Both PSObject
    and HashTable input types are supported. HashTables follow the same subkey
    conventions as PSObjects but do not support -UseFirstPropertyAsKey when multiple
    HashTables are provided.

    When multiple objects are piped in, each object is written to its own subkey. By
    default, subkeys are named numerically starting at 0. Use -UseLeadingZeros to pad
    numeric names to a consistent width, or -UseFirstPropertyAsKey to name each subkey
    after the value of the object's first property. By default, unique key names are
    ensured by appending numeric suffixes when using -UseFirstPropertyAsKey but the
    -AllowOverwrite parameter may be used to reverse this behavior.

    When only a single object is piped in and neither -UseFirstPropertyAsKey nor
    -UseLeadingZeros is specified, properties are written directly to the target key
    with no subkey created.

    This function requires appropriate permissions to write to the target hive. Writing
    to HKLM typically requires an elevated session.

.PARAMETER InputObject
    One or more PSObjects to serialize into the registry. Accepts pipeline input.
    Each object's properties become registry values. Nested objects and arrays are
    not supported as property values; such properties should be flattened before
    passing to this function.

.PARAMETER Hive
    The registry hive to target. Accepted values are 'HKLM' (HKEY_LOCAL_MACHINE)
    and 'HKCU' (HKEY_CURRENT_USER). Defaults to 'HKLM'.

    Writing to HKLM requires an elevated (Administrator) session.

.PARAMETER KeyName
    The name of the top-level key to create or target under the hive's SOFTWARE key.
    For example, a value of 'MyApp' resolves to HKLM:\SOFTWARE\MyApp.

.PARAMETER SubKeyNames
    An optional ordered array of subkey names representing the path beneath KeyName
    where data will be written. For example, passing @('Config', 'Network') resolves
    to HKLM:\SOFTWARE\MyApp\Config\Network.

    Defaults to an empty array, writing directly under KeyName.

.PARAMETER ResetKey
    When specified, the target registry key and all of its subkeys and values are
    deleted and recreated before any objects are written. This ensures that stale
    data from a previous write is not retained when the new dataset is smaller than
    the original.

    This applies equally to both naming approaches: with numeric enumeration, subkeys
    beyond the new dataset's count will persist without a reset; with -UseFirstPropertyAsKey,
    any named subkey not present in the new dataset will similarly remain with its
    original data intact.

    This operation is irreversible. Use -WhatIf to preview the deletion before
    committing. Writing to HKLM requires an elevated session.

.PARAMETER UseFirstPropertyAsKey
    When specified, the value of each object's first property is used as the subkey
    name for that object rather than a numeric index. This is useful when objects
    have a natural identifier property such as a name or ID.

    By default, multiple objects sharing the same first property value will receive suffixes
    to ensure their uniqueness. You may use -AllowOverwrite to prevent that behavior.

.PARAMETER AllowOverwrite
    When specified alongside -UseFirstPropertyAsKey, a numeric suffix is not appended to
    any subkey name that would otherwise collide with an already-written key. For example,
    if two objects both have a first property value of 'Service', the second will be
    written to the same subkey as the first, overwriting values from the first.

    Has no effect when -UseFirstPropertyAsKey is not specified.

.PARAMETER UseLeadingZeros
    When specified, numeric subkey names are left-padded with zeros to a consistent
    width determined by the total number of input objects. For example, with 100 objects,
    keys will be named 000, 001, 002 ... 099, 100 rather than 0, 1, 2 ... 99, 100.

    This ensures subkeys sort correctly in tools such as regedit and when enumerated
    by the registry provider.

    Has no effect when -UseFirstPropertyAsKey is specified.

.INPUTS
    System.Management.Automation.PSObject
        Any PSObject or PSCustomObject. Objects are accepted from the pipeline or
        passed directly via -InputObject.

    System.Collections.Hashtable
        A HashTable whose keys and values are written as registry value names and
        data. Multiple HashTables are written to numeric subkeys following the same
        conventions as PSObject input. Cannot be mixed with PSObject input in the
        same pipeline call.

.OUTPUTS
    None
        This function does not return output to the pipeline. All writes are performed
        against the registry.

.EXAMPLE
    [PSCustomObject]@{ Theme = 'Dark'; FontSize = 14; AutoSave = $true } |
        Add-PSObjectToRegistry -Hive HKCU -KeyName MyApp

    Writes a single object's properties directly to HKCU:\SOFTWARE\MyApp. No subkey
    is created because only one object was piped and -UseFirstPropertyAsKey was not
    specified.

.EXAMPLE
    Get-Service | Select-Object Name, DisplayName, Status |
        Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp -SubKeyNames @('Services') -UseFirstPropertyAsKey

    Writes each service object to a subkey named after its Name property, e.g.
    HKLM:\SOFTWARE\MyApp\Services\Spooler, HKLM:\SOFTWARE\MyApp\Services\WinRM, etc.

.EXAMPLE
    Get-Service | Select-Object Name, DisplayName, Status |
        Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp -SubKeyNames @('Services') `
            -UseFirstPropertyAsKey -AllowOverwrite

    Same as the previous example, but if any two services share the same Name value,
    the second occurrence is written to the same subkey as the first, silently
    overwriting data from the first write.

.EXAMPLE
    $freshData | Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp `
        -SubKeyNames @('Items') -UseLeadingZeros -ResetKey

    Deletes HKLM:\SOFTWARE\MyApp\Items and all of its contents before writing,
    ensuring no subkeys from a previous run persist in the refreshed dataset.

.EXAMPLE
    $data = Import-Csv .\Items.csv
    $data | Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp -SubKeyNames @('Items') -UseLeadingZeros

    Imports a CSV and writes each row to a zero-padded numeric subkey. With 42 rows,
    keys will be named 00 through 41, ensuring consistent sort order.

.EXAMPLE
    @{ Server = 'SRV01'; Port = 8080; Enabled = $true } |
        Add-PSObjectToRegistry -Hive HKCU -KeyName MyApp -SubKeyNames @('Connection')

    Writes a single HashTable's keys and values directly to
    HKCU:\SOFTWARE\MyApp\Connection with no subkey created.

.EXAMPLE
    $configs | Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp -SubKeyNames @('Configs') -UseLeadingZeros

    Writes multiple HashTables to zero-padded numeric subkeys under
    HKLM:\SOFTWARE\MyApp\Configs, one subkey per HashTable.

.EXAMPLE
    Get-Process | Select-Object Name, Id, CPU |
        Add-PSObjectToRegistry -Hive HKLM -KeyName MyApp -WhatIf

    Previews all registry operations without writing any data. Use this to validate
    the target path and key structure before committing changes.

.NOTES
    - This function is not transactional. If an error occurs mid-write, previously
      written keys and values will not be rolled back. When using the -ResetKey
      switch, all existing data under the target key is permanently deleted before
      writing begins. If an error occurs during the subsequent write, the deleted
      data cannot be recovered.
    - Writing to HKLM requires an elevated session. If the session is not elevated,
      a permissions error will be raised when the key is created.
    - Nested objects, arrays, and null values as property values are not supported
      and should be resolved or filtered out before calling this function.

.LINK
    https://www.powershellgallery.com/packages/PSObjectToRegistry

#>

    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Position = 0, Mandatory, ValueFromPipeline)]
        [object[]]$InputObject,

        [Parameter(Position = 1)]
        [ValidateSet('HKLM', 'HKCU')]
        [string]$Hive = 'HKLM',

        [Parameter(Position = 2, Mandatory)]
        [string]$KeyName,

        [Parameter(Position = 3)]
        [string[]]$SubKeyNames = @(),

        [Parameter(Position = 4)]
        [switch]$ResetKey,

        [Parameter(Position = 5)]
        [switch]$UseFirstPropertyAsKey,

        [Parameter(Position = 6)]
        [switch]$AllowOverwrite,

        [Parameter(Position = 7)]
        [switch]$UseLeadingZeros
    )
    begin {
        $DetectedType = $null
        [int]$ItemNumber = 0
        $UsedKeyNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $Items = [System.Collections.Generic.List[object]]::new()

        $PathParts = @('SOFTWARE', $KeyName) + $SubKeyNames | Where-Object { $_ }
        $BasePath = Join-Path "${Hive}:" ($PathParts -join '\')

    }
    process {
        foreach ($Item in $InputObject) {
            $CurrentType = if ($Item -is [hashtable]) { 'hashtable' } else { 'psobject' }

            if ($null -eq $DetectedType) {
                $DetectedType = $CurrentType
            }
            elseif ($DetectedType -ne $CurrentType) {
                throw 'InputObject must contain either HashTables or PSObjects, not both.'
            }

            $Items.Add($Item)
        }
    }
    end {
        [string]$DValue = 'D' + $Items.Count.ToString().Length

        if ($UseFirstPropertyAsKey -and $DetectedType -eq 'hashtable') {
            Write-Verbose '-UseFirstPropertyAsKey has no effect when the input is a HashTable.'
        }

        if ($ResetKey -and (Test-Path $BasePath)) {
            if ($PSCmdlet.ShouldProcess($BasePath, 'Delete registry key and all subkeys')) {
                Remove-Item -Path $BasePath -Recurse -Force
            }
        }

        if ($PSCmdlet.ShouldProcess($BasePath, 'Create registry key')) {
            if (-not (Test-Path $BasePath)) {
                New-Item -Path $BasePath -Force | Out-Null
            }
        }

        foreach ($Item in $Items) {
            if ($Item -is [hashtable]) {
                if ($Items.Count -eq 1) {
                    $TargetPath = $BasePath
                }
                elseif ($UseLeadingZeros) {
                    $TargetPath = Join-Path $BasePath ($ItemNumber++).ToString($DValue)
                }
                else {
                    $TargetPath = Join-Path $BasePath ($ItemNumber++)
                }

                if ($TargetPath -ne $BasePath) {
                    if ($PSCmdlet.ShouldProcess($TargetPath, 'Create registry subkey')) {
                        New-Item -Path $TargetPath -Force | Out-Null
                    }
                }

                foreach ($Entry in $Item.GetEnumerator()) {
                    if ($PSCmdlet.ShouldProcess($TargetPath, "Set property '$($Entry.Key)'")) {
                        Set-ItemProperty -Path $TargetPath -Name $Entry.Key -Value $Entry.Value
                    }
                }
                continue
            }
            if ($Items.Count -eq 1 -and -not $UseFirstPropertyAsKey) {
                $TargetPath = $BasePath
            }
            elseif ($UseFirstPropertyAsKey) {
                $ItemKeyName = ($Item.PSObject.Properties | Select-Object -First 1).Value
                $TargetPath = Join-Path $BasePath $ItemKeyName

                if (-not $AllowOverwrite) {
                    [int]$Suffix = 0
                    while ($UsedKeyNames.Contains($ItemKeyName)) {
                        $ItemKeyName = "$ItemKeyName`_$Suffix"
                        $Suffix++
                    }
                    $TargetPath = Join-Path $BasePath $ItemKeyName
                    $UsedKeyNames.Add($ItemKeyName) | Out-Null
                }
            }
            elseif ($UseLeadingZeros) {
                $TargetPath = Join-Path $BasePath ($ItemNumber++).ToString($DValue)
            }
            else {
                $TargetPath = Join-Path $BasePath ($ItemNumber++)
            }

            if ($TargetPath -ne $BasePath) {
                if ($PSCmdlet.ShouldProcess($TargetPath, 'Create registry subkey')) {
                    New-Item -Path $TargetPath -Force | Out-Null
                }
            }

            foreach ($Prop in $Item.PSObject.Properties) {
                if ($PSCmdlet.ShouldProcess($TargetPath, "Set property '$($Prop.Name)'")) {
                    Set-ItemProperty -Path $TargetPath -Name $Prop.Name -Value $Prop.Value
                }
            }
        }
    }
}

Export-ModuleMember -Function Add-PSObjectToRegistry
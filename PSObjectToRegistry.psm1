function Add-PSObjectToRegistry {
    <#
.SYNOPSIS
    Writes PSObject properties to Windows registry keys from the pipeline.

.DESCRIPTION
    Add-PSObjectToRegistry accepts one or more PSObjects from the pipeline and serializes
    their properties as registry values under a specified key path. The target path is
    always rooted under HKLM:\SOFTWARE or HKCU:\SOFTWARE.

    When multiple objects are piped in, each object is written to its own subkey. By
    default, subkeys are named numerically starting at 0. Use -UseLeadingZeros to pad
    numeric names to a consistent width, or -UseFirstPropertyAsKey to name each subkey
    after the value of the object's first property.

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

.PARAMETER RootKeyName
    The name of the top-level key to create or target under the hive's SOFTWARE key.
    For example, a value of 'MyApp' resolves to HKLM:\SOFTWARE\MyApp.

.PARAMETER SubKeyNames
    An optional ordered array of subkey names representing the path beneath RootKeyName
    where data will be written. For example, passing @('Config', 'Network') resolves
    to HKLM:\SOFTWARE\MyApp\Config\Network.

    Defaults to an empty array, writing directly under RootKeyName.

.PARAMETER UseFirstPropertyAsKey
    When specified, the value of each object's first property is used as the subkey
    name for that object rather than a numeric index. This is useful when objects
    have a natural identifier property such as a name or ID.

    If multiple objects share the same first property value, their data will collide
    under the same subkey. Use -EnsureUniqueness to handle this automatically.

.PARAMETER EnsureUniqueness
    When specified alongside -UseFirstPropertyAsKey, appends a numeric suffix to any
    subkey name that would otherwise collide with an already-written key. For example,
    if two objects both have a first property value of 'Service', the second will be
    written to a subkey named 'Service_1'.

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

.OUTPUTS
    None
        This function does not return output to the pipeline. All writes are performed
        against the registry.

.EXAMPLE
    [PSCustomObject]@{ Theme = 'Dark'; FontSize = 14; AutoSave = $true } |
        Add-PSObjectToRegistry -Hive HKCU -RootKeyName MyApp

    Writes a single object's properties directly to HKCU:\SOFTWARE\MyApp. No subkey
    is created because only one object was piped and -UseFirstPropertyAsKey was not
    specified.

.EXAMPLE
    Get-Service | Select-Object Name, DisplayName, Status |
        Add-PSObjectToRegistry -Hive HKLM -RootKeyName MyApp -SubKeyNames @('Services') -UseFirstPropertyAsKey

    Writes each service object to a subkey named after its Name property, e.g.
    HKLM:\SOFTWARE\MyApp\Services\Spooler, HKLM:\SOFTWARE\MyApp\Services\WinRM, etc.

.EXAMPLE
    Get-Service | Select-Object Name, DisplayName, Status |
        Add-PSObjectToRegistry -Hive HKLM -RootKeyName MyApp -SubKeyNames @('Services') `
            -UseFirstPropertyAsKey -EnsureUniqueness

    Same as the previous example, but if any two services share the same Name value,
    the second occurrence is written to a subkey with a numeric suffix rather than
    silently overwriting the first.

.EXAMPLE
    $data = Import-Csv .\Items.csv
    $data | Add-PSObjectToRegistry -Hive HKLM -RootKeyName MyApp -SubKeyNames @('Items') -UseLeadingZeros

    Imports a CSV and writes each row to a zero-padded numeric subkey. With 42 rows,
    keys will be named 00 through 41, ensuring consistent sort order.

.EXAMPLE
    Get-Process | Select-Object Name, Id, CPU |
        Add-PSObjectToRegistry -Hive HKLM -RootKeyName MyApp -WhatIf

    Previews all registry operations without writing any data. Use this to validate
    the target path and key structure before committing changes.

.NOTES
    - Writing to HKLM requires an elevated session. If the session is not elevated,
      a permissions error will be raised when the key is created.
    - Nested objects, arrays, and null values as property values are not supported
      and should be resolved or filtered out before calling this function.
    - This function is not transactional. If an error occurs mid-write, previously
      written keys and values will not be rolled back.

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
        [switch]$UseFirstPropertyAsKey,

        [Parameter(Position = 5)]
        [switch]$EnsureUniqueness,

        [Parameter(Position = 6)]
        [switch]$UseLeadingZeros
    )
    begin {
        [int]$ItemNumber = 0
        $Items = [System.Collections.Generic.List[object]]::new()

        $PathParts = @('SOFTWARE', $KeyName) + $SubKeyNames | Where-Object { $_ }
        $BasePath = Join-Path "${Hive}:" ($PathParts -join '\')

        if ($PSCmdlet.ShouldProcess($BasePath, 'Create registry key')) {
            if (-not (Test-Path $BasePath)) {
                New-Item -Path $BasePath -Force | Out-Null
            }
        }
    }
    process {
        foreach ($Item in $InputObject) { $Items.Add($Item) }
    }
    end {
        [string]$DValue = "D", ($Items.Count.ToString() -replace '\.|-|\+').Length -join ''
        foreach ($Item in $Items) {
            if ($Items.Count -eq 1 -and -not $UseFirstPropertyAsKey) {
                $TargetPath = $BasePath
            }
            elseif ($UseFirstPropertyAsKey -and $EnsureUniqueness) {
                $KeyName = ($Item.PSObject.Properties | Select-Object -First 1).Value
                $TargetPath = Join-Path $BasePath $KeyName
                if (Test-Path $TargetPath) {
                    $TargetPath = $TargetPath, ($ItemNumber++) -join '_'
                }
            }
            elseif ($UseFirstPropertyAsKey -and -not $EnsureUniqueness) {
                $KeyName = ($Item.PSObject.Properties | Select-Object -First 1).Value
                $TargetPath = Join-Path $BasePath $KeyName
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

            foreach ($prop in $Item.PSObject.Properties) {
                if ($PSCmdlet.ShouldProcess($TargetPath, "Set property '$($prop.Name)'")) {
                    Set-ItemProperty -Path $TargetPath -Name $prop.Name -Value $prop.Value
                }
            }
        }
    }
}

Export-ModuleMember -Function Add-PSObjectToRegistry
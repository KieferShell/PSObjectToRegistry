#Requires -Modules Pester
<#
    .SYNOPSIS
        Pester 5 test suite for the PSObjectToRegistry module.

    .DESCRIPTION
        Tests all permutations of input types, hives, cardinality, and switch
        combinations. HKLM tests are skipped automatically when the session is
        not elevated.

    .NOTES
        Requirements:
            Pester 5.x      Install-Module Pester -Force
            PowerShell 5.1+ or PowerShell 7+

        Usage:
            # Run all tests with detailed output
            Invoke-Pester -Path .\PSObjectToRegistry.Tests.ps1 -Output Detailed

            # Run only HKCU tests (no elevation required)
            Invoke-Pester -Path .\PSObjectToRegistry.Tests.ps1 -Tag 'HKCU' -Output Detailed

            # Run only HKLM tests (requires elevation)
            Invoke-Pester -Path .\PSObjectToRegistry.Tests.ps1 -Tag 'HKLM' -Output Detailed
#>

# Evaluated at discovery time - must be outside BeforeAll
$CurrentUser       = [Security.Principal.WindowsIdentity]::GetCurrent()
$CurrentPrincipal  = [Security.Principal.WindowsPrincipal]$CurrentUser
$IsElevated        = $CurrentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

BeforeAll {
    $ModulePath = Join-Path $PSScriptRoot 'PSObjectToRegistry.psm1'
    Import-Module $ModulePath -Force

    # Dedicated test root keys - isolated from any real application data.
    $script:HKCURoot = 'HKCU:\SOFTWARE\PSObjectToRegistry_Tests'
    $script:HKLMRoot = 'HKLM:\SOFTWARE\PSObjectToRegistry_Tests'

    # ------------------------------------------------------------------
    # Reusable dummy data sets
    # ------------------------------------------------------------------

    $script:SinglePSObject = [PSCustomObject]@{
        Name    = 'TestItem'
        Value   = 42
        Enabled = $true
    }

    $script:MultiplePSObjects = @(
        [PSCustomObject]@{ Name = 'Alpha'; Value = 1; Enabled = $true },
        [PSCustomObject]@{ Name = 'Beta'; Value = 2; Enabled = $false },
        [PSCustomObject]@{ Name = 'Gamma'; Value = 3; Enabled = $true }
    )

    # Two objects share the first property value 'Alpha' - used for collision tests.
    $script:DuplicatePSObjects = @(
        [PSCustomObject]@{ Name = 'Alpha'; Value = 1 },
        [PSCustomObject]@{ Name = 'Alpha'; Value = 2 },
        [PSCustomObject]@{ Name = 'Beta'; Value = 3 }
    )

    $script:SingleHashTable = @{
        Name    = 'TestItem'
        Value   = 42
        Enabled = $true
    }

    $script:MultipleHashTables = @(
        @{ Name = 'Alpha'; Value = 1; Enabled = $true },
        @{ Name = 'Beta'; Value = 2; Enabled = $false },
        @{ Name = 'Gamma'; Value = 3; Enabled = $true }
    )

    # Ten objects - triggers two-digit padding with UseLeadingZeros.
    $script:TenPSObjects = 0..9 | ForEach-Object {
        [PSCustomObject]@{ Index = $_; Label = "Item$_" }
    }

    # One hundred objects - triggers three-digit padding with UseLeadingZeros.
    $script:HundredPSObjects = 0..99 | ForEach-Object {
        [PSCustomObject]@{ Index = $_; Label = "Item$_" }
    }
}

AfterAll {
    if (Test-Path $script:HKCURoot) {
        Remove-Item -Path $script:HKCURoot -Recurse -Force
    }
    if ($IsElevated -and (Test-Path $script:HKLMRoot)) {
        Remove-Item -Path $script:HKLMRoot -Recurse -Force
    }
}

# ---------------------------------------------------------------------------
# Helper: wipe the test key so each test starts from a known clean state.
# ---------------------------------------------------------------------------
function Reset-TestKey {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force
    }
    New-Item -Path $Path -Force | Out-Null
}


# ===========================================================================
Describe 'Parameter Validation' -Tag 'Validation' {
    # ===========================================================================

    It 'Throws when an invalid Hive value is supplied' {
        { $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCC -KeyName 'Test' } |
        Should -Throw
    }

    It 'Throws when KeyName is omitted' {
        { $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU } |
        Should -Throw
    }

    It 'Throws when no InputObject is supplied' {
        { Add-PSObjectToRegistry -Hive HKCU -KeyName 'Test' } |
        Should -Throw
    }

    It 'Throws when PSObjects and HashTables are mixed in the same pipeline call' {
        $Mixed = @(
            [PSCustomObject]@{ Name = 'Alpha' },
            @{ Name = 'Beta' }
        )
        { $Mixed | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests' } |
        Should -Throw -ExpectedMessage '*HashTables or PSObjects, not both*'
    }

    It 'Emits a verbose message when UseFirstPropertyAsKey is used with multiple HashTables' {
        $VerboseOutput = $script:MultipleHashTables |
        Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests' `
            -UseFirstPropertyAsKey -Verbose 4>&1

        $VerboseOutput | Where-Object { $_ -match 'no effect' } | Should -Not -BeNullOrEmpty
    }

    It 'UseFirstPropertyAsKey has no effect when multiple HashTables are piped - data is still written' {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey

        # Should create numeric subkeys as normal, not named ones
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '0'
        $SubKeys | Should -Contain '1'
        $SubKeys | Should -Not -Contain 'Alpha'
    }
    
}


# ===========================================================================
Describe 'HKCU - Single PSObject' -Tag 'HKCU', 'PSObject', 'Single' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates the base key when it does not yet exist' {
        Remove-Item -Path $script:HKCURoot -Recurse -Force
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        Test-Path $script:HKCURoot | Should -BeTrue
    }

    It 'Writes all properties directly to the base key' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $Written = Get-ItemProperty -Path $script:HKCURoot
        $Written.Name  | Should -Be 'TestItem'
        $Written.Value | Should -Be 42
    }

    It 'Does not create any subkeys' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKCURoot -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'Produces no output to the pipeline' {
        $Result = $script:SinglePSObject |
        Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $Result | Should -BeNullOrEmpty
    }

    It 'Creates a named subkey when UseFirstPropertyAsKey is specified' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        Test-Path (Join-Path $script:HKCURoot 'TestItem') | Should -BeTrue
    }

    It 'Writes properties to the named subkey when UseFirstPropertyAsKey is specified' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        (Get-ItemProperty (Join-Path $script:HKCURoot 'TestItem')).Value | Should -Be 42
    }
}


# ===========================================================================
Describe 'HKCU - Single PSObject - SubKeyNames' -Tag 'HKCU', 'PSObject', 'Single', 'SubKeyNames' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates the correct nested path when a single SubKeyName is provided' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('Level1')
        Test-Path (Join-Path $script:HKCURoot 'Level1') | Should -BeTrue
    }

    It 'Creates all intermediate keys when multiple SubKeyNames are provided' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('L1', 'L2', 'L3')
        Test-Path (Join-Path $script:HKCURoot 'L1\L2\L3') | Should -BeTrue
    }

    It 'Writes data at the deepest SubKeyNames path' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('L1', 'L2')
        (Get-ItemProperty (Join-Path $script:HKCURoot 'L1\L2')).Name | Should -Be 'TestItem'
    }

    It 'Does not write data to intermediate keys, only to the deepest path' {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('L1', 'L2')
        $ShallowProps = Get-ItemProperty (Join-Path $script:HKCURoot 'L1') |
        Select-Object -ExcludeProperty PS*
        ($ShallowProps.PSObject.Properties |
        Where-Object { $_.MemberType -eq 'NoteProperty' }).Count | Should -Be 0
    }
}


# ===========================================================================
Describe 'HKCU - Multiple PSObjects - Numeric Subkeys' -Tag 'HKCU', 'PSObject', 'Multiple', 'Numeric' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates one numeric subkey per object starting at 0' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '0'
        $SubKeys | Should -Contain '1'
        $SubKeys | Should -Contain '2'
    }

    It 'Creates exactly as many subkeys as there are input objects' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 3
    }

    It 'Writes each object''s properties to the correct numeric subkey' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        (Get-ItemProperty (Join-Path $script:HKCURoot '0')).Name | Should -Be 'Alpha'
        (Get-ItemProperty (Join-Path $script:HKCURoot '1')).Name | Should -Be 'Beta'
        (Get-ItemProperty (Join-Path $script:HKCURoot '2')).Name | Should -Be 'Gamma'
    }

    It 'Writes all properties for each object, not just the first' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $Key0 = Get-ItemProperty (Join-Path $script:HKCURoot '0')
        $Key0.Value   | Should -Be 1
        $Key0.Enabled | Should -Be 'True'   # Registry stores booleans as DWord 1/0
    }
}


# ===========================================================================
Describe 'HKCU - Multiple PSObjects - UseLeadingZeros' -Tag 'HKCU', 'PSObject', 'Multiple', 'LeadingZeros' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Does not pad subkey names when fewer than 10 objects are piped' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '0'
        $SubKeys | Should -Not -Contain '00'
    }

    It 'Pads to two digits when exactly 10 objects are piped' {
        $script:TenPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '00'
        $SubKeys | Should -Contain '09'
        $SubKeys | Should -Not -Contain '0'
    }

    It 'Pads to three digits when 100 objects are piped' {
        $script:HundredPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '000'
        $SubKeys | Should -Contain '099'
        $SubKeys | Should -Not -Contain '00'
    }

    It 'Writes correct data to zero-padded subkeys' {
        $script:TenPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        (Get-ItemProperty (Join-Path $script:HKCURoot '00')).Index | Should -Be 0
        (Get-ItemProperty (Join-Path $script:HKCURoot '09')).Index | Should -Be 9
    }

    It 'Creates exactly as many subkeys as there are input objects' {
        $script:TenPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 10
    }
}


# ===========================================================================
Describe 'HKCU - Multiple PSObjects - UseFirstPropertyAsKey' -Tag 'HKCU', 'PSObject', 'Multiple', 'FirstProperty' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates a subkey named after the first property value of each object' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain 'Alpha'
        $SubKeys | Should -Contain 'Beta'
        $SubKeys | Should -Contain 'Gamma'
    }

    It 'Writes each object''s properties under its named subkey' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Alpha')).Value | Should -Be 1
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Beta')).Value  | Should -Be 2
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Gamma')).Value | Should -Be 3
    }

    It 'Creates exactly as many subkeys as there are input objects when names are unique' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 3
    }

    It 'Appends a suffix to ensure uniqueness when duplicate first property values exist' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        ($SubKeys | Where-Object { $_ -like 'Alpha*' }).Count | Should -Be 2
    }

    It 'Creates a total subkey count equal to the number of input objects even with duplicates' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 3
    }

    It 'Preserves unique entries alongside suffixed duplicate entries' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain 'Beta'
    }
}


# ===========================================================================
Describe 'HKCU - Multiple PSObjects - UseFirstPropertyAsKey + AllowOverwrite' -Tag 'HKCU', 'PSObject', 'Multiple', 'FirstProperty', 'AllowOverwrite' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Overwrites the subkey when duplicate first property values are present' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite
        ($SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName |
        Where-Object { $_ -like 'Alpha*' }).Count | Should -Be 1
    }

    It 'The last object written wins when AllowOverwrite is used on a duplicate key' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite
        # Alpha appears twice: Value=1 then Value=2 - the second write should win
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Alpha')).Value | Should -Be 2
    }

    It 'Does not affect non-duplicate entries' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Beta')).Value | Should -Be 3
    }

    It 'Results in fewer subkeys than input objects when duplicates exist' {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite
        # 3 input objects but 2 unique first-property values: Alpha and Beta
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 2
    }
}


# ===========================================================================
Describe 'HKCU - ResetKey' -Tag 'HKCU', 'ResetKey' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Does not throw when ResetKey is used and the base key does not yet exist' {
        Remove-Item -Path $script:HKCURoot -Recurse -Force
        { $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKCU `
                -KeyName 'PSObjectToRegistry_Tests' -ResetKey } | Should -Not -Throw
    }

    It 'Recreates the base key after deletion so new data can be written' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $script:SinglePSObject   | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -ResetKey
        Test-Path $script:HKCURoot | Should -BeTrue
    }

    It 'Removes stale numeric subkeys beyond the new dataset count' {
        # Write 5 items, then replace with 2 - subkeys 2, 3, 4 must be purged.
        0..4 | ForEach-Object { [PSCustomObject]@{ Index = $_ } } |
        Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'

        0..1 | ForEach-Object { [PSCustomObject]@{ Index = $_ } } |
        Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests' -ResetKey

        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 2
    }

    It 'Removes stale named subkeys when using UseFirstPropertyAsKey' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey

        [PSCustomObject]@{ Name = 'Delta'; Value = 99 } | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -ResetKey

        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Not -Contain 'Alpha'
        $SubKeys | Should -Not -Contain 'Beta'
        $SubKeys | Should -Not -Contain 'Gamma'
        $SubKeys | Should -Contain 'Delta'
    }

    It 'Combines correctly with UseLeadingZeros - stale wide-padded keys are removed' {
        # Write 100 items (three-digit padding), then reset with 10 (two-digit padding).
        $script:HundredPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros

        $script:TenPSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros -ResetKey

        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys.Count    | Should -Be 10
        $SubKeys          | Should -Contain '00'
        $SubKeys          | Should -Not -Contain '000'
    }

    It 'Combines correctly with UseFirstPropertyAsKey and AllowOverwrite after a reset' {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey

        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite -ResetKey

        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        # Gamma from the first write must be gone; Alpha and Beta from second write remain.
        $SubKeys | Should -Not -Contain 'Gamma'
        $SubKeys | Should -Contain 'Alpha'
        $SubKeys | Should -Contain 'Beta'
    }
}


# ===========================================================================
Describe 'HKCU - Single HashTable' -Tag 'HKCU', 'HashTable', 'Single' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates the base key when it does not yet exist' {
        Remove-Item -Path $script:HKCURoot -Recurse -Force
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        Test-Path $script:HKCURoot | Should -BeTrue
    }

    It 'Writes all HashTable entries as registry values to the base key' {
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $Written = Get-ItemProperty -Path $script:HKCURoot
        $Written.Name  | Should -Be 'TestItem'
        $Written.Value | Should -Be 42
    }

    It 'Does not create any subkeys for a single HashTable' {
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKCURoot -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'Writes a single HashTable to the correct path when SubKeyNames are provided' {
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('Config')
        (Get-ItemProperty (Join-Path $script:HKCURoot 'Config')).Name | Should -Be 'TestItem'
    }

    It 'Does not throw when UseFirstPropertyAsKey is specified with a single HashTable' {
        { $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU `
                -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey } | Should -Not -Throw
    }

    It 'Still writes data to the base key when UseFirstPropertyAsKey is specified with a single HashTable' {
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        (Get-ItemProperty -Path $script:HKCURoot).Value | Should -Be 42
    }
}


# ===========================================================================
Describe 'HKCU - Multiple HashTables - Numeric Subkeys' -Tag 'HKCU', 'HashTable', 'Multiple', 'Numeric' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Creates one numeric subkey per HashTable starting at 0' {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '0'
        $SubKeys | Should -Contain '1'
        $SubKeys | Should -Contain '2'
    }

    It 'Creates exactly as many subkeys as there are input HashTables' {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 3
    }

    It 'Writes all entries from each HashTable to its respective subkey' {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'
        # HashTable enumeration order is not guaranteed - verify by collecting all written Name values.
        $AllNames = Get-ChildItem -Path $script:HKCURoot | ForEach-Object {
            (Get-ItemProperty -Path $_.PSPath).Name
        }
        $AllNames | Should -Contain 'Alpha'
        $AllNames | Should -Contain 'Beta'
        $AllNames | Should -Contain 'Gamma'
    }
}


# ===========================================================================
Describe 'HKCU - Multiple HashTables - UseLeadingZeros' -Tag 'HKCU', 'HashTable', 'Multiple', 'LeadingZeros' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Pads subkey names to two digits when exactly 10 HashTables are piped' {
        $TenTables = 0..9 | ForEach-Object { @{ Index = $_; Label = "Item$_" } }
        $TenTables | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '00'
        $SubKeys | Should -Contain '09'
        $SubKeys | Should -Not -Contain '0'
    }

    It 'Does not pad subkey names when fewer than 10 HashTables are piped' {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKCU `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKCURoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '0'
        $SubKeys | Should -Not -Contain '00'
    }
}


# ===========================================================================
Describe 'HKCU - Multiple HashTables - ResetKey' -Tag 'HKCU', 'HashTable', 'ResetKey' {
    # ===========================================================================

    BeforeEach { Reset-TestKey -Path $script:HKCURoot }

    It 'Removes stale numeric HashTable subkeys beyond the new dataset count' {
        $Large = 0..4 | ForEach-Object { @{ Index = $_ } }
        $Large | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests'

        $Small = 0..1 | ForEach-Object { @{ Index = $_ } }
        $Small | Add-PSObjectToRegistry -Hive HKCU -KeyName 'PSObjectToRegistry_Tests' -ResetKey

        (Get-ChildItem -Path $script:HKCURoot).Count | Should -Be 2
    }
}


# ===========================================================================
Describe 'HKLM - All Permutations' -Tag 'HKLM' {
    # ===========================================================================

    BeforeAll {
        if (-not $IsElevated) {
            Write-Warning 'HKLM tests require an elevated (Administrator) session and will be skipped.'
        }
    }

    AfterAll {
        if ($IsElevated -and (Test-Path $script:HKLMRoot)) {
            Remove-Item -Path $script:HKLMRoot -Recurse -Force
        }
    }

    BeforeEach {
        if ($IsElevated) { Reset-TestKey -Path $script:HKLMRoot }
    }

    It 'Creates the base key under HKLM' -Skip:(-not $IsElevated) {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'
        Test-Path $script:HKLMRoot | Should -BeTrue
    }

    It 'Writes a single PSObject flat to the HKLM base key' -Skip:(-not $IsElevated) {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'
        $Written = Get-ItemProperty -Path $script:HKLMRoot
        $Written.Name  | Should -Be 'TestItem'
        $Written.Value | Should -Be 42
    }

    It 'Creates numeric subkeys under HKLM for multiple PSObjects' -Skip:(-not $IsElevated) {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKLMRoot).Count | Should -Be 3
    }

    It 'Applies UseLeadingZeros correctly under HKLM' -Skip:(-not $IsElevated) {
        $script:TenPSObjects | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseLeadingZeros
        $SubKeys = Get-ChildItem -Path $script:HKLMRoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain '00'
        $SubKeys | Should -Contain '09'
        $SubKeys | Should -Not -Contain '0'
    }

    It 'Applies UseFirstPropertyAsKey correctly under HKLM' -Skip:(-not $IsElevated) {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        $SubKeys = Get-ChildItem -Path $script:HKLMRoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Contain 'Alpha'
        $SubKeys | Should -Contain 'Beta'
        $SubKeys | Should -Contain 'Gamma'
    }

    It 'Appends uniqueness suffixes under HKLM when duplicate first property values exist' -Skip:(-not $IsElevated) {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey
        $SubKeys = Get-ChildItem -Path $script:HKLMRoot | Select-Object -ExpandProperty PSChildName
        ($SubKeys | Where-Object { $_ -like 'Alpha*' }).Count | Should -Be 2
    }

    It 'AllowOverwrite collapses duplicate named subkeys under HKLM' -Skip:(-not $IsElevated) {
        $script:DuplicatePSObjects | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -AllowOverwrite
        (Get-ChildItem -Path $script:HKLMRoot).Count | Should -Be 2
    }

    It 'ResetKey removes stale numeric subkeys under HKLM' -Skip:(-not $IsElevated) {
        0..4 | ForEach-Object { [PSCustomObject]@{ Index = $_ } } |
        Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'

        0..1 | ForEach-Object { [PSCustomObject]@{ Index = $_ } } |
        Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests' -ResetKey

        (Get-ChildItem -Path $script:HKLMRoot).Count | Should -Be 2
    }

    It 'ResetKey removes stale named subkeys under HKLM' -Skip:(-not $IsElevated) {
        $script:MultiplePSObjects | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey

        [PSCustomObject]@{ Name = 'Delta'; Value = 99 } | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -UseFirstPropertyAsKey -ResetKey

        $SubKeys = Get-ChildItem -Path $script:HKLMRoot | Select-Object -ExpandProperty PSChildName
        $SubKeys | Should -Not -Contain 'Alpha'
        $SubKeys | Should -Contain 'Delta'
    }

    It 'Writes a single HashTable flat to the HKLM base key' -Skip:(-not $IsElevated) {
        $script:SingleHashTable | Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'
        (Get-ItemProperty -Path $script:HKLMRoot).Name | Should -Be 'TestItem'
    }

    It 'Creates numeric subkeys under HKLM for multiple HashTables' -Skip:(-not $IsElevated) {
        $script:MultipleHashTables | Add-PSObjectToRegistry -Hive HKLM -KeyName 'PSObjectToRegistry_Tests'
        (Get-ChildItem -Path $script:HKLMRoot).Count | Should -Be 3
    }

    It 'Writes to a SubKeyNames path under HKLM' -Skip:(-not $IsElevated) {
        $script:SinglePSObject | Add-PSObjectToRegistry -Hive HKLM `
            -KeyName 'PSObjectToRegistry_Tests' -SubKeyNames @('Config', 'Network')
        $TargetPath = Join-Path $script:HKLMRoot 'Config\Network'
        (Get-ItemProperty -Path $TargetPath).Name | Should -Be 'TestItem'
    }
}
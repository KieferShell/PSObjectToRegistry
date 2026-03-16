# PSObjectToRegistry

PowerShell module for transposing PSObjects from the pipeline to the Windows registry.

## Install

### PowerShell Gallery Install

```powershell
Install-Module -Name PSObjectToRegistry
```

See the [PowerShell Gallery](http://www.powershellgallery.com/packages/PSObjectToRegistry/) for the complete details and
instructions.

## Quick Start (After Installation)

PSObjectToRegistry was built to accept collections of PSObjects or HashTables, as well as single/individual PSObjects or HashTables. For example, let's write some information about Windows services to the registry using Add-PSObjectToRegistry:

```powershell
Get-Service | Select-Object Name, DisplayName, Status | Add-PSObjectToRegistry -Hive HKLM -KeyName MyOrg -SubKeyNames @('Services') -UseFirstPropertyAsKey
```

The above example does the following:
1) Gathers all local Windows services via Get-Service
2) Sends the collection of services down the pipeline to Select-Object where we select the 'Name', 'DisplayName' and 'Status' properties
3) Sends the collection of selected properties down the pipeline to Add-PSObjectToRegistry where we specify the following parameters:
  a) -Hive: **HKLM** (This selects the HKEY_LOCAL_MACHINE hive)
  b) -KeyName: **MyOrg** (This determines the base key used after HKLM:\Software)
  c) -SubKeyName **@('Services')** (This provides an ordered array of subkeys to use within HKLM:\Software\MyOrg to further sort your data)
  d) -UseFirstPropertyAsKey (This specifies that the first property, in our case 'Name' will be used as the key name when adding the data to the registry path specified with the prior parameters)

The result is a series of keys within HKLM:\Software\MyOrg\Services named for each service 'Name' property containing the selected properties 'Name', 'DisplayName' and 'Status' as values within each key, e.g.

HKLM:\SOFTWARE\MyOrg\Services\Spooler, HKLM:\SOFTWARE\MyOrg\Services\WinRM, etc.

## Getting Help

```powershell
Get-Help Add-PSObjectToRegistry -Full
```

## License

- see [LICENSE](LICENSE) file




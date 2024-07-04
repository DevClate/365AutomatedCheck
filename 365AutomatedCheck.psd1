#
# Module manifest for module '365AutomatedCheck'
#
# Generated by: Clayton Tyger
#
# Generated on: 6/14/2024
#

@{

# Script module or binary module file associated with this manifest.
RootModule = '365AutomatedCheck.psm1'

# Version number of this module.
ModuleVersion = '0.0.7'

# Supported PSEditions
CompatiblePSEditions = 'Core'

# ID used to uniquely identify this module
GUID = 'a745cd66-f76c-4767-8b3c-69993a604b80'

# Author of this module
Author = 'Clayton Tyger'

# Company or vendor of this module
CompanyName = 'Clayton Tyger'

# Copyright statement for this module
Copyright = '(c) Clayton Tyger. All rights reserved.'

# Description of the functionality provided by this module
Description = 'This module checks for your fields in your Office 365 tenant to see if they meet company standards.'

# Minimum version of the PowerShell engine required by this module
PowerShellVersion = '7.1'

# Name of the PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
# ClrVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
RequiredModules = @(
    @{ ModuleName='ImportExcel'; ModuleVersion='7.8.2' }
    @{ ModuleName='ExchangeOnlineManagement'; ModuleVersion='2.0.6' }
    @{ ModuleName='Microsoft.Graph.Users'; ModuleVersion='1.17.0' }
    @{ ModuleName='Microsoft.Graph.Groups'; ModuleVersion='1.17.0' }
    @{ ModuleName='Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion='1.17.0' }
    @{ ModuleName='Microsoft.Graph.Users.Actions'; ModuleVersion='1.17.0' }
    @{ ModuleName='PSFramework'; ModuleVersion='1.8.289' }
    @{ ModuleName='Pester'; ModuleVersion='5.3.0' }
)

# Assemblies that must be loaded prior to importing this module
# RequiredAssemblies = @()

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
FunctionsToExport = @(
    'Invoke-365AutomatedCheck',
    'Test-365ACMobilePhone',
    'Export-365ACResultToHtml',
    'Test-365ACCompanyName',
    'Export-365ACResultToExcel',
    'Test-365ACDepartment',
    'Test-365ACFaxNumber',
    'Test-365ACJobTitle',
    'Convert-365ACXmlToHtml',
    'Get-365ACPesterConfiguration', 
    'Set-365ACDefaultOutputPath'
)

# Cmdlets to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no cmdlets to export.
CmdletsToExport = @()

# Variables to export from this module
VariablesToExport = '*'

# Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
AliasesToExport = @()

# DSC resources to export from this module
# DscResourcesToExport = @()

# List of all modules packaged with this module
# ModuleList = @()

# List of all files packaged with this module
# FileList = @()

# Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
PrivateData = @{

    PSData = @{

        # Tags applied to this module. These help with module discovery in online galleries.
        # Tags = @()

        # A URL to the license for this module.
        # LicenseUri = ''

        # A URL to the main website for this project.
        # ProjectUri = ''

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

        # Prerelease string of this module
        # Prerelease = ''

        # Flag to indicate whether the module requires explicit user acceptance for install/update/save
        # RequireLicenseAcceptance = $false

        # External dependent modules of this module
        # ExternalModuleDependencies = @()

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}


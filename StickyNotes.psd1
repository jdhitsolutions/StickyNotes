#
# Module manifest for module 'StickyNotes'
#
#

@{

# Script module or binary module file associated with this manifest.
RootModule = 'StickyNotes.psm1'

# Version number of this module.
ModuleVersion = '1.2.0.3'

# ID used to uniquely identify this module
GUID = 'ed5922e6-152d-482b-bd7b-060fc3679246'

# Author of this module
Author = 'Jeff Hicks @(JeffHicks)'

# Company or vendor of this module
CompanyName = 'JDH Information Technology Solutions, Inc'

# Copyright statement for this module
Copyright = '(c) 2016 Jeff Hicks @(JeffHicks). All rights reserved.'

# Description of the functionality provided by this module
Description = 'A set of PowerShell commands for working with desktop sticky notes'

# Minimum version of the Windows PowerShell engine required by this module
PowerShellVersion = '5.0'

# Name of the Windows PowerShell host required by this module
# PowerShellHostName = ''

# Minimum version of the Windows PowerShell host required by this module
# PowerShellHostVersion = ''

# Minimum version of Microsoft .NET Framework required by this module
# DotNetFrameworkVersion = ''

# Minimum version of the common language runtime (CLR) required by this module
# CLRVersion = ''

# Processor architecture (None, X86, Amd64) required by this module
# ProcessorArchitecture = ''

# Modules that must be imported into the global environment prior to importing this module
# RequiredModules = @()

# Assemblies that must be loaded prior to importing this module
RequiredAssemblies = 'System.Windows.Forms','Microsoft.VisualBasic','PresentationCore','System.Drawing'

# Script files (.ps1) that are run in the caller's environment prior to importing this module.
# ScriptsToProcess = @()

# Type files (.ps1xml) to be loaded when importing this module
# TypesToProcess = @()

# Format files (.ps1xml) to be loaded when importing this module
# FormatsToProcess = @()

# Modules to import as nested modules of the module specified in RootModule/ModuleToProcess
# NestedModules = @()

# Functions to export from this module
FunctionsToExport = 'New-StickyNote', 'Remove-StickyNote', 'Set-StickyNote', 'Find-StickyNote', 'Restore-StickyNote'

# Cmdlets to export from this module
# CmdletsToExport = '*'

# Variables to export from this module
# VariablesToExport = '*'

# Aliases to export from this module
AliasesToExport = 'nn', 'rn', 'sn','kn'

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
        Tags = 'Desktop', 'Notes', 'StickyNote'

        # A URL to the license for this module.
        LicenseUri = 'https://github.com/jdhitsolutions/StickyNotes/blob/master/License.txt'

        # A URL to the main website for this project.
        ProjectUri = 'https://github.com/jdhitsolutions/StickyNotes'

        # A URL to an icon representing this module.
        # IconUri = ''

        # ReleaseNotes of this module
        # ReleaseNotes = ''

    } # End of PSData hashtable

} # End of PrivateData hashtable

# HelpInfo URI of this module
# HelpInfoURI = ''

# Default prefix for commands exported from this module. Override the default prefix using Import-Module -Prefix.
# DefaultCommandPrefix = ''

}


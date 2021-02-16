@{
    # Version number of this module.
    ModuleVersion        = '1.0.3'

    # Script module or binary module file associated with this manifest.
    RootModule           = 'OutSpeech.psm1'

    # Supported PSEditions
    CompatiblePSEditions = 'Desktop'

    # ID used to uniquely identify this module
    GUID                 = 'cc906516-9201-4bc4-83e4-547e4326c4a2'

    # Author of this module
    Author               = 'Mike Campbell'
    CompanyName          = 'themodulecollective'
    Copyright            = '2021'

    # Description of the functionality provided by this module
    Description          = 'Provides Out-Speech and Export-Speech functions along with customized SpeechConfiguration support.  Uses System.Speech.Synthesis.SpeechSynthesizer from .net framework.'

    # Minimum version of the PowerShell engine required by this module
    PowerShellVersion    = '5.1'

    # Name of the PowerShell host required by this module
    # PowerShellHostName = ''

    # Minimum version of the PowerShell host required by this module
    # PowerShellHostVersion = ''

    # Minimum version of Microsoft .NET Framework required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # DotNetFrameworkVersion = ''

    # Minimum version of the common language runtime (CLR) required by this module. This prerequisite is valid for the PowerShell Desktop edition only.
    # CLRVersion = ''

    # Type files (.ps1xml) to be loaded when importing this module
    # TypesToProcess = @()

    # Format files (.ps1xml) to be loaded when importing this module
    # FormatsToProcess = @()

    # Functions to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no functions to export.
    FunctionsToExport    = @(
        'Enable-SpeechConfiguration'
        'Set-SpeechConfiguration'
        'Disable-SpeechConfiguration'
        'Out-Speech'
        'Export-Speech'
        'Get-SpeechVoice'
        'Get-SpeechConfiguration'
    )

    # Variables to export from this module
    #VariablesToExport = @()

    # Aliases to export from this module, for best performance, do not use wildcards and do not delete the entry, use an empty array if there are no aliases to export.
    #AliasesToExport = @()

    # Private data to pass to the module specified in RootModule/ModuleToProcess. This may also contain a PSData hashtable with additional module metadata used by PowerShell.
    PrivateData          = @{

        PSData = @{

            # Tags applied to this module. These help with module discovery in online galleries.
            Tags       = @('Speech', 'Synthesis', 'Voice')

            # A URL to the license for this module.
            LicenseUri = 'https://github.com/themodulecollective/OutSpeech/blob/master/LICENSE'

            # A URL to the main website for this project.
            ProjectUri = 'https://github.com/themodulecollective/OutSpeech'

            # A URL to an icon representing this module.
            # IconUri = ''

            # ReleaseNotes of this module
            # ReleaseNotes = ''

        } # End of PSData hashtable

    } # End of PrivateData hashtable

    # HelpInfo URI of this module
    # HelpInfoURI = ''
}


Function Enable-SpeechConfiguration
{
    <#
    .SYNOPSIS
        Creates a SpeechConfiguration (System.Speech.Synthesis.SpeechSynthesizer object) with the specified values
    .DESCRIPTION
        Creates a SpeechConfiguration (System.Speech.Synthesis.SpeechSynthesizer object) with the specified values for Rate, Volume, and Voice.
        SpeechConfiguration objects are stored in the module variable SpeechConfigurations and can be retrieved, set, or deleted using the
        Get-SpeechConfiguration, Set-SpeechConfiguration, or Disable-SpeechConfiguration functions.
    .EXAMPLE
        PS C:\Enable-SpeechConfiguration -Rate 3 -Volume 50 -Voice 'Microsoft David Desktop' -ConfigurationName 'DavidR3V50'
    .PARAMETER Rate
        Specifies the speech rate with higher values being faster.  Valid range is -10 through 10.  Default is 0.
    .PARAMETER Volume
        Specifies the speech volume with higher values being louder.  Valid range is 1 through 100.  Default is 100.
    .PARAMETER Voice
        Specifies the speech voice to user.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
    .PARAMETER ConfigurationName
        Specifies the ConfigurationName to create.  Default is 'Default'.
    #>
    [cmdletbinding()]
    param
    (
        [parameter()]
        [string]$ConfigurationName = 'Default'
        ,
        [parameter()]
        [string]$Voice
        ,
        [parameter()]
        [ValidateRange(-10, 10)]
        [Int]$Rate
        ,
        [parameter()]
        [ValidateRange(1, 100)]
        $Volume
    )
    if ($Script:SpeechConfigurations.ContainsKey($ConfigurationName))
    {
        throw ("SpeechConfiguration $ConfigurationName already exists.  Use Disable-Speech to remove it or Set-Speech to modify it.")
        Return
    }
    $InitializeSpeechParams =
    @{
        ErrorAction         = 'Stop'
        ConfigurationName = $ConfigurationName
    }
    foreach ($param in $PSBoundParameters.Keys)
    {
        if ($param -in 'Rate', 'Volume', 'Voice', 'Verbose')
        {
            $InitializeSpeechParams.$param = $($PSBoundParameters.$param)
        }
    }
    Initialize-Speech @InitializeSpeechParams
}#function Enable-Speech
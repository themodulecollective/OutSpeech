Function Set-SpeechConfiguration
{
        <#
    .SYNOPSIS
        Modifies a SpeechConfiguration (System.Speech.Synthesis.SpeechSynthesizer object) with the specified values
    .DESCRIPTION
        Modifies a SpeechConfiguration (System.Speech.Synthesis.SpeechSynthesizer object) with the specified values for Rate, Volume, and Voice.
        SpeechConfiguration objects are stored in the module variable SpeechConfigurations and can be retrieved, created, or deleted using the
        Get-SpeechConfiguration, Enable-SpeechConfiguration, or Disable-SpeechConfiguration functions.
    .EXAMPLE
        PS C:\Enable-SpeechConfiguration -Voice 'Microsoft David Desktop' -ConfigurationName 'DavidR3V50'
        PS C:\Set-SpeechConfiguration -Rate 3 -Volume 50 -Voice 'Microsoft David Desktop' -ConfigurationName 'DavidR3V50'

        Modifies the Speech Configuration 'DavidR3V50' with the specifed Rate, Volume, and Voice values.
    .PARAMETER Rate
        Specifies the speech rate with higher values being faster.  Valid range is -10 through 10.  Default is 0.
    .PARAMETER Volume
        Specifies the speech volume with higher values being louder.  Valid range is 1 through 100.  Default is 100.
    .PARAMETER Voice
        Specifies the speech voice to user.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
    .PARAMETER ConfigurationName
        Specifies the ConfigurationName to modify.  Default is 'Default'.
    #>
    [cmdletbinding()]
    param
    (
        [parameter()]
        $ConfigurationName = 'Default'
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
    if ($script:SpeechConfigurations.ContainsKey($ConfigurationName))
    {
        $SpeechConfiguration = $script:SpeechConfigurations.$ConfigurationName
    }
    else
    {
        throw ("SpeechConfiguration $ConfigurationName does not exist.  Use Enable-SpeechConfiguration to create it.")
        Return
    }
    If ($PSBoundParameters.ContainsKey('Volume'))
    {
        Write-Verbose "Setting SpeechConfiguration $ConfigurationName volume to $Volume"
        $SpeechConfiguration.Volume = $Volume
    }
    Else
    {
        Write-Verbose "SpeechConfiguration $ConfigurationName volume = $($SpeechConfiguration.volume)"
    }
    If ($PSBoundParameters.ContainsKey('Rate'))
    {
        Write-Verbose "Setting SpeechConfiguration $ConfigurationName rate to $Rate"
        $SpeechConfiguration.Rate = $Rate
    }
    Else
    {
        Write-Verbose "SpeechConfiguration $ConfigurationName rate =  $($SpeechConfiguration.rate)"
    }
    If ($PSBoundParameters.ContainsKey('Voice'))
    {
        Write-Verbose "Setting SpeechConfiguration voice to $Voice"
        $SpeechConfiguration.SelectVoice($Voice)
    }
    Else
    {
        Write-Verbose "SpeechConfiguration $ConfigurationName voice = $($SpeechConfiguration.voice.name)"
    }
}

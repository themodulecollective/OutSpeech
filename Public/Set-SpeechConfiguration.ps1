Function Set-SpeechConfiguration
{
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

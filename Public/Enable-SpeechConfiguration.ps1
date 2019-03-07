
Function Enable-SpeechConfiguration
{
    [cmdletbinding()]
    param
    (
        [parameter()]
        [string]$ConfigurationName
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
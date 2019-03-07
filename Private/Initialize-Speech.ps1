function Initialize-Speech
{
    <#
    .SYNOPSIS
        Creates a System.Speech.Synthesis.SpeechSynthesizer object with the specified values
    .DESCRIPTION
        Creates a System.Speech.Synthesis.SpeechSynthesizer object and configures the specified values for Rate, Volume, and Voice
    .EXAMPLE
        PS C:\Initialize-Speech -Rate 3 -Volume 50 -Voice 'Microsoft David Desktop' -ConfigurationName 'DavidR3V50'
    .PARAMETER Rate
        Specifies the speech rate with higher values being faster.  Valid range is -10 through 10.  Default is 0.
    .PARAMETER Volume
        Specifies the speech volume with higher values being louder.  Valid range is 1 through 100.  Default is 100.
    .PARAMETER Voice
        Specifies the speech voice to user.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
    .PARAMETER ConfigurationName
        Specifies the ConfigurationName to use.  Default is NULL.
    #>
    [cmdletbinding()]
    param(
        $Rate
        ,
        $Volume
        ,
        $Voice
        ,
        $ConfigurationName
    )
    Write-Verbose "Loading System.speech assembly"
    Add-Type -AssemblyName System.speech -ErrorAction Stop
    Write-Verbose "Creating Speech object"
    $SpeechConfiguration = New-Object System.Speech.Synthesis.SpeechSynthesizer -ErrorAction Stop
    foreach ($param in $PSBoundParameters.Keys)
    {
        if ($param -in 'Rate', 'Volume')
        {
            $SpeechConfiguration.$param = $($PSBoundParameters.$param)
        }
        if ($param -in 'Voice')
        {
            $SpeechConfiguration.SelectVoice($($PSBoundParameters.$param))
        }
    }
    $Script:SpeechConfigurations.$ConfigurationName = $SpeechConfiguration
}

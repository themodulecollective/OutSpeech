function Initialize-Speech
{
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
}#function Initialize-Speech
Function Disable-SpeechConfiguration
{
    <#
    .SYNOPSIS
        Disables (deletes) a SpeechConfiguration from the OutSpeech module's SpeechConfigurations variable.
    .DESCRIPTION
        Disables (deletes) a SpeechConfiguration from the OutSpeech module's SpeechConfigurations variable.
    .EXAMPLE
        PS C:\> Enable-SpeechConfiguration -ConfigurationName 'TooFastTooLoud' -Rate 10 -Volume 100
        PS C:\> Get-SpeechConfiguration -ConfigurationName 'TooFastTooLoud'

        State Rate Volume Voice
        ----- ---- ------ -----
        Ready   10    100 System.Speech.Synthesis.VoiceInfo

        PS C:\> Disable-SpeechConfiguration -ConfigurationName 'TooFastTooLoud'
        PS C:\> Get-SpeechConfiguration -ConfigurationName 'TooFastTooLoud'

        # No output expected
    .PARAMETER ConfigurationName
    The name of an existing SpeechConfiguration to Disable (delete).
    #>
    [cmdletbinding()]
    param
    (
        [Parameter()]
        [string[]]$ConfigurationName
    )
    foreach ($cn in $ConfigurationName){
        if ($Script:SpeechConfigurations.ContainsKey($cn))
        {
            $Script:SpeechConfigurations.$cn.dispose()
            $Script:SpeechConfigurations.remove($cn)
        }
        Else
        {
            Write-Warning "SpeechConfiguration $cn does not exist."
        }
    }
}#Function Disable-SpeechConfiguration
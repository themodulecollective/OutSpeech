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
    .PARAMETER All
        Specifying All causes Disable-SpeechConfiguration to Disable (delete) all currently configured SpeechConfigurations.
    #>
    [cmdletbinding(DefaultParameterSetName = 'NamedConfig')]
    param
    (
        [Parameter(ParameterSetName = 'NamedConfig')]
        [string[]]$ConfigurationName
        ,
        [Parameter(ParameterSetName = 'All')]
        [switch]$All
    )
    switch ($PSCmdlet.ParameterSetName)
    {
        'NamedConfig'
        {
            foreach ($cn in $ConfigurationName)
            {
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
        }
        'All'
        {
            $Keys = $Script:SpeechConfigurations.Keys | ForEach-Object { $_ } #disconnect the Keys from the actual hashtable object
            foreach ($k in $Keys)
            {
                $Script:SpeechConfigurations.$k.dispose()
                $Script:SpeechConfigurations.remove($k)
            }
        }
    }
}#Function Disable-SpeechConfiguration
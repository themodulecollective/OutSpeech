Function Export-Speech
{
    <#
    .SYNOPSIS
        Exports the specified input text or objects (converting them to a string with out-string) to a specified Wave file.
    .DESCRIPTION
        Exports the specified input text or objects (converting them to a string with out-string) to a specified Wave file using the Path parameter.
        A SpeechConfiguration can be configured with the ConfigurationName parameter or the rate, volume, or voice parameters can be used to adjust the output.
    .EXAMPLE
        PS C:\> Enable-SpeechConfiguration -ConfigurationName 'TooFastTooLoud' -Rate 10 -Volume 100
        PS C:\> Export-Speech -Path MyWaveFile.wav -ConfigurationName 'TooFastTooLoud' -InputObject "Here is a sentence read by the computer."

        PS C:\> Get-Item MyWaveFile.wav
        PS C:\> Invoke-Item MyWaveFile.wav
    .PARAMETER InputObject
        Strings or objects that will be converted to speech and sent to the default audio device.
    .PARAMETER Path
        Specify a valid path to a wav file to store the audio output export.  File does not need to exist but an existing file will be clobbered without warning.
    .PARAMETER Rate
        Specifies the speech rate with higher values being faster.  Valid range is -10 through 10.  Default is 0.
        If an existing SpeechConfiguration is specified with the ConfigurationName parameter this will modify that SpeechConfiguration's rate value.
    .PARAMETER Volume
        Specifies the speech volume with higher values being louder.  Valid range is 1 through 100.  Default is 100.
        If an existing SpeechConfiguration is specified with the ConfigurationName parameter this will modify that SpeechConfiguration's volume value.
    .PARAMETER Voice
        Specifies the speech voice to use.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
        If an existing SpeechConfiguration is specified with the ConfigurationName parameter this will modify that SpeechConfiguration's voice.
    .PARAMETER ConfigurationName
        Specifies the ConfigurationName to use.  Default is 'Default'.
        If Rate, Volume, or Voice are specified the associated SpeechConfiguration's value for that attribute will be modified.
    #>
    [cmdletbinding()]
    param
    (
        [Parameter(ValueFromPipeline = 'True')]
        [string[]]$inputobject
        ,
        [Parameter()]
        $ConfigurationName = 'Default'
        ,
        [Parameter()]
        [string]$voice
        ,
        [Parameter()]
        [ValidateRange(-10, 10)]
        [Int]$Rate
        ,
        [Parameter()]
        [ValidateRange(1, 100)]
        $Volume
        ,
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    begin
    {
        $SpeechParams = @{
            ErrorAction       = 'Stop'
            ConfigurationName = $ConfigurationName
        }
        foreach ($param in $PSBoundParameters.Keys)
        {
            if ($param -in 'Rate', 'Volume', 'Voice', 'Verbose')
            {
                $SpeechParams.$param = $($PSBoundParameters.$param)
            }
        }
        switch ($script:SpeechConfigurations.ContainsKey($ConfigurationName))
        {
            $true
            {
                #set
                Set-SpeechConfiguration @SpeechParams
            }
            $false
            {
                #enable
                Enable-SpeechConfiguration @SpeechParams
            }
        }
        $SpeechConfiguration = $script:SpeechConfigurations.$ConfigurationName
        Write-Verbose "Setting speech to output to wavfile: $Path"
        $SpeechConfiguration.SetOutputToWaveFile($Path)
    }#begin
    Process
    {
        ForEach ($line in $inputobject)
        {
            Write-Verbose "Exporting: $line"
            $null = $SpeechConfiguration.Speak(($line | Out-String))
        }
    }#Process
    End
    {
        $SpeechConfiguration.SetOutputToNull()
        $SpeechConfiguration.SetOutputToDefaultAudioDevice()
    }
}#Function Export-Speech
Function Export-Speech
{
    #.PARAMETER Path
    # output to a Waveform audio format file
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
            ErrorAction         = 'Stop'
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
        $SpeechConfiguration.SetOutputToDefaultAudioDevice()
    }
}#Function Export-Speech
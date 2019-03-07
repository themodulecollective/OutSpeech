
Function Out-Speech
{
<#
.SYNOPSIS
    Accepts string or object input (non-string objects will be converted to strings using out-string) and outputs it to speech.
.DESCRIPTION
    Accepts string or object input (non-string objects will be converted to strings using out-string) and outputs it to speech
    using the specified SpeechConfiguration and/or Rate, Volume, and Voice. The default output method is asynchronous (audio outputs while PowerShell continues).
.PARAMETER InputObject
    Strings or objects that will be converted to speech and sent to the default audio device.
.PARAMETER Rate
    Specifies the speech rate with higher values being faster.  Valid range is -10 through 10.  Default is 0.
.PARAMETER Volume
    Specifies the speech volume with higher values being louder.  Valid range is 1 through 100.  Default is 100.
.PARAMETER Voice
    Specifies the speech voice to user.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
.PARAMETER ConfigurationName
    Specifies the ConfigurationName to use.  Default is 'Default'.
.PARAMETER SynchronousOutput
    Makes the audio output Synchronous (PowerShell pauses until the audio output has completed).

.EXAMPLE
    "This is a test" | Out-Speech

    Description
    -----------
    Speaks the string that was given to the function in the pipeline.
.EXAMPLE
    Enable-SpeechConfiguration -ConfigurationName 'TooSlowTooSoft' -rate -2 -volume 5
    "Today's date is $((get-date).toshortdatestring())" | Out-Speech -ConfigurationName 'TooSlowTooSoft'

    Description
    -----------
    Sends 'Today's date is ______' to the default audio device.

.EXAMPLE
    Enable-SpeechConfiguration -ConfigurationName 'TooSlowTooSoft' -rate -2 -volume 5
    "Today's date is $((get-date).toshortdatestring())" | Out-Speech -ConfigurationName 'TooSlowTooSoft' -SynchronousOutput

    Description
    -----------
    Sends 'Today's date is ______' to the default audio device and PowerShell waits for the output to complete.
#>
    [cmdletbinding(DefaultParameterSetName = 'ASync')]
    Param
    (
        [parameter(ParameterSetName = 'ASync', ValueFromPipeline = $true, Position = 1)]
        [parameter(ParameterSetName = 'Sync', ValueFromPipeline = $true, Position = 1)]
        [string[]]$InputObject
        ,
        [parameter(ParameterSetName = 'ASync')]
        [parameter(ParameterSetName = 'Sync')]
        $ConfigurationName = 'Default'
        ,
        [parameter(ParameterSetName = 'ASync')]
        [parameter(ParameterSetName = 'Sync')]
        [ValidateRange(-10, 10)]
        [Int]$Rate
        ,
        [parameter(ParameterSetName = 'ASync')]
        [parameter(ParameterSetName = 'Sync')]
        [ValidateRange(1, 100)]
        [int]$Volume
        ,
        [parameter(ParameterSetName = 'ASync')]
        [parameter(ParameterSetName = 'Sync')]
        [string]$voice
        ,
        [parameter(ParameterSetName = 'Sync')]
        [switch]$SynchronousOutput
    )
    Begin
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
    }
    Process
    {
        ForEach ($line in $inputobject)
        {
            Write-Verbose "Speaking: $line"
            if ($PSCmdlet.ParameterSetName -eq 'Sync')
            {
                $null = $SpeechConfiguration.Speak(($line | Out-String))
            }
            else
            {
                $null = $SpeechConfiguration.SpeakAsync(($line | Out-String))
            }
        }
    }
}
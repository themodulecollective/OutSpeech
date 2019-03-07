
Function Out-Speech
{
    <#
.SYNOPSIS
    Used to allow PowerShell to speak to you.

.DESCRIPTION
    Used to allow PowerShell to speak to you.

.PARAMETER InputObject
    Data that will be spoken or sent to a WAV file.

.PARAMETER Rate
    Sets the speaking rate

.PARAMETER Volume
    Sets the output volume

.EXAMPLE
    "This is a test" | Out-Voice

    Description
    -----------
    Speaks the string that was given to the function in the pipeline.

.EXAMPLE
    "Today's date is $((get-date).toshortdatestring())" | Out-Voice

    Description
    -----------
    Says todays date

    .EXAMPLE
    "Today's date is $((get-date).toshortdatestring())" | Out-Voice -ToWavFile "C:\temp\test.wav"

    Description
    -----------
    Says todays date

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
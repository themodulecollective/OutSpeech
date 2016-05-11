<#
inspired by https://learn-powershell.net/2013/12/04/give-powershell-a-voice-using-the-speechsynthesizer-class/ 
#and
#https://gallery.technet.microsoft.com/scriptcenter/Out-Voice-1be16d5e
.NOTES 
    Name: OutSpeech.psm1
    Author: Mike Campbell
    DateCreated: 2016-05-07
#>
##ToDo: Add a Set-Speech function for persistent changes to the Speech object for voice, rate, volume, etc.
##ToDo: Make the -voice parameter of each function that has it a dynamic parameter with currently installed voices or add validation within the function for voice specified
##ToDo: Allow multiple speech configuration objects for different configurations beyond just Speech and ExportSpeech
##ToDo: Export-Speech: Support Audio Format Info:  https://msdn.microsoft.com/en-us/library/ms586885(v=vs.110).aspx
##ToDo: Add inline documentation to functions
###############################################################################################
#Module Variables and Variable Functions
###############################################################################################
function Get-OutSpeechVariable
{
param
(
[string]$Name
)
    Get-Variable -Scope Script -Name $name 
}
function Get-OutSpeechVariableValue
{
param
(
[string]$Name
)
    Get-Variable -Scope Script -Name $name -ValueOnly
}
function Set-OutSpeechVariable
{
param
(
[string]$Name
,
$Value
)
    Set-Variable -Scope Script -Name $Name -Value $value  
}
function New-OutSpeechVariable
{
param 
(
[string]$Name
,
$Value
)
    New-Variable -Scope Script -Name $name -Value $Value
}
function Remove-OutSpeechVariable
{
param
(
[string]$Name
)
    Remove-Variable -Scope Script -Name $name
}
###############################################################################################
#Core Voice Functions
###############################################################################################
Function Enable-Speech
{
[cmdletbinding()]
param
(
    [parameter()]
    [string]$Voice
    ,
    [parameter()]
    [ValidateRange(-10,10)]
    [Int]$Rate
    ,
    [parameter()]
    [ValidateRange(1,100)]
    $Volume
    ,
    [switch]$passthru
    ,
    [switch]$resetSpeechObject
    ,
    [parameter()]
    [ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
    [string]$SpeechConfiguration = 'DefaultSpeech'
)
$SpeechVariablePath = 'variable:script:' + $SpeechConfiguration
$VoiceVariableInitialized = Test-Path $SpeechVariablePath
$InitializeSpeechParams = 
@{
    ErrorAction = 'Stop'
    SpeechConfiguration = $SpeechConfiguration
}
foreach ($param in $PSBoundParameters.Keys)
{
    if ($param -in 'Rate','Volume','Voice','Verbose')
    {
        $InitializeSpeechParams.$param = $($PSBoundParameters.$param)
    }
}
if ($VoiceVariableInitialized)
{
    if ($resetSpeechObject) 
    {
        Disable-Speech -SpeechConfiguration $SpeechConfiguration
        Try
        {
            Initialize-Speech @InitializeSpeechParams
        }
        Catch
        {
            $status = $false
            $_
        }
    } else
    {
        $status = $true
        #nothing to do
    }
} else
{
    Try
    {
        Initialize-Speech @InitializeSpeechParams
        $status = $true
    }
    Catch
    {
        $status = $false
        $_
    }
}#else if speech was not already initialized
if ($passthru)
{Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly}
else
{$status}
}#function Enable-Speech
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
    $passthru
    ,
    $SpeechConfiguration
)
    Write-Verbose "Loading System.speech assembly"
    Add-Type -AssemblyName System.speech -ErrorAction Stop
    Write-Verbose "Creating Speech object (Module variable 'Speech')"
    New-Variable -Name $SpeechConfiguration -Value $(New-Object System.Speech.Synthesis.SpeechSynthesizer) -Scope Script -Description 'Object for *-Speech cmdlets' -ErrorAction Stop 
    $speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly
    foreach ($param in $PSBoundParameters.Keys)
    {
        if ($param -in 'Rate','Volume')
        {
            $speechObject.$param = $($PSBoundParameters.$param)
        }
        if ($param -in 'Voice')
        {
            $speechObject.SelectVoice($($PSBoundParameters.$param))
        }
    }
}#function Initialize-Speech
Function Set-Speech
{
[cmdletbinding()]
param
(
[parameter()]
[ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
$SpeechConfiguration = 'DefaultSpeech'
,
[parameter()]
[string]$Voice
,
[parameter()]
[ValidateRange(-10,10)]
[Int]$Rate
,
[parameter()]
[ValidateRange(1,100)]
$Volume
,
[switch]$passthru
)
$speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly
if (Enable-Speech -SpeechConfiguration $SpeechConfiguration) {
    If ($PSBoundParameters.ContainsKey('Volume'))
    {
        Write-Verbose "Setting speech object $SpeechConfiguration volume to $Volume"
        $speechObject.Volume = $Volume
    } Else
    {
        Write-Verbose "Speech object $SpeechConfiguration volume = $($speechObject.volume)"
    }
    If ($PSBoundParameters.ContainsKey('Rate'))
    {
        Write-Verbose "Setting speech object $SpeechConfiguration rate to $Rate"
        $speechObject.Rate = $Rate
    } Else {
        Write-Verbose "Speech object $SpeechConfiguration rate =  $($speechObject.rate)"
    }
    If ($PSBoundParameters.ContainsKey('Voice'))
    {
        Write-Verbose "Setting speech voice to $Voice"
        $speechObject.SelectVoice($Voice)
    } Else {
        Write-Verbose "Speech object $SpeechConfiguration voice = $($speechObject.voice.name)"
    }
}
else
{
    Throw "Unable to Initialize Speech Assembly or Speech Object."
}
if ($passthru)
{$speechObject}
}#Function Set-SpeechConfiguration
Function Get-InstalledVoice
{
[cmdletbinding()]
param
(
[parameter()]
[ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
$SpeechConfiguration = 'DefaultSpeech'
)
if (Enable-Speech -SpeechConfiguration $SpeechConfiguration)
{
    $speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly    
    $Voices = $speechObject.GetInstalledVoices() 
    foreach ($voice in $Voices)
    {
        $CustomOutputObject = [pscustomobject]@{
            Name = $voice.VoiceInfo.Name
            Age = $voice.VoiceInfo.Age
            Gender = $voice.VoiceInfo.Gender
            Culture = $voice.VoiceInfo.Culture
            Id = $voice.VoiceInfo.Id
            Description = $voice.VoiceInfo.Description
            SupportedAudioFormats = $voice.VoiceInfo.SupportedAudioFormats
            AdditionalInfo = $voice.VoiceInfo.AdditionalInfo
            Enabled = $voice.Enabled
        }
        $CustomOutputObject
    }
}#if Enable-Speech
else
{
    Throw "Unable to Initialize Speech Assembly or Speech Object."    
}
}#function
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
[parameter(ParameterSetName = 'ASync',ValueFromPipeline = $true, Position = 1)]
[parameter(ParameterSetName = 'Sync',ValueFromPipeline = $true, Position = 1)]
[string[]]$InputObject
,
[parameter(ParameterSetName = 'ASync')]
[parameter(ParameterSetName = 'Sync')]
[ValidateRange(-10,10)]
[Int]$Rate
,
[parameter(ParameterSetName = 'ASync')]
[parameter(ParameterSetName = 'Sync')]
[ValidateRange(1,100)]
$Volume
,
[parameter(ParameterSetName = 'ASync')]
[parameter(ParameterSetName = 'Sync')]
[string]$voice
,
[parameter(ParameterSetName = 'ASync')]
[parameter(ParameterSetName = 'Sync')]
[ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
$SpeechConfiguration = 'DefaultSpeech'
,
[parameter(ParameterSetName = 'Sync')]
[switch]$SynchronousOutput
,
[parameter(ParameterSetName = 'Sync')]
[switch]$RestoreSpeechObjectSettings
)
Begin
{
$EnableSpeechParams = 
@{
    ErrorAction = 'Stop'
    SpeechConfiguration = $SpeechConfiguration
}
foreach ($param in $PSBoundParameters.Keys)
{
    if ($param -in 'Rate','Volume','Voice','Verbose')
    {
        $EnableSpeechParams.$param = $($PSBoundParameters.$param)
    }
}
    if (Enable-Speech @EnableSpeechParams)
    {
        $speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly
        $originalVolume = $speechObject.volume
        $originalRate = $speechObject.rate
        $originalVoice = $speechObject.Voice.Name
        If ($PSBoundParameters.ContainsKey('Volume') -and $Volume -ne $originalVolume)
        {
            Write-Verbose "Setting speech volume to $Volume"
            $speechObject.Volume = $Volume
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech volume $originalVolume"
        }
        If ($PSBoundParameters.ContainsKey('Rate') -and $Rate -ne $originalRate)
        {
            Write-Verbose "Setting speech rate to $Rate"
            $speechObject.Rate = $Rate
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech rate $originalRate"
        }
        If ($PSBoundParameters.ContainsKey('Voice') -and $voice -ne $originalVoice)
        {
            Write-Verbose "Setting speech voice to $Voice"
            $speechObject.SelectVoice($Voice)
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech voice $originalVoice"
        }
        if ($speechObjectChanged -and (-not $RestoreSpeechObjectSettings))
        {
            Write-Warning -Message "Changes to the Rate, Volume, or Voice properties of Speech Configuration: $SpeechConfiguration will persist with future uses of Speech Configuration: $SpeechConfiguration"
        }
    }
    else
    {
        Throw "Unable to Initialize Speech Assembly or Speech Object."
    }
}
Process
{
    ForEach ($line in $inputobject)
    {
        Write-Verbose "Speaking: $line"
        if ($PSCmdlet.ParameterSetName -eq 'Sync')
        {
            $speechObject.Speak(($line | Out-String)) | Out-Null
        } else
        {
            $speechObject.SpeakAsync(($line | Out-String)) | Out-Null
        }
        
    }
}
End
{
if ($RestoreSpeechObjectSettings)
{
    #reset Speech object to original values
    If ($PSBoundParameters.ContainsKey('Volume'))
    {
        Write-Verbose "Setting speech volume back to $originalVolume"
        $speechObject.volume = $originalVolume 
    }
    If ($PSBoundParameters.ContainsKey('Rate'))
    {
        Write-Verbose "Setting speech rate back to $originalRate"
        $speechObject.rate = $originalRate
    }
    If ($PSBoundParameters.ContainsKey('Voice'))
    {
        Write-Verbose "Setting speech voice back to $originalVoice"
        $speechObject.SelectVoice($originalVoice)
    }
}
}#end
}#function Out-Speech
Function Export-Speech
{
#.PARAMETER Path
# output to a Waveform audio format file 
[cmdletbinding()]
param
(
[Parameter(ValueFromPipeline='True')]
[string[]]$inputobject
,
[Parameter()]
[ValidateRange(-10,10)]
[Int]$Rate
,
[Parameter()]
[ValidateRange(1,100)]
$Volume
,
[Parameter()]
[string]$voice
,
[Parameter(Mandatory = $true)]
[string]$Path
,
[Parameter()]
[ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
[ValidateScript({-not (Test-Path $('variable:script:' + $_))})]
$SpeechConfiguration = 'DefaultExportSpeech'
)
begin
{
$EnableSpeechParams = 
@{
    ErrorAction = 'Stop'
    SpeechConfiguration = $SpeechConfiguration
}
foreach ($param in $PSBoundParameters.Keys)
{
    if ($param -in 'Rate','Volume','Voice','Verbose')
    {
        $EnableSpeechParams.$param = $($PSBoundParameters.$param)
    }
}
    if (Enable-Speech @EnableSpeechParams)
    {
        $speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly
        $originalVolume = $speechObject.volume
        $originalRate = $speechObject.rate
        $originalVoice = $speechObject.Voice.Name
        Write-Verbose "Setting speech to output to wavfile: $Path"
        $speechObject.SetOutputToWaveFile($Path)
        If ($PSBoundParameters.ContainsKey('Volume') -and $Volume -ne $originalVolume)
        {
            Write-Verbose "Setting speech volume to $Volume"
            $speechObject.Volume = $Volume
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech volume $originalVolume"
        }
        If ($PSBoundParameters.ContainsKey('Rate') -and $Rate -ne $originalRate)
        {
            Write-Verbose "Setting speech rate to $Rate"
            $speechObject.Rate = $Rate
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech rate $originalRate"
        }
        If ($PSBoundParameters.ContainsKey('Voice') -and $voice -ne $originalVoice)
        {
            Write-Verbose "Setting speech voice to $Voice"
            $speechObject.SelectVoice($Voice)
            $speechObjectChanged = $true
        } Else
        {
            Write-Verbose "Using speech voice $originalVoice"
        }
    }
    else
    {
        Throw "Unable to Initialize Speech Assembly or Speech Object."
    }
}#begin
Process
{
    ForEach ($line in $inputobject)
    {
        Write-Verbose "Exporting: $line"       
        $speechObject.Speak(($line | Out-String)) | Out-Null
    }
}#Process
End
{
    Disable-Speech -SpeechConfiguration $SpeechConfiguration
}
}#Function Export-Speech
Function Disable-Speech
{
[cmdletbinding()]
param
(
[Parameter()]
[ValidatePattern('^[A-Za-z0-9][A-Za-z0-9_-]{0,25}')]
[ValidateScript({(Test-Path $('variable:script:' + $_))})]
$SpeechConfiguration
)
$speechObject = Get-Variable -Name $SpeechConfiguration -Scope Script -ValueOnly
$speechObject.dispose()
Remove-Variable -Name $SpeechConfiguration -Scope Script -Force
}#Function Disable-Speech


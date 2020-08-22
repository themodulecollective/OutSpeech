Function Get-SpeechVoice
{
    <#
    .SYNOPSIS
        Gets all or specified currently available voices that can be specified for use with SpeechConfigurations, Out-Speech, or Export-Speech
    .DESCRIPTION
        Gets all or specified currently available voices that can be specified for use with SpeechConfigurations, Out-Speech, or Export-Speech.
        Voices can be specified by name, age, gender, or culture.
    .PARAMETER Name
        Specify the name of the voice to get. Simple wildcards supported. Run 'Get-SpeechVoice | Select-Object -ExpandProperty Name' to see all available names.
    .PARAMETER Age
        Specify the age of the voice to get. Valid values are 'Adult','Child','NotSet','Senior','Teen'
    .PARAMETER Gender
        Specfiy the gender of the voice to get. Valid values are 'Female','Male','Neutral','NotSet'
    .PARAMETER Culture
        Specify the culture of the voice to get. Simple wildcards supported. Like 'en-*' or 'es-*'
    .EXAMPLE
        PS C:\> Get-SpeechVoice

        Gets all currently available voices
    .EXAMPLE
        PS C:\> Get-SpeechVoice -Name 'Microsoft Zira Desktop'

        Gets the voice with the name specified if it is available on the system.
    .EXAMPLE
        PS C:\> Get-SpeechVoice -Name '*Desktop'

        Gets the available voices with matching names.
    .EXAMPLE
        PS C:\> Get-SpeechVoice -Culture en-*

        Gets the available voices with a matching culture.
    .EXAMPLE
        PS C:\> Get-SpeechVoice -Age Adult

        Gets the available voices with a Adult age designation.
    .EXAMPLE
        PS C:\> Get-SpeechVoice -Gender Female

        Gets the available voices with a Female gender designation.
    #>
    [cmdletbinding()]
    param
    (
        [string[]]$Name
        ,
        [parameter()]
        [validateset('Adult', 'Child', 'NotSet', 'Senior', 'Teen')]
        [string[]]$Age
        ,
        [parameter()]
        [validateset('Female', 'Male', 'Neutral', 'NotSet')]
        [string[]]$Gender
        ,
        [parameter()]
        [string[]]$Culture # like 'en-US'
    )
    if (-not $script:SpeechConfigurations.containskey('Default'))
    {
        Enable-SpeechConfiguration -ConfigurationName 'Default'
    }
    $SpeechConfiguration = $script:SpeechConfigurations.Default
    $Voices = $SpeechConfiguration.GetInstalledVoices()
    foreach ($voice in $Voices)
    {
        $CustomOutputObject = [pscustomobject]@{
            Name                  = $voice.VoiceInfo.Name
            Age                   = $voice.VoiceInfo.Age
            Gender                = $voice.VoiceInfo.Gender
            Culture               = $voice.VoiceInfo.Culture
            Id                    = $voice.VoiceInfo.Id
            Description           = $voice.VoiceInfo.Description
            SupportedAudioFormats = $voice.VoiceInfo.SupportedAudioFormats
            AdditionalInfo        = $voice.VoiceInfo.AdditionalInfo
            Enabled               = $voice.Enabled
        }
        $CustomOutputObject | Where-Object -FilterScript {
            ($null -eq $Name -or $_.Name -in $Name -or $_.Name -like $Name) -and
            ($null -eq $Age -or $_.Age -in $Age) -and
            ($null -eq $Gender -or $_.Gender -in $Gender) -and
            ($null -eq $Culture -or $_.Culture -in $Culture -or $_.Culture -like $Culture)
        }
    }
}
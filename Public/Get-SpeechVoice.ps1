Function Get-SpeechVoice
{
    [cmdletbinding()]
    param
    (
        [string[]]$Name
        ,
        [parameter()]
        [validateset('Adult','Child','NotSet','Senior','Teen')]
        [string[]]$Age
        ,
        [parameter()]
        [validateset('Female','Male','Neutral','NotSet')]
        [string[]]$Gender
        ,
        [parameter()]
        [string[]]$Culture # like 'en-US'
    )
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
        $CustomOutputObject | Where-object -FilterScript {
            ($null -eq $Name -or $_.Name -in $Name) -and
            ($null -eq $Age -or $_.Age -in $Age) -and
            ($null -eq $Gender -or $_.Gender -in $Gender) -and
            ($null -eq $Culture -or $_.Culture -in $Culture)
        }
    }
}
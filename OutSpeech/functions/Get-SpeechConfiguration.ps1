function Get-SpeechConfiguration
{
    <#
.SYNOPSIS
    Gets all existing SpeechConfigurations or those that match the values provided with the ConfigurationName parameter.
.DESCRIPTION
    Gets all existing SpeechConfigurations or those that match the values provided with the ConfigurationName parameter.
.PARAMETER ConfigurationName
    Specify the name (string) of an existing SpeechConfiguration to get.  'Default' should always exist unless it has been disabled with Disable-SpeechConfiguration.
.PARAMETER Rate
    Specifies the speech rate of the configurations to return.  Valid range is -10 through 10.
.PARAMETER Volume
    Specifies the speech volume of the configurations to return.  Valid range is 1 through 100.
.PARAMETER Voice
    Specifies the speech voice of the configurations to return.  Run Get-SpeechVoice to see valid values - use the name attribute.  Default depends on the system language/culture settings.
.EXAMPLE
    PS C:\> Enable-SpeechConfiguration -ConfigurationName 'Test'
    PS C:\> Get-SpeechConfiguration

    Name                           Value
    ----                           -----
    Test                           System.Speech.Synthesis.SpeechSynthesizer
    Default                        System.Speech.Synthesis.SpeechSynthesizer

    In this example you see Get-SpeechConfiguration returning both the Default SpeechConfiguration and the Test SpeechConfiguration.

.EXAMPLE
    PS C:\> Enable-SpeechConfiguration -ConfigurationName 'Test'
    PS C:\> Get-SpeechConfiguration

    Name                           Value
    ----                           -----
    Test                           System.Speech.Synthesis.SpeechSynthesizer

    In this example you see Get-SpeechConfiguration returning the Test SpeechConfiguration.
#>
    [CmdletBinding()]
    param (
        [parameter()]
        [string[]]$ConfigurationName
        ,
        [parameter()]
        [string[]]$Voice
        ,
        [parameter()]
        [ValidateRange(-10, 10)]
        [Int[]]$Rate
        ,
        [parameter()]
        [ValidateRange(1, 100)]
        [int[]]$Volume
    )

    $Script:SpeechConfigurations.keys.foreach( {
            $Script:SpeechConfigurations.$_ |
            Add-Member -MemberType NoteProperty -Name ConfigurationName -Value $_ -PassThru -Force
        }).where(
        {
            ($null -eq $ConfigurationName -or $_.ConfigurationName -in $ConfigurationName) -and
            ($null -eq $Voice -or $_.Voice.Name -in $Voice) -and
            ($null -eq $Rate -or $_.Rate -in $Rate) -and
            ($null -eq $Volume -or $_.Volume -in $Volume)
        }
    )
}
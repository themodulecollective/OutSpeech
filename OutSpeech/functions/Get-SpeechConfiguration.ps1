function Get-SpeechConfiguration
{
    <#
.SYNOPSIS
    Gets all existing SpeechConfigurations or those that match the values provided with the ConfigurationName parameter.
.DESCRIPTION
    Gets all existing SpeechConfigurations or those that match the values provided with the ConfigurationName parameter.
.PARAMETER ConfigurationName
    Specify the name (string) of an existing SpeechConfiguration to get.  'Default' should always exist unless it has been disabled with Disable-SpeechConfiguration.
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
    )
    if ($null -eq $ConfigurationName)
    {
        $Script:SpeechConfigurations.keys | ForEach-Object {
            $Script:SpeechConfigurations.$_ |
            Add-Member -MemberType NoteProperty -Name ConfigurationName -Value $_ -PassThru -Force
        }
    }
    foreach ($cn in $ConfigurationName)
    {
        try
        {
            $Script:SpeechConfigurations.$cn |
            Add-Member -MemberType NoteProperty -Name ConfigurationName -Value $cn -PassThru -Force -ErrorAction Stop
        }
        catch
        {
            Write-Error -Message "SpeechConfiguration $cn is not found."
        }

    }
}
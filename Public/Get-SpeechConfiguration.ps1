function Get-SpeechConfiguration {
    [CmdletBinding()]
    param (
        [parameter()]
        [string[]]$ConfigurationName
    )
    if ($null -eq $ConfigurationName)
    {
        $Script:SpeechConfigurations.getenumerator() | ForEach-Object {$_}
    }
    foreach ($cn in $ConfigurationName)
    {
        $Script:SpeechConfigurations.$cn
    }
}
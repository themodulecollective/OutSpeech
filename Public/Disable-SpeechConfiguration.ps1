Function Disable-SpeechConfiguration
{
    [cmdletbinding()]
    param
    (
        [Parameter()]
        [string[]]$ConfigurationName
    )
    foreach ($cn in $ConfigurationName){
        if ($Script:SpeechConfigurations.ContainsKey($cn))
        {
            $Script:SpeechConfigurations.$cn.dispose()
            $Script:SpeechConfigurations.remove($cn)
        }
        Else
        {
            Write-Warning "SpeechConfiguration $cn does not exist."
        }
    }
}#Function Disable-SpeechConfiguration
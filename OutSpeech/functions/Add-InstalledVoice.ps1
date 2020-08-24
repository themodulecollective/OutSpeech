$sourcePath = 'HKLM:\software\Microsoft\Speech_OneCore\Voices\Tokens' #Where the OneCore voices live
$destinationPath = 'HKLM:\SOFTWARE\Microsoft\Speech\Voices\Tokens' #For 64-bit apps
#$destinationPath2 = 'HKLM:\SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens' #For 32-bit apps
#Set-Location $destinationPath
$SourceVoices = Get-ChildItem $sourcePath
$DestinationVoices = Get-ChildItem -Path $destinationPath
foreach ($voice in $SourceVoices.where( { $_.pschildname -notin $DestinationVoices.pschildname }))
{
    $source = $voice.PSPath #Get the path of this voices key
    Copy-Item -Path $source -Destination $destinationPath -Recurse
    #Copy-Item -Path $source -Destination $destinationPath2 -Recurse
}
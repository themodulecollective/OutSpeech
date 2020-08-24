#Requires -Version 5.1 -psedition Desktop

$ModuleFolder = Split-Path $PSCommandPath -Parent

$Scripts = Join-Path -Path $ModuleFolder -ChildPath 'scripts'
$Functions = Join-Path -Path $ModuleFolder -ChildPath 'functions'

$Script:ModuleFiles = @(
    $(Join-Path -Path $Scripts -ChildPath 'Initialize.ps1')
    # Load Functions
    $(Join-Path -Path $functions -ChildPath 'Enable-SpeechConfiguration.ps1')
    $(Join-Path -Path $functions -ChildPath 'Disable-SpeechConfiguration.ps1')
    $(Join-Path -Path $functions -ChildPath 'Set-SpeechConfiguration.ps1')
    $(Join-Path -Path $functions -ChildPath 'Export-Speech.ps1')
    $(Join-Path -Path $functions -ChildPath 'Get-SpeechConfiguration.ps1')
    $(Join-Path -Path $functions -ChildPath 'Get-SpeechVoice.ps1')
    $(Join-Path -Path $functions -ChildPath 'Initialize-Speech.ps1')
    $(Join-Path -Path $functions -ChildPath 'Out-Speech.ps1')
    # Finalize / Run any Module Functions defined above
    $(Join-Path -Path $Scripts -ChildPath 'RunFunctions.ps1')
)
foreach ($f in $ModuleFiles)
{
    . $f
}

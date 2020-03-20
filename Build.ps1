[CmdletBinding()]
param(

)

$Script:ModuleInstallScope = 'CurrentUser'
$Script:ModuleName = 'OutSpeech'

'Starting build...'
'Installing module dependencies...'
if ((Get-Module -Name 'InvokeBuild').count -lt 1)
{
    Install-Module -Name 'InvokeBuild' -Scope $Script:ModuleInstallScope -Force -SkipPublisherCheck
}


$Error.Clear()

"Invoking build action [$Task]"

Invoke-Build -Result 'Result'
if ($Result.Error)
{
    $Error[-1].ScriptStackTrace | Out-String
    exit 1
}

exit 0
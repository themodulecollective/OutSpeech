Write-Verbose "Importing Functions"

# Import everything in sub folders folder
foreach ( $folder in @( 'Private', 'Public', 'Classes','Functions' ) )
{
    $root = Join-Path -Path $PSScriptRoot -ChildPath $folder
    if ( Test-Path -Path $root )
    {
        Write-Verbose "processing folder $root"
        $files = Get-ChildItem -Path $root -Filter *.ps1

        # dot source each file
        $files | where-Object { $_.name -NotLike '*.Tests.ps1' } |
            ForEach-Object { Write-Verbose $_.name; . $_.FullName }
    }
}

$script:SpeechConfigurations = @{}
$script:SpeechConfigurations
Enable-SpeechConfiguration -ConfigurationName 'Default'

if ($null -ne $PSVersionTable -and $PSVersionTable.PSVersion.Major -ge 5)
{
    Register-ArgumentCompleter -CommandName 'Enable-SpeechConfiguration','Get-SpeechVoice','Export-Speech','Out-Speech','Set-SpeechConfiguration' -ParameterName Voice -ScriptBlock {
        param($CommandName,$ParameterName,$WordToComplete,$CommandAST,$FakeBoundParameter)
        $choices = @(Get-SpeechVoice | Select-Object -ExpandProperty Name)
        ForEach ($c in $choices)
        {
                [System.Management.Automation.CompletionResult]::new("'$c'", "'$c'", 'ParameterValue', "'$c'")
        }
    }
    Register-ArgumentCompleter -CommandName 'Enable-SpeechConfiguration','Get-SpeechVoice','Export-Speech','Out-Speech','Set-SpeechConfiguration','Disable-SpeechConfiguration','Get-SpeechConfiguration' -ParameterName ConfigurationName -ScriptBlock {
        param($CommandName,$ParameterName,$WordToComplete,$CommandAST,$FakeBoundParameter)
        $choices = @($script:SpeechConfigurations.keys)
        ForEach ($c in $choices)
        {
                [System.Management.Automation.CompletionResult]::new("'$c'", "'$c'", 'ParameterValue', "'$c'")
        }
    }
}

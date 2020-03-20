$Script:ModuleName = 'OutSpeech'
$CommandName = $MyInvocation.MyCommand.Name.Replace(".Tests.ps1", "")
Write-Information -MessageData "Command is $CommandName" -InformationAction Continue
Write-Information -MessageData "Module Name is $Script:ModuleName" -InformationAction Continue
$Script:ProjectRoot = $(Split-Path -Path $PSScriptRoot -Parent)
Write-Information -MessageData "ProjectRoot is $Script:ProjectRoot" -InformationAction Continue
$Script:ModuleRoot = $(Join-Path -Path $(Split-Path -Path $PSScriptRoot -Parent) -ChildPath $Script:ModuleName)
Write-Information -MessageData "Module Root is $script:ModuleRoot" -InformationAction Continue
$Script:ModuleFile = $Script:ModuleFile = Join-Path -Path $($Script:ModuleRoot) -ChildPath $($Script:ModuleName + '.psm1')
Write-Information -MessageData "Module File is $($script:ModuleFile)" -InformationAction Continue
$Script:ModuleSettingsFile = Join-Path -Path $($Script:ModuleRoot) -ChildPath $($Script:ModuleName + '.psd1')
Write-Information -MessageData "Module Settings File is $($script:ModuleSettingsFile)" -InformationAction Continue

Describe "$ModuleName Unit Tests" -Tag 'UnitTests' {
    Context "Validate Top Level Files" {
        [string[]]$moduleFileNames = (Get-ChildItem $ModuleRoot -File).Name
        $expectedFileNames = @($($ModuleName + '.psd1'), $($ModuleName + '.psm1'))
        It "Should contain expected files $($expectedFileNames -join ', ')" {
            ( (Compare-Object -ReferenceObject $expectedFileNames -DifferenceObject $moduleFileNames -IncludeEqual | Where-Object SideIndicator -eq "==").Count ) | Should Be $expectedFileNames.Count
        }
        <#         It "Should only contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        } #>
    }
}
Describe "$ModuleName Project Unit Tests" -Tag 'UnitTests' {
    Context "Validate Top Level Files" {
        [string[]]$ProjectFileNames = (Get-ChildItem $ProjectRoot -File).Name
        $expectedFileNames = @($('README.md', 'license', 'ScriptAnalyzerSettings.psd1'))
        It "Should contain expected files $($expectedFileNames -join ', ')" {
            ( (Compare-Object -ReferenceObject $expectedFileNames -DifferenceObject $ProjectFileNames -IncludeEqual | Where-Object SideIndicator -eq "==").Count ) | Should Be $expectedFileNames.Count
        }
        <#         It "Should only contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        } #>
    }
}

#Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
#Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue
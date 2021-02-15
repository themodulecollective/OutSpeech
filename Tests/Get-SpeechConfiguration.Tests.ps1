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

Write-Information -MessageData "Removing Module $Script:ModuleName" -InformationAction Continue
Remove-Module -Name $Script:ModuleName -Force -ErrorAction SilentlyContinue
Write-Information -MessageData "Import Module $Script:ModuleName" -InformationAction Continue
Import-Module -Name $Script:ModuleSettingsFile -Force

Describe "$CommandName Unit Tests" -Tag 'UnitTests' {
    Context "Validate parameters" {
        $defaultParamCount = 11
        [object[]]$params = (Get-ChildItem "function:\$CommandName").Parameters.Keys
        $knownParameters = 'ConfigurationName', 'Voice', 'Rate', 'Volume'
        $paramCount = $knownParameters.Count
        It "Should contain specific parameters" {
            ( (Compare-Object -ReferenceObject $knownParameters -DifferenceObject $params -IncludeEqual | Where-Object SideIndicator -EQ "==").Count ) | Should Be $paramCount
        }
        It "Should only contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        }
    }
}

Describe "$commandname Integration Tests" -Tags "IntegrationTests" {
    BeforeAll {
        Disable-SpeechConfiguration -All
        Enable-SpeechConfiguration
        Enable-SpeechConfiguration -Voice 'Microsoft David Desktop' -ConfigurationName 'David'
        Enable-SpeechConfiguration -Rate 5 -ConfigurationName 'Rate'
        Enable-SpeechConfiguration -Volume 15 -ConfigurationName 'Volume'
    }
    Context "Get SpeechConfiguration(s)" {
        It "Gets a SpeechConfiguration with the specified Voice" {
            $SpecifiedVoiceConfiguration = Get-SpeechConfiguration -Voice 'Microsoft David Desktop'
            ($SpecifiedVoiceConfiguration).count | Should BeExactly 4
            $SpecifiedVoiceConfiguration[0].voice.name | Should BeExactly 'Microsoft David Desktop'
        }
        It "Gets a SpeechConfiguration with the specified Rate" {
            $SpeechConfiguration = Get-SpeechConfiguration -rate 5
            $SpeechConfiguration.Rate | Should BeExactly 5
        }
        It "Gets a SpeechConfiguration with the specified Volume" {
            $SpeechConfiguration = Get-SpeechConfiguration -Volume 15
            $SpeechConfiguration.Volume | Should BeExactly 15
        }
    }
}

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
        $knownParameters = 'VoiceId', 'Name', 'Age', 'Gender', 'Culture'
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
    }
    Context "Get SpeechVoice(s)" {
        It "Gets a SpeechVoice with the specified Name" {
            $SpecifiedVoice = @(Get-SpeechVoice -Name 'Microsoft David Desktop')
            $SpecifiedVoice.count | Should BeExactly 1
            $SpecifiedVoice.name | Should BeExactly 'Microsoft David Desktop'
        }
        It "Gets a SpeechVoice with the specified Age" {
            $SpecifiedVoice = @(Get-SpeechVoice -Age 'Adult')
            $SpecifiedVoice[0].age | Should BeExactly 'Adult'
        }
        It "Gets a SpeechVoice with the specified Gender" {
            $SpecifiedVoice = @(Get-SpeechVoice -Gender 'Male')
            $SpecifiedVoice[0].Gender | Should BeExactly 'Male'
        }
        It "Gets a SpeechVoice with the specified Culture" {
            $SpecifiedVoice = @(Get-SpeechVoice -Culture 'en-US')
            $SpecifiedVoice[0].Culture | Should BeExactly 'en-US'
        }
        It "Gets a SpeechVoice with the specified ID" {
            $SpecifiedVoice = @(Get-SpeechVoice -VoiceId 'TTS_MS_EN-US_DAVID_11.0')
            $SpecifiedVoice[0].Id | Should BeExactly 'TTS_MS_EN-US_DAVID_11.0'
        }
    }
}

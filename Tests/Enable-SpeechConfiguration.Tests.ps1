$CommandName = $MyInvocation.MyCommand.Name.Replace(".Tests.ps1", "")

Describe "$CommandName Unit Tests" -Tag 'UnitTests' {
    Context "Validate parameters" {
        $defaultParamCount = 11
        [object[]]$params = (Get-ChildItem "function:\$CommandName").Parameters.Keys
        $knownParameters = 'ConfigurationName','Voice','Rate','Volume'
        $paramCount = $knownParameters.Count
        It "Should contain specific parameters" {
            ( (Compare-Object -ReferenceObject $knownParameters -DifferenceObject $params -IncludeEqual | Where-Object SideIndicator -eq "==").Count ) | Should Be $paramCount
        }
        It "Should only contain $paramCount parameters" {
            $params.Count - $defaultParamCount | Should Be $paramCount
        }
    }
}

Describe "$commandname Integration Tests" -Tags "IntegrationTests" {
    BeforeAll {
        Disable-SpeechConfiguration -All
    }
    Context "Enables SpeechConfiguration(s)" {
        It "Enables a Default SpeechConfiguration without error"{
            {Enable-SpeechConfiguration} | Should Not Throw
            $SpeechConfiguration = Get-SpeechConfiguration -ConfigurationName Default
            $SpeechConfiguration.ConfigurationName | Should BeExactly 'Default'
        }
        It "Throws if a SpeechConfiguration already exists with the ConfigurationName"{
            {Enable-SpeechConfiguration} | Should Throw
        }
        It "Enables a SpeechConfiguration with the specified Voice"{
            {Enable-SpeechConfiguration -Voice 'Microsoft David Desktop' -ConfigurationName 'David'} | Should Not Throw
            $SpeechConfiguration = Get-SpeechConfiguration -ConfigurationName 'David'
            $SpeechConfiguration.Voice.Name | Should BeExactly 'Microsoft David Desktop'
        }
        It "Enables a SpeechConfiguration with the specified Rate"{
            {Enable-SpeechConfiguration -Rate 5 -ConfigurationName 'Rate'} | Should Not Throw
            $SpeechConfiguration = Get-SpeechConfiguration -ConfigurationName 'Rate'
            $SpeechConfiguration.Rate| Should BeExactly 5
        }
        It "Enables a SpeechConfiguration with the specified Volume"{
            {Enable-SpeechConfiguration -Volume 15 -ConfigurationName 'Volume'} | Should Not Throw
            $SpeechConfiguration = Get-SpeechConfiguration -ConfigurationName 'Volume'
            $SpeechConfiguration.Volume| Should BeExactly 15
        }
    }
}

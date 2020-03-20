###############################################################################################
# Import User's Configuration
###############################################################################################
#Import-OutSpeechConfiguration
$script:SpeechConfigurations = @{ }
###############################################################################################
# Setup Tab Completion
###############################################################################################
# Tab Completions for OutSpeech

if ($null -ne $PSVersionTable -and $PSVersionTable.PSVersion.Major -ge 5)
{
  Register-ArgumentCompleter -CommandName 'Enable-SpeechConfiguration', 'Get-SpeechVoice', 'Export-Speech', 'Out-Speech', 'Set-SpeechConfiguration' -ParameterName Voice -ScriptBlock {
    param($CommandName, $ParameterName, $WordToComplete, $CommandAST, $FakeBoundParameter)
    $choices = @(Get-SpeechVoice | Select-Object -ExpandProperty Name)
    ForEach ($c in $choices)
    {
      [System.Management.Automation.CompletionResult]::new("'$c'", "'$c'", 'ParameterValue', "'$c'")
    }
  }
  Register-ArgumentCompleter -CommandName 'Enable-SpeechConfiguration', 'Get-SpeechVoice', 'Export-Speech', 'Out-Speech', 'Set-SpeechConfiguration', 'Disable-SpeechConfiguration', 'Get-SpeechConfiguration' -ParameterName ConfigurationName -ScriptBlock {
    param($CommandName, $ParameterName, $WordToComplete, $CommandAST, $FakeBoundParameter)
    $choices = @($script:SpeechConfigurations.keys)
    ForEach ($c in $choices)
    {
      [System.Management.Automation.CompletionResult]::new("'$c'", "'$c'", 'ParameterValue', "'$c'")
    }
  }
}
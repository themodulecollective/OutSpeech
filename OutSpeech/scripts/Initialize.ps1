#Requires -Version 5.1
###############################################################################################
# Module Variables
###############################################################################################
$ModuleVariableNames = ('OutSpeechConfiguration', 'SpeechConfigurations')
$ModuleVariableNames.ForEach( { Set-Variable -Scope Script -Name $_ -Value $null })

###############################################################################################
# Module Removal
###############################################################################################
#Clean up objects that will exist in the Global Scope due to no fault of our own . . . like PSSessions

$OnRemoveScript = {
  # perform cleanup
  Write-Verbose -Message 'Removing Module Items from Global Scope'
}

$ExecutionContext.SessionState.Module.OnRemove += $OnRemoveScript

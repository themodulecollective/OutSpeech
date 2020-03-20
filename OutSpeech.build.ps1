Task . InstallDependencies, CleanTestResults, Tests, CleanArtifacts, BuildModuleFiles

Task InstallDependencies {
  if ((Get-Module -Name Pester -ListAvailable).count -lt 1)
  {
    Install-Module -Name Pester -Scope CurrentUser -Force
  }
  if ((Get-Module -Name PSScriptAnalyzer -ListAvailable).count -lt 1)
  {
    Install-Module -Name PSScriptAnalyzer -Scope CurrentUser -Force
  }
  if ((Get-Module -Name PSFramework -ListAvailable).count -lt 1)
  {
    Install-Module -Name PSFramework -Scope CurrentUser -Force
  }
}

Task CleanTestResults {
  $TestResults = $(Join-Path -Path $BuildRoot -ChildPath 'TestResults')

  if (Test-Path -Path $TestResults)
  {
    Remove-Item $TestResults -Recurse -Force
  }

  New-Item -ItemType Directory -Path $TestResults -Force
}

Task Tests {

  Import-Module Pester
  $TestResults = $(Join-Path -Path $BuildRoot -ChildPath 'TestResults')

  $invokePesterParams = @{
    Strict       = $true
    PassThru     = $true
    Verbose      = $false
    EnableExit   = $false
    OutputFormat = 'NUnitXml'
    OutputFile   = $(Join-Path -Path $TestResults -childPath 'Test-Pester.XML')
    CodeCoverage = "$(Join-Path -Path $(Join-Path -Path $BuildRoot -ChildPath 'InstallManager') -ChildPath 'functions')*.ps1"
  }

  # Publish Test Results as NUnitXml


  $testResults = Invoke-Pester @invokePesterParams

  $numberFails = $testResults.FailedCount
  Assert($numberFails -eq 0) ('Failed "{0}" unit tests.' -f $numberFails)

}

Task CleanArtifacts {
  $Artifacts = $(Join-Path -Path $BuildRoot -ChildPath 'artifacts')

  if (Test-Path -Path $Artifacts)
  {
    Remove-Item $Artifacts -Recurse -Force
  }

  New-Item -ItemType Directory -Path $Artifacts -Force
}

Task BuildModuleFiles {

  $Artifacts = $(Join-Path -Path $BuildRoot -ChildPath 'artifacts')
  Write-Information -MessageData "Artifacts Path = $Artifacts" -InformationAction Continue
  $ModuleName = $(Split-Path $BuildFile -Leaf).split('.')[0]
  Write-Information -MessageData "ModuleName = $ModuleName" -InformationAction Continue
  $ModuleManifestFileName = $ModuleName + '.psd1'
  $ModuleFolder = $(Join-Path -Path $BuildRoot -ChildPath $ModuleName)
  #Write-Information -MessageData "ModuleFolder = $ModuleFolder" -InformationAction Continue
  $ModuleManifest = $(Join-Path -Path $ModuleFolder -ChildPath $ModuleManifestFileName)
  #Write-Information -MessageData "ModuleManifest = $ModuleManifest" -InformationAction Continue
  $Module = Import-Module -FullyQualifiedName $ModuleManifest -PassThru
  $PSM1SourceFiles = $Module.Invoke( { Get-Variable -ValueOnly -Name ModuleFiles })
  $Module = $null
  Remove-Module $ModuleName
  $PSM1Content =
  foreach ($sf in $PSM1SourceFiles)
  {
    Get-Content -Path $sf -Raw
  }
  $PSM1FilePath = $(Join-Path -Path $Artifacts -ChildPath $($ModuleName + '.psm1'))
  $PSM1Content | Out-File -FilePath $PSM1FilePath
  Copy-Item -Destination $Artifacts -Path $ModuleManifest
  Copy-Item -Destination $Artifacts -Path $(Join-Path -path $BuildRoot -ChildPath 'LICENSE')
  Copy-Item -Destination $Artifacts -Path $(Join-Path -path $BuildRoot -ChildPath 'README.md')
  Copy-Item -Destination $Artifacts -Path $(Join-Path -path $BuildRoot -ChildPath 'ReleaseNotes.md')
}
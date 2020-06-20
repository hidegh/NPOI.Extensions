param(
  [Parameter(Mandatory=$false)]
  [string]$list = "pack-list.txt"
  ,
  [Parameter(Mandatory=$false)]
  [string]$nugetFolder = "nupkg"
  ,
  [Parameter(Mandatory=$false)]
  [string]$nugetSource = "https://api.nuget.org/v3/index.json"
  ,
  [Parameter(Mandatory=$true)]
  [string]$nugetApiKey
)

Write-Host "Publishing NuGet from projects..."

$fileExists = [System.IO.File]::Exists($list)
if (-not $fileExists)
{ 
  Write-Host "File '${list}' does not exists, can't continue publishing."
  Write-Host "ERROR!";
  exit -1;
}

$lines = [System.IO.File]::ReadLines($list) | ? {$_.trim() -ne "" }

if ($lines.Length -eq 0) {
  Write-Host "No project(s) listed in file '${list}', can't continue publishing."
  Write-Host "ERROR!";
  exit -1;
}

$lines | ForEach-Object {
  $source = $_
  $packageSourceFolder = "$($source)/$($nugetFolder)"

  Write-Host "`n`n`n"
  Write-Host "Project       : $(${source})"
  Write-Host "Package folder: $(${packageSourceFolder})"
  Write-Host "---"

  # as *.nupkg does not works (in Github actions)
  # dotnet nuget push $packageSourceFolder/*.nupkg -s $nugetSource -k $nugetApiKey  --skip-duplicate 2>&1 | Write-Host

  $nupkgFileName = @(gci $packageSourceFolder/*.nupkg)[0].Name
  dotnet nuget push $packageSourceFolder/$nupkgFileName -s $nugetSource -k $nugetApiKey  --skip-duplicate 2>&1 | Write-Host
  if ($LASTEXITCODE -ne 0) { Write-Host "ERROR!"; exit -1; }
}

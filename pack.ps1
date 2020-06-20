param(
  [Parameter(Mandatory=$false)]
  [string]$list = "pack-list.txt"
  ,
  [Parameter(Mandatory=$false)]
  [string]$nugetFolder = "nupkg"
)

Write-Host "Packaging NuGet from projects..."

$fileExists = [System.IO.File]::Exists($list)
if (-not $fileExists)
{ 
  Write-Host "File '${list}' does not exists, can't continue packing."
  Write-Host "ERROR!";
  exit -1;
}

$lines = [System.IO.File]::ReadLines($list) | ? {$_.trim() -ne "" }

if ($lines.Length -eq 0) {
  Write-Host "No project(s) listed in file '${list}', can't continue packing."
  Write-Host "ERROR!";
  exit -1;
}

$lines | ForEach-Object {
  $source = $_
  $destination = "$($source)/$($nugetFolder)"

  Write-Host "`n`n`n"
  Write-Host "Project    : $(${source})"
  Write-Host "NugetFolder: $(${destination})"
  Write-Host "---"

  dotnet pack $source --configuration Release -o $destination 2>&1 | Write-Host
  if ($LASTEXITCODE -ne 0) { Write-Host "ERROR!"; exit -1; }
}

param(
  [Parameter(Mandatory=$true)]
  [string]$versionParameter
  ,
  [Parameter(Mandatory=$false)]
  [string]$prefix
  ,
  [Parameter(Mandatory=$false)]
  [string]$suffix
)

Write-Host "---"
Write-Host "Calculating new tag, setting into: $env:VERSION_TAG"

Write-Host "Build version: $versionParameter"
Write-Host "Prefix       : $prefix"
Write-Host "Suffix       : $suffix"

$newVersionTag = $versionParameter;
  
if (![string]::IsNullOrWhiteSpace($prefix)) {
  $newVersionTag = $prefix + $newVersionTag
}

if (![string]::IsNullOrWhiteSpace($suffix)) {
  $newVersionTag = $newVersionTag + $suffix
}

Write-Host "New TAG: $newVersionTag"

echo "::set-env name=VERSION_TAG::$newVersionTag"
Write-Host "Environment variable: VERSION_TAG was set"

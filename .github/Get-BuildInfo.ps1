#Requires -Version 7.0
#Requires -PSEdition Core
#
# Generates deployment information for the build.
#

param (
  [String]
  [Parameter()]
  $ref,
  [String]
  [Parameter()]
  $event_name,
  [String]
  [Parameter()]
  $configuration
)

function Write-GitHubVariable {
  param ($name, $value)
  Write-Output "$name=$value" >> $env:GITHUB_OUTPUT
  Write-Host "  steps.${env:GITHUB_ACTION}.outputs.$name=$value" -ForegroundColor Cyan
}

[xml]$project = Get-Content (Join-Path -Path $PSScriptRoot -ChildPath '../src/NetOfficeExtension/source.extension.vsixmanifest')
$app_version = $project.PackageManifest.Metadata.Identity.Version

$sign_binaries = 'false'
$publish_vsix = 'false'
$app_version_suffix = "preview${env:GITHUB_RUN_NUMBER}"

if ($configuration -ieq 'release') {
  if ($event_name -eq 'push' -and $ref -like 'refs/tags/v*') {
    # temporary disabled as we don't have a valid certificate
    $sign_binaries = 'false'
  }

  if ($ref -like 'refs/tags/v*') {
    $publish_vsix = 'true'
    $app_version_suffix = ''
  }
}

$app_version_full = $app_version
if ($app_version_suffix -ne '') {
  $app_version_full += '-' + $app_version_suffix
}

Write-GitHubVariable "app_version" $app_version
Write-GitHubVariable "app_version_suffix" $app_version_suffix
Write-GitHubVariable "app_version_full" $app_version_full
Write-GitHubVariable "sign_binaries" $sign_binaries
Write-GitHubVariable "publish_vsix" $publish_vsix

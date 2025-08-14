Function Info($msg) {
  Write-Host -ForegroundColor DarkGreen "`nINFO: $msg`n"
}

Function Error($msg) {
  Write-Host `n`n
  Write-Error $msg
  exit 1
}

Function CheckReturnCodeOfPreviousCommand($msg) {
  if(-Not $?) {
    Error "${msg}. Error code: $LastExitCode"
  }
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
Add-Type -AssemblyName System.IO.Compression.FileSystem
$root = Resolve-Path $PSScriptRoot
$buildDir = "$root/out/release"

Info "Create build directory"
New-Item -ItemType Directory -Path $buildDir -ErrorAction SilentlyContinue | Out-Null

if (-Not (Test-Path "$buildDir/ahk-v2.zip")) {
  Info "Download AutoHotkey v2"
  Invoke-WebRequest -Uri "https://github.com/AutoHotkey/AutoHotkey/releases/download/v2.0.19/AutoHotkey_2.0.19.zip" -OutFile "$buildDir/ahk-v2.zip"
  [System.IO.Compression.ZipFile]::ExtractToDirectory("$buildDir/ahk-v2.zip", "$buildDir/ahk-v2")
}

Info "Find Visual Studio installation path"
$vswhereCommand = Get-Command -Name "${Env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe"
$installationPath = & $vswhereCommand -prerelease -latest -property installationPath

Info "Open Visual Studio 2022 Developer PowerShell"
& "$installationPath\Common7\Tools\Launch-VsDevShell.ps1" -Arch amd64

Info "Cmake generate cache"
cmake `
  -S $root `
  -B $buildDir `
  -G Ninja `
  -D CMAKE_BUILD_TYPE=Release `
  -D CMAKE_TOOLCHAIN_FILE="$env:VCPKG_ROOT/scripts/buildsystems/vcpkg.cmake" `
  -D VCPKG_TARGET_TRIPLET=x64-windows-static-release `
CheckReturnCodeOfPreviousCommand "cmake cache failed"

Info "Cmake build"
cmake --build $buildDir
CheckReturnCodeOfPreviousCommand "cmake build failed"

Info "Run tests"
Copy-Item -Path "$root/src/test/test.ahk2" -Destination $buildDir -Force
& "$buildDir/ahk-v2/AutoHotkey64.exe" $buildDir/test.ahk2 | Out-Host
CheckReturnCodeOfPreviousCommand "tests failed"

Info "Create a zip archive from DLL"
Compress-Archive -Force -Path $buildDir/AutoHotkey-ProcessRunner.dll -DestinationPath $buildDir/AutoHotkey-ProcessRunner.zip

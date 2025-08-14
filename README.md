# AutoHotkey-ProcessRunner-dll
A DLL for [AutoHotkey v2](https://www.autohotkey.com/) that provides `ProcessRunner` class to run executables.

# Features
* Runs an executable without showing a console window.
* Retrieves `StdOut` and `StdErr` without creating temporary files.
* Supports Windows 10 x64 or higher.

# How to use
* Download the DLL from the [latest release](https://github.com/PolarGoose/AutoHotkey-ProcessRunner-dll/releases) and place it next to your `.ahk2` script.
* Look at the example below to see how to use the `ProcessRunner` class.

# Method description
```
ProcessRunnerResult {
  string StdOut,   ; Standard output of the process
  string StdErr,   ; Standard error output of the process
  int ExitCode     ; Exit code of the process
}

ProcessRunnerResult ProcessRunner.Run(
  string executable,
  <optional> AhkArrayOfStrings arguments,
  <optional> string workingDirectory)
```

# Example
```ahk
#Requires AutoHotkey v2.0

; Load the DLL. Specify the path where the DLL is located.
#DllLoad "%A_ScriptDir%/AutoHotkey-ProcessRunner.dll"

; Create the ProcessRunner class. You can do it only once and reuse it. It is also possible to create multiple instances if needed.
processRunner := ComObjFromPtr(DllCall("AutoHotkey-ProcessRunner\CreateProcessRunner"))

; You can run a command without any arguments.
; Run method waits until the executable is finished and returns stdout, stderr, and exit code.
result := processRunner.Run("C:/Program Files/Git/usr/bin/echo.exe")

; StdOut, StdErr are strings and ExitCode is an integer
result.StdOut
result.StdErr
result.ExitCode

; You can run a command with arguments.
; The arguments are passed as an array of strings. One argument per array element.
result := processRunner.Run("C:/Program Files/Git/usr/bin/echo.exe", ["arg1 with spaces", "arg2"])
result := processRunner.Run("C:/Program Files/Git/usr/bin/file.exe", ["--mime-type", "--brief", "image.jpg"])

; You can specify the working directory for the command.
result := processRunner.Run("C:/Program Files/Git/usr/bin/echo.exe", [], "C:/Program Files")

; You can use environment variables in the command path and working directory.
result := processRunner.Run("%ProgramW6432%/Git/usr/bin/echo.exe", [], "%ProgramW6432%/Git")
```

# How to build
## Build requirements
* You can use `Visual Studio 2022` to work with the codebase.
* To build the release, run `build.ps1` script.

## Notes
* `vcpkg_overlay_ports\boost-asio` is an original port with the custom `increase_pipe_buffer.diff` patch. This is a workaround for the issue: [[Windows] Setting the capacity of the underline pipe #470
](https://github.com/boostorg/process/issues/470)

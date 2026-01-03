@echo off
echo ========================================
echo  VDK Tool Build Script
echo ========================================
echo.

cd /d "%~dp0src"

:: Try to find MSBuild or dotnet
where dotnet >nul 2>&1
if %errorlevel% equ 0 (
    echo Using dotnet CLI...
    dotnet build -c Release
    if %errorlevel% equ 0 (
        echo.
        echo Build successful!
        echo Output: bin\Release\net48\VDK_Tool.exe

        :: Copy to parent directory
        if not exist "..\bin" mkdir "..\bin"
        copy /Y "bin\Release\net48\VDK_Tool.exe" "..\bin\" >nul 2>&1
        copy /Y "bin\Release\net48\*.dll" "..\bin\" >nul 2>&1
        copy /Y "bin\Release\net48\*.config" "..\bin\" >nul 2>&1
        echo Copied to: ..\bin\VDK_Tool.exe
    ) else (
        echo Build failed!
    )
    goto :end
)

:: Try MSBuild from VS2022
set "MSBUILD="
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe"
)
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Professional\MSBuild\Current\Bin\MSBuild.exe"
)
if exist "%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe" (
    set "MSBUILD=%ProgramFiles%\Microsoft Visual Studio\2022\Enterprise\MSBuild\Current\Bin\MSBuild.exe"
)

if defined MSBUILD (
    echo Using MSBuild...
    "%MSBUILD%" VDKTool.csproj /p:Configuration=Release /v:minimal
    if %errorlevel% equ 0 (
        echo.
        echo Build successful!
    ) else (
        echo Build failed!
    )
    goto :end
)

echo ERROR: Could not find dotnet CLI or MSBuild.
echo Please install .NET SDK or Visual Studio 2022.
echo.
echo Download .NET SDK: https://dotnet.microsoft.com/download

:end
echo.
pause

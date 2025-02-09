@echo off
setlocal

set currentPath=%~dp0
set basePath=%currentPath:~0,-1%
set outputPath=%basePath%\nupkg
set ExcelDnaVersion=%1

if exist "%outputPath%\*.nupkg" del "%outputPath%\*.nupkg"

if not exist "%outputPath%" mkdir "%outputPath%"

echo on
nuget.exe pack "%basePath%\ExcelDna.Registration\ExcelDna.Registration.nuspec" -BasePath "%basePath%\ExcelDna.Registration" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Registration.FSharp\ExcelDna.Registration.FSharp.nuspec" -BasePath "%basePath%\ExcelDna.Registration.FSharp" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

nuget.exe pack "%basePath%\ExcelDna.Registration.VisualBasic\ExcelDna.Registration.VisualBasic.nuspec" -BasePath "%basePath%\ExcelDna.Registration.VisualBasic" -OutputDirectory "%outputPath%" -Verbosity detailed -NonInteractive -Prop ExcelDnaVersion="%ExcelDnaVersion%"
@if errorlevel 1 goto end

:end

setlocal

set PackageVersion=%1
set MSBuildPath=%2

cd ..\NuGet
call package.cmd %PackageVersion%
@if errorlevel 1 goto end

:end

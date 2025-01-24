@echo off
set /p fodler="Zadejte složku: "
dotnet script dates.csx "%fodler%"
pause

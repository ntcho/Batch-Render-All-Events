:: replace the build file to your original script folder location
SET "FILENAME=%~n0.cs"

SET "BUILDFILE=%~dp0\%FILENAME%"

echo f | xcopy /f /y %BUILDFILE% "C:\Program Files\VEGAS\VEGAS Pro 15.0\Script Menu\%FILENAME%"
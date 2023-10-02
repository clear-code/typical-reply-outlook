@REM ==================
@REM Compile C# sources
@REM ==================
msbuild /p:Configuration=Release /t:Rebuild

@REM ==================
@REM Build an installer
@REM ==================
iscc.exe /Odest TypicalReplyOutlook.iss

@REM ==================================
@REM Create package with default config
@REM ==================================
powershell -C ^
"$version=(Get-ItemProperty bin\release\TypicalReply.dll).VersionInfo.FileVersion; ^
 Compress-Archive ^
   -Path dest\TypicalReplySetup-${version}.exe, DefaultConfig ^
   -DestinationPath dest\TypicalReplyOutlook-${version}-with-Default-Config.zip ^
   -Force"

@REM ==================
@REM Compress templates
@REM ==================
powershell -C "Compress-Archive  -DestinationPath dest\TypicalReplyADMX.zip policy"
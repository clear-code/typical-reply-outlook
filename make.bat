@REM ==================
@REM Compile C# sources
@REM ==================
msbuild /p:Configuration=Release /t:Rebuild /p:Platform="Any CPU"

@REM ==================
@REM Build an installer
@REM ==================
iscc.exe /Odest TypicalReplyOutlook.iss

@REM ==================================
@REM Create package with default config
@REM ==================================
powershell -C ^
"$version=(Get-ItemProperty bin\release\TypicalReply.dll).VersionInfo.FileVersion; ^
 Remove-Item -Recurse dest\TypicalReplyOutlook-${version}-with-Default-Config; ^
 mkdir dest\TypicalReplyOutlook-${version}-with-Default-Config; ^
 Copy-Item -Recurse -Path dest\TypicalReplySetup-${version}.exe, DefaultConfig dest\TypicalReplyOutlook-${version}-with-Default-Config;^
 Compress-Archive ^
   -Path dest\TypicalReplyOutlook-${version}-with-Default-Config ^
   -DestinationPath dest\TypicalReplyOutlook-${version}-with-Default-Config.zip ^
   -Force"

@REM ==================
@REM Compress templates
@REM ==================
powershell -C "Compress-Archive  -DestinationPath dest\TypicalReplyADMX.zip policy"
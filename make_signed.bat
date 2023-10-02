@REM Defines the certificate to use for signatures.
@REM Use Powershell to find the available certs:
@REM
@REM PS> Get-ChildItem -Path Cert:CurrentUser\My
@REM
set cert=73E7B9D1F72EDA033E7A9D6B17BC37A96CE8513A
set timestamp=http://timestamp.sectigo.com

@REM ==================
@REM Compile C# sources
@REM ==================
msbuild /p:Configuration=Release /t:Rebuild

@REM ==================
@REM Build an installer
@REM ==================
iscc.exe /Odest TypicalReplyOutlook.iss

@REM ==================
@REM Sign the installer
@REM ==================
signtool sign /t %timestamp% /fd SHA256 /sha1 %cert% TypicalReplyOutlook*.exe

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
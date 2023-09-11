@REM ==================
@REM Compile C# sources
@REM ==================
msbuild /p:Configuration=Release

@REM ==================
@REM Build an installer
@REM ==================
iscc.exe /Odest TypicalReplyOutlook.iss

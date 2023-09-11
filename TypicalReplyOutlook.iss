﻿[Setup]
AppName=TypicalReply
AppVerName=TypicalReply
VersionInfoVersion=22.2.2.0
AppPublisher=ClearCode Inc.
AppVersion=22.2.2
UninstallDisplayIcon={app}\tr.ico
DefaultDirName={commonpf}\TypicalReply
ShowLanguageDialog=no
Compression=lzma2
SolidCompression=yes
OutputDir=dest
OutputBaseFilename=TypicalReplySetup-{#SetupSetting("AppVersion")}
VersionInfoDescription=TypicalReplySetup
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Registry]
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\TypicalReply"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\TypicalReply"; ValueType: string; ValueName: "Description"; ValueData: "Addon for quick replying or transfering with template"
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\TypicalReply"; ValueType: string; ValueName: "FriendlyName"; ValueData: "TypicalReply"
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\TypicalReply"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\TypicalReply.vsto|vstolocal"
Root: HKLM; Subkey: "Software\Microsoft\Office\Outlook\Addins\TypicalReply"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3

; Install for 32bit Office as well
Root: HKLM32; Subkey: "Software\Microsoft\Office\Outlook\Addins\FlexConfirmMail"; Flags: uninsdeletekey
Root: HKLM32; Subkey: "Software\Microsoft\Office\Outlook\Addins\FlexConfirmMail"; ValueType: string; ValueName: "Description"; ValueData: "Addon for quick replying or transfering with template"
Root: HKLM32; Subkey: "Software\Microsoft\Office\Outlook\Addins\FlexConfirmMail"; ValueType: string; ValueName: "FriendlyName"; ValueData: "TypicalReply"
Root: HKLM32; Subkey: "Software\Microsoft\Office\Outlook\Addins\FlexConfirmMail"; ValueType: string; ValueName: "Manifest"; ValueData: "file:///{app}\TypicalReply.vsto|vstolocal"
Root: HKLM32; Subkey: "Software\Microsoft\Office\Outlook\Addins\FlexConfirmMail"; ValueType: dword; ValueName: "LoadBehavior"; ValueData: 3

; Prevent Outlook from disabling .NET addon
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "TypicalReply"; ValueData: 1
Root: HKCU; Subkey: "Software\Microsoft\Office\13.0\Outlook\Resiliency\DoNotDisableAddinList"; ValueType: dword; ValueName: "TypicalReply"; ValueData: 1

[Languages]
Name: en; MessagesFile: "compiler:Default.isl"
Name: jp; MessagesFile: "compiler:Languages\Japanese.isl"

[Files]
Source: "bin\Release\*.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\TypicalReply.dll.manifest"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\TypicalReply.dll.config"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\TypicalReply.vsto"; DestDir: "{app}"; Flags: ignoreversion
Source: "bin\Release\TypicalReply.pdb"; DestDir: "{app}"; Flags: ignoreversion
Source: "Resources\tr.ico"; DestDir: "{app}"; Flags: ignoreversion
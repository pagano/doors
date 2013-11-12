[Setup]
AppName=Gestar Remote DCE
AppVerName=Gestar Remote DCE 7.0.84
DefaultDirName={pf}\Gestar
Compression=zip/9
PrivilegesRequired=admin
OutputBaseFilename=rdcesetup
DisableStartupPrompt=true
DisableProgramGroupPage=false
InternalCompressLevel=max
AllowNoIcons=true
DefaultGroupName=Gestar

[Files]
; begin VB system files
Source: ..\..\..\..\dep\vbrun60sp6\advpack.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
Source: ..\..\..\..\dep\vbrun60sp6\asycfilt.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
Source: ..\..\..\..\dep\vbrun60sp6\comcat.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver allowunsafefiles
Source: ..\..\..\..\dep\vbrun60sp6\msvbvm60.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: ..\..\..\..\dep\vbrun60sp6\oleaut32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: ..\..\..\..\dep\vbrun60sp6\olepro32.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: ..\..\..\..\dep\vbrun60sp6\stdole2.tlb; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: ..\..\..\..\dep\vbrun60sp6\vb6es.dll; DestDir: {sys}; Flags: restartreplace uninsneveruninstall sharedfile
; end VB system files

; msxml4
Source: ..\..\..\..\dep\msxml4r.dll; DestDir: {sys}; Flags: sharedfile restartreplace
Source: ..\..\..\..\dep\msxml4.dll; DestDir: {sys}; Flags: regserver sharedfile restartreplace

; dlls y ocxs
Source: ..\..\..\..\dep\cmax40.dll; DestDir: {sys}; Flags: regserver sharedfile restartreplace
Source: ..\..\..\..\dep\tabctl32.ocx; DestDir: {sys}; Flags: regserver sharedfile restartreplace
Source: ..\..\..\..\dep\tabctes.dll; DestDir: {sys}; Flags: sharedfile restartreplace
Source: ..\..\..\..\dep\mscomctl.ocx; DestDir: {sys}; Flags: regserver sharedfile restartreplace
Source: ..\..\..\..\dep\mscmces.dll; DestDir: {sys}; Flags: sharedfile restartreplace
Source: ..\..\..\..\dep\tlbinf32.dll; DestDir: {sys}; Flags: regserver sharedfile restartreplace
Source: ..\..\..\..\dep\msscript.ocx; DestDir: {sys}; Flags: regserver sharedfile restartreplace

; bin
Source: dapihttp.dll; DestDir: {sys}; Flags: regserver sharedfile restartreplace
Source: ..\rdce.exe; DestDir: {app}\bin
Source: ..\vbscript.lng; DestDir: {app}\bin

[Icons]
Name: {group}\Gestar Remote DCE; Filename: {app}\bin\rdce.exe

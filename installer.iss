; ============================================================
;  מתקין Inno Setup ל-Davidoff Physics Trajectory Log
;  דרישה מוקדמת: dist\TrajectoryLogQA.exe  (נבנה ע"י build_exe.bat או GitHub Actions)
;  שימוש: פתח קובץ זה ב-Inno Setup ולחץ Build > Compile (F9).
;  תוצאה: Output\TrajectoryLogQA_Setup.exe
; ============================================================

#define MyAppName "Davidoff Physics Trajectory Log"
#define MyAppVersion "1.0"
#define MyAppExeName "TrajectoryLogQA.exe"

[Setup]
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
OutputDir=Output
OutputBaseFilename=TrajectoryLogQA_Setup
Compression=lzma2
SolidCompression=yes
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
DisableProgramGroupPage=yes
WizardStyle=modern

[Languages]
Name: "hebrew"; MessagesFile: "compiler:Languages\Hebrew.isl"
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "צור קיצור דרך בשולחן העבודה"; GroupDescription: "קיצורי דרך:"

[Files]
Source: "dist\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "קרא_אותי.txt"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\הסר התקנה"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "הפעל את התוכנה כעת"; Flags: nowait postinstall skipifsilent

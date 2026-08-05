; ===================================================================
;  INSTALADOR AVES AGRO (Inno Setup)
; ===================================================================
;  Como usar:
;  1. Baixe e instale o Inno Setup (gratuito): https://jrsoftware.org/isdl.php
;  2. Coloque este arquivo (instalador.iss) na MESMA pasta dos arquivos:
;     aves-vivas.html, servidor.ps1, iniciar.vbs, agro-benevenuto.ico,
;     version.txt
;  3. Abra o instalador.iss com o Inno Setup (duplo clique) e clique em
;     Compilar (ou aperte F9).
;  4. O instalador pronto aparece em uma subpasta "Output", chamado
;     AvesAgro-Instalador.exe — é esse arquivo que você distribui.
;  5. Quem receber só precisa dar duplo clique nele: instala tudo e já
;     cria o atalho "Aves Agro" na área de trabalho, com a logo.
; ===================================================================

[Setup]
AppName=Aves Agro
AppVersion=1.0
AppPublisher=Agro Benevenuto
DefaultDirName={localappdata}\AvesAgro
DefaultGroupName=Aves Agro
UninstallDisplayIcon={app}\agro-benevenuto.ico
OutputBaseFilename=AvesAgro-Instalador
OutputDir=Output
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
DisableProgramGroupPage=yes
DisableWelcomePage=no
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"

[Files]
Source: "aves-vivas.html"; DestDir: "{app}"; Flags: ignoreversion
Source: "servidor.ps1"; DestDir: "{app}"; Flags: ignoreversion
Source: "iniciar.vbs"; DestDir: "{app}"; Flags: ignoreversion
Source: "agro-benevenuto.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "version.txt"; DestDir: "{app}"; Flags: onlyifdoesntexist

[Icons]
Name: "{autodesktop}\Aves Agro"; Filename: "wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\agro-benevenuto.ico"; IconIndex: 0
Name: "{group}\Aves Agro"; Filename: "wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; IconFilename: "{app}\agro-benevenuto.ico"; IconIndex: 0
Name: "{group}\Desinstalar Aves Agro"; Filename: "{uninstallexe}"

[Run]
Filename: "wscript.exe"; Parameters: """{app}\iniciar.vbs"""; WorkingDir: "{app}"; Description: "Abrir o Aves Agro agora"; Flags: postinstall nowait skipifsilent

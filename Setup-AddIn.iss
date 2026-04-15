[Setup]
AppName=Add-insSupport
AppVersion=1.0.0.1
AppPublisher=ToolSupport
DefaultDirName={autopf}\Add-insSupport
DefaultGroupName=Add-insSupport
OutputDir=.\installer_output
OutputBaseFilename=Add-insSupport_Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
; Khong can quyen Admin, chi dung HKCU

[Files]
; Copy TOAN BO file tu bin\Release (bao gom cac dependency DLL Private=True
; vi du: Microsoft.Office.Tools.Common.v4.0.Utilities.dll, v.v.)
; Neu thieu bat ky DLL nao, add-in se load loi tham va ribbon khong hien.
Source: "bin\Release\*"; DestDir: "{app}"; Flags: ignoreversion; Excludes: "*.pdb,*.ps1"

[Icons]
Name: "{group}\Uninstall Add-insSupport"; Filename: "{uninstallexe}"

[Messages]
FinishedLabel=Cai dat thanh cong. Vui long khoi dong lai Excel de ap dung add-in.

[Code]
// Dung Code section de tao file URI voi forward slash dung chuan
// va tro vao file .vsto (deployment manifest) voi |vstolocal

procedure CurStepChanged(CurStep: TSetupStep);
var
  AppPath: String;
  ManifestUri: String;
  RegKey: String;
begin
  if CurStep = ssPostInstall then
  begin
    AppPath := ExpandConstant('{app}');
    // Chuyen backslash thanh forward slash cho file URI
    StringChangeEx(AppPath, '\', '/', True);
    // Tro vao .vsto (deployment manifest) + |vstolocal de bo qua kiem tra cert
    ManifestUri := 'file:///' + AppPath + '/Add-insSupport.vsto|vstolocal';
    RegKey := 'Software\Microsoft\Office\Excel\Addins\Add-insSupport';
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'Description',  'Add-insSupport Excel Add-in');
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'FriendlyName', 'Add-insSupport');
    RegWriteDWordValue (HKEY_CURRENT_USER, RegKey, 'LoadBehavior',  3);
    RegWriteStringValue(HKEY_CURRENT_USER, RegKey, 'Manifest',     ManifestUri);
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usPostUninstall then
    RegDeleteKeyIncludingSubkeys(HKEY_CURRENT_USER,
      'Software\Microsoft\Office\Excel\Addins\Add-insSupport');
end;
